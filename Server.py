# coding=utf-8
"""
307attoube冷却水监控报警程序-server

  %%需配合花生壳运行在80端口%%

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Copyright: Copyright (c) 2018
%Created on 2018-11-23 
%Author:MengDa (github:pilidili)
%Version 1.0 
%Title: 307attoube冷却水监控报警程序-server
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  
"""
import itchatmp, modbus_tk, os, random, re, serial, struct, time, xlrd
import modbus_tk.defines as cst
from modbus_tk import modbus_rtu
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.schedulers.background import BlockingScheduler


#软水压力表通信端口
soft_water_Port='COM6';
#自来水水表通信端口
city_water_Port='COM5'

#全局状态变量定义
RTP_IsON=0   #水压表是否开启
Water_src=0  #冷却水源
SW_RTP=0.0   #软水实时水压
CW_Flow=0.0  #自来水实时流量
CW_Cons=0.0  #自来水总用量


Havesend=0   #群发标志 0-当日未群发 1-当日已群发
IsWarking=0  #循环水工作标志 0-未工作 1-正工作

#设置微信公众号
itchatmp.update_config(itchatmp.WechatConfig(
    token='***',
    appId = '***',
    appSecret = '***'))

#回复消息生成
def reply_msg():
    global RTP_IsON
    global Water_src
    global SW_RTP
    global CW_Flow
    global CW_Cons
    
    if RTP_IsON==0:
        CoolingWater_msg='冷却水压力表控制器未开启\n 截止当前，自来水总用量：%f m³' %(CW_Cons)
    else:
        if int(Water_src)==0:
            Water_src_str='软水'
        else:
            Water_src_str='自来水'

        CoolingWater_msg='冷却水：\n  水源：%s \n  软水压力：%f MPa\n  自来水流量：%f L/h\n  自来水总用量：%f m³' % (Water_src_str,SW_RTP,CW_Flow,CW_Cons)
    return_msg=CoolingWater_msg
    if SW_RTP<0.15 and CW_Flow<100:
        return_msg='**********设备断水**********\n'+return_msg
    return return_msg

#公众号群发消息
def mass_all(masage_mass):
    itchatmp.messages.send_all(TEXT, masage_mass)

#公众号消息自动回复
@itchatmp.msg_register(itchatmp.content.INCOME_MSG)
def reply(msg):
    return reply_msg()


# RS485-Modbus通讯
def Modbus_485(PORT,Regist_Address,Regist_Number):
    try:
        #Connect to the slave
        master = modbus_rtu.RtuMaster(
            serial.Serial(port=PORT, baudrate=9600, bytesize=8, parity='N', stopbits=1, xonxoff=0)
        )
        master.set_timeout(5.0)
        master.set_verbose(True)
        print("%s  connected"% datetime.now())
 
        return(master.execute(1, cst.READ_HOLDING_REGISTERS, Regist_Address, Regist_Number))
 
        #send some queries
        #logger.info(master.execute(1, cst.READ_COILS, 0, 10))
        #logger.info(master.execute(1, cst.READ_DISCRETE_INPUTS, 0, 8))
        #logger.info(master.execute(1, cst.READ_INPUT_REGISTERS, 0, 10))
        #logger.info(master.execute(1, cst.READ_HOLDING_REGISTERS, 100, 12))
        #logger.info(master.execute(1, cst.WRITE_SINGLE_COIL, 7, output_value=1))
        #logger.info(master.execute(1, cst.WRITE_SINGLE_REGISTER, 100, output_value=54))
        #logger.info(master.execute(1, cst.WRITE_MULTIPLE_COILS, 0, output_value=[1, 1, 0, 1, 1, 0, 1, 1]))
        #logger.info(master.execute(1, cst.WRITE_MULTIPLE_REGISTERS, 100, output_value=xrange(12)))
 
    except modbus_tk.modbus.ModbusError as exc:
        print("%s- Code=%d", exc, exc.get_exception_code())
        




#软水压力表信息采集
def SW_collector():
    
    #采集数据 实时水压（0-1，单位：PSI）、量程（2-3）、下限（4-5，单位：PSI）、上限（6-7，单位：PSI）、单位（8）和开关机（9）继电器状态（10）
    SW_Val=Modbus_485(soft_water_Port,0,11)
    
    global SW_RTP
    global RTP_IsON
    global Water_src

    #实时水压-MPa-4位有效数字
    SW_RTP=struct.unpack('f',struct.pack('H',SW_Val[0])+struct.pack('H',SW_Val[1]))[0]
        
    #水压表是否开启 0-未开启 1-开启
    RTP_IsON=SW_Val[10]
        
    #继电器状态  0-软水  1-自来水
    Water_src=SW_Val[9]

    
    
#自来水水表信息采集
def CW_collector():
    
    global CW_Flow  #自来水实时流量
    global CW_Cons  #自来水总用量
    
    CW_Val=Modbus_485(city_water_Port,0,4)
    
    CW_Cons =                          \
    int(str(hex(CW_Val[0]))[2:])+      \
    int(str(hex(CW_Val[1]))[2:])*0.0001

    CW_Flow =                          \
    int(str(hex(CW_Val[2]))[2:])*100+      \
    int(str(hex(CW_Val[3]))[2:])*0.01

    
    

#发送确认信息，确认网络畅通
def Conf_msg_sent():
    global Havesend
    if Havesend==0:
        mass_all('今日冷却水无异常\n当前状态:\n'+reply_msg())
    else:
        Havesend=0

#更新数据文件
def refresh_data():
    global Havesend
    global IsWarking
    global RTP_IsON
    global Water_src
    global SW_RTP
    global CW_Flow
    global CW_Cons
    
    SW_collector()
    CW_collector()
    with open('attoDRY.txt', 'w') as f:
        f.write('%d\n%.3f\n%.4f\n%.4f\n%d\n%d' %(Water_src,SW_RTP,CW_Flow,CW_Cons,RTP_IsON,Havesend))
    if SW_RTP<0.15 and CW_Flow<100:
        if IsWarking==1:
            IsWarking=0
            if Havesend==0:
                mass_all('！！！冷却水断水！！！')
                Havesend=1
    else:
        IsWarking==1




#建立时间表计划   
scheduler = BackgroundScheduler()
#scheduler.add_job(tick, 'interval', seconds=3)
#scheduler.add_job(tick, 'date', run_date='2016-02-14 15:01:05')
#每日23:30发微信确认网络通畅 
scheduler.add_job(Conf_msg_sent, 'cron',hour='23', minute='30')
#每两秒更新一次数据
scheduler.add_job(refresh_data, 'cron', second='*/2')
'''
        year (int|str) – 4-digit year
        month (int|str) – month (1-12)
        day (int|str) – day of the (1-31)
        week (int|str) – ISO week (1-53)
        day_of_week (int|str) – number or name of weekday (0-6 or mon,tue,wed,thu,fri,sat,sun)
        hour (int|str) – hour (0-23)
        minute (int|str) – minute (0-59)
        second (int|str) – second (0-59)
        
        start_date (datetime|str) – earliest possible date/time to trigger on (inclusive)
        end_date (datetime|str) – latest possible date/time to trigger on (inclusive)
        timezone (datetime.tzinfo|str) – time zone to use for the date/time calculations (defaults to scheduler timezone)
    
        *    any    Fire on every value
        */a    any    Fire every a values, starting from the minimum
        a-b    any    Fire on any value within the a-b range (a must be smaller than b)
        a-b/c    any    Fire every c values within the a-b range
        xth y    day    Fire on the x -th occurrence of weekday y within the month
        last x    day    Fire on the last occurrence of weekday x within the month
        last    day    Fire on the last day within the month
        x,y,z    any    Fire on any matching expression; can combine any number of any of the above expressions
    '''
scheduler.start()    #这里的调度任务是独立的一个线程
#启动微信公众号服务程序
itchatmp.run()