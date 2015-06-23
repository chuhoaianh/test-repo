"""
pond.py
    variable #      Description
    00              Setpoint Temperature 
    01              Slew Rate -- setpoint rate of change Deg. C/ hour
    02              Cycle to setpoint
    03              Memory 0 Setpoint Temperature
    04              Tolerance range at Setpoints for Dwell Time
    05              Dwell Counter (in seconds)
    06              Time Scale
    07              GPIB Primary Address
    09              Thermoelectric offset
    10              Current Measured Stage Temeprature (Deg. C)
    11              Cycle Flag
    12              Setpoint Increment
    13              Cycles to do
    14              Cycles done
    15              First Cycle Setpoint
    16              Second Cycle Setpoint
    17              First Cycle Dwell Time
    18              Second Cycle Dwell Time
    19              Cycle Slew Rate
    20              ACCESS CODE for protected variables

Begin ascii communication with the pond by issuing '?' via raw ascii over the serial port at 9600 baud.    
There are two commands supported when communicating with the pond in raw ASCII
1) 'r'
    example: 'r00' reads variable 0. All variables must be entered with 2 digits
2) 'w'
    example 'w00,-5' . Write to variable 00 with a value of -5  

Note about slew rate (Important)

'01': enable slew rate. 
        + If = 0, the POND will get value in '00' , update display in RAMP SP (but '02' will not change, '02' is no use in this case) 
        and the temp will go to the set point fastest. 
        '00' now is a fixed set point and POND will go there as fast as it can.
        '02' has no use in this case.
        
        + If = 1, the POND will get the current value of '02' and let the temp go to that value plus update the RAMP SP. 
        The value of '00' now is an auto self update real time set point 
        '00' Automatically update to the next set point (calculated based on the slew rate) when the temp move closely to that set point.
        '02' is now a fixed set point that '00' wants to move to.    
      
"""
import sys
import time
import Queue
from datetime import datetime
import threading
import traceback


try:
    import wmi
except:
    print "Dont have wmi module"
    pass
    
pondPort = 1
pondBaud = 9600
sp = None

CyclesToDo = 0
SetTemp1 = 0
SetTemp2 = 0
Dwell1 = 0
Dwell2 = 0
Slew = 0

SeekMode = ''
SetPoint = 0

def scanPort():
    ports = wmi.WMI()   
    portlist = []
        
    for port in ports.Win32_PnPEntity():        
        if port!=0:            
            if str(port.Name).count('COM'):                              
                portlist.append(str(port.Name))              
                print "-- Ports Name: " + str(port.Name) + "---";   
    return portlist
            
def AutoStart():
    global pondPort, pondBaud, sp
    """
    Auto scan the list of all ports then auto open each port. Send a sample to each port to make sure that it connect to the right port (Pond port)
    """
    ## ANH's code
    ports = [];
    ports = scanPort()
    for i in range (0,len(ports)):        
        ports[i] = ports[i][ports[i].index('M')+1:len(ports[i])-1];   #return the string 'xx' to ports[i];            
    ######
        print '------'
        print i;
        sp = Serial(int(ports[i]), pondBaud)
        
        #rtn = pyport.open(int(ports[i]), pondBaud);
        print sp
        if (sp==0):
            print 'unable to open COM%s' %(ports[i])
        else:
            print "COM%s Open!" %(ports[i])           
            try:
                print 'ready to send'
                rtn = snd('?')
                print rtn
                if not rtn.lower().count('pond'):
                    print 'failed to start communication with the pond.'
                else:           
                    print "Connect to COM%s"%(ports[i])
                    return 1
            except:
                print "Can not send"

def cycle(SetTemp1, SetTemp2, Slew, Dwell1, Dwell2, CyclesToDo):    
    snd('w14,%0.1f' %CyclesToDo)
    snd('w15,%0.1f' %SetTemp1)
    snd('w16,%0.1f' %SetTemp2)    
    snd('w17,%0.1f' %Dwell1)
    snd('w18,%0.1f' %Dwell2)
    snd('w19,%0.1f' %Slew)
    
    snd('w11,1')      #turn on cycle flag to enable slew rate
    
    DwellTimer = abs(Dwell2 - Dwell1);
    tempRange1 = [SetTemp1-1,SetTemp1+1];
    tempRange2 = [SetTemp2-1,SetTemp2+1];
    
    slewSeek(tempRange1, True, False);
    
def getInfo():
    global CyclesToDo, SetTemp1, SetTemp2, Dwell1, Dwell2, Slew, SeekMode, SetPoint;
    CyclesToDo = snd('r14')       
    SetTemp1 = snd('r15')
    SetTemp2 = snd('r16')
    Dwell1 = snd('r17')
    Dwell2 = snd('r18')
    Slew = snd('r19')
    SeekMode = snd('r11')
        
    CyclesToDo = CyclesToDo[6:18]
    SetTemp1 = SetTemp1[6:18]
    SetTemp2 = SetTemp2[6:18]
    Dwell1 = Dwell1[6:18]
    Dwell2 = Dwell2[6:18]    
    Slew = Slew[6:18]
    SeekMode = SeekMode[6:18]
    
    
    if (int(float(SeekMode)) == 0):
        SetPoint = sp.Xfer('r00\r', '00\r\n')#snd('r00');
        SetPoint = SetPoint[6:18]
    else:
        SetPoint = sp.XferAndWait('r02\r','02\r\n')#snd('r02');
        SetPoint = SetPoint[6:18]
    
    try:
        a = int(float(SetPoint))
    except:
        return False
    
    print "===================="
    print "CyclesToDo: " + repr(CyclesToDo)
    print "SetTemp1: " + repr(SetTemp1)
    print "SetTemp2: " + repr(CyclesToDo)
    print "Slew: " + repr(Slew)
    print "SeekMode: " + repr(SeekMode)
    print "Setpoint: " + repr(SetPoint)
    print "===================="
        
    return True
     
def snd(cmd):
    #global sp
    #print sp
    """
    send the string to the pond + '\r'
    return the return from the pond
    """
    rtn = cmd.replace('r','')
    return sp.XferAndWait('%s\r' % cmd, rtn+'\r\n')
    #return "read: " + sp.serQueue.get()
    #sp.XferAndWait('%s\r' % cmd,'\r\n', 5)
    #return pyport.sendOnline('%s\r' % astring,5,1)
          
def start():
    global pondPort, pondBaud, sp
    """
    Initialize communications with the pond.
    """
    #pondBaud = 9600;
    
    print "Pond Port: " + str(pondPort)
    print "Pond Baud: " + str(pondBaud)
    
    sp = SerialComm(pondPort, pondBaud)
    sp.setDaemon(True)        
    sp.start()
                   
    rtn = sp.XferAndWait('?\r','Thermal Conditioner\r\n', 5)
    #print "return: " + rtn
    if not rtn.count('Thermal Conditioner\r\n'):
        print 'failed to start communication with the pond.'
        stop()
        return 0
    
    #sp.XferAndWait('r14\r','\r\n', 5)
    return 1
    
def stop():
    """
    Terminate communcations with the pond.
    """
    global sp
    print pondPort
    
    global serStop, serStopFlag
    serStop = True        
    print "Closing Serial Thread" 
    while (serStopFlag==False):
        print ".",
        time.sleep(1)        
    sp.join()
        
    #pyport.close(pondPort)
    return 1
    
def setSlew(slew):
    """
    Set the pond's slew rate
    """
    #set the desired slew rate
    print snd('w19,%s' % slew)
    print snd('r19')               #may want to do a check here, for now assuming that everything went ok on the write    

def getSetTemperature():
    SeekMode = snd('r11');
    SeekMode = SeekMode[6:18];
    if (int(float(SeekMode)) == 0):
        SetPoint = snd('r00');
        SetPoint = int(float(SetPoint[6:18]));        
    else:
        SetPoint = snd('r02');
        SetPoint = int(float(SetPoint[6:18]));   
    return SetPoint;     
    
def getTemperature(retries = 3):
    """
    return the measured temperature of the pond
    """
    for retry in range(retries):
        try:
            time.sleep(1)
            rtn = snd('r10')
            tokens = rtn.split()
            for token in tokens:
                if token.count('e'):
                    return float(token)
        except KeyboardInterrupt:
            print 'keyboard interrupt, returning'
            return
        except:
            continue
    traceback.print_exc()
    raise Exception('failed to gather the measured temperature of the pond')

    
def _seek(temperatureRange, useSlew, returnImmediate = False):
    """Note:    
    '01': enable slew rate. 
        + If = 0, the POND will get value in '00' , update display in RAMP SP (but '02' will not change, '02' is no use in this case) 
        and the temp will go to the set point fastest. '00' now is a fixed set point
        + If = 1, the POND will get the current value of '02' and let the temp go to that value plus update the RAMP SP. The value of '00' now is an auto self update real time set point 
        '00' Automatically update to the next set point (calculated based on the slew rate) when the temp move closely to that set point        
    """
    if useSlew:
        print snd('w11,1')      #turn on cycle flag to enable slew rate
        print snd('w00,%.1f' % getTemperature())                                                            #set a real set point to the current measured value. '00' is a real time set point.  
        print snd('w02,%.1f' % (float((max(temperatureRange)) + float(min(temperatureRange))) / 2.0))       #Update the display. '02' will update a display, and doing the move to whatever value put in
        
    else:
        print snd('w00,%.1f' % (float((max(temperatureRange)) + float(min(temperatureRange))) / 2.0))       #'00' will update a display, and doing the move fast to whatever value put in
        print snd('w02,%.1f' % (float((max(temperatureRange)) + float(min(temperatureRange))) / 2.0))       #'02': has no use in this case
    
    if returnImmediate:
        return
    
    currentTemp = -1000
    outstr = ''
    while (currentTemp < min(temperatureRange) or currentTemp > max(temperatureRange)):
        time.sleep(1)
        currentTemp = getTemperature()
        if len(outstr):
            sys.stdout.write('\b' * len(outstr))
        outstr = 'temperature = %.1f, seeking to range %s' % (currentTemp, temperatureRange)
        sys.stdout.write(outstr)
    print ''
    print 'temperature seek complete'
        
def fastSeek(temperatureRange, returnImmediate = False):
    """
    Seek to the desired temperature range as fast as possible.
    return immediately after setting the seek point on the pond if the 'returnImmediate' paramter is set to True
    """
    _seek(temperatureRange, False, returnImmediate)

def slewSeek(temperatureRange, slew = None, returnImmediate = False):
    """
    Seek to the desired temperature range at a given slew rate. If no slew rate is given, use the current
    slew setting.
    
    if returnImmediate = True, return from the function as soon as the seek setpoint has been sent to the pond.
    
    example: pond.slewSeek([40,42], 100) #seek to temperature range [40,42] with slew Degrees C / hour
    example: pond.slewSeek([40,42], None, True) #set the temperature range [40,42], use the current pond slew
                                                 rate, return from the function immediately.
    
    """
    if slew:
        setSlew(slew)
    _seek(temperatureRange, True, returnImmediate)
    
#-----------------------------------------------------------------------------------------------------
#-------------------------------------- Synchronous Serial class -------------------------------------
#-----------------------------------------------------------------------------------------------------
# Do NOT modify or remove this copyright and confidentiality notice!
#
# Copyright (c) 2001 - $Date: 2013/12/18 $ Seagate Technology, LLC.
#
# The code contained herein is CONFIDENTIAL to Seagate Technology, LLC.
# Portions are also trade secret. Any use, duplication, derivation, distribution
# or disclosure of this code, for any reason, not expressly authorized is
# prohibited. All other rights are expressly reserved by Seagate Technology, LLC.
#
"""
simple windows Serial port with no Overlapped IO
"""
from win32file import *     # The base COM port and file IO functions.
from win32event import *    # Use events and the WaitFor[Multiple]Objects functions.
import win32con             # constants.
import numpy
import time
import serial
class Serial:    
    def __init__(self, port, baud):
        print "serial port: " + str(port)
        self.ser = serial.Serial(port = port-1, baudrate = baud, parity=serial.PARITY_NONE, bytesize=serial.EIGHTBITS)   
        print self.ser.baudrate
        print self.ser.port
    
    def inWaiting(self):
        return self.ser.inWaiting()
            
    def open(self):
        self.ser.open()        
    
    def close(self):
        self.ser.close()    
        print self.ser.isOpen()    
        
    def getTimeout(self):
        return ( self.ser.timeout )
            
    def setTimeout(self, newTimeout):
        newTimeout = int( newTimeout )
        self.ser.timeout = newTimeout
    
    def flush(self):
        self.ser.flush()    
    
    def write(self, chars):
        self.ser.write(chars) 

    def read(self, bytesToRead):        
        data = self.ser.read(bytesToRead)        
        return data
    
    def putc(self, data, timeout=2):
        try:
            self.ser.timeout = timeout
            self.ser.write(data)            
            return len(data)
        except TimeOutError:
            return 0

    def getc(self, size, timeout=2):
        try:
            self.ser.timeout = timeout
            x = self.ser.read(size)
            return x
        except TimeOutError:
            print "TimeoutError"
    
    def WaitForPrompt(self, prompt, retData = True, timeout = 2):
        InData = ""
        timer = time.time()
        while (InData.find(prompt)<0):
            if self.ser.inWaiting():
                InData += self.ser.read(self.ser.inWaiting())
            if (time.time() - timer >= timeout):
                print InData
                raise TimeOutError(timeout)
        if retData:
            return InData

    def TransmitAndWait(self, command, prompt, retData = True, timeout = 2):
        #Clear any pending read data
        if self.ser.inWaiting():
            self.ser.read(self.ser.inWaiting())
        self.write(command)
        return self.WaitForPrompt(prompt, retData, timeout)
        
#Custom Timout Exception Class
class TimeOutError(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr(self.value)
#-----------------------------------------------------------------------------------------------------
#------------------------------------ End Synchronous Serial class -----------------------------------
#-----------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------
#--------------------------------------- Serial Listener class ---------------------------------------
#-----------------------------------------------------------------------------------------------------
class SerialComm(threading.Thread):
    def __init__(self, COMPort, baudrate):
        threading.Thread.__init__(self)        
        #self.output_window = output_window
        
        try:
            self.sp =  Serial(COMPort, baudrate)
        except Exception,e:
            print str(traceback.format_exc())
                    
        self.sp.setTimeout(0.01)        
        self.serData = ""
        self.serQueue = Queue.Queue() 
        self.serCmdQueue = Queue.Queue()        
        self.InData = ""
        self.rtnData = ""
        self.io_timer = IOTimer()
        self.prompt = "A"
        
    def run(self):        
        global serStop,serStopFlag
        serStop = False
        serStopFlag = False        
        print "SerialListener started"
        while serStop==False:                        
            while(self.sp.inWaiting()!=None):
                time.sleep(0.1)
                self.serData += self.sp.read(self.sp.inWaiting())   
                #print "ser: " + repr(self.serData)
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
                timestamp = timestamp[:-3]   
                                            
                if(self.serData.count(self.prompt)!=0): 
                    #print "***************** Found prompt: " + repr(self.prompt)                   
                    break                
                if(serStop==True):                    
                    serStopFlag = True
                    break                    
            
            self.serData.replace("\r", "")
            self.serData.replace("\n", "")  
            
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
            timestamp = timestamp[:-3]       
            
            print timestamp + "\tsrd\t" + repr(self.serData) 
            self.serQueue.put(self.serData)                        
            self.serData = ""
        
        print "SerStop = " + repr(serStop)
        serStopFlag = True
        print "SerialListener Thread is killed"
        
        self.serCmdQueue.put("break")
        self.sp.close()
    
    def flushserQueue(self):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        timestamp = timestamp[:-3]
        #print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        print "~~~~~~~~Check Queue 3: " + str(self.serQueue.empty()) + "\n" 
        #if self.serQueue.empty()==False:
        temp = self.serQueue.get()
        print timestamp + "\ttemp: " +repr(temp)
            
        print timestamp + "\t~~~~~~~~Check Queue 4: " + str(self.serQueue.empty()) + "\n"
        #print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
    
    #def UpdateGUI(self, msg, label = "log"):
    #    evt = UpdateDisplayEvent(DisplayEventType, -1)  #initialize update display event
    #    evt.UpdateText(str(msg), label)                         #update display event
    #    wx.PostEvent(self.output_window, evt)           #After the event trigger,
          
    def WaitForRtn(self, rtn, timeout=2, printBuf = True):  
        #global InData     
        self.InData = ""
        self.rtnData = ""
        self.prompt = rtn      
        
        timer = time.time()   
        
        elpasedtime = time.time() - timer        
        
        while self.InData.count(self.prompt)==0:            
            if (elpasedtime < timeout):                
                self.InData = self.serQueue.get()
                self.rtnData += self.InData
                #if printBuf:
                    #timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
                    #timestamp = timestamp[:-3]                
                    #self.InData = self.InData.replace("\r", "")
                    #self.InData = self.InData.replace("\n", "")
                    #self.UpdateGUI(self.InData, "srd")
            else:
                #print "failed: Elapsed Time = %s / %s"  %(str(elpasedtime),str(timeout))
                raise Exception("Timeout while waiting for: " + repr(self.prompt))
            #print "Stuck 5"
            elpasedtime = time.time() - timer
            #print str(elpasedtime)     
        
        
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        timestamp = timestamp[:-3]#return first 3 digits in milliseconds
        #print "+++++++++++++++++++++ Indata: " + repr(self.InData) 
        #print "+++++++++++++++++++++ rtnData: " + repr(self.rtnData) 
        self.InData = ""
        return self.rtnData   

    def XferAndWait(self, command, prompt, timeout=2, printBuf = True):       
        global SerialPort#Thread stuff
           
        self.Xfer(command)
        
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        timestamp = timestamp[:-3]#return first 3 digits in milliseconds         
        
        rtnData = self.WaitForRtn(prompt, timeout, printBuf)      
        #print "+++++++++++++++++++++ rtnData: " + repr(rtnData)   
        return rtnData    
        
    def Xfer(self, command):   
        #print "\n============ SEND %s ============\n" %repr(command)
        self.sp.write(command)
        
    def ReadBuf(self):
        self.sp.read(self.sp.inWaiting())
       
            
class IOTimer(threading.Thread):
    """
    Thread object to control a timer for time IO operations
    """
    def __init__(self):
        threading.Thread.__init__(self)
        self._timer_complete = False # flag to indicate timer has completed
        self._active = False         # flag to indicate the timed IO is active
        self.time_int = 0
        self._lock = threading.Lock()
        
    def setTimeDuration(self, time_interval):
        """
        Sets the timer duration
        """
        self.time_int = time_interval         
                
    def run(self):
        """
        Waits for the specified duration
        """
        print self.time_int
        time.sleep(self.time_int)
        with self._lock:
            # timer has completed
            self._timer_complete = True        

    def start(self):
        """
        Starts the timer
        """
        self.start_time = time.time()
        self._active = True
        with self._lock:
            self._timer_complete = False
        threading.Thread.start(self)
        
    def stop(self):
        """
        Sets the self._active flag to False to indicate timed IO has stopped
        """
        self._active = False
    
    def skipIO(self):
        """
        Returns whether IO commands should be skipped
        
        Will return True only if the timer has completed and the time IO is
        still active
        """
        return self._timer_complete and self._active        
    