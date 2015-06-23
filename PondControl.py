import sys;
sys.path.append('C:/Program Files (x86)/Seacon');
sys.path.append('C:/Program Files (x86)/Seacon/macros/scripts');
sys.path.append('C:/Program Files (x86)/Seacon/macros/cnrlib');
sys.path.append('C:/Program Files (x86)/Seacon/macros/scripts/libraries');
sys.path.append('C:/Users/475177/Documents/SeaCon/User Scripts');
sys.path.append('C:/Python27/Scripts');

try:
	import os;
except:
	print "Error importing os"

try:
	import wx;
except:
	print "Error importing wx"

try:
	import threading;
except:
	print "Error importing threading"

try:
	import time;	
except:
	print "Error importing time"

try:
	import traceback;
except:
	print "Error importing traceback";
	
try:
	import math;
except:
	print "Error importing math";

try:
	import logging;
except:
	print "Error importing math";

try:
	import datetime;
except:
	print "Error importing math";

try:
	import csv;
except:
	print "Error importing csv";
	

#try:
print "Importing a_pond";
import a_pond;
#except:
#    print "Missing a_pond module"
    
try:
	print "Importing PowerBrick";
	import PowerBrick;
except: 
    print "Missing PowerBrick module"

global WMI_Flag;
WMI_Flag = False; 
try:
    import wmi;
    WMI_Flag = True;    
except:
    WMI_Flag = False;    
    print "Dont have wmi module"
    pass


ButtonSize = (120,23);
InputTemp = 0;
InputCOMPort = 1;
x=10;
y=25;
Temp1Input = '';    #will not be saved in POND
Temp2Input = '';    #will not be saved in POND
SlewInput = '';    #will be saved in POND through CycleThread
Dwel1Input = '';    #will be saved in POND through CycleThread
Dwel2Input = '';    #will be saved in POND through CycleThread
CyclesToDo = '';
Display_line1 = "RAMP SP = 0 DEG C \n";
Display_line2 = " STAGE = 0 DEG C";
global CurrentTemp;
CurrentTemp = 0;
SetTemp = 0;
logmsg = ""

SeekMode = "";  #there are 2 seek modes. Fast mode ('01' = 0)and Slew mode ('01' = 1)

SelectedBricks = [];

#GetInfoThread variables
choice = ""; #OnThreadUpdate Update choice
GetInfoThread_Done = False;

#Ram list for manual temperature cycle (TVM ReliKey_Generation kind of)
Ram=[];
ManualCyclesToDo = 2;
RamCounting = 0;

################## THREAD UPDATE GUI ######################
#1. Create new custom event to update the display
DisplayEventType = wx.NewEventType();
EVT_DISPLAY = wx.PyEventBinder(DisplayEventType, 1);


global GetTempThreadPause, GetTempThreadAlreadyPaused, GetTempThreadStop, CycleThreadStop, ThreadCompletyStop, ManualCycleThreadStop;                    
GetTempThreadPause = False; # Temporarily PAUSE the GetTempThread to work on the CycleThread. Not a complete thread stop, his will resume after CycleThread done the work
GetTempThreadStop = False;  # Completely STOP the GetTempThread after hitting "Stop" button to stop.
GetTempThreadAlreadyPaused = False;
CycleThreadStop = False;    # Completely STOP the CycleThread.
ThreadCompletyStop = False;
ManualCycleThreadStop = False;

def GetTempThreadStart(window):
    GetTempThread(window);
    
def CycleThreadStart(window):
    CycleThread(window);    

def ManualCycleThreadStart(window):
    ManualCycleThread(window);    

def SaveSettingThreadStart(window):
    SaveSettingThread(window);

def GetInfoThreadStart(window):
    GetInfoThread(window);

def CycleFunction(self,TempInput, SlewInput):
    global CurrentTemp, SetTemp, CycleThreadStop, choice;
    a_pond.getTemperature();
    print "-- Cycle Thread get temp...";
    print str(math.floor(CurrentTemp));
    print str(math.floor(TempInput));
    
    if(SlewInput != 0):
        a_pond.slewSeek([TempInput - 1, TempInput + 1], SlewInput, True);        
    else:
        a_pond.fastSeek([TempInput - 1, TempInput + 1], returnImmediate=True);
    
    #In case if the current temp =24.9487...C, just return and stay for dwel time and go to the next set point
    #it will be very slow to put 24.646C to 25C      
    
    if(math.fabs(float(CurrentTemp) - TempInput) < 0.2):
        return;
    
    choice = "TempDisplay";
    
    if(math.fabs(CurrentTemp - TempInput) < 0.5):
        pass
        
    #while (math.floor(CurrentTemp)!=math.floor(TempInput)):  
    while (math.fabs(CurrentTemp - TempInput) > 0.2):             
        try:
            print "CycleThreadStop = " + str(CycleThreadStop);
            if(CycleThreadStop == False):
                CurrentTemp = a_pond.getTemperature(3);
                Display_line1 = "RAMP SP= " + str(SetTemp) + " DEG C \n";
                Display_line2 = " STAGE = " + str(CurrentTemp) + " DEG C";
                choice = "TempDisplay";                    
                evt = UpdateDisplayEvent(DisplayEventType, -1); #initialize update display event
                evt.UpdateText(Display_line1 + Display_line2);  #update display event
                wx.PostEvent(self.output_window, evt);          #After the event trigger,
                        
                time.sleep(1);  
            else:
                return;  
        except:
            print "Can not get measure"
            trace = traceback.format_exc()
            print trace
            return; 
              
#Define event
class UpdateDisplayEvent(wx.PyCommandEvent):    
    def __init__(self, evtType, id):
        wx.PyCommandEvent.__init__(self, evtType, id)
        self.msg = "";        
    
    def UpdateText(self,text):
        #print "Update MSG";
        self.msg = text;
    
    def GetText(self):
        #print "Get MSG";
        return self.msg;

#GetTempThread will get temperature from the pond for every second
#When Cycling thread starts, this thread will be paused and will be resumed when Cycling is finish    
class GetTempThread(threading.Thread):
    def __init__(self, output_window):
        threading.Thread.__init__(self)
        self.output_window = output_window
        print "Thread started"
        self.start()
    
    def run(self):
        global CurrentTemp, GetInfoThread_Done, choice, SetTemp, GetTempThreadAlreadyPaused;
        print "In Temp thread"  
        
        retries = 0
        while (GetInfoThread_Done == False):
            pass;
        #time.sleep(0.5);
        choice = "TempDisplay";
        
        while(GetTempThreadStop==False):    #if "Stop" button is not hit, just need to pause the GetTempThread. This will resume after CycleThread done the work
            while(GetTempThreadPause==False):   #Just pause the thread, will resume after cyclethread is done the work
                #print str(GetTempThreadPause);                  
                if(GetTempThreadStop == True):
                    print "Please wait";
                    return;
                else:                    
                    try:
                        print "- Temp Thread get temp..."
                        if(GetTempThreadStop == True): 
                            print "Inside 1";                           
                            break;
                        
                        CurrentTemp = a_pond.getTemperature(3);
                        
                        if(GetTempThreadStop == True):
                            print "Inside 2";                                                      
                            break; 
                        
                        SetTemp = a_pond.getSetTemperature();
                        Display_line1 = "RAMP SP= " + str(SetTemp) + " DEG C \n";
                        Display_line2 = " STAGE = " + str(CurrentTemp) + " DEG C";
                        
                        choice = "TempDisplay";
                        #print Display_line1 + Display_line2;
                        evt = UpdateDisplayEvent(DisplayEventType, -1); #initialize update display event
                        evt.UpdateText(Display_line1 + Display_line2);  #update display event
                        wx.PostEvent(self.output_window, evt);          #After the event trigger,
                        
                        if(GetTempThreadStop == True):
                            print "Inside 3";                                                      
                            break; 
                        
                        time.sleep(0.5);                        
                    except:
                        print "\nCan not get measure - Retrying\n"
                        trace = traceback.format_exc()
                        print trace
                        retries = retries + 1
                        pass
                                                            
            print "\nGet Temp Thread PAUSED!";
            GetTempThreadAlreadyPaused = True;
            while(GetTempThreadPause==True):    #Just stay there and wait until the thread is resumed      
                if (GetTempThreadStop==True):
                    print "Get Temp Thread STOPPED!"
                    global ThreadCompletyStop;
                    ThreadCompletyStop = True;
                    return;                     
                pass; 

#CycleThread will cycle the temperature from Temp1 -> Temp2 with the rate xC/hour (Slew) 
#Dwel1: time in seconds stay at Temp1 (same for Dwell2)
#When this thread starts, it will paused the GetTempThread, and will resume that thread when done the number of cycles (CyclesToDo) 
class CycleThread(threading.Thread):
    def __init__(self, output_window):
        threading.Thread.__init__(self)
        self.output_window = output_window
        print "Cycle Thread started"
        self.start()     
    
    def run(self):
        global GetTempThreadPause, CycleThreadStop;
        GetTempThreadPause = True;
        time.sleep(3);
        CycleThreadStop = False;
        print str(CycleThreadStop);
        
        while(CycleThreadStop==False):
            global SetTemp, CurrentTemp;                             
            print "============== CYCLING START =============="
            print "== Temp 1 = " + str(Temp1Input);
            print "== Temp 2 = " + str(Temp2Input);   
            print "== Slew = %s C/hour" %(str(SlewInput));
                                    
            if(CyclesToDo != 0):                                
                for i in range(0, CyclesToDo):       
                    print "Cycle " + str(i) + ":"             
                    if(CycleThreadStop == False):                        
                        SetTemp = Temp1Input;                                                
                        CycleFunction(self,Temp1Input, SlewInput);
                        
                        print "Stay at %sC for %s(s)" %(str(Temp1Input),str(Dwel1Input));
                        start = time.time();
                        stop = time.time();                        
                        
                        while(stop - start < Dwel1Input):
                            print "Elapsed time: %s(s)\r" %(str(stop - start));
                            time.sleep(0.2);
                            if(CycleThreadStop==True):
                                print "Cycle STOPPED";
                                return;
                            stop = time.time();
                        print "Wake Up";
                        
                        SetTemp = Temp2Input;
                        #a_pond.slewSeek([Temp2Input - 1, Temp2Input + 1], SlewInput, True);                                         
                        CycleFunction(self,Temp2Input, SlewInput);
                        #CycleFunction(self,Temp2Input);        
                        print "Stay at %sC for %s(s)" %(str(Temp2Input),str(Dwel2Input));
                        startt = time.time();
                        stopt = time.time();
                        
                        while(stopt - startt < Dwel2Input):
                            print "Elapsed time: %s(s)\r" %(str(stop - start));
                            time.sleep(0.2);
                            if(CycleThreadStop==True):
                                print "Cycle STOPPED";
                                return;
                            stopt = time.time();                                                  
                    else:
                        return;
                print "Finished %s cycle(s)" %(str(CyclesToDo));
                GetTempThreadPause = False;  #resume normal state
                CycleThreadStop = True;
                return;         #finish the number of cycles to do    
        print "Cycle Thread is STOPPED"    

class ManualCycleThread(threading.Thread):
    def __init__(self, output_window):
        threading.Thread.__init__(self)
        self.output_window = output_window
        print "Manual Cycle Thread started"
        self.start();
    
    def run(self):
        global Temp1Input, Temp2Input, SlewInput, Dwel1Input, Dwel2Input, CyclesToDo, Temp1Range, Temp2Range, CycleThreadStop, RamCounting;
        ManualCyclesToDo = 2;
        RamCounting = 0;
        
        while(CycleThreadStop==False):
            for j in range(0,ManualCyclesToDo):            
                for i in range(1,len(Ram)):
                    Temp1Input = int(Ram[i][0]);
                    Temp2Input = int(Ram[i][1]);
                    Dwel1Input = int(Ram[i][2]);
                    Dwel2Input = int(Ram[i][3]);
                    if(int(Ram[i][4])!=0):
                        SlewInput = 3600*math.fabs(Temp2Input - Temp1Input)/int(Ram[i][4]); #time is in second 3600s = 60min
                    else:
                        SlewInput = 0;                                
                    Temp1Range = [Temp1Input - 1, Temp1Input + 1];
                    Temp2Range = [Temp2Input - 1, Temp2Input + 1];
                                    
                    CyclesToDo = 1;
                    
                    #print "============== CYCLE %s START ==============" %(str(j));
                    #print "== Temp 1 = " + str(Temp1Input);
                    #print "== Temp 2 = " + str(Temp2Input);   
                    #print "== Slew = %s C/hour" %(str(SlewInput));
                    self.UpdateFunction("\n\nRAM %s START\n" %(str(j)), "Log"); 
                    self.UpdateFunction("* Temp 1 = %s for %d(s)\n" %(str(Temp1Input),Dwel1Input), "Log"); 
                    self.UpdateFunction("* Temp 2 = %s for %d(s)\n" %(str(Temp2Input),Dwel2Input), "Log");                     
                    self.CycleFunc()
                    
                    if(CycleThreadStop == True):
                    	print "Return to main"
                    	return
                    
                    RamCounting = i + 1;
                
                RamCounting = 0;            
        print "Finished!!!!!!!!!!!!!!!";
        
    def UpdateFunction(self, msg, choice_local):
        global choice;
        
        choice = choice_local;
        print "*" + choice + "--" + choice_local;
        evt = UpdateDisplayEvent(DisplayEventType, -1); #initialize update display event
        evt.UpdateText(str(msg));  #update display event
        wx.PostEvent(self.output_window, evt);          #After the event trigger
    
    def CycleFunc(self):
        global GetTempThreadPause, CycleThreadStop;
        GetTempThreadPause = True;
        time.sleep(3);
        CycleThreadStop = False;
        print str(CycleThreadStop);
        
        while(CycleThreadStop==False):
            global SetTemp, CurrentTemp;                                    
            if(CyclesToDo != 0):                                
                for i in range(0, CyclesToDo): 
                    print "Ram " + str(RamCounting) + ":"                    
                                 
                    if(CycleThreadStop == False):                        
                        SetTemp = Temp1Input;                                                                        
                        CycleFunction(self,Temp1Input, SlewInput);
                        
                        print "Stay at %sC for %s(s)" %(str(Temp1Input),str(Dwel1Input));
                        start1 = time.time();
                        stop1 = time.time();                        
                        
                        while(stop1 - start1 < Dwel1Input):
                            print "Elapsed time: %s(s)\r" %(str(stop1 - start1));
                            time.sleep(0.2);
                            if(CycleThreadStop==True):
                                print "Cycle STOPPED";
                                return;
                            stop1 = time.time();                       
                        print "Wake Up";
                        
                        #Go to temp 2
                        SetTemp = Temp2Input;
                        CycleFunction(self,Temp2Input, SlewInput);
                            
                        print "Stay at %sC for %s(s)" %(str(Temp2Input),str(Dwel2Input));
                        start2 = time.time();
                        stop2 = time.time();
                        
                        while(stop2 - start2 < Dwel2Input):
                            print "Elapsed time: %s(s)\r" %(str(stop2 - start2));
                            time.sleep(0.2);
                            if(CycleThreadStop==True):
                                print "Cycle STOPPED";
                                return;
                            stop2 = time.time();                                                  
                    else:
                        return;
                GetTempThreadPause = False;  #resume normal state
                CycleThreadStop = True;
                return;         #finish the number of cycles to do    
        print "Cycle Thread is STOPPED"

class SaveSettingThread(threading.Thread):
    def __init__(self, output_window):
        threading.Thread.__init__(self)
        self.output_window = output_window
        print "Save Setting Thread started"
        self.start()
    
    def run(self):
        global GetTempThreadPause;
        GetTempThreadPause = True;  #Pause GetTempThread
        time.sleep(3);
        
        a_pond.snd('w14,%s' %CyclesToDo);
        print a_pond.snd('r14');            
        a_pond.snd('w15,%s' %Temp1Input);
        print a_pond.snd('r15');
        a_pond.snd('w16,%s' %Temp2Input);
        print a_pond.snd('r16');        
        a_pond.snd('w17,%s' %Dwel1Input);
        print a_pond.snd('r17');
        a_pond.snd('w18,%s' %Dwel2Input);
        print a_pond.snd('r18');
        a_pond.snd('w19,%s' %SlewInput);
        print a_pond.snd('r19');
        print "Saving done!";
        
        GetTempThreadPause = False; #Resume GetTempThread

class GetInfoThread(threading.Thread):
    def __init__(self, output_window):
        threading.Thread.__init__(self)
        self.output_window = output_window
        print "Get Info Thread started"
        self.start();
        
    def run(self):
        global Temp1Input, Temp2Input, SlewInput, Dwel1Input, Dwel2Input, CyclesToDo, choice, GetInfoThread_Done, SeekMode, SetTemp; 
        
        rtn = False        
        while rtn==False:	
        	try:
				rtn = a_pond.getInfo()
        	except:
        		pass
        
        SlewInput = int(float(a_pond.Slew));    #Slew rate is set for Go to Temp
        self.UpdateFunction(SlewInput, "Slew1");
        time.sleep(0.1);
               
        Temp1Input = int(float(a_pond.SetTemp1));    #will not be saved in POND
        self.UpdateFunction(Temp1Input, "Temp1");
        time.sleep(0.1);
        
        Temp2Input = int(float(a_pond.SetTemp2));    #will not be saved in POND
        self.UpdateFunction(Temp2Input, "Temp2");
        time.sleep(0.1);
        
        SlewInput = int(float(a_pond.Slew));    #will be saved in POND through CycleThread
        self.UpdateFunction(SlewInput, "Slew");
        time.sleep(0.1);
        
        Dwel1Input = int(float(a_pond.Dwell1));    #will be saved in POND through CycleThread
        self.UpdateFunction(Dwel1Input, "Dwel1");
        time.sleep(0.1);
        
        Dwel2Input = int(float(a_pond.Dwell2));    #will be saved in POND through CycleThread
        self.UpdateFunction(Dwel1Input, "Dwel2");
        time.sleep(0.1);
        
        CyclesToDo = int(float(a_pond.CyclesToDo));       
        self.UpdateFunction(CyclesToDo, "Cycles"); 
        time.sleep(0.1);
        
        #Update status bar to display the Seek mode
        SeekMode = str(int(float(a_pond.SeekMode)));
        if(SeekMode == '0'):
            SeekMode = "Fast Seek";
        else:
            SeekMode = "Slew Seek";
        self.UpdateFunction(SeekMode, "SeekMode");
        time.sleep(0.1);
        
        #Update display line RAMP SP = SetTemp
        SetTemp = int(float(a_pond.SetPoint));
        self.UpdateFunction(SetTemp, "SetTemp");
        time.sleep(0.1);
        
        self.UpdateFunction("Connected\n", "Log"); 
        time.sleep(0.1);
        
        GetInfoThread_Done = True;
        
    def UpdateFunction(self, msg, choice_local):
        global choice;
        
        choice = choice_local;
        print "*" + choice + "--" + choice_local;
        evt = UpdateDisplayEvent(DisplayEventType, -1); #initialize update display event
        evt.UpdateText(str(msg));  #update display event
        wx.PostEvent(self.output_window, evt);          #After the event trigger       
        
########################################

class MyPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1);
        self.Log = "";        
        
        #Auto scan and connect to POND port
        self.Start_Button_Status = "Auto Connect";
        self.Pond_Start_Button = wx.Button(self, -1, self.Start_Button_Status, pos = (10,10), size = ButtonSize);        
        self.Bind(wx.EVT_BUTTON, self.Pond_Start_Button_Click, self.Pond_Start_Button);
        
        #Manual connect to POND port
        self.ManualStart_Button_Status = "Connect to COM";
        self.Pond_ManualStart_Button = wx.Button(self, -1, self.ManualStart_Button_Status, pos = (10,10 + y), size = (90,23));        
        self.Bind(wx.EVT_BUTTON, self.Pond_ManualStart_Button_Click, self.Pond_ManualStart_Button);
        
        self.InputCOMPort_Text = wx.TextCtrl(self,-1,str(InputCOMPort), pos = (105,10 + y), size=(25,23));
        
        #Go to set point
        self.Pond_GoToTemp_Button = wx.Button(self, -1, "Go to", pos = (10,10 + 2*y), size = (40,23));        
        self.Bind(wx.EVT_BUTTON, self.Pond_GoToTemp_Button_Click, self.Pond_GoToTemp_Button);
        
        self.InputTemp_Text = wx.TextCtrl(self,-1,str(InputTemp), pos = (53,10 + 2*y), size=(23,23));
        self.InputTemp_Text.SetMaxLength(5)
        self.With_StaticText = wx.StaticText(self,-1,"rate", pos = (79,10 + 2.2*y));
        self.SlewInput1_Text = wx.TextCtrl(self,-1,str(SlewInput), pos = (105,10 + 2*y), size=(25,23));
        self.SlewInput1_Text.SetMaxLength(3)

        
        #-- Cycle setup
        self.Cycle_StaticText = wx.StaticText(self,-1,"------- Cycles setup -------", pos = (10,10 + 3*y + 5));
        self.Temp1_StaticText = wx.StaticText(self,-1,"Temp 1: ", pos = (10,10 + 4*y));
        self.Temp1Input_Text = wx.TextCtrl(self,-1,str(Temp1Input), pos = (95,5 + 4*y), size=(35,23));
        self.Temp2_StaticText = wx.StaticText(self,-1,"Temp 2: ", pos = (10,10 + 5*y));
        self.Temp2Input_Text = wx.TextCtrl(self,-1,str(Temp2Input), pos = (95,5 + 5*y), size=(35,23));
        self.Slew_StaticText = wx.StaticText(self,-1,"Slew (C/hour): ", pos = (10,10 + 6*y));
        self.SlewInput_Text = wx.TextCtrl(self,-1,str(SlewInput), pos = (95,5 + 6*y), size=(35,23));
        self.Dwel1_StaticText = wx.StaticText(self,-1,"Dwel 1: ", pos = (10,10 + 7*y));
        self.Dwel1Input_Text = wx.TextCtrl(self,-1,str(Dwel1Input), pos = (95,5 + 7*y), size=(35,23));
        self.Dwel1_StaticText = wx.StaticText(self,-1,"Dwel 2: ", pos = (10,10 + 8*y));
        self.Dwel2Input_Text = wx.TextCtrl(self,-1,str(Dwel2Input), pos = (95,5 + 8*y), size=(35,23));
        self.Dwel1_StaticText = wx.StaticText(self,-1,"Cycles to do: ", pos = (10,10 + 9*y));
        self.CyclesToDo_Text = wx.TextCtrl(self,-1,str(CyclesToDo), pos = (95,5 + 9*y), size=(35,23));
        #-----------------------
        
        #-- Save System Setting
        self.Pond_SaveSetting_Button = wx.Button(self, -1, "Save Setting to Pond", pos = (10,10 + 10*y), size = ButtonSize);        
        self.Bind(wx.EVT_BUTTON, self.Pond_SaveSetting_Button_Click, self.Pond_SaveSetting_Button);
        #-----------------------
        
        #-- Start/Stop cycle buttons--
        self.CycleStatus_Button = "Start Cycle";        
        self.Pond_StartCycle_Button = wx.Button(self, -1, self.CycleStatus_Button, pos = (10,10 + 11*y), size = ButtonSize);        
        self.Bind(wx.EVT_BUTTON, self.Pond_StartCycle_Button_Click, self.Pond_StartCycle_Button);
        #-----------------------
        
        #-- Start/Stop Manual cycle buttons--
        self.ManualCycleStatus_Button = "Start Temp Profile";        
        self.Pond_ManualCycle_Button = wx.Button(self, -1, self.ManualCycleStatus_Button, pos = (10,10 + 13*y), size = ButtonSize);        
        self.Bind(wx.EVT_BUTTON, self.Pond_ManualCycle_Button_Click, self.Pond_ManualCycle_Button);
        #-----------------------
        
        #--Log and Temperature Display--
        self.Display_Text = wx.TextCtrl(self,-1,Display_line1 + Display_line2, pos = (150,10), size=(250,35),style=wx.TE_MULTILINE);
        self.Log_Text = wx.TextCtrl(self,-1,self.Log, pos = (150,60), size=(250,192),style=wx.TE_MULTILINE);
        #-----------------------
        '''
        #-- Power cycle control
        self.PowerBrick_StaticText = wx.StaticText(self,-1,"-- Power Brick control --", pos = (150,10 + 10*y));
        BrickList = self.PowerBrickScan();
        self.Brick_Choice = wx.Choice(self,-1,pos = (150,10 + 11*y), size=(120,35), choices = BrickList);
        self.Bind(wx.EVT_CHOICE, self.PowerBrickChoice, self.Brick_Choice)
        self.PowerSwitch_Status = "Turn On";
        self.PowerSwitch_Button = wx.Button(self,-1,self.PowerSwitch_Status,pos = (275,10 + 11*y), size=(60,23));
        self.Bind(wx.EVT_BUTTON, self.PowerSwitch, self.PowerSwitch_Button);
        self.PowerSwitch_Button.Disable();
        #-----------------------
        '''
        self.Bind(EVT_DISPLAY, self.OnThreadUpdate);                
        self.Pond_StartCycle_Button.Disable();
        
        if (WMI_Flag == False): #if WMI module is not found in the system, Disable the auto start
            self.Pond_Start_Button.Disable();

#---------------Power Brick Control
    def PowerBrickScan(self):
        self.BrickList = [];        
        self.enumeratedBricks = [];
        try:
            print self.enumeratedBricks;
            self.enumeratedBricks = PowerBrick.enumerate();    
            print self.enumeratedBricks;
            index = 0;
            for brick in self.enumeratedBricks:
                self.BrickList.append('Brick ID %s, index %d' % (brick.id, index));                
                index += 1;     
            print self.BrickList;                   
        except:
            print "Power Brick scan error";       
        
        return self.BrickList;
    
    def PowerBrickChoice(self,event):       
        print event.GetSelection();
        self.UserSelectedBrick = self.enumeratedBricks[event.GetSelection()];
        self.PowerSwitch_StatusCheck();
        pass;
    
    def PowerSwitch(self,evnt):   
        if (self.PowerSwitch_Status == "Turn On"):
            self.UserSelectedBrick.switch(self.UserSelectedBrick.ON);
            self.PowerSwitch_Status = "Turn Off";
            self.PowerSwitch_Button.Destroy();
            self.PowerSwitch_Button = wx.Button(self,-1,self.PowerSwitch_Status,pos = (275,10 + 11*y), size=(60,23))
            self.Bind(wx.EVT_BUTTON, self.PowerSwitch, self.PowerSwitch_Button)
        else:
            self.UserSelectedBrick.switch(self.UserSelectedBrick.OFF);
            self.PowerSwitch_Button.Destroy();
            self.PowerSwitch_Status = "Turn On";
            self.PowerSwitch_Button = wx.Button(self,-1,self.PowerSwitch_Status,pos = (275,10 + 11*y), size=(60,23))
            self.Bind(wx.EVT_BUTTON, self.PowerSwitch, self.PowerSwitch_Button)             
   
    def PowerSwitch_StatusCheck(self):
        self.UserSelectedBrick.identify();
        
        if self.UserSelectedBrick.isPoweredOn():
            self.PowerSwitch_Status = "Turn Off";
        else:
            self.PowerSwitch_Status = "Turn On";
                
        self.PowerSwitch_Button.Destroy();
        self.PowerSwitch_Button = wx.Button(self,-1,self.PowerSwitch_Status,pos = (275,10 + 11*y), size=(60,23))
        self.Bind(wx.EVT_BUTTON, self.PowerSwitch, self.PowerSwitch_Button)
#---------------            
                           
#--------------- Auto Start
    def Pond_Start_Button_Click(self, event): 
        global GetTempThreadPause, GetTempThreadStop, ThreadCompletyStop;  #have to define GetTempThreadPause to be a global variable so it can be passed to GetTempThread func     
        GetTempThreadPause = False;  
        GetTempThreadStop = False;
        ThreadCompletyStop = False;
        
        self.Start_Button_Status = "Stop";
        self.Start_Button_StatusCheck();
        
        self.Log += "Scanning the port list, please be patience.... \n";
        self.Log_Text.SetValue(self.Log);
         
        temp = a_pond.AutoStart();
        if (temp==1):
            self.Pond_ManualStart_Button.Disable();   #Disable Manual start
            self.InputCOMPort_Text.Disable();
            self.Pond_GoToTemp_Button.Enable();
            self.Pond_StartCycle_Button.Enable();  
            self.Pond_SaveSetting_Button.Enable();
            
            GetInfoThreadStart(self);        
            #Start thread and tell thread to send the event back to "self" (which is MyPanel) so MyPanel can receive the
            #event when it happens
            GetTempThreadStart(self);      
        else:
            Status = "Can not connect";
            self.Start_Button_Status = "Auto Connect";
            self.Start_Button_StatusCheck();
            
    def Start_Button_StatusCheck(self):
        if(self.Start_Button_Status == "Auto Connect"):
            self.Pond_Start_Button.Destroy();
            self.Pond_Start_Button = wx.Button(self, -1, self.Start_Button_Status, pos = (10,10), size = ButtonSize);
            self.Bind(wx.EVT_BUTTON, self.Pond_Start_Button_Click, self.Pond_Start_Button);
        else:
            self.Pond_Start_Button.Destroy();    
            self.Pond_Start_Button = wx.Button(self, -1, self.Start_Button_Status, pos = (10,10), size = ButtonSize);        
            self.Bind(wx.EVT_BUTTON, self.Pond_Stop_Button_Click, self.Pond_Start_Button);      
        
        print self.Start_Button_Status;
    
#--------------- Manual Start
    def Pond_ManualStart_Button_Click(self,event):
        global GetTempThreadPause, GetTempThreadStop, ThreadCompletyStop, GetInfoThread_Done;  #have to define GetTempThreadPause to be a global variable so it can be passed to GetTempThread func     
        GetTempThreadPause = False;  
        GetTempThreadStop = False;
        ThreadCompletyStop = False;
        GetInfoThread_Done = False;
        
        retries = 3;
                        
        self.Log += "Connecting, please be patience.... \n";
        self.Log_Text.SetValue(self.Log);
        a_pond.pondPort = int(self.InputCOMPort_Text.GetValue());
        
                
        for retry in range(retries):
            temp = a_pond.start();
            print temp
            
            if (temp==1):
                print temp;                        
                self.Pond_Start_Button.Disable();   #Disable auto start
                self.Pond_GoToTemp_Button.Enable();
                self.Pond_StartCycle_Button.Enable();   
                self.Pond_SaveSetting_Button.Enable();
                
                self.ManualStart_Button_Status = "Disconnect COM";
                self.ManualStart_Button_StatusCheck();         
                           
                GetInfoThreadStart(self);     
                GetTempThreadStart(self);                
                return;      
            else:
                print "Can not connect" + str(temp);
                Status = "Can not connect";
                self.Log += Status;
                self.Log_Text.SetValue(self.Log);
                self.ManualStart_Button_Status = "Connect to COM";
                self.ManualStart_Button_StatusCheck();
    
    def ManualStart_Button_StatusCheck(self):
        if(self.ManualStart_Button_Status == "Connect to COM"):
            self.Pond_ManualStart_Button.Destroy();
            self.Pond_ManualStart_Button = wx.Button(self, -1, self.ManualStart_Button_Status, pos = (10,10 + y), size = (90,23));  
            self.Bind(wx.EVT_BUTTON, self.Pond_ManualStart_Button_Click, self.Pond_ManualStart_Button);
        else:
            self.Pond_ManualStart_Button.Destroy();    
            self.Pond_ManualStart_Button = wx.Button(self, -1, self.ManualStart_Button_Status, pos = (10,10 + y), size = (90,23));        
            self.Bind(wx.EVT_BUTTON, self.Pond_Stop_Button_Click, self.Pond_ManualStart_Button);
    
    def Pond_Stop_Button_Click(self, event):        
        global GetTempThreadPause, GetTempThreadStop, CycleThreadStop; #have to define GetTempThreadPause to be a global variable so it can be passed to GetTempThread func        
        GetTempThreadPause = True;
        GetTempThreadStop = True
        CycleThreadStop = True;
        #print ThreadCompletyStop;
        count = 0;
        text = ".";
        print "Thread is closing",
        while(ThreadCompletyStop == False):            
            if((count%20000)==1):
                print text,;
            count = count + 1;                      
        
        print "\n" + str(count) + "\n";
        a_pond.stop()
        if(WMI_Flag):
            print "Found WMI"
            if(self.Pond_Start_Button.IsEnabled()):
                print "Auto"
                self.Start_Button_Status = "Auto Connect";
                self.Start_Button_StatusCheck();
                self.Pond_ManualStart_Button.Enable();
                self.InputCOMPort_Text.Enable();
            else:
                print "Manual"
                self.ManualStart_Button_Status = "Connect to COM";
                self.ManualStart_Button_StatusCheck();
                self.Pond_Start_Button.Enable();
                self.InputCOMPort_Text.Enable();
        else:
            print "Manual"
            self.ManualStart_Button_Status = "Connect to COM";
            self.ManualStart_Button_StatusCheck();
             
        self.Pond_GoToTemp_Button.Disable();
        self.Pond_StartCycle_Button.Disable();
        self.Pond_SaveSetting_Button.Disable();
        
        Status = "Disconnected";
        self.Log += Status + "\n";
        self.Log_Text.SetValue(self.Log)
        
   
    def OnThreadUpdate(self, event):
        global choice
        msg = event.GetText();
        #self.Display_Text.SetValue(msg);
        print choice + "OnThreadUpdate";
        
        if (choice == "TempDisplay"):
            self.Display_Text.SetValue(msg);
        elif (choice == "Log"):
        	self.Log += msg
        	self.Log_Text.SetValue(self.Log);
        elif(choice == "Temp1"):
            self.Temp1Input_Text.SetValue(msg);
        elif (choice == "Temp2"):
            self.Temp2Input_Text.SetValue(msg);
        elif (choice == "Slew"):
            self.SlewInput_Text.SetValue(msg);
        elif (choice == "Dwel1"):
            self.Dwel1Input_Text.SetValue(msg);
        elif (choice == "Dwel2"):
            self.Dwel2Input_Text.SetValue(msg);
        elif (choice == "Cycles"):
            self.CyclesToDo_Text.SetValue(msg);
        elif (choice == "Slew1"):
            self.SlewInput1_Text.SetValue(msg);
        elif (choice == "SeekMode"):            
            StatusBar = self.GetParent()
            StatusBar.SetStatusText("Seek Mode: " + msg);

#--------------- Default cycle
    def Pond_StartCycle_Button_Click(self, event):  
        global Temp1Input, Temp2Input, SlewInput, Dwel1Input, Dwel2Input, CyclesToDo, Temp1Range, Temp2Range;
        self.CycleStatus_Button = "Stop Cycle";
        self.Cycle_Button_StatusCheck();
        
        Temp1Input = int(self.Temp1Input_Text.GetValue());        
        Temp2Input = int(self.Temp2Input_Text.GetValue());            
        SlewInput = int(self.SlewInput_Text.GetValue());
        Dwel1Input = int(self.Dwel1Input_Text.GetValue());
        Dwel2Input = int(self.Dwel2Input_Text.GetValue());
        CyclesToDo = int(self.CyclesToDo_Text.GetValue());
        
        Temp1Range = [Temp1Input - 1, Temp1Input + 1];
        Temp2Range = [Temp2Input - 1, Temp2Input + 1];
        Status = "Cycling";
        self.Log += Status + "\n";
        self.Log_Text.SetValue(self.Log);
        #self.Pond_StopCycle_Button.Enable();
        
        CycleThreadStart(self);
        
    def Pond_StopCycle_Button_Click(self, event):        
        global GetTempThreadPause, CycleThreadStop;  #have to define GetTempThreadPause to be a global variable so it can be passed to GetTempThread func     
        CycleThreadStop = True; 
        
        time.sleep(1.2);
        GetTempThreadPause = False;
        self.CycleStatus_Button = "Start Cycle";
        self.Cycle_Button_StatusCheck();
    
    def Cycle_Button_StatusCheck(self):        
        if (self.CycleStatus_Button=="Start Cycle"):            
            self.Pond_StartCycle_Button.Destroy();
            self.Pond_StartCycle_Button = wx.Button(self, -1, self.CycleStatus_Button, pos = (10,10 + 11*y), size = ButtonSize);        
            self.Bind(wx.EVT_BUTTON, self.Pond_StartCycle_Button_Click, self.Pond_StartCycle_Button);
        else:            
            print self.CycleStatus_Button;
            self.Pond_StartCycle_Button.Destroy();
            self.Pond_StartCycle_Button = wx.Button(self, -1, self.CycleStatus_Button, pos = (10,10 + 11*y), size = ButtonSize);        
            self.Bind(wx.EVT_BUTTON, self.Pond_StopCycle_Button_Click, self.Pond_StartCycle_Button);       
                
#--------------- Goto Temp        
    def Pond_GoToTemp_Button_Click(self, event):  
        global SetTemp, SlewInput, GetTempThreadPause, GetTempThreadAlreadyPaused;  
        
        GetTempThreadPause = True;  #Pause GetTempThread        
        print GetTempThreadPause;
        
        while(GetTempThreadAlreadyPaused == False):
            #print GetTempThreadAlreadyPaused;
            pass;
        GetTempThreadAlreadyPaused = False; #Reset the thread pause
                    
        SetTemp = int(self.InputTemp_Text.GetValue());
        SetSlew = int(self.SlewInput1_Text.GetValue());
        LowRange = SetTemp - 1;
        HighRange = SetTemp + 1;
        
        #time.sleep(0.01);
        #a_pond.fastSeek([LowRange,HighRange], returnImmediate=True)
        
        if (SetSlew == 0):
            a_pond.fastSeek([LowRange,HighRange], returnImmediate=True);
            SeekMode = "Fast mode"
            self.GetParent().SetStatusText = "Seek Mode: " + SeekMode;
        else:
            a_pond.slewSeek([LowRange,HighRange], SlewInput, returnImmediate=True);
            SeekMode = "Fast mode"
            self.GetParent().SetStatusText("Seek Mode: " + SeekMode);

        #self.SlewInput_Text.SetValue(str(SlewInput));                
        self.Log += "Set Temp = " + str(SetTemp) + "\n";
        self.Log_Text.SetValue(self.Log);        
        GetTempThreadPause = False; #Resume GetTempThread
        
#--------------- Save setting    
    def Pond_SaveSetting_Button_Click(self, event):
        global Temp1Input, Temp2Input, SlewInput, Dwel1Input, Dwel2Input, CyclesToDo;
        Temp1Input = int(self.Temp1Input_Text.GetValue());        
        Temp2Input = int(self.Temp2Input_Text.GetValue());            
        SlewInput = int(self.SlewInput_Text.GetValue());
        Dwel1Input = int(self.Dwel1Input_Text.GetValue());
        Dwel2Input = int(self.Dwel2Input_Text.GetValue());
        CyclesToDo = int(self.CyclesToDo_Text.GetValue());
        
        TempRecommendRange = range(-5,76);
        SlewRecommendRange = range(0,101);        
        
        if Temp1Input in TempRecommendRange:
            pass;
        else:           
            dlg = wx.MessageDialog(self, 'Temp1 is not in range. Recommended range -5 -> 75 degree C','POND GUI - Error',wx.OK | wx.ICON_ERROR);
            dlg.ShowModal();
            dlg.Destroy();
            return;
                    
        if Temp2Input in TempRecommendRange:
            pass;            
        else:
            dlg = wx.MessageDialog(self, 'Temp2 is not in range. Recommended range -5 -> 75 degree C','POND GUI - Error',wx.OK | wx.ICON_ERROR);
            dlg.ShowModal();
            dlg.Destroy();
            return;
        
        if SlewInput in SlewRecommendRange:
            pass;            
        else:
            dlg = wx.MessageDialog(self, 'Slew is not in range. Recommended range 0 -> 100 degree C/hour','POND GUI - Error',wx.OK | wx.ICON_ERROR);
            dlg.ShowModal();
            dlg.Destroy();
            return;
        
        SaveSettingThread(self);
        
       
#--------------- Manual Cycle     
    def Pond_ManualCycle_Button_Click(self,event):
        global Temp1Input, Temp2Input, SlewInput, Dwel1Input, Dwel2Input, CyclesToDo, Temp1Range, Temp2Range, CycleThreadStop
        global GetTempThreadPause
        
        ButtonStatus = self.Pond_ManualCycle_Button_StatusCheck();
        
        if(ButtonStatus == "Start"):
            dlg = wx.FileDialog(self, message = "Open cycle profile", defaultDir = os.getcwd(), defaultFile = "*.cyc", style = wx.OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                filepath = dlg.GetPath();
                file = csv.reader(open(filepath));
                for row in file:
                    Ram.append(row);
                                
            dlg.Destroy();
            
            IsFileLoaded = False;
            try:
                t = Ram[0];
                IsFileLoaded = True;
                pass
            except:
                print "Please load .cyc file"
            
            #ManualCyclesToDo variable is different from CyclesToDo. CyclesToDo is a variable set in the POND and it only has 1 ram from temp1->temp2
            #ManualCyclesToDo can have as many ram as user define:
            #Ex: Cycle 1 
            #Ram 1: 15C-25C  
            #Ram 2: 25C-60C
            #Ram 3: 60C-5C...
            #Cycle 2: repeat od cycle 1
            ManualCyclesToDo = 1;   
                    
            #Parsing data from file to variables
            if(IsFileLoaded == True):
                i = 0;
                NumberOfRam = len(Ram) - 1;
                for i in range(1,len(Ram)):
                    self.Log += "\n" + "Ram " + str(i) + ": " + str(Ram[i]);
                    self.Log_Text.SetValue(self.Log);
                
                self.ManualCycleStatus_Button = "Stop Temp Profile";
                status = self.Pond_ManualCycle_Button_StatusCheck();
                print status;
                
                ManualCycleThreadStart(self);
        else:
            CycleThreadStop = True;        
            time.sleep(1.2);            
            self.ManualCycleStatus_Button = "Start Temp Profile";
            self.Pond_ManualCycle_Button_StatusCheck();
            GetTempThreadPause = False

            
    def Pond_ManualCycle_Button_StatusCheck(self):        
        if (self.ManualCycleStatus_Button=="Start Temp Profile"):            
            self.Pond_ManualCycle_Button.Destroy();
            self.Pond_ManualCycle_Button = wx.Button(self, -1, self.ManualCycleStatus_Button, pos = (10,10 + 13*y), size = ButtonSize);        
            self.Bind(wx.EVT_BUTTON, self.Pond_ManualCycle_Button_Click, self.Pond_ManualCycle_Button);
            return "Start";
        else:            
            print self.ManualCycleStatus_Button;
            self.Pond_ManualCycle_Button.Destroy();
            self.Pond_ManualCycle_Button = wx.Button(self, -1, self.ManualCycleStatus_Button, pos = (10,10 + 13*y), size = ButtonSize);        
            self.Bind(wx.EVT_BUTTON, self.Pond_ManualCycle_Button_Click, self.Pond_ManualCycle_Button);
            return "Stop";
            
                                              
    
#--------- Run directly from SEACON drop down menu ---------
class MyFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, -1, 'Pond Control GUI', size = (430,500));
                
        #Define file menu bar
        menubar = wx.MenuBar();
        FileMenu = wx.Menu();
        HelpMenu = wx.Menu();
        
        self.COMmenu = wx.Menu();
        self.CreateStatusBar();        

        FileMenu.Append(wx.ID_OPEN,'Load cycle profile', 'Load file');           
        FileMenu.Append(wx.ID_EXIT,'Quit', 'Quit App');
        HelpMenu.Append(wx.ID_HELP,'Help me!!!','Help');        
        HelpMenu.Append(wx.ID_ABOUT,'About','About!!!');
        menubar.Append(FileMenu,'&File');
        menubar.Append(HelpMenu,'&Help!!!')
        
        #EVENTS
        self.Bind(wx.EVT_MENU, self.LoadFile_Menu, id = wx.ID_OPEN);
        
        self.Bind(wx.EVT_MENU, self.Help_Menu, id = wx.ID_HELP);
        self.Bind(wx.EVT_MENU, self.About_Menu, id = wx.ID_ABOUT);
                
        
        #self.COMPorts = [None]*len(self.ports);    #initialize an array
#        self.COMPorts = [];                         #initialize a list
#        
#        for i in range(0,len(self.ports)):           
#            self.COMPorts.append(self.COMmenu.AppendRadioItem(i, self.ports[i]));            
#            self.Bind(wx.EVT_MENU, self.COMPortChoosen, self.COMPorts[i])            
#
#        menubar.Append(self.COMmenu,'&COMPort');        
        self.SetMenuBar(menubar);
                
        p = MyPanel(self);     
    
    def COMPortChoosen(self, event):        
        for item in self.COMmenu.GetMenuItems():
            if item.IsChecked():
                #print str(item.GetId());
                FullPortName = self.ports[item.GetId()];
                print FullPortName;
                PondPort = FullPortName[FullPortName.index('M')+1:len(FullPortName)-1];
                print PondPort;
                a_pond.start(int(PondPort));
                #self.connect(PondPort);

    def LoadFile_Menu(self,event):        
        dlg = wx.FileDialog(self, message = "Open cycle profile", defaultDir = os.getcwd(), defaultFile = "*.cyc", style = wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            filepath = dlg.GetPath();
            file = csv.reader(open(filepath));
            for row in file:
                Ram.append(row);
#            file = open(filepath);
#            while True:
#                line = file.readline();                                
#                if not line:
#                    break;
#                Ram.append(line);            
        dlg.Destroy();
        
    def Help_Menu(self,event):
        Help_String  = "* Auto Connect: the application will auto scan COM port and find the POND port. When the POND port is found, the application will automatically connect to the POND PORT.\nRequired: wmi and pyWIN32 modules. This button will be disabled when those 2 modules are not found. \n\n";
        Help_String += "* Connect to COM: manually connect to COM port chosen by the user. Do not required wmi and pyWIN32 module like auto connect. This button is always available. \n\n";
        Help_String += "* Go to temp: Input desired temperature and POND will seek to that set point temperature with the Slew rate set in Cycles Setup section. \n";
		#Help_String += "* Rate: 0 is the fastest rate, 1-100: is xx degree per hour \n\n";
        Help_String += "----- Cycles Setup -----\n";
        Help_String += "* Temp1: First Cycle set point temperature (-5 to 75 degree C).\n";
        Help_String += "* Temp2: Second Cycle set point temperature (-5 to 75 degree C).\n";
        Help_String += "* Slew: Cycle slew rate (0 to 100 degree C/hour).\n";
        Help_String += "* Dwel 1: First cycle dwell time. The duration in seconds that POND will stay at Temp 1 (seconds).\n";
        Help_String += "* Dwel 2: Second cycle dwell time. The duration in seconds that POND will stay at Temp 2 (Seconds).\n";
        Help_String += "* Cycles to do: number of cycle the user want to perform.\n\n"
        Help_String += "* Save Setting to POND: save all the above variables to POND.\n\n";
        Help_String += "* Start cycle: Start the cycle.\n\n"
        Help_String += "* The temperature display will be refresh for every 1 second.\n";
        Help_String += "  * RAMP SP : current set point temperature.\n";
        Help_String += "  * STAGE : current temperature.\n";
        dlg = wx.MessageDialog(self, Help_String,'POND GUI - Help',wx.OK | wx.ICON_INFORMATION);
        dlg.ShowModal();
        dlg.Destroy();
    
    def About_Menu(self,event):       
        from wx.lib.wordwrap import wordwrap;          
        info = wx.AboutDialogInfo();
        info.Name = "POND GUI";
        info.Version = "1.0.0";
        info.Copyright = "(C) 2012 SEAGATE Programmers and Coders."
        info.Description = wordwrap("GUI for controlling POND Model No. K49.",350,wx.ClientDC(self));
        info.Developers = ["Anh Chu","Eric Mayers","SEAGATE Programmers and Coders"];
        wx.AboutBox(info);        
    
    def connect(self,PondPort):
        try:
            pyport.open(int(PondPort),9600);
        except:
            print 'unable to open com %s, baud %s. quitting.' % (PondPort, 9600);
            return 0;
        rtn = a_pond.sp.XferAndWait('?','Thermal Conditioner\r\n\r\n')
        print rtn
        if not rtn.lower().count('pond'):
            print 'failed to start communication with the pond.'
            a_pond.stop()
            return 0
        return 1             
    
def run():
	try:
		app = wx.PySimpleApp();
		frame = MyFrame(None);
		frame.CenterOnScreen();
		frame.Show();
		app.MainLoop();
	except:
		print "error";
	
if __name__ == '__main__':    
	try:
		run();
	except:		
		print "error";