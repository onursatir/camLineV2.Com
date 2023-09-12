'*************************************************************************
'******    PROJECT                    : MRA2_IMSY
'******    PROJECT NUMBER             : Z:\projekte\carbody\Daimler_MRA2\03_EOL
'******    AUTHOR                     : Michael Kainzbauer, A CEP HUB1 RGB IE22
'******    TESTSTATION (optional)     : RBT3227S, RBT3226S, RBT3224S
'******    TEMPLAT VERSION			  : v1_02 : 30.06.16; Judmann Stefan 	
'*************************************************************************
'******    Changelist
'******    vx_xx : <Date>  ;<Name> ;<Comment>
'******    v1_00 : initial version
'******    v1_01 : 15.05.2019; MKA; Avtivate new MSA handling
'******    v1_02 : 18.07.2019; MKA; Correction of MSA handling in participation with new ATI statistic module
'******          : 09.08.2019; LST; Insert PDI Handling
'******    v1_03 : 20.09.2019; LST; Implement WipMama
'******    v1_04 : 24.02.2020; MKA; Implementation of StartButton flag for WipMama init
'******    v1_05
'****** 
'******    v1_00 : 04.05.2020; L.P.; Script for VW MQB19 GW  based on MRA2 V1_05
'****** 

'*************************************************************************

'To prepare this script for the top11 protocol needed 
'change or check the source code line which are marked with '!!Change here
'_________________________________________________________________________
'_________________________________________________________________________
'
' necessary ini file entries:
' 
' in axscripting.ini
' under the section [Scripts]
' VBScript = Path to this file
' under the section [Scriptable Objects]
' Top11=ATICOMM.Top11
'
' APC Event Handling:
'  -ATICOMM ocx version 1.04 is needed to be registered
'  -ATITPCNET.dll must be installed by ATISetup and enabled in Option dialog
'  -The hostname / process step(s) of the testsystem must be registered on TPC Server
'
' necessary testplan symbols (these symbols will be accessed by this script)
' SequenceLocals.OpuiLoopCount  write access (First Test = 1, Second Test = 2, >2 in MSA)
' SequenceLocals.ReadyToCheck read/write access if CreateReadyToCheck = True
'
'_________________________________________________________________________
'
' in opuiterminal.ini 
' under the section [Testplan]
' Testplan=testplan1;.......
' under the section [ModeOfSelectingTestplan/ModuleType]
' Mode=SelectTestplanFirst
' under the section [Options]
' EnableSPS=1
' 
'_________________________________________________________________________
'
'Global control variables and constants

Option Explicit

'Set the constants here         '!!Change here check all Const
Const RepeatOnFail = True		'Set here if we have repeat on fail. 
                                'RepeatOnFail = True: a failed module is tested a second time
Const Top11Port = 6				'Set here the serial port of the Top11 communication
Const Top11BaudRate = 19200     'Set here the baud rate of the Top11 communication
Const CreateReadyToCheck = TRUE 'Set here if you want to enable ready to check button
Const RunInstanz = False	 	'Set here true if this script runs under an instanz of more than one TestExecs
Const ResetInstruments = True   'Set here false if you don't want to reset the instruments during handling time of the dut
Const ProvideAPCEvents = False 	'Set here true if you want to provide APC Events. 
Const PDI_Test_activ = False	'Set here if you want to test PDI in script
Const WipMama_Test_activ = True'Set here if you want to test WipMama in script
								'This feature can only be used if the TPC / Jidoka ATI-Modul is installed 
								'and if the PLC send the information about the APC events

Dim ValidTestplanLoaded 		'False: no valid mutiple up testplan loaded
Dim Top11State          		'Contains the actual top11 state      
Dim Top11QuitRequest    		'Top11QuitRequest 1: Stop Top11
Dim MSA                 		'False: MSA off, True: MSA on
'Dim MSA_Counter		    		'MSA_Counter from ATI_Modul, introduced in v1_02
Dim OpuiLoopCount       		'current OpuiLoopCount (First Test = 1, Second Test = 2, Repair Test = 3) 
								'is stored in in the testplan symbol SequenceLocals.OpuiLoopCount 
Dim ModuleFailed 				'ModuleFailed 1: the module failed
Dim Response(4)					'buffers for the response of the top11 slave, increase the number of array elements for APC Events
Dim StateId						'response of the parameter s_id of the top11 state command
Dim StatisticType				'response of the parameter s_statistic_type of the top11 state command
Dim Result						'buffer to set the result string for the result command
Dim CheckStartEvent				'true: the start button was pressed
Dim InstProcess					'contains ProcessStep if RunInstanz = True
Dim ProcessStep					'contains the ProcessStep retrieved by event OnProvideSystemInfo
Dim Hostname					'contains the Hostname retrieved by event OnProvideSystemInfo

Dim PDI_Config					'PDI config file
Dim PDI_Init_State				'PDI init status
Dim CL_ResultCode				'CamLine result code
Dim CL_ResultMsg				'CamLine result message
Dim CL_serialNumberType			'CamLine serial number type
Dim CL_InitFail                 'Flag indicates CamLine init error  
Dim StartButtonPressed			'Flag=1 if Start Button in Opui is pressed

If WipMama_Test_activ Then
	Dim camLine
	Set camline = CreateObject("Continental.camLine.MES")
	Const CL_ProcessStep = "EP" 	'ProcessStep for camLine init
End If

'Const symbols to use for call ATI.CreateExternalDataReference 
Const T_INT32 = 1 
Const T_REAL64 = 2
Const T_STRING = 3 
Const T_INT32ARRAY = 4  
Const T_REAL64ARRAY = 5 
Const T_STRINGARRAY = 6 
Const T_INT32ARRAYDIM2 = 7
Const T_REAL64ARRAYDIM2 = 8
Const T_STRINGARRAYDIM2 = 9

'variables to used for script and testplan handling
Dim strHostname         		'hostname used in Top11_Init
Dim TestStationNumber			'in i_id_equipment used teststation no. for EP 001-004
Dim strid_equipment     'i_id_equipment in init command depends on hostname (station) 
Dim strid_equ_resp      'response parameter for id_equipment 

Dim strStatus			'Response(2) of state command
Dim strStationID		'station id chars 1-2 of s_id 
Dim strAdapterNumber 	'adapter number chars 5-6 of s_id 
Dim strPalletNumber		'pallet number chars 9-12 of sid
Dim strHousingDMC		'housing DMC char 15-35 in s_id (20 char length) 
Dim strTestType         'Default or SBZ (screening) out of s_statistic-type

'_________________________________________________________________________
'Event handlers of ActiveX Controls

'this event is fired by the Top11 control in case of a communication error
Sub Top11_OnError(Code, Description)
   ATI.SendReportMessage "Top11 Error " & Code & ": " & Description
   ATI.SendMessage "Top11 Error " & Code & ": " & Description, "OpuiTerminal", "StateLine"
   Top11State = "Top11_Error"
End Sub


Sub Top11_OnSend(Message)
   'if you debug the script in this event you can get an top11 error (char timeout)
   ATI.SendReportMessage "Send: " + Message
   ATI.SendMessage "Send: " + Message, "OpuiTerminal", "StateLine"
End Sub


Sub Top11_OnReceive(Message)
   ATI.SendReportMessage "Receive: " + Message
   ATI.SendMessage "Receive: " + Message, "OpuiTerminal", "StateLine"
End Sub


'_________________________________________________________________________
'Startup code of script, loaded after the testplan is loaded

'first available ATI event
'create testplan symbol reference needed
ATI.CreateExternalDataReference "OpuiLoopCount", "SequenceLocals" + Chr(10) + "OpuiLoopCount", T_INT32 , False
If CreateReadyToCheck then
   ATI.CreateExternalDataReference "ReadyToCheck", "SequenceLocals" + Chr(10) + "ReadyToCheck", T_INT32, False
End if
'ATI.CreateExternalDataReference "strStationID", "SequenceLocals" + Chr(10) + "strStationID", T_STRING, False
ATI.CreateExternalDataReference "strAdapterNumber", "SequenceLocals" + Chr(10) + "strAdapterNumber", T_STRING, False
ATI.CreateExternalDataReference "strPalletNumber", "SequenceLocals" + Chr(10) + "strPalletNumber", T_STRING, False
ATI.CreateExternalDataReference "strHousingDMC", "SequenceLocals" + Chr(10) + "strHousingDMC", T_STRING, False
ATI.CreateExternalDataReference "strTestType", "SequenceLocals" + Chr(10) + "strTestType", T_STRING, False
ATI.CreateExternalDataReference "ModuleType", "System" + Chr(10) + "ModuleType", 3, FALSE
ATI.CreateExternalDataReference "Selected_Version", "SequenceLocals" + Chr(10) + "Selected_Version", 3, FALSE

'set default value of the internal buffers
MSA = False 
ValidTestplanLoaded = False
Top11.Port = Top11Port
Top11.BaudRate = Top11BaudRate
CheckStartEvent = False
StartButtonPressed = 0

'get the instanz process step
InstProcess = "ATI Script"
If RunInstanz then
   Dim Buffer
   Buffer = ATI.GetATIIni
   Dim Found1, Found2
   Found1 = InStrRev(Buffer, "\")
   If Found1 > 0 then
	  Found2 = InStrRev(Buffer, "\", Found1-1)
	  If Found2 > 0 then
		 Found1 = InStrRev(Buffer, "\", Found2-1)
		 If Found1 > 0 then
		   InstProcess = "ATI Script: " & Mid(Buffer,Found1+1, Found2-Found1-1)
		 End if 
	  End if
   End if
End If

If WipMama_Test_activ Then
	camline.LoadDll CL_ResultCode, CL_ResultMsg
	If CL_ResultCode <> 0 Then
		MsgBox "Error during loading CamLine.dll; ResultCode: " & CL_ResultCode &" message: " & CL_ResultMsg
	End If
End If

'check if we have a valid testplan loaded
ValidTestplanLoaded = True

If ATI.IsValidExternalDataReference("OpuiLoopCount")=0 then
    MsgBox "No valid testplan loaded: SequenceLocals.OpuiLoopCount (INT32) is missing",0,InstProcess 
	ValidTestplanLoaded = False
End if 
'If ATI.IsValidExternalDataReference("strStationID")=0 then
 ' MsgBox "No valid multiple up testplan loaded: SequenceLocals.strStationID (STRING) is missing" ,0,InstProcess
 ' ValidTestplanLoaded = False
'End if
If ATI.IsValidExternalDataReference("strAdapterNumber")=0 then
  MsgBox "No valid multiple up testplan loaded: SequenceLocals.strAdapterNumber (STRING) is missing" ,0,InstProcess
  ValidTestplanLoaded = False
End if
If ATI.IsValidExternalDataReference("strPalletNumber")=0 then
  MsgBox "No valid multiple up testplan loaded: SequenceLocals.strPalletNumber (STRING) is missing" ,0,InstProcess
  ValidTestplanLoaded = False
End if
If ATI.IsValidExternalDataReference("strHousingDMC")=0 then
  MsgBox "No valid multiple up testplan loaded: SequenceLocals.strHousingDMC (STRING) is missing" ,0,InstProcess
  ValidTestplanLoaded = False
End if
If ATI.IsValidExternalDataReference("strTestType")=0 then
  MsgBox "No valid multiple up testplan loaded: SequenceLocals.strTestType (STRING) is missing" ,0,InstProcess
  ValidTestplanLoaded = False
End if
IF ATI.IsValidExternalDataReference("ModuleType")=0 THEN
  MsgBox "No valid multiple up testplan loaded: System.ModuleType (STRING) is missing" ,0,InstProcess
  ValidTestplanLoaded = FALSE
END IF		
IF ATI.IsValidExternalDataReference("Selected_Version")=0 THEN
  MsgBox "No valid multiple up testplan loaded: SequenceLocals.Selected_Version (STRING) is missing" ,0,InstProcess
  ValidTestplanLoaded = FALSE
END IF

'Delete ReadyToCheck Button
If CreateReadyToCheck then  
 ATI.SendMessage "ToolButtonRemove", "OpuiTerminal", ""
 If ATI.IsValidExternalDataReference("ReadyToCheck")<>0 then
   'Create Button ReadyToCheck 
   ATI.SetInt32 "ReadyToCheck", 0
   ATI.SendMessage "ToolButtonAdd", "OpuiTerminal", "Ready to check is OFF !|ReadyToCheck|0||AXScripting|OnEvent"	
 Else
    MsgBox "No valid testplan loaded: SequenceLocals.ReadyToCheck (INT32) is missing",0,InstProcess 
	ValidTestplanLoaded = False
 End If
End If

'_________________________________________________________________________
'ATI event handlers

Sub OnEvent_SequenceBegin(hPB)
   'this event is fired by the TestExecutive at the start of a testplan run
   'Nothing to do here
End Sub


Sub OnEvent_SequenceEnd(hPB)
   'this event is fired by the TestExecutive at the end of a testplan run
   If ValidTestplanLoaded = False then
	Exit Sub
   End if

   If CheckStartEvent = False then
       MsgBox "Testplan was started directly without Top11 communication" & vbCrLf & "To run this script Enable SPS Event in OpuiTerminal options",0,InstProcess
       Exit Sub
   End if

   Dim State
   ATI.PBGetInt32 hPB, "State", State

   Select case State
      Case 1: 'Sequencer is idle
         'open session again
         If Top11.OpenSession <> 0 then
            Top11State = "Top11_Error"
            ATI.SendReportMessage "Top11.OpenSession failed!"
            MsgBox "Top11Script OnEvent_SequenceEnd: Top11.OpenSession failed!",0,InstProcess
	        Exit Sub
         Else
            ATI.SendReportMessage "Top11.OpenSession is ok!"
            ATI.SendMessage "Top11.OpenSession is ok!", "OpuiTerminal", "StateLine"
         End If

         'decontact module at test end
       
         ATI.SendReportMessage "Top11 State = " + Top11State
         ATI.SendMessage "Top11 State = " + Top11State, "OpuiTerminal", "StateLine"

		 Result = "R*" ' default recontact
    'MSA on?
	If (MSA <> True) Then 	' Or (MSA = True And MSA_Counter < 50 And (MSA_Counter MOD 10 = 0)) Then
         If ModuleFailed = 1 then
		   If RepeatOnFail = False Or (RepeatOnFail = True And OpuiLoopCount > 1) then
		     Result = "F*" 'failed
           End If
         Else 'module passed
		   Result = "P*" 'passed
         End If
	 End If	 

	If ProvideAPCEvents then
	  'APC Events are provided by PLC in the additional top11 parameter
	  'always 60 characters with the format APCEventType = 10 characters, APCEventMessage = 50 characters
	  'Top11.Top11ResultDecontactEx StateId, StatisticType, Result, "************************************************************", _
	   '		Response(0), Response(1), Response(2), Response(3)  '!!Change here
	   Top11.Top11ResultDecontact StateId, StatisticType, Result, Response(0), Response(1), Response(2) '!!Change here
	Else        
	  'send command result
	   Top11.Top11ResultDecontact StateId, StatisticType, Result, Response(0), Response(1), Response(2) '!!Change here
	End If

	If ResetInstruments = True then
	   ATI.ResetInstruments
	End if
			
    If Result <> "R*" then
	   If Top11State <> "Top11_Error" Then
	      'check if echoed?
	      If Response(0) = StateId And Response(1) = StatisticType And Response(2) = Result Then '!!Change here
		   OpuiLoopCount = 0
		   Top11State = "Top11_StateContact" '!!Change here
	      Else
		  MsgBox "Critical Error in Top11 Result Command: parameter not echoed!",0,InstProcess
		  ATI.SendReportMessage "Error in Top11 Result Command: parameter not echoed!"
                  ATI.SendMessage "Error in Top11 Result Command: parameter not echoed!", "OpuiTerminal", "StateLine"
		  Top11State = "Top11_Error"
	      End If
	    End If
        Else ' recontacted
	    If Top11State <> "Top11_Error" Then
		'check if echoed?
		If Response(0) = StateId And Response(1) = StatisticType And Response(2) = Result Then '!!Change here
		    Top11State = "Top11_StateContact" '!!Change here
		Else
		    MsgBox "Critical Error in Top11 Result Command: parameter not echoed!",0,InstProcess
		    ATI.SendReportMessage "Error in Top11 Result Command: parameter not echoed!"
                    ATI.SendMessage "Error in Top11 Result Command: parameter not echoed!", "OpuiTerminal", "StateLine"
		    Top11State = "Top11_Error"
		End If
	    End If 		
	End if
		
	'Handle APCEvents
	If ProvideAPCEvents then
	    SendAPCEventsToTPCServer Response(3)
	End if
		
      Case Else: 'Sequencer ended with error state
         Top11State = "Top11_Quit"
         ATI.SendReportMessage "Sequencer ended with error state: Top11 State = " + Top11State
         ATI.SendMessage "Sequencer ended with error state: Top11 State = " + Top11State, "OpuiTerminal", "StateLine"
   End select
End Sub


Sub OnEvent_OpuiStartTestingRequest(hPB)
   ' this event is fired by the OpuiTerminal when the operator presses the Start button

   If ValidTestplanLoaded = False then
	MsgBox "No valid testplan loaded!",0,InstProcess
	Exit Sub
   End if
   
   StartButtonPressed=1
   'MsgBox "StartButtonPressed value: " +  Cstr(StartButtonPressed)
   
   Dim SingleStep
   ATI.PBGetInt32 hPB, "StartInTraceMode", SingleStep
   ATI.SendReportMessage "SingleStep = " + CStr(SingleStep)
   Top11QuitRequest = 0

   CheckStartEvent = True
   'start top 11 loop
   Run_Top11 SingleStep
End Sub


Sub OnEvent_OpuiStopTestingRequest(hPB)
   'this event is fired by the OpuiTerminal when the operator presses the Stop button
   Top11QuitRequest = 1
   StartButtonPressed=0
   'MsgBox "StartButtonPressed value: " + CStr(StartButtonPressed)
   ATI.SendReportMessage "Stop testing request, please wait until testplan has finished!"
   ATI.SendMessage "Stop testing request, please wait until testplan has finished!", "OpuiTerminal", "StateLine"
End Sub


Sub OnEvent_ProvideModuleInfo(hPB)
   'this event is fired by the TestExecutive to provide pass-fail information
   'get the test result info from 
   ATI.PBGetInt32Array hPB, "FailFlags", 0, ModuleFailed

End Sub

'added in v1_01 for APC Events
Sub OnEvent_ProvideSystemInfo(hPB)
   ' get parameter values needed for APC Events
   Hostname = ATI.PBGetString(hPB, "TeststationID")
   ProcessStep = ATI.PBGetString(hPB, "ProcessStep")
End Sub


Sub OnEvent_StatisticTesttypeChange(hPB)
   Dim TestType
   Dim Counter

   ATI.PBGetInt32 hPB, "TestType", TestType
   ATI.PBGetInt32 hPB, "Counter", Counter

   'MSA_Counter = Counter - 1
   'MSA activated?
   If TestType = 1 AND Counter > 0 Then
      ATI.SendReportMessage "TestType = MSA (" + CStr(Counter) + ")"
      ATI.SendMessage "TestType Changed = MSA (" + CStr(Counter) + ")", "OpuiTerminal", "StateLine"
      MSA = True
   Else
      ATI.SendReportMessage "TestType changed to normal mode"
      ATI.SendMessage "TestType changed to normal mode", "OpuiTerminal", "StateLine"
      MSA = False
   End If
End Sub


Sub OnEvent_ReadyToCheck(hPB)
 Dim ReadyToCheck
 If Top11State <> "Top11_RunTestplan" And CreateReadyToCheck then
   ATI.GetInt32 "ReadyToCheck", ReadyToCheck
   If ReadyToCheck = 1 Then
     ATI.SetInt32 "ReadyToCheck", 0
     ATI.SendMessage "ToolButtonChangeText","OpuiTerminal","Ready to check is ON !|Ready to check is OFF !"    	
   Else
     ATI.SetInt32 "ReadyToCheck", 1
     ATI.SendMessage "ToolButtonChangeText","OpuiTerminal","Ready to check is OFF !|Ready to check is ON !"	     
   End If
 End If
End Sub

'added in v1_03
Dim FirstFailedTestErrorCode
Dim FirstFailedTestErrorMessage
Dim FirstFailedTestResult
Dim FirstFailedTestHighLimit
Dim FirstFailedTestLowLimit
Dim FirstFailedTestUnit

'collect the first failed test information over all uut positions, single test -> 0 array index
Sub OnEvent_OpuiProvideErrorCodes(hPB)
   FirstFailedTestErrorCode = ATI.PBGetStringArray(hPB, "ErrorCodes", 0)
   FirstFailedTestErrorMessage = ATI.PBGetStringArray(hPB, "ErrorMessages", 0)	
   FirstFailedTestHighLimit = ATI.PBGetStringArray(hPB, "HighLimits", 0)
   FirstFailedTestLowLimit = ATI.PBGetStringArray(hPB, "LowLimits", 0)
   FirstFailedTestUnit = ATI.PBGetStringArray(hPB, "TestUnits", 0)
   FirstFailedTestResult = ATI.PBGetStringArray(hPB, "TestResults", 0)		   
End Sub
'End of ATI event handlers
'_________________________________________________________________________

'helper functions and subroutines

Sub Run_Top11 (SingleStep)
   'this procedure handles the complete top11 protocol
   Dim i					'loop variable
   Dim Dummy
   Dim ModuleType			'contains the ModuleType which is read from p_type out of PDI
   Dim Selected_Version		'contains the Selected_Version which is read from p_dtp_version out of PDI
   
   'try to open the session
   If Top11.OpenSession <> 0 then
      'exit this sub with error
      ATI.SendReportMessage "Top11.OpenSession failed!"
      MsgBox "Top11Script OnEvent_OpuiStartTestingRequest: Top11.OpenSession failed!",0,InstProcess
      Exit Sub
   End if

   'go on with TOP11 communications
   ATI.SendReportMessage "Top11.OpenSession is ok!"

   'initialize TOP11 command scheduler 
   Top11State = "Top11_Init"

   Dim Running
   Running = True

   ATI.SendReportMessage "Starting Top11 communication"
   ATI.SendMessage "Starting Top11 communication", "OpuiTerminal", "StateLine"

   'TOP11 command scheduler
   '============================
   While Running = True 

     ' reset response variables to default value
     For i=0 to 2
        Response(i) = "*"
     Next

    'allow to quit loop only in Top11_Init or Top11_StateContact and in MSA in Top11_ResultDecontact
     If Top11QuitRequest=1 And  (Top11State = "Top11_Init"  Or _
                                 Top11State = "Top11_StateContact" Or _
				 (Top11State = "Top11_ResultDecontact" And MSA = True)) Then   '!!Change here
	Top11State = "Top11_Quit"
     End If

     'TOP11 state machine
     '-------------------
     Select case Top11State

	'_____________________________________________________________________________________________________
        Case "Top11_Init":
		'init command
       '0001##i_id_equipment##i_reference#<eot>
       '0002##i_id_equipment##i_reference#<eot>
       '0001##F001##00#<eot>
       '0002##S001##00#<eot>
       ' 
       'i_id_equipment:	4 character in sum
	   '				1 character: functional test			F
       '				3 character  numerical identification 	001: EP1
	   '														002: EP2
	   '														
       '
       'i_reference:		2 characters:	specific data of the system, default = 00 

          'initialize OpuiLoopCount for new test
	  OpuiLoopCount = 0

	   'send command
	   'Get hostname defined in ATI.ini
	   strHostname = ATI.GetString ("strComputerName")
  
	   Select Case strHostname
	   Case "RBT4170N":               		  'EOL1
			TestStationNumber="001"
	   Case "RBT4171N":				  'EOL2
			TestStationNumber="002"
			
	   Case Else
			MsgBox "This hostname is in Top11Init not allowed: "&Hostname,0,InstProcess
	   End Select
	   
	   strid_equipment="F"&TestStationNumber
	   
	   Top11.Top11Init strid_equipment, "00", Response(0), Response(1)  '!!Change here

	
        If ResetInstruments = True then
	      ATI.ResetInstruments
	   End if
			
	   'no error?
	   If Top11State <> "Top11_Error" then
	     'reply OK?
	     strid_equ_resp="S"&TestStationNumber
		 If Response(0) = strid_equ_resp AND Response(1) = "00" then  '!!Change here
			Top11State = "Top11_StateContact"  '!!Change here
	     Else
		 ATI.SendReportMessage "Error in Top11 Response of Init"
		 ATI.SendMessage "Error in Top11 Response of Init", "OpuiTerminal", "StateLine"
		 Top11State = "Top11_Error"
	     End If
	   End If
	   
		If PDI_Test_activ Then
            ' PDI init for MRA2_IMSY
            PDI_Config = "k:\USERS\VW\VW_MQB_19_GW\\verriegelung\vw_mqb19_gw_pdiconfig.cfg"
            PDI_Init_State = PDI.PdiInit_StandardRbg01 (PDI_Config, "EP", "PDI_TEST_PLAN_ID_REV", "OFF")
		End If
		If WipMama_Test_activ Then
			If (StartButtonPressed=1) Or (CL_InitFail=1) Then
				'MsgBox "StartButton pressed, WipMama wird durchgefuehrt!"
				CL_ResultCode = -1
				CL_ResultMsg = "Unset"
				CL_InitFail = 1
				' GHP Client must be configured and running for your hostname.
				camline.Init_StandardRbg01 CL_ProcessStep, CL_ResultCode, CL_ResultMsg
				If CL_ResultCode <> 0 then
					ATI.SendReportMessage "Error in camLine.Init: " & CL_ResultCode & "(" &CL_ResultMsg &")"
					ATI.SendMessage  "Error in camLine.Init: " & CL_ResultCode & "(" &CL_ResultMsg &")", "OpuiTerminal", "StateLine"
					Top11State = "Top11_Error"
					MsgBox "Der Init für WIPMAMA ist fehlgeschlagen. Wurde der GHP-Client aktiviert?"
					CL_InitFail=1
				Else
					StartButtonPressed=0
					CL_InitFail= 0
				End If
			End If
			If CL_InitFail<>1 Then
				CL_ResultCode = -1
				CL_ResultMsg = "Unset"
				camline.Reset_StandardRbg01 CL_ResultCode, CL_ResultMsg
				If CL_ResultCode <> 0 then
					ATI.SendReportMessage "Error in camLine.Reset: " & CL_ResultCode & "(" &CL_ResultMsg &")"
					ATI.SendMessage  "Error in camLine.Reset: " & CL_ResultCode & "(" &CL_ResultMsg &")", "OpuiTerminal", "StateLine"
					Top11State = "Top11_Error"
				End If
			End If	
		End If
		
         '____________________________________________________________________________________________________
         Case "Top11_StateContact":
        'status command
        '0023##s_id##s_statistic_type##s_status#<eot>
        '0024##s_id##s_statistic_type##s_status#<eot>
        '0023##****************************************##**##**#<eot>   without UUT
        '0024##123456789012345678901234******W010011234##F*##X*#<eot>
		
		's_id: 40 chars :   30 character serial number (housing dmc), rest filled with *
		'					6 character pallet number (W00001-W99999); ******=no pallet
		'					4 character adapter number (0001-9999); ****=no pallet	
        '
        's_statistic_type:	2 characters:	F* = first test
        '									R* = repair test
		'									S* = Special Test -> Line specific testsequence
        '
        's_status:		    2 characters:	** = device not present
        '       			                E* = Error during contaction (cylinder movement)
        '       			                X* = unknown state -> start test
		'


	    'send command
		
	    If ProvideAPCEvents then
   	       'APC Events are provided by PLC in the additional top11 parameter
	       'always 60 characters with the format APCEventType = 10 characters, APCEventMessage = 50 characters
	       Top11.Top11StateContactEx "********************************************", "**", "**", "************************************************************", _
		                         Response(0), Response(1), Response(2), Response(3)  '!!Change here
	    Else
	       Top11.Top11StateContact "****************************************", "**", "**", Response(0), Response(1), Response(2)  '!!Change here
	    End If
		   
	    'no error?
	    If Top11State <> "Top11_Error" then
				'uut's ready to test?
				If Not(Response(0)="" AND Response(1)="" AND Response(2)="") Then  '!!Change here
					'check if response has the right format
					If Len(Response(0))<>40 Or Len(Response(1))<>2 Or Len(Response(2))<>2 Then  '!!Change here
						ATI.SendReportMessage "Error in Top11 Response of StateContact"
						ATI.SendMessage "Error in Top11 Response of StateContact", "OpuiTerminal", "StateLine"
						Top11State = "Top11_Error"
					Else
						'buffer response for the result command
						StateId = Response(0)
						StatisticType = Response(1)
						strStatus=Response(2)
			
						'Set testplan symbols if needed '!!Change here
						If strStatus ="X*" Then
							strHousingDMC=Left(StateId,30)
							strPalletNumber=Mid(StateId,31,6)
							'strStationID=Left(StateId,2)
							strAdapterNumber=Mid(StateId,37,4)
							
							
							
							'MsgBox (StateId)
							'MsgBox (strAdapterNumber)
							'MsgBox (strHousingDMC)
							
											
							Select case StatisticType
							Case "S*":
								strTestType="SBZ"
							Case Else
								strTestType="Default"
							End Select	
				
							'ATI.SetString "strStationID", strStationID
							ATI.SetString "strAdapterNumber",strAdapterNumber
							ATI.SetString "strPalletNumber", strPalletNumber
							ATI.SetString "strHousingDMC", strHousingDMC
							ATI.SetString "strTestType", strTestType
			
							'ATI.SendReportMessage "strStationID: "&strStationID
							ATI.SendReportMessage "strAdapterNumber: "&strAdapterNumber
							ATI.SendReportMessage "strPalletNumber: "&strPalletNumber
							ATI.SendReportMessage "strHousingDMC: "&strHousingDMC
							ATI.SendReportMessage "strTestType: "&strTestType
																
							If PDI_Test_activ Or WipMama_Test_activ Then
								Top11State = "CHECK_PDI_WIPMAMA" 		'check and get data with PDI of WIPMAMA
							Else
								Top11State = "Top11_RunTestplan"  '!!Change here if you have a ProgramLoad after StateContact
							End If
						End If
					End If     'repeat until a valid response exists
			    End If
	      'Handle APCEvents
	      If ProvideAPCEvents then
			SendAPCEventsToTPCServer Response(3)
	      End if
	    End If
         '___________________________________________________________________________________________________
        Case "CHECK_PDI_WIPMAMA":
            Top11State = "Top11_RunTestplan"  
            If PDI_Test_activ = True then
				Dummy = PDI.PdiReset_StandardRbg01()
                Dummy = PDI.PdiTest_StandardRbg01(strHousingDMC, 0, 0)
				
				Dummy = PDI.PdiGetParam(strHousingDMC, "EP", "p_type", ModuleType)
                ATI.SendReportMessage "PDI p_type " & ModuleType & " <Return=" & Dummy & ">"
				Dummy = PDI.PdiGetParam(strHousingDMC, "EP", "p_dtp_version", Selected_Version)
                ATI.SendReportMessage "PDI p_dtp_version " & Selected_Version & " <Return=" & Dummy & ">"
			Else
				CL_ResultCode = -1
				CL_ResultMsg = "Unset"
				CL_serialNumberType = ""

				camline.UnitCheck_StandardRbg01 strHousingDMC, CL_serialNumberType, CL_ResultCode, CL_ResultMsg
				If (CL_ResultCode <> 0) Then
					'ATI.SendReportMessage "position" & CStr(i + 1) & " disabled wegen WIPMAMA. SerNo: " & strHousingDMC
					'Top11State = "Top11_ResultFailed" 					
				Else
					' ModuleType
					' Für MLFB/A2C (früher "p_type") gibt es noch keine Funktion !!!!
					'Dummy = PDI.PdiGetParam(UutSerialNumbers(i), "PROGRAM", "p_type", ModuleType)
					'ATI.SendReportMessage "WIPMAMA ModuleType " & ModuleType & " <ResultCode=" & CL_ResultCode & " " & CL_ResultMsg& ">"
					' evtl. workaround temporär p_type im attachement??????
					camline.GetAttachmentParameter_StandardRbg01 strHousingDMC, CL_serialNumberType, "p_type", ModuleType, 1024, CL_ResultCode, CL_ResultMsg
					ATI.SendReportMessage "WIPMAMA p_type " & ModuleType & " <ResultCode=" & CL_ResultCode & " " & CL_ResultMsg& ">"
			
					' DTP selection infos
					camline.GetAttachmentParameter_StandardRbg01 strHousingDMC, CL_serialNumberType, "p_dtp_version", Selected_Version, 1024, CL_ResultCode, CL_ResultMsg
					ATI.SendReportMessage "WIPMAMA p_dtp_version " & Selected_Version & " <ResultCode=" & CL_ResultCode & " " & CL_ResultMsg& ">"
				End If
			End If 
            ATI.SetString "ModuleType", ModuleType						' read out of PDI(p_type)
			ATI.SetString "Selected_Version", Selected_Version			' read out of PDI(p_dtp_version)
			
	    'If Top11State <> "Top11_ResultFailed" then			
			Case "Top11_RunTestplan":
				Top11State = "Top11_WAIT"
				'default test state is failed   
				ModuleFailed = 1
	
				'close session to allow operating commands during test sequence
				Top11.CloseSession

				'Increase LoopCount only if not MSA
				If MSA <> True Then
					'increment "OpuiLoopCount"
					OpuiLoopCount = OpuiLoopCount + 1
				End If

				ATI.SetInt32 "OpuiLoopCount", OpuiLoopCount

				If SingleStep = 1 then
					ATI.StepSequence
				Else
					ATI.SendMessage "PreRun: Loading Dlls, please wait", "OpuiTerminal", "StateLine"
					ATI.RunSequence
				End if
  		'End if
		
		'__________________________________________________________________________________________________
	 'Case "Top11_ResultFailed"
		' Do Nothing

		'__________________________________________________________________________________________________
	 Case "Top11_WAIT"
	   ' Do Nothing

  	 '__________________________________________________________________________________________________
	 Case "Top11_Quit"
	   Running = False
	   ATI.SendReportMessage "Quitting Top11 communication"
	   ATI.SendMessage "Quitting Top11 communication", "OpuiTerminal", "StateLine"

	'__________________________________________________________________________________________________
	Case "Top11_Error"
	   Top11State = "Top11_Init"
	   ATI.SendReportMessage "Reinitializing Top11 due to communication error"
	   ATI.SendMessage "Reinitializing Top11 due to communication error", "OpuiTerminal", "StateLine"
	   Wait 1

    End Select
    'end of TOP11 state machine
    '--------------------------
    'make sure that other events will be handled
     ATI.DoEvents

  Wend

  Top11.CloseSession
  CheckStartEvent = False

End Sub

'Sub to send an APC event to the TPC Server added in v1_01
Sub SendAPCEventsToTPCServer(APCEventFromTop11)
    'Check if we have all information to send APC Events
    If Hostname <> "" And ProcessStep <> "" then
	'Split up the Top11Response (depends on the SPS implementation)
	'and remove the fill character *
	Dim EventType, EventMessage
	EventType = RTrim(Replace(Left(APCEventFromTop11 , 10), "*", " "))
	EventMessage = RTrim(Replace(Right(APCEventFromTop11 , 50), "*", " "))
	'send event only if we got one
	If EventType <> "" then
		'Send APC Event to TPC Server
		ATI.SendMessage "ProvideEventData","ATITPCNET", Hostname & "|" & ProcessStep & "|" & EventType & "|" & EventMessage 
	End if
    End If
End Sub

'This sub waits sec seconds
Sub Wait(WaitTime)
   Dim StartTime
   Dim StopTime
   Dim CurrentTime

   'Timer returns the number of seconds since midnight
   StartTime = Timer()
   Do 
      ATI.DoEvents
      StopTime = Timer()
      If StartTime <= StopTime then 
	CurrentTime = StopTime - StartTime
      Else
	CurrentTime = 86400 - StartTime + StopTime
      End If      
   Loop While WaitTime > CurrentTime
End Sub

'End of helper functions and subroutines
'_________________________________________________________________________
'End of file


