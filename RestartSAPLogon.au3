
;-Begin------------------------------------------------------------------
 
   ;-Sub RestartProcess--------------------------------------------------
     Func RestartProcess($ProcessName)
 
       $PID = ProcessExists($ProcessName)
       If $PID Then
 
         $oWMI = ObjGet("winmgmts:\\.\root\CIMV2")
         If IsObj($oWMI) Then
           $Process = $oWMI.ExecQuery("Select * From win32_process " & _
             "Where Name = '" & $ProcessName & "'")
           If IsObj($Process) Then
             $ProcessPath = $Process.ItemIndex(0).ExecutablePath
           EndIf
         EndIf
 
         ProcessClose($PID)
         ProcessWait($PID)
 
         Run($ProcessPath)
 
       EndIf
 
     EndFunc
 
   ;-Sub Main------------------------------------------------------------
     Func Main()
 
       RestartProcess("saplogon.exe")
       WinWait("SAP Logon ")
       Sleep(4096)
       ClipPut("Restart SAP Logon successful")
 
       $SAPROT = ObjCreate("SapROTWr.SAPROTWrapper")
       If Not IsObj($SAPROT) Then
         Exit
       EndIf
 
       $SapGuiAuto = $SAPROT.GetROTEntry("SAPGUI")
       If Not IsObj($SapGuiAuto) Then
         Exit
       EndIf
 
       $application = $SapGuiAuto.GetScriptingEngine()
       If Not IsObj($application) Then
         Exit
       EndIf
 
       $connection = $application.Openconnection("NSP", True)
       If Not IsObj($connection) Then
         Exit
       EndIf
 
       $session = $connection.Children(0)
       If Not IsObj($session) Then
         Exit
       EndIf
 
       $session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "BCUSER"
       $session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = InputBox("Password", "Enter your password", "", "*")
       $session.findById("wnd[0]").sendVKey(0)
       $session.findById("wnd[0]/tbar[0]/okcd").text = "/nse38"
       $session.findById("wnd[0]").sendVKey(0)
       $session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = "ZRESTARTSAPLOGON"
       $session.findById("wnd[0]").sendVKey(8)
 
     EndFunc
 
   ;-Main----------------------------------------------------------------
     Main()
 
;-End--------------------------------------------------------------------
