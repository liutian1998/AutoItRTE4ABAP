
"-Begin------------------------------------------------------------------
   Program ZRESTARTSAPLOGON.
 
     "-Variables---------------------------------------------------------
       Data lo_AutoIt Type Ref To ZCL_AUTOIT.
       Data lv_WorkDir Type String.
       Data lv_Str Type String.
 
     "-Main--------------------------------------------------------------
       Try.
         Create Object lo_AutoIt.
         lv_Str = lo_AutoIt->GetClipBoard( ).
         If lv_Str <> 'Restart SAP Logon successful'.
           lo_AutoIt->ProvideRTE( i_DeleteExistRTE = ABAP_FALSE ).
           lv_WorkDir = lo_AutoIt->GETWORKDIR( ).
           lo_AutoIt->StoreInclAsFile( i_InclName = 'ZRESTARTSAPLOGONAU3'
             i_FileName = lv_WorkDir && '\RestartSAPLogon.au3').
           lo_AutoIt->Execute( i_FileName = 'RestartSAPLogon.au3'
             i_WorkDir = lv_WorkDir ).
         EndIf.
         lo_AutoIt->PutClipBoard( '' ).
 
         Write: / 'Restart successful'.
 
       Catch cx_root.
         Write: / 'An error occured'.
       EndTry.
 
"-End-------------------------------------------------------------------
