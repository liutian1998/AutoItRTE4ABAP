
"-Begin-----------------------------------------------------------------
  Program ZTESTAUTOITRTE.

    Data:
      lo_AutoIt Type Ref To ZCL_AUTOIT,
      lv_WorkDir Type String,
      lv_Str Type String
      .

    "-Main--------------------------------------------------------------
    Create Object lo_AutoIt.
    Try.
      lo_AutoIt->ProvideRTE( i_DeleteExistRTE = ABAP_TRUE ).

      "-Get path of SAP work directory----------------------------------
      lv_WorkDir = lo_AutoIt->GETWORKDIR( ).

      "-Read AutoIt code from include and store it as file--------------
      lo_AutoIt->StoreInclAsFile( i_InclName = 'ZINPUTBOXAU3'
        i_FileName = lv_WorkDir && '\InputBox.au3').

      "-Execute AutoIt with InputBox.au3--------------------------------
      lo_AutoIt->Execute( i_FileName = 'InputBox.au3'
          i_WorkDir = lv_WorkDir ).

      "-Get result from clipboard---------------------------------------
      lv_Str = lo_AutoIt->GetClipBoard( ).

      Write: / lv_Str.

      lo_AutoIt->DeleteFile( i_filename = lv_WorkDir && '\InputBox.au3' ).
      lo_AutoIt->DeleteExistingRTE( ).

    Catch cx_Root.
      Write: / 'An error occured'.
    EndTry.

"-End-------------------------------------------------------------------
