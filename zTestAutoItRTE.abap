
"-Begin-----------------------------------------------------------------
  Program ZTESTAUTOITRTE.

    "-Variables---------------------------------------------------------
      Data lo_AutoIt Type Ref To ZCL_AUTOIT.
      Data lv_WorkDir Type String.
      Data lv_Str Type String.

    "-Main--------------------------------------------------------------
      Try.

        "-Copy AutoIt3 runtime environemnt------------------------------
          Create Object lo_AutoIt.
          lo_AutoIt->ProvideRTE( i_DeleteExistRTE = ABAP_TRUE ).

        "-Get path of SAP work directory--------------------------------
          lv_WorkDir = lo_AutoIt->GETWORKDIR( ).

        "-Read AutoIt code from include and store it as file------------
          lo_AutoIt->StoreInclAsFile( i_InclName = 'ZINPUTBOXAU3'
            i_FileName = lv_WorkDir && '\InputBox.au3').

        "-Execute AutoIt with InputBox.au3------------------------------
          lo_AutoIt->Execute( i_FileName = 'InputBox.au3'
            i_WorkDir = lv_WorkDir ).

        "-Get result from clipboard-------------------------------------
          lv_Str = lo_AutoIt->GetClipBoard( ).

        Write: / lv_Str.

      Catch cx_root.
        Write: / 'An error occured'.
      EndTry.

"-End-------------------------------------------------------------------
