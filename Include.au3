
;-Begin-----------------------------------------------------------------

  ;-Directives----------------------------------------------------------
    AutoItSetOption("MustDeclareVars", 1)

  ;-Constants-----------------------------------------------------------
    Const $RAR_OM_EXTRACT = 1
    Const $ERAR_SUCCESS = 0
    Const $RAR_EXTRACT = 2

  ;-Sub UnPackInclude---------------------------------------------------
    Func UnPackInclude($FileName)

      ;-Variables-------------------------------------------------------
        Local $hUnRAR, $RAROpenArchiveDataEx, $RARHeaderDataEx
        Local $RARFileName, $hArc, $Ret, $Test

      $hUnRAR = DLLOpen("unrar.dll")
      If $hUnRAR <> -1 Then

        $RAROpenArchiveDataEx = DllStructCreate("Struct;" & _
          "Ptr ArcName;" & _
          "Ptr ArcNameW;" & _
          "uInt OpenMode;" & _
          "uInt OpenResult;" & _
          "Ptr CmtBuf;" & _
          "uInt CmtBufSize;" & _
          "uInt CmtSize;" & _
          "uInt CmtState;" & _
          "uInt Flags;" & _
          "Ptr Callback;" & _
          "lParam UserData;" & _
          "uInt Reserved[28];" & _
          "EndStruct")

        $RARHeaderDataEx = DllStructCreate("Struct;" & _
          "Char ArcName[1024];" & _
          "wChar ArcNameW[1024];" & _
          "Char FileName[1024];" & _
          "wChar FileNameW[1024];" & _
          "uInt Flags;" & _
          "uInt PackSize;" & _
          "uInt PackSizeHigh;" & _
          "uInt UnpSize;" & _
          "uInt UnpSizeHigh;" & _
          "uInt HostOS;" & _
          "uInt FileCRC;" & _
          "uInt FileTime;" & _
          "uInt UnpVer;" & _
          "uInt Method;" & _
          "uInt FileAttr;" & _
          "Ptr CmtBuf;" & _
          "uInt CmtBufSize;" & _
          "uInt CmtSize;" & _
          "uInt CmtState;" & _
          "uInt DictSize;" & _
          "uInt HashType;" & _
          "Char Hash[32];" & _
          "uInt RedirType;" & _
          "Ptr RedirName;" & _
          "uInt RedirNameSize;" & _
          "uInt DirTarget;" & _
          "uInt Reserved[994];" & _
          "EndStruct")

        $RARFileName = DllStructCreate("Char[65535]")
        DllStructSetData($RARFileName, 1, _
          @ScriptDir & "\" & $FileName & Chr(0))

        DllStructSetData($RAROpenArchiveDataEx, "ArcName", _
          DllStructGetPtr($RARFileName))
        DllStructSetData($RAROpenArchiveDataEx, "OpenMode", _
          $RAR_OM_EXTRACT)

        $hArc = DllCall($hUnRAR, "Int", "RAROpenArchiveEx", _
          "Ptr", DllStructGetPtr($RAROpenArchiveDataEx))

        If $hArc[0] <> 0 And DllStructGetData($RAROpenArchiveDataEx, _
          "OpenResult") = $ERAR_SUCCESS Then

          Do
            DllCall($hUnRAR, "Int", "RARProcessFileW", "Handle", _
              $hArc[0], "Int", $RAR_EXTRACT, "Ptr", 0, "Ptr", 0)
            $Ret = DllCall($hUnRAR, "Int", "RARReadHeaderEx", "Handle", _
              $hArc[0], "Ptr", DllStructGetPtr($RARHeaderDataEx))
          Until $Ret[0] <> $ERAR_SUCCESS

          DllCall($hUnRAR, "Int", "RARCloseArchive", "Handle", $hArc[0])
        EndIf

        DLLClose($hUnRAR)
      EndIf

    EndFunc

  ;-Sub Main------------------------------------------------------------
    Func Main()
      UnPackInclude("Include.rar")
    EndFunc

  ;-Main----------------------------------------------------------------
    Main()

;-End-------------------------------------------------------------------
