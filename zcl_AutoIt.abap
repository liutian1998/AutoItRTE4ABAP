
"-Begin-----------------------------------------------------------------

Class ZCL_AUTOIT Definition Public.

  Public Section.

    Methods ProvideRTE
      Importing i_DeleteExistRTE Type ABAP_BOOL
      Exceptions Error.

    Methods DeleteExistingRTE
      Exceptions Error.

    Methods GetWorkDir
      Returning Value(r_WorkDir) Type String
      Exceptions Error.

    Methods ReadInclAsString
      Importing i_InclName Type SOBJ_NAME
      Returning Value(r_strIncl) Type String
      Exceptions Error.

    Methods StoreInclAsFile
      Importing i_InclName Type SOBJ_NAME i_FileName Type String
      Returning Value(r_FileLength) Type i
      Exceptions Error.

    Methods Execute
      Importing i_FileName Type String i_WorkDir Type String
      Exceptions Error.

    Methods GetClipBoard
      Returning Value(r_ClipData) Type String
      Exceptions Error.

    Methods PutClipBoard
      Importing i_ClipData Type String
      Returning Value(r_Success) Type i
      Exceptions Error.

    Methods ReadFile
      Importing i_FileName Type String
      Returning Value(r_FileData) Type String
      Exceptions Error.

    Methods WriteFile
      Importing i_FileName Type String i_FileData Type String
      Returning Value(r_FileLength) Type i
      Exceptions Error.

    Methods AppendFile
      Importing i_FileName Type String i_FileData Type String
      Returning Value(r_FileLength) Type i
      Exceptions Error.

    Methods DeleteFile
      Importing i_FileName Type String
      Returning Value(r_Success) Type i
      Exceptions Error.

    Methods Flush
      Exceptions Error.

EndClass.


Class ZCL_AUTOIT Implementation.

  Method ProvideRTE."---------------------------------------------------

    "-Variables---------------------------------------------------------
      Data lv_WorkDir Type String.
      Data lv_InclCode Type String.
      Data lt_FileData Type Table Of String.
      Data lv_RC Type i.
      Data lv_Result Type ABAP_BOOL.
      Data lo_IPC Type Ole2_Object.

    lv_WorkDir = Me->GetWorkDir( ).

    If i_DeleteExistRTE = ABAP_TRUE.
      Me->DeleteExistingRTE( ).
    EndIf.

    "-Provide AutoIt runtime environment--------------------------------
      Call Function 'ZAUTOIT3EXE'.
      Call Function 'ZUNRARDLL'.
      Call Function 'ZINCLUDERAR'.

    "-Read AutoIt code from include file--------------------------------
      Me->StoreInclAsFile( i_InclName = 'ZINCLUDEAU3' 
        i_FileName = lv_WorkDir && '\Include.au3' ).

    "-Execute AutoIt to unpack Include directory------------------------
      Call Method cl_gui_frontend_services=>execute
        Exporting
          APPLICATION = lv_WorkDir && '\AutoIt3.exe'
          PARAMETER = 'Include.au3'
          DEFAULT_DIRECTORY = lv_WorkDir
          SYNCHRONOUS = 'X'
        Exceptions
          CNTL_ERROR = 1
          ERROR_NO_GUI = 2
          BAD_PARAMETER = 3
          FILE_NOT_FOUND = 4
          PATH_NOT_FOUND = 5
          FILE_EXTENSION_UNKNOWN = 6
          ERROR_EXECUTE_FAILED  = 7
          SYNCHRONOUS_FAILED = 8
          NOT_SUPPORTED_BY_GUI = 9
          Others = 10.

      If sy-subrc <> 0.
        Raise Error.
      EndIf.

    "-Delete Include.rar archive----------------------------------------
      Call Method cl_gui_frontend_services=>file_delete
        Exporting
          FILENAME = lv_WorkDir && '\Include.rar'
        Changing
          RC = lv_RC
        Exceptions
          FILE_DELETE_FAILED = 1
          CNTL_ERROR = 2
          ERROR_NO_GUI = 3
          FILE_NOT_FOUND = 4
          ACCESS_DENIED = 5
          UNKNOWN_ERROR = 6
          NOT_SUPPORTED_BY_GUI = 7
          WRONG_PARAMETER = 8
          Others = 9.

      If sy-subrc <> 0.
        Raise Error.
      EndIf.

    "-Delete AutoIt code------------------------------------------------
      Call Method cl_gui_frontend_services=>file_delete
        Exporting
          FILENAME = lv_WorkDir && '\Include.au3'
        Changing
          RC = lv_RC
        Exceptions
          FILE_DELETE_FAILED = 1
          CNTL_ERROR = 2
          ERROR_NO_GUI = 3
          FILE_NOT_FOUND = 4
          ACCESS_DENIED = 5
          UNKNOWN_ERROR = 6
          NOT_SUPPORTED_BY_GUI = 7
          WRONG_PARAMETER = 8
          Others = 9.

      If sy-subrc <> 0.
        Raise Error.
      EndIf.

    "-Delete unrar.dll library------------------------------------------
      Call Method cl_gui_frontend_services=>file_delete
        Exporting
          FILENAME = lv_WorkDir && '\unrar.dll'
        Changing
          RC = lv_RC
        Exceptions
          FILE_DELETE_FAILED = 1
          CNTL_ERROR = 2
          ERROR_NO_GUI = 3
          FILE_NOT_FOUND = 4
          ACCESS_DENIED = 5
          UNKNOWN_ERROR = 6
          NOT_SUPPORTED_BY_GUI = 7
          WRONG_PARAMETER = 8
          Others = 9.

      If sy-subrc <> 0.
        Raise Error.
      EndIf.

  EndMethod.


  Method DeleteExistingRTE."--------------------------------------------

    "-Variables---------------------------------------------------------
      Data lv_Result Type ABAP_BOOL.
      Data lt_FileData Type Table Of String.
      Data lv_AutoItCmd Type String.
      Data lv_RC Type i.
      Data lv_WorkDir Type String.

    lv_WorkDir = Me->GetWorkDir( ).

    "-Delete AutoIt3.exe if exists--------------------------------------
      Call Method cl_gui_frontend_services=>file_exist
        Exporting
          FILE = lv_WorkDir && '\AutoIt3.exe'
        Receiving
          RESULT = lv_Result
        Exceptions
          CNTL_ERROR = 1
          ERROR_NO_GUI = 2
          WRONG_PARAMETER = 3
          NOT_SUPPORTED_BY_GUI = 4
          Others = 5.

      If sy-subrc <> 0.
        Raise Error.
      EndIf.

      If lv_Result = ABAP_TRUE.

        "-Delete Include directory if exists----------------------------
          Call Method cl_gui_frontend_services=>directory_exist
            Exporting
              DIRECTORY = lv_WorkDir && '\Include'
            Receiving
              RESULT = lv_Result
            Exceptions
              CNTL_ERROR = 1
              ERROR_NO_GUI = 2
              WRONG_PARAMETER = 3
              NOT_SUPPORTED_BY_GUI = 4
              Others = 5.

          If sy-subrc <> 0.
            Raise Error.
          EndIf.

          If lv_Result = ABAP_TRUE.

            Clear lt_FileData.
            lv_AutoItCmd = 'DirRemove("' && lv_WorkDir &&
              '\Include", 1)'.
            Append lv_AutoItCmd To lt_FileData.

            "-Store AutoIt code-----------------------------------------
              Call Method cl_gui_frontend_services=>gui_download
                Exporting
                  FILENAME = lv_WorkDir && '\DelInclude.au3'
                Changing
                  DATA_TAB = lt_FileData
                Exceptions
                  FILE_WRITE_ERROR = 1
                  NO_BATCH = 2
                  GUI_REFUSE_FILETRANSFER = 3
                  INVALID_TYPE = 4
                  NO_AUTHORITY = 5
                  UNKNOWN_ERROR = 6
                  HEADER_NOT_ALLOWED = 7
                  SEPARATOR_NOT_ALLOWED = 8
                  FILESIZE_NOT_ALLOWED = 9
                  HEADER_TOO_LONG = 10
                  DP_ERROR_CREATE = 11
                  DP_ERROR_SEND = 12
                  DP_ERROR_WRITE = 13
                  UNKNOWN_DP_ERROR = 14
                  ACCESS_DENIED = 15
                  DP_OUT_OF_MEMORY = 16
                  DISK_FULL = 17
                  DP_TIMEOUT = 18
                  FILE_NOT_FOUND = 19
                  DATAPROVIDER_EXCEPTION = 20
                  CONTROL_FLUSH_ERROR = 21
                  NOT_SUPPORTED_BY_GUI = 22
                  ERROR_NO_GUI = 23
                  Others = 24.

              If sy-subrc <> 0.
                Raise Error.
              EndIf.

            "-Execute AutoIt to delete the Include directory------------
              Call Method cl_gui_frontend_services=>execute
                Exporting
                  APPLICATION = lv_WorkDir && '\AutoIt3.exe'
                  PARAMETER = 'DelInclude.au3'
                  DEFAULT_DIRECTORY = lv_WorkDir
                  SYNCHRONOUS = 'X'
                Exceptions
                  CNTL_ERROR = 1
                  ERROR_NO_GUI = 2
                  BAD_PARAMETER = 3
                  FILE_NOT_FOUND = 4
                  PATH_NOT_FOUND = 5
                  FILE_EXTENSION_UNKNOWN = 6
                  ERROR_EXECUTE_FAILED  = 7
                  SYNCHRONOUS_FAILED = 8
                  NOT_SUPPORTED_BY_GUI = 9
                  Others = 10.

              If sy-subrc <> 0.
                Raise Error.
              EndIf.

            "-Delete AutoIt code----------------------------------------
              Call Method cl_gui_frontend_services=>file_delete
                Exporting
                  FILENAME = lv_WorkDir && '\DelInclude.au3'
                Changing
                  RC = lv_RC
                Exceptions
                  FILE_DELETE_FAILED = 1
                  CNTL_ERROR = 2
                  ERROR_NO_GUI = 3
                  FILE_NOT_FOUND = 4
                  ACCESS_DENIED = 5
                  UNKNOWN_ERROR = 6
                  NOT_SUPPORTED_BY_GUI = 7
                  WRONG_PARAMETER = 8
                  Others = 9.

              If sy-subrc <> 0.
                Raise Error.
              EndIf.

          EndIf.

        "-Delete AutoIt3.exe executable---------------------------------
          Call Method cl_gui_frontend_services=>file_delete
            Exporting
              FILENAME = lv_WorkDir && '\AutoIt3.exe'
            Changing
              RC = lv_RC
            Exceptions
              FILE_DELETE_FAILED = 1
              CNTL_ERROR = 2
              ERROR_NO_GUI = 3
              FILE_NOT_FOUND = 4
              ACCESS_DENIED = 5
              UNKNOWN_ERROR = 6
              NOT_SUPPORTED_BY_GUI = 7
              WRONG_PARAMETER = 8
              Others = 9.

          If sy-subrc <> 0.
            Raise Error.
          EndIf.

      EndIf.

  EndMethod.


  Method GetWorkDir."---------------------------------------------------

    "-Variables---------------------------------------------------------
      Data lv_WorkDir Type String.
      Data lo_SAPGUI Type OBJ_RECORD.
      Data lv_UserName Type String.

    Call Method cl_gui_frontend_services=>get_sapgui_workdir
      Changing
        SAPWORKDIR = lv_WorkDir
      Exceptions
        GET_SAPWORKDIR_FAILED = 1
        CNTL_ERROR = 2
        ERROR_NO_GUI = 3
        NOT_SUPPORTED_BY_GUI  = 4
        Others = 5.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

    If lv_WorkDir Is Initial.
      Create Object lo_SAPGUI 'Sapgui.InfoCtrl.1'.
      If sy-subrc = 0 And lo_SAPGUI-HANDLE > 0 And lo_SAPGUI-TYPE = 'OLE2'.
        Get Property Of lo_SAPGUI 'GetUserName' = lv_UserName.
        Free Object lo_SAPGUI.
      EndIf.
      lv_WorkDir = 'C:\Users\' && lv_UserName && '\Documents\SAP\SAP GUI'.
    EndIf.

    r_WorkDir = lv_WorkDir.

  EndMethod.


  Method ReadInclAsstring."---------------------------------------------

    "-Variables---------------------------------------------------------
      Data lt_TADIR Type TADIR.
      Data lt_Incl Type Table Of String.
      Data lv_InclLine Type String.
      Data lv_retIncl Type String.

    Select Single * From TADIR Into lt_TADIR 
      Where OBJ_NAME = i_InclName.
    If sy-subrc = 0.
      Read Report i_InclName Into lt_Incl.
      If sy-subrc = 0.
        Loop At lt_Incl Into lv_InclLine.
          lv_retIncl = lv_retIncl && lv_InclLine &&
            cl_abap_char_utilities=>cr_lf.
          lv_InclLine = ''.
        EndLoop.
      EndIf.
    Else.
      Raise Error.
    EndIf.
    r_strIncl = lv_retIncl.

  EndMethod.


  Method StoreInclAsFile."----------------------------------------------

    "-Variables---------------------------------------------------------
      Data lv_InclCode Type String.
      Data lt_FileData Type Table Of String.

    lv_InclCode = Me->ReadInclAsString( i_InclName ).

    Split lv_InclCode At cl_abap_char_utilities=>cr_lf
      Into Table lt_FileData.

    Call Method cl_gui_frontend_services=>gui_download
      Exporting
        FILENAME = i_FileName
      Importing
        FILELENGTH = r_FileLength
      Changing
        DATA_TAB = lt_FileData
      Exceptions
        FILE_WRITE_ERROR = 1
        NO_BATCH = 2
        GUI_REFUSE_FILETRANSFER = 3
        INVALID_TYPE = 4
        NO_AUTHORITY = 5
        UNKNOWN_ERROR = 6
        HEADER_NOT_ALLOWED = 7
        SEPARATOR_NOT_ALLOWED = 8
        FILESIZE_NOT_ALLOWED = 9
        HEADER_TOO_LONG = 10
        DP_ERROR_CREATE = 11
        DP_ERROR_SEND = 12
        DP_ERROR_WRITE = 13
        UNKNOWN_DP_ERROR = 14
        ACCESS_DENIED = 15
        DP_OUT_OF_MEMORY = 16
        DISK_FULL = 17
        DP_TIMEOUT = 18
        FILE_NOT_FOUND = 19
        DATAPROVIDER_EXCEPTION = 20
        CONTROL_FLUSH_ERROR = 21
        NOT_SUPPORTED_BY_GUI = 22
        ERROR_NO_GUI = 23
        Others = 24.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

  EndMethod.


  Method Execute."------------------------------------------------------

    Call Method cl_gui_frontend_services=>execute
      Exporting
        APPLICATION = i_WorkDir && '\AutoIt3.exe'
        PARAMETER = i_FileName
        DEFAULT_DIRECTORY = i_WorkDir
        SYNCHRONOUS = 'X'
      Exceptions
        CNTL_ERROR = 1
        ERROR_NO_GUI = 2
        BAD_PARAMETER = 3
        FILE_NOT_FOUND = 4
        PATH_NOT_FOUND = 5
        FILE_EXTENSION_UNKNOWN = 6
        ERROR_EXECUTE_FAILED  = 7
        SYNCHRONOUS_FAILED = 8
        NOT_SUPPORTED_BY_GUI = 9
        Others = 10.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

  EndMethod.


  Method GetClipBoard."-------------------------------------------------

    "-Type--------------------------------------------------------------
      Types t_ClipData Type c Length 8192.

    "-Variables---------------------------------------------------------
      Data lt_ClipData Type Standard Table Of t_ClipData.
      Data lv_Lines Type i.
      Data lv_Str Type String.
      Data lv_Ret Type String.

    Call Method cl_gui_frontend_services=>clipboard_import
      Importing
        DATA = lt_ClipData
      Exceptions
        CNTL_ERROR = 1
        ERROR_NO_GUI = 2
        NOT_SUPPORTED_BY_GUI = 3
        Others = 4.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

    lv_Lines = Lines( lt_ClipData ).
    Loop At lt_ClipData Into lv_Str.
      If sy-tabix = lv_Lines.
        lv_Ret = lv_Ret && lv_Str.
      Else.
        lv_Ret = lv_Ret && lv_Str &&  cl_abap_char_utilities=>cr_lf.
      EndIf.
    EndLoop.
    r_ClipData = lv_Ret.

  EndMethod.


  Method PutClipBoard."-------------------------------------------------

    "-Type--------------------------------------------------------------
      Types t_ClipData Type c Length 8192.

    "-Variables---------------------------------------------------------
      Data lt_ClipData Type Standard Table Of t_ClipData.

    Split i_ClipData At cl_abap_char_utilities=>cr_lf
      Into Table lt_ClipData.

    Call Method cl_gui_frontend_services=>clipboard_export
      Importing
        DATA = lt_ClipData
      Changing
        RC = r_Success
      Exceptions
        CNTL_ERROR = 1
        ERROR_NO_GUI = 2
        NOT_SUPPORTED_BY_GUI = 3
        NO_AUTHORITY = 4
        Others = 5.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

  EndMethod.


  Method ReadFile."-----------------------------------------------------

    "-Variables---------------------------------------------------------
      Data lt_FileData Type Table Of String.
      Data lv_Lines Type i.
      Data lv_Str Type String.
      Data lv_Ret Type String.

    Call Method cl_gui_frontend_services=>gui_upload
      Exporting
        FILENAME = i_FileName
      Changing
        DATA_TAB = lt_FileData
      Exceptions
        FILE_OPEN_ERROR = 1
        FILE_READ_ERROR = 2
        NO_BATCH = 3
        GUI_REFUSE_FILETRANSFER = 4
        INVALID_TYPE = 5
        NO_AUTHORITY = 6
        UNKNOWN_ERROR = 7
        BAD_DATA_FORMAT = 8
        HEADER_NOT_ALLOWED = 9
        SEPARATOR_NOT_ALLOWED = 10
        HEADER_TOO_LONG = 11
        UNKNOWN_DP_ERROR = 12
        ACCESS_DENIED = 13
        DP_OUT_OF_MEMORY = 14
        DISK_FULL = 15
        DP_TIMEOUT = 16
        NOT_SUPPORTED_BY_GUI = 17
        ERROR_NO_GUI = 18
        Others = 19.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

    lv_Lines = Lines( lt_FileData ).
    Loop At lt_FileData Into lv_Str.
      If sy-tabix = lv_Lines.
        lv_Ret = lv_Ret && lv_Str.
      Else.
        lv_Ret = lv_Ret && lv_Str &&  cl_abap_char_utilities=>cr_lf.
      EndIf.
    EndLoop.
    r_FileData = lv_Ret.

  EndMethod.


  Method WriteFile."----------------------------------------------------

    "-Variables---------------------------------------------------------
      Data lt_FileData Type Table Of String.

    Split i_FileData At cl_abap_char_utilities=>cr_lf
      Into Table lt_FileData.

    Call Method cl_gui_frontend_services=>gui_download
      Exporting
        FILENAME = i_FileName
      Importing
        FILELENGTH = r_FileLength
      Changing
        DATA_TAB = lt_FileData
      Exceptions
        FILE_WRITE_ERROR = 1
        NO_BATCH = 2
        GUI_REFUSE_FILETRANSFER = 3
        INVALID_TYPE = 4
        NO_AUTHORITY = 5
        UNKNOWN_ERROR = 6
        HEADER_NOT_ALLOWED = 7
        SEPARATOR_NOT_ALLOWED = 8
        FILESIZE_NOT_ALLOWED = 9
        HEADER_TOO_LONG = 10
        DP_ERROR_CREATE = 11
        DP_ERROR_SEND = 12
        DP_ERROR_WRITE = 13
        UNKNOWN_DP_ERROR = 14
        ACCESS_DENIED = 15
        DP_OUT_OF_MEMORY = 16
        DISK_FULL = 17
        DP_TIMEOUT = 18
        FILE_NOT_FOUND = 19
        DATAPROVIDER_EXCEPTION = 20
        CONTROL_FLUSH_ERROR = 21
        NOT_SUPPORTED_BY_GUI = 22
        ERROR_NO_GUI = 23
        Others = 24.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

  EndMethod.


  Method AppendFile."---------------------------------------------------

    "-Variables---------------------------------------------------------
      Data lt_FileData Type Table Of String.

    Split i_FileData At cl_abap_char_utilities=>cr_lf
      Into Table lt_FileData.

    Call Method cl_gui_frontend_services=>gui_download
      Exporting
        FILENAME = i_FileName
        APPEND = ABAP_TRUE
      Importing
        FILELENGTH = r_FileLength
      Changing
        DATA_TAB = lt_FileData
      Exceptions
        FILE_WRITE_ERROR = 1
        NO_BATCH = 2
        GUI_REFUSE_FILETRANSFER = 3
        INVALID_TYPE = 4
        NO_AUTHORITY = 5
        UNKNOWN_ERROR = 6
        HEADER_NOT_ALLOWED = 7
        SEPARATOR_NOT_ALLOWED = 8
        FILESIZE_NOT_ALLOWED = 9
        HEADER_TOO_LONG = 10
        DP_ERROR_CREATE = 11
        DP_ERROR_SEND = 12
        DP_ERROR_WRITE = 13
        UNKNOWN_DP_ERROR = 14
        ACCESS_DENIED = 15
        DP_OUT_OF_MEMORY = 16
        DISK_FULL = 17
        DP_TIMEOUT = 18
        FILE_NOT_FOUND = 19
        DATAPROVIDER_EXCEPTION = 20
        CONTROL_FLUSH_ERROR = 21
        NOT_SUPPORTED_BY_GUI = 22
        ERROR_NO_GUI = 23
        Others = 24.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

  EndMethod.


  Method DeleteFile."---------------------------------------------------

    Call Method cl_gui_frontend_services=>file_delete
      Exporting
        FILENAME = i_FileName
      Changing
        RC = r_Success
      Exceptions
        FILE_DELETE_FAILED = 1
        CNTL_ERROR = 2
        ERROR_NO_GUI = 3
        FILE_NOT_FOUND = 4
        ACCESS_DENIED = 5
        UNKNOWN_ERROR = 6
        NOT_SUPPORTED_BY_GUI = 7
        WRONG_PARAMETER = 8
        Others = 9.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

  EndMethod.


  Method Flush."--------------------------------------------------------

    Call Method CL_GUI_CFW=>Flush
      EXCEPTIONS
        CNTL_SYSTEM_ERROR = 1
        CNTL_ERROR = 2
        Others = 3.

    If sy-subrc <> 0.
      Raise Error.
    EndIf.

  EndMethod.

EndClass.

"-End-------------------------------------------------------------------
