VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const BUFFERSIZE = 255

Private hOpen As Long
Private hConnection As Long
Private hFile As Long
Private dwType As Long
Private dwSeman As Long

Private szErrorMessage As String

Private mDirCol As New cDirList
Public Enum FTP_CONNECTION_STATES
    FTP_CONNECTION_RESOLVING_HOST
    FTP_CONNECTION_HOST_RESOLVED
    FTP_CONNECTION_CONNECTED
    FTP_CONNECTION_AUTHENTICATION
    FTP_USER_LOGGED
    FTP_ESTABLISHING_DATA_CONNECTION
    FTP_DATA_CONNECTION_ESTABLISHED
    FTP_RETRIEVING_DIRECTORY_INFO
    FTP_DIRECTORY_INFO_COMPLETED
    FTP_TRANSFER_STARTING
    FTP_TRANSFER_COMLETED
End Enum
'all possible reply codes that can be sent by ftp server
Private Enum FTP_RESPONSE_CODES
    FTP_RESPONSE_RESTART_MARKER_REPLY = 110
    FTP_RESPONSE_SERVICE_READY_IN_MINUTES = 120
    FTP_RESPONSE_DATA_CONNECTION_ALREADY_OPEN = 125
    FTP_RESPONSE_FILE_STATUS_OK = 150
    FTP_RESPONSE_COMMAND_OK = 200
    FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED_SUPERFLUOUS_AT_THIS_SITE = 202 'superfluous at this site
    FTP_RESPONSE_SYSTEM_STATUS_OR_SYSTEM_HELP_REPLY = 211
    FTP_RESPONSE_DIRECTORY_STATUS = 212
    FTP_RESPONSE_FILE_STATUS = 213
    FTP_RESPONSE_HELP_MESSAGE = 214
    FTP_RESPONSE_NAME_SYSTEM_TYPE = 215
    FTP_RESPONSE_SERVICE_READY_FOR_NEW_USER = 220
    FTP_RESPONSE_SERVICE_CLOSING_CONTROL_CONNECTION = 221
    FTP_RESPONSE_DATA_CONNECTION_OPEN = 225
    FTP_RESPONSE_CLOSING_DATA_CONNECTION = 226
    FTP_RESPONSE_ENTERING_PASSIVE_MODE = 227
    FTP_RESPONSE_USER_LOGGED_IN = 230
    FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED = 250
    FTP_RESPONSE_PATHNAME_CREATED = 257
    FTP_RESPONSE_USER_NAME_OK_NEED_PASSWORD = 331
    FTP_RESPONSE_NEED_ACCOUNT_FOR_LOGIN = 332
    FTP_RESPONSE_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO = 350
    FTP_RESPONSE_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION = 421
    FTP_RESPONSE_CANNOT_OPEN_DATA_CONNECTION = 425
    FTP_RESPONSE_CONNECTION_CLOSED_TRANSFER_ABORTED = 426
    FTP_RESPONSE_REQUESTED_FILE_ACTION_NOT_TAKEN = 450
    FTP_RESPONSE_REQUESTED_ACTION_ABORTED = 451
    FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN = 452
    FTP_RESPONSE_SYNTAX_ERROR_COMMAND_UNRECOGNIZED = 500
    FTP_RESPONSE_SYNTAX_ERROR_IN_PARAMETERS_OR_ARGUMENTS = 501
    FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED = 502
    FTP_RESPONSE_BAD_SEQUENCE_OF_COMMANDS = 503
    FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED_FOR_THAT_PARAMETER = 504
    FTP_RESPONSE_NOT_LOGGED_IN = 530
    FTP_RESPONSE_NEED_ACCOUNT_FOR_STORING_FILES = 532
    FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN_FILE_UNAVAILABLE = 550
    FTP_RESPONSE_REQUESTED_ACTION_ABORTED_PAGE_TYPE_UNKNOWN = 551
    FTP_RESPONSE_REQUESTED_FILE_ACTION_ABORTED_EXCEEDED_STORAGE_ALLOCATION = 552
    FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN_FILE_NAME_NOT_ALLOWED = 553
End Enum
Public Enum FtpErrors
    ERROR_FTP_WINSOCK_AddressInUse
    ERROR_FTP_WINSOCK_AddressNotAvailable
    ERROR_FTP_WINSOCK_AlreadyComplete
    ERROR_FTP_WINSOCK_AlreadyConnected
    ERROR_FTP_WINSOCK_BadState
    ERROR_FTP_WINSOCK_ConnectAborted
    ERROR_FTP_WINSOCK_ConnectionRefused
    ERROR_FTP_WINSOCK_ConnectionReset
    ERROR_FTP_WINSOCK_GetNotSupported
    ERROR_FTP_WINSOCK_HostNotFound
    ERROR_FTP_WINSOCK_HostNotFoundTryAgain
    ERROR_FTP_WINSOCK_InProgress
    ERROR_FTP_WINSOCK_InvalidArg
    ERROR_FTP_WINSOCK_InvalidArgument
    ERROR_FTP_WINSOCK_InvalidOp
    ERROR_FTP_WINSOCK_InvalidPropertyValue
    ERROR_FTP_WINSOCK_MsgTooBig
    ERROR_FTP_WINSOCK_NetReset
    ERROR_FTP_WINSOCK_NetworkSubsystemFailed
    ERROR_FTP_WINSOCK_NetworkUnreachable
    ERROR_FTP_WINSOCK_NoBufferSpace
    ERROR_FTP_WINSOCK_NoData
    ERROR_FTP_WINSOCK_NonRecoverableError
    ERROR_FTP_WINSOCK_NotConnected
    ERROR_FTP_WINSOCK_NotInitialized
    ERROR_FTP_WINSOCK_NotSocket
    ERROR_FTP_WINSOCK_OpCanceled
    ERROR_FTP_WINSOCK_OutOfMemory
    ERROR_FTP_WINSOCK_OutOfRange
    ERROR_FTP_WINSOCK_PortNotSupported
    ERROR_FTP_WINSOCK_SetNotSupported
    ERROR_FTP_WINSOCK_SocketShutdown
    ERROR_FTP_WINSOCK_Success
    ERROR_FTP_WINSOCK_Timedout
    ERROR_FTP_WINSOCK_Unsupported
    ERROR_FTP_WINSOCK_WouldBlock
    ERROR_FTP_WINSOCK_WrongProtocol
    ERROR_FTP_PROTOCOL_SERVICE_READY_IN_MINUTES
    ERROR_FTP_PROTOCOL_USER_NAME_OK_NEED_PASSWORD
    ERROR_FTP_PROTOCOL_NEED_ACCOUNT_FOR_LOGIN
    ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO
    ERROR_FTP_PROTOCOL_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION
    ERROR_FTP_PROTOCOL_CANNOT_OPEN_DATA_CONNECTION
    ERROR_FTP_PROTOCOL_CONNECTION_CLOSED_TRANSFER_ABORTED
    ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_NOT_TAKEN
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_ABORTED
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN
    ERROR_FTP_PROTOCOL_SYNTAX_ERROR_COMMAND_UNRECOGNIZED
    ERROR_FTP_PROTOCOL_SYNTAX_ERROR_IN_PARAMETERS_OR_ARGUMENTS
    ERROR_FTP_PROTOCOL_COMMAND_NOT_IMPLEMENTED
    ERROR_FTP_PROTOCOL_BAD_SEQUENCE_OF_COMMANDS
    ERROR_FTP_PROTOCOL_COMMAND_NOT_IMPLEMENTED_FOR_THAT_PARAMETER
    ERROR_FTP_PROTOCOL_NOT_LOGGED_IN
    ERROR_FTP_PROTOCOL_NEED_ACCOUNT_FOR_STORING_FILES
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN_FILE_UNAVAILABLE
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_ABORTED_PAGE_TYPE_UNKNOWN
    ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_ABORTED_EXCEEDED_STORAGE_ALLOCATION
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN_FILE_NAME_NOT_ALLOWED
    ERROR_FTP_USER_TIMEOUT
    ERROR_FTP_USER_TRANSFER_IN_PROGRESS
End Enum
Public Event StateChanged(State As FTP_CONNECTION_STATES)
Public Event FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)

Property Get Directory() As cDirList
   Set Directory = mDirCol
End Property

Private Sub ErrorOut(ByVal dwError As Long, ByRef szFunc As String)
    Dim dwRet As Long
    Dim dwTemp As Long
    Dim szString As String * 2048
    dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                      GetModuleHandle("wininet.dll"), dwError, 0, _
                      szString, 256, 0)
    szErrorMessage = szFunc & " error code: " & dwError & " Message: " & szString
    If (dwError = 12003) Then
        ' Extended error information was returned
        dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
        szErrorMessage = szString
    End If
End Sub

Property Get GetLastErrorMessage() As String
    GetLastErrorMessage = szErrorMessage
End Property

Public Sub CloseConnection()
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = 0
End Sub

Public Function SimpleFTPPutFile(sLocal As String, sRemote As String) As Boolean
    If (FTPPutFile(hConnection, sLocal, sRemote, dwType, 0) = False) Then
        ErrorOut Err.LastDllError, "SimpleFtpPutFile"
        SimpleFTPPutFile = False
        Exit Function
    Else
        SimpleFTPPutFile = True
    End If
End Function
 
Public Function RenameFTPFile(sExisting As String, sNewName As String) As Boolean
    If (FtpRenameFile(hConnection, sExisting, sNewName) = False) Then
        ErrorOut Err.LastDllError, "RenameFTPFile"
        RenameFTPFile = False
        Exit Function
    Else
        RenameFTPFile = True
    End If
End Function

Public Function CreateFTPDirectory(sDirectory As String) As Boolean
    If (FtpCreateDirectory(hConnection, sDirectory) = False) Then
        ErrorOut Err.LastDllError, "CreateFTPDirectory"
        CreateFTPDirectory = False
        Exit Function
    Else
        CreateFTPDirectory = True
    End If
End Function

Public Function RemoveFTPDirectory(sDirectory As String) As Boolean
    If (FtpRemoveDirectory(hConnection, sDirectory) = False) Then
        ErrorOut Err.LastDllError, "RemoveFTPDirectory"
        RemoveFTPDirectory = False
        Exit Function
    Else
        RemoveFTPDirectory = True
    End If
End Function
 
Public Function DeleteFTPFile(sRemote As String) As Boolean
    If (FtpDeleteFile(hConnection, sRemote) = False) Then
        ErrorOut Err.LastDllError, "DeleteFTPFile"
        DeleteFTPFile = False
        Exit Function
    Else
        DeleteFTPFile = True
    End If
End Function

Public Function OpenConnection(sServer As String, sPort As String, sUser As String, sPassword As String) As Boolean
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    
    hConnection = InternetConnect(hOpen, sServer, sPort, sUser, sPassword, INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then
        ErrorOut Err.LastDllError, "InternetConnect"
        OpenConnection = False
        Exit Function
    Else
        OpenConnection = True

    End If
End Function

Public Function FTPUploadFile(sLocal As String, sRemote As String) As Boolean
    Dim Data(BUFFERSIZE - 1) As Byte
    Dim Written As Long
    Dim Size As Long
    Dim Sum As Long
    Dim lBlock As Long
    
    Sum = 0
    lBlock = 0
    sLocal = Trim(sLocal)
    sRemote = Trim(sRemote)
    
    If sLocal <> "" And sRemote <> "" Then
      hFile = FtpOpenFile(hConnection, sRemote, GENERIC_WRITE, dwType, 0)
      If hFile = 0 Then
          ErrorOut Err.LastDllError, "FtpOpenFile:PutFile"
          FTPUploadFile = False
          Exit Function
      End If
      
      Open sLocal For Binary Access Read As #1
      Size = LOF(1)
      For lBlock = 1 To Size \ BUFFERSIZE
          Get #1, , Data
          If (InternetWriteFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
              FTPUploadFile = False
              ErrorOut Err.LastDllError, "InternetWriteFile"
              Exit Function
          End If
          DoEvents
          Sum = Sum + BUFFERSIZE
          RaiseEvent FileTransferProgress(Sum, Size)
      Next lBlock
      
      Get #1, , Data
      If (InternetWriteFile(hFile, Data(0), Size Mod BUFFERSIZE, Written) = 0) Then
          FTPUploadFile = False
          ErrorOut Err.LastDllError, "InternetWriteFile2"
          Exit Function
      End If
      
      Sum = Sum + (Size Mod BUFFERSIZE)
      Size = Sum
      RaiseEvent FileTransferProgress(Sum, Size)
      Close #1
      InternetCloseHandle (hFile)
      FTPUploadFile = True
   End If
End Function

Public Function FTPDownloadFile(sLocal As String, sRemote As String) As Boolean
    Dim Data(BUFFERSIZE - 1) As Byte ' array of 100 elements 0 to 99
    Dim Written As Long
    Dim Size As Long
    Dim Sum As Long
    Dim lBlock As Long
    DoEvents
    DoEvents
    FTPDownloadFile = False
    DoEvents
    DoEvents
           RaiseEvent StateChanged(FTP_TRANSFER_STARTING)
           DoEvents
           DoEvents
    Sum = 0
    lBlock = 0
    DoEvents
    DoEvents
    sLocal = sLocal
    sRemote = sRemote
    DoEvents
    DoEvents
    If sLocal <> "" And sRemote <> "" Then
      Size = frmmain.ListView3.SelectedItem.SubItems(4)
      DoEvents
      DoEvents
      If Size >= 0 Then
      DoEvents
          hFile = FtpOpenFile(hConnection, sRemote, GENERIC_READ, dwType, 0)
                DoEvents
      DoEvents
          If hFile = 0 Then
              ErrorOut Err.LastDllError, "FtpOpenFile:GetFile"
              Exit Function
          End If
          DoEvents
          Open sLocal For Binary Access Write As #1
                DoEvents
      DoEvents
          Seek #1, 1
          Sum = 1
                DoEvents
      DoEvents
          For lBlock = 1 To Size \ BUFFERSIZE
              If (InternetReadFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
                  ErrorOut Err.LastDllError, "InternetReadFile"
                  Close #1
                  Exit Function
              End If
                    DoEvents
      DoEvents
              Put #1, , Data
              DoEvents
              Sum = Sum + BUFFERSIZE
              DoEvents
              RaiseEvent FileTransferProgress(Sum, Size)
                DoEvents
      DoEvents
          Next lBlock
          
          ReDim Data2((Size Mod BUFFERSIZE) - 1) As Byte
                DoEvents
      DoEvents
          If (InternetReadFile(hFile, Data2(0), Size Mod BUFFERSIZE, Written) = 0) Then
              ErrorOut Err.LastDllError, "InternetReadFile2"
              Close #1
                    DoEvents
      DoEvents
              Exit Function
          End If
                      
          Put #1, , Data2
                DoEvents
      DoEvents
          Sum = Sum + (Size Mod BUFFERSIZE)
          Size = Sum
          RaiseEvent FileTransferProgress(Sum, Size)
                DoEvents
      DoEvents
          Close #1
          DoEvents
          DoEvents
          InternetCloseHandle (hFile)
          FTPDownloadFile = True
          DoEvents
          DoEvents
          RaiseEvent StateChanged(FTP_TRANSFER_COMLETED)
      End If
   End If
End Function

Public Function SimpleFTPGetFile(sLocal As String, sRemote As String) As Boolean
   ' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
    If (FTPGetFile(hConnection, sRemote, sLocal, False, FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
        ErrorOut Err.LastDllError, "SimpleFtpGetFile"
        SimpleFTPGetFile = False
        Exit Function
    Else
        SimpleFTPGetFile = True
    End If
End Function

Public Function GetFTPDirectory() As String
    Dim szDir As String
    szDir = String(1024, Chr$(0))
    If (FtpGetCurrentDirectory(hConnection, szDir, 1024) = False) Then
        ErrorOut Err.LastDllError, "FtpGetCurrentDirectory"
        Exit Function
    Else
        GetFTPDirectory = Left(szDir, InStr(1, szDir, String(1, 0), vbBinaryCompare) - 1)
    End If
End Function

Public Function SetFTPDirectory(sDir As String)
    If (FtpSetCurrentDirectory(hConnection, sDir) = False) Then
        ErrorOut Err.LastDllError, "FtpSetCurrentDirectory"
        SetFTPDirectory = False
        Exit Function
    Else
        SetFTPDirectory = True
    End If
End Function

Public Function GetFTPFileSize(sFile As String) As Long
    Dim szDir As String
    Dim hFind As Long
    Dim nLastError As Long
    Dim pData As WIN32_FIND_DATA
    
    hFind = FtpFindFirstFile(hConnection, sFile, pData, 0, 0)
    nLastError = Err.LastDllError
    If hFind = 0 Then
        If (nLastError = ERROR_NO_MORE_FILES) Then
            GetFTPFileSize = -1  ' File not found
        Else
            GetFTPFileSize = -2  ' Other error
            ErrorOut Err.LastDllError, "FtpFindFirstFile"
        End If
    End If
           
    GetFTPFileSize = pData.nFileSizeLow
    InternetCloseHandle (hFind)
End Function

Public Function GetDirectoryListing(sFilter As String) As cDirList
    Dim szDir As String
    Dim hFind As Long
    Dim nLastError As Long
    Dim dError As Long
    Dim ptr As Long
    Dim pData As WIN32_FIND_DATA
    Dim sFilename As String
        
    Set mDirCol = Nothing
    DoEvents
    hFind = FtpFindFirstFile(hConnection, sFilter, pData, 0, 0)
    DoEvents
    nLastError = Err.LastDllError
    DoEvents
    If hFind = 0 Then
        If (nLastError <> ERROR_NO_MORE_FILES) Then
            ErrorOut Err.LastDllError, "FtpFindFirstFile"
        End If
        Exit Function
    End If
    
    dError = NO_ERROR
    Dim bRet As Boolean

    sFilename = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    mDirCol.Add pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, sFilename
    Do
        DoEvents
        pData.cFileName = String(MAX_PATH, 0)
        bRet = InternetFindNextFile(hFind, pData)
        DoEvents
        If Not bRet Then
            dError = Err.LastDllError
            If dError = ERROR_NO_MORE_FILES Then
                Exit Do
            Else
                ErrorOut Err.LastDllError, "InternetFindNextFile"
                InternetCloseHandle (hFind)
                Exit Function
            End If
            DoEvents
        Else
            sFilename = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            mDirCol.Add pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, sFilename
        End If
        DoEvents
    Loop
       DoEvents
    Set GetDirectoryListing = mDirCol
    DoEvents
    InternetCloseHandle (hFind)
End Function

Public Sub SetTransferASCII()
    dwType = FTP_TRANSFER_TYPE_ASCII
          DoEvents
      DoEvents
End Sub

Public Sub SetTransferBinary()
    dwType = FTP_TRANSFER_TYPE_BINARY
          DoEvents
      DoEvents
End Sub

Public Sub SetModeActive()
    dwSeman = 0
          DoEvents
      DoEvents
End Sub

Public Sub SetModePassive()
    dwSeman = INTERNET_FLAG_PASSIVE
          DoEvents
      DoEvents
End Sub

Private Sub Class_Initialize()
    hOpen = InternetOpen("SMC FTP Client", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        ErrorOut Err.LastDllError, "InternetOpen"
    End If
    dwType = FTP_TRANSFER_TYPE_ASCII
    dwSeman = 0
    hConnection = 0
End Sub

Private Sub Class_Terminate()
    InternetCloseHandle hOpen
End Sub
