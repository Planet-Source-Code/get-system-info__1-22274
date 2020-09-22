VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If Win16 Then
  Private Declare Function ExitWindows Lib "User" (ByVal dwReturnCode _
    As Long, ByVal uReserved As Integer) As Integer
  Private Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, _
    ByVal nSize As Integer) As Integer
  Private Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, _
    ByVal nSize As Integer) As Integer
        
  Const EW_REBOOTSYSTEM = &H43
  Const EW_RESTARTWINDOWS = &H42
  Const EW_EXITWINDOWS = 0
  
#Else

  Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, _
    ByVal dwReserved As Long) As Boolean
  Private Declare Function GetWindowsDirectory32 Lib "kernel32" Alias _
    "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) _
    As Long
  Private Declare Function GetSystemDirectory32 Lib "kernel32" Alias _
    "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) _
    As Long
  
  Const EWX_LOGOFF = 0
  Const EWX_SHUTDOWN = 1
  Const EWX_REBOOT = 2
  Const EWX_FORCE = 4
#End If

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTickCount& Lib "kernel32" ()

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
  (ByVal sDrive As String) As Long

'- Drive Type Constants...
Const DRIVE_CDROM = 5
Const DRIVE_FIXED = 3
Const DRIVE_RAMDISK = 6
Const DRIVE_REMOTE = 4
Const DRIVE_REMOVABLE = 2

' These represent the possible returns errors from API.
Const ERROR_BAD_DEVICE           As Long = 1200
Const ERROR_CONNECTION_UNAVAIL   As Long = 1201
Const ERROR_EXTENDED_ERROR       As Long = 1208
Const ERROR_MORE_DATA            As Long = 234
Const ERROR_NOT_SUPPORTED        As Long = 50
Const ERROR_NO_NET_OR_BAD_PATH   As Long = 1203
Const ERROR_NO_NETWORK           As Long = 1222
Const ERROR_NOT_CONNECTED        As Long = 2250
Const NO_ERROR                   As Long = 0

' This API declaration is used to return the    ' UNC path from a Drive letter.
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" _
  (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Public Sub ShutdownWindows(ByVal intParamater As Integer)

#If Win16 Then    '16 bit
    
  Dim intRetVal As Integer
  
  intRetVal = ExitWindows(intParamater, 0)
  
#Else    '32 bit

  Dim blnReturn As Boolean
  
  blnReturn = ExitWindowsEx(intParamater, 0)

#End If

End Sub

Public Function Get_Windows_Directory() As String

Dim strBuffer As String * 256

  #If Win16 Then    '16 bit code
  
  Dim intReturn As Integer
    
    intReturn = GetWindowsDirectory(strBuffer, Len(strBuffer))
    ' Trim Buffer and return string
    Get_Windows_Directory = Left$(strBuffer, intReturn)
    
  #Else    '32 bit code
  
  Dim lngReturn As Integer
  
    lngReturn = GetWindowsDirectory32(strBuffer, Len(strBuffer))
    ' Trim Buffer and return string
    Get_Windows_Directory = Left$(strBuffer, lngReturn)
    
  #End If

End Function

Public Function Get_Windows_System_Directory() As String
    
Dim strBuffer As String * 256

  #If Win16 Then    '16 bit code
    
    Dim intReturn As Integer
    intReturn = GetSystemDirectory(strBuffer, Len(strBuffer))
    'trim C style String
    Get_Windows_System_Directory = Left$(strBuffer, intReturn)
    
  #Else    '32 bit code
  
    Dim lngReturn As Long
    lngReturn = GetSystemDirectory32(strBuffer, Len(strBuffer))
    'trim C style String
    Get_Windows_System_Directory = Left$(strBuffer, lngReturn)
    
  #End If
  
End Function

Public Function GetSysUserName() As String

Dim UserName As String
Dim BufSize As Long
Dim RetCode As Long
Dim NullCharPos As Long

  UserName = Space(80)
  BufSize = Len(UserName)
  
  '---call WINAPI
  RetCode = GetUserName(UserName, BufSize)
  
  '---search for the end of the string
  NullCharPos = InStr(UserName, Chr(0))
  
  If NullCharPos > 0 Then
    UserName = Left(UserName, NullCharPos - 1)
  Else
    UserName = ""
  End If
  
  GetSysUserName = UserName
  
End Function

Public Function GetSysComputerName() As String

Dim Computer As String
Dim BufSize As Long
Dim RetCode As Long
Dim NullCharPos As Long

  Computer = Space(80)
  BufSize = Len(Computer)
  
  '---call WINAPI
  RetCode = GetComputerName(Computer, BufSize)
  
  '---search for the end of the string
  NullCharPos = InStr(Computer, Chr(0))
  If NullCharPos > 0 Then
    Computer = Left(Computer, NullCharPos - 1)
  Else
    Computer = ""
  End If
  
  GetSysComputerName = Computer
  
End Function

Public Function TimeWindowsRunning()
'
Dim Minutes As Single
Dim Seconds As Single
Dim Hours As Single
  '
  Minutes = 0
  Hours = 0
  Seconds = GetTickCount / 1000
  '
  Do Until Seconds < 60
    Seconds = Seconds - 60
    Minutes = Minutes + 1
  Loop
  '
  Do Until Minutes < 60
    Minutes = Minutes - 60
    Hours = Hours + 1
  Loop
  '
  TimeWindowsRunning = Hours & "hr : " & Minutes & "min : " & Round(Seconds, 0) & "sec"
  '
End Function

Public Function GetDriveDesc(strDrive As String) As String

Dim RC As Long
  
  RC = GetDriveType(strDrive)  ' Get Drive Type
  
  Select Case RC
    Case DRIVE_CDROM
      GetDriveDesc = "[CD Rom Drive]"
    Case DRIVE_FIXED
      GetDriveDesc = "[Fixed Drive]"
    Case DRIVE_REMOTE
      GetDriveDesc = "[Network Drive]"
    Case DRIVE_RAMDISK
      GetDriveDesc = "[Ram Disk Drive]"
    Case DRIVE_REMOVABLE
      GetDriveDesc = "[Removable Disk Drive]"
    Case Else
      GetDriveDesc = "[Unknown drive type or error]"
  End Select
    
End Function

Public Function GetUNCPath(ByVal strDriveLetter As String, _
  ByRef strUNCPath As String) As Long

On Local Error GoTo GetUNCPath_Err

Dim Msg             As String
Dim lngReturn       As Long
Dim strLocalName    As String
Dim strRemoteName   As String
Dim lngRemoteName   As Long

  strLocalName = strDriveLetter
  strRemoteName = String$(255, Chr$(32))
  lngRemoteName = Len(strRemoteName)
  lngReturn = WNetGetConnection(strLocalName, strRemoteName, lngRemoteName)
  
  If lngReturn = NO_ERROR Then
    '// Return the UNC through the function.
    GetUNCPath = NO_ERROR
    strUNCPath = Trim$(strRemoteName)
    strUNCPath = Left$(strUNCPath, Len(strUNCPath) - 1)
  Else
    '// Return the original driveletter & error.
    GetUNCPath = lngReturn
    strUNCPath = strDriveLetter & "\"
  End If
  
GetUNCPath_End:
  
  Exit Function
  
GetUNCPath_Err:
  
  GetUNCPath = ERROR_NOT_SUPPORTED
  strUNCPath = strDriveLetter
  Resume GetUNCPath_End
  
End Function

