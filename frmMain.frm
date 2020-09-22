VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPU Info"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4440
      Top             =   120
   End
   Begin VB.Label lblIP 
      Caption         =   "1"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "RAM (Available)"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label15 
      Caption         =   "RAM (Total)"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Processor"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Processor Vendor"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblRamAvail 
      Caption         =   "1"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblRamTotal 
      Caption         =   "1"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblNormSpeed 
      Caption         =   "1"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblRawSpeed 
      Caption         =   "1"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblProcessor 
      Caption         =   "1"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblVendor 
      Caption         =   "1"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblSP 
      Caption         =   "1"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblOS 
      Caption         =   "1"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Processor Speed(Normal)"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Processor Speed(Raw)"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Service Pack"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "OS"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub GetInfo()

Dim memoryInfo As MEMORYSTATUS
Dim sCpu As String, sVendor As String
Dim sL2Cache As String
Dim sRawSpeed As String
Dim sNormSpeed As String
Dim dl&, s$
Dim mySys As SYSTEM_INFO

  GlobalMemoryStatus memoryInfo
  lblRamTotal.Caption = Round(memoryInfo.dwTotalPhys / 1043321, 0)
  lblRamAvail.Caption = Round(memoryInfo.dwAvailPhys / 1043321, 0)
  
  sCpu = String(255, 0)
  sRawSpeed = String(255, 0)
  sNormSpeed = String(255, 0)
  sVendor = String(255, 0)
  sL2Cache = String(255, 0)
  
  GetProcessor sCpu, sVendor, sL2Cache
  GetProcessorRawSpeed sRawSpeed
  GetProcessorNormSpeed sNormSpeed

  lblProcessor.Caption = StripZero(sCpu)
  lblVendor.Caption = StripZero(sVendor)
  lblRawSpeed.Caption = StripZero(sRawSpeed)
  lblNormSpeed.Caption = StripZero(sNormSpeed)
  
  myVer.dwOSVersionInfoSize = 148
  dl& = GetVersionEx&(myVer)
  s$ = LPSTRToVBString(myVer.szCSDVersion)
  
  lblSP.Caption = s$
  
  If myVer.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    s$ = "Windows95 "
  ElseIf myVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
    s$ = "Windows NT "
  End If
  
  lblOS.Caption = s$ & myVer.dwMajorVersion & "." & myVer.dwMinorVersion & " Build " & (myVer.dwBuildNumber And &HFFFF&)
  
  'lblIP = Winsock1.LocalIP
  lblIP = GetIPAddress

End Sub

Public Function StripZero(sInput As String) As String

Dim nPos As Integer
Dim x As New clsWinAPI

  nPos = InStr(1, sInput, Chr(0))
  
  If nPos <> 0 Then
    StripZero = Left$(sInput, nPos - 1)
  Else
    StripZero = sInput
  End If
  
  frmMain.Caption = "CPU Info for " & x.GetSysComputerName
  
End Function

Public Function LPSTRToVBString$(ByVal s$)
  
Dim nullpos&

  nullpos& = InStr(s$, Chr$(0))
  
  If nullpos > 0 Then
      LPSTRToVBString = Left$(s$, nullpos - 1)
  Else
      LPSTRToVBString = ""
  End If
  
End Function

Private Sub Form_Load()

  GetInfo

End Sub

Private Sub Timer1_Timer()
  
Dim memoryInfo As MEMORYSTATUS
Dim sRawSpeed As String
Dim sNormSpeed As String

  GlobalMemoryStatus memoryInfo
  lblRamTotal.Caption = Round(memoryInfo.dwTotalPhys / 1043321, 0)
  lblRamAvail.Caption = Round(memoryInfo.dwAvailPhys / 1043321, 0)
  
  sRawSpeed = String(255, 0)
  sNormSpeed = String(255, 0)
  
  GetProcessorRawSpeed sRawSpeed
  GetProcessorNormSpeed sNormSpeed

  lblRawSpeed.Caption = StripZero(sRawSpeed)
  lblNormSpeed.Caption = StripZero(sNormSpeed)

End Sub
