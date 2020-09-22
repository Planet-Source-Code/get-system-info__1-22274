Attribute VB_Name = "modCPUInfo"
Option Explicit

Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        wProcessorLevel As Integer
        wProcessorRevision As Integer
End Type

Public Type OSVERSIONINFO ' 148 bytes
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Public Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type


Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


Public Declare Function GetProcessor Lib "getcpu.dll" (ByVal strCpu As String, ByVal strVendor As String, ByVal strL2Cache As String) As Long
Public Declare Function GetProcessorRawSpeed Lib "getcpu.dll" (ByVal RawSpeed As String) As Long
Public Declare Function GetProcessorNormSpeed Lib "getcpu.dll" (ByVal NormSpeed As String) As Long

#If Win32 Then
Public Const VER_PLATFORM_WIN32_NT& = 2
Public Const VER_PLATFORM_WIN32_WINDOWS& = 1
#End If

Global myVer As OSVERSIONINFO

