Attribute VB_Name = "mLabel"
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Public Const REG_DWORD = 4
Public Const HKEY_DYN_DATA = &H80000006

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim SystemInfo As SYSTEM_INFO

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Private Const PROCESSOR_ALPHA_21064 = 21064
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_MIPS_R4000 = 4000

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0

Private Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type gLabelSession
  'Image display
  iImageCount As Long
  iImagePath() As String
  
  'Text variables
  sDisplayText As String
  sDisplayFontName As String
  sDisplayFontSize As Single
  
  'Name of the profile
  sProfileName As String
  
  'Background options
  bBackTransparent As Boolean
  sBackgroundImage As String
  bTiledPattern As Boolean
End Type

Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2

Public Session As gLabelSession

Public Sub DrawPict(iPicture As StdPicture, pForm As Form)
pForm.Picture = iPicture
End Sub

Public Sub DrawTiledPic(iPicture As StdPicture, pForm As Form)
Dim wX As Long, hY As Long
Dim cX As Long, cY As Long
'frmLbl.pTmp is set to AutoSize itself. Useful
'to get the height and width of the picture.
frmLbl.pTmp.Picture = iPicture
cX = frmLbl.pTmp.ScaleWidth
cY = frmLbl.pTmp.ScaleHeight
If cX >= pForm.Width And cY >= pForm.Height Then pForm.Picture = iPicture: Exit Sub
'above we check if it's smaller, if not just display it, no use tiling
For wX = 0 To pForm.Width \ cX
  For hY = 0 To pForm.ScaleHeight \ cY
    pForm.PaintPicture iPicture, wX * cX, hY * cY
  Next hY
Next wX
End Sub

Public Sub Main()
'Does a whole lot of things
On Error Resume Next
Dim i As Long 'looper
If Command$ = "" Then
  MsgBox "Profile to load must be specified as a command-line argument." & vbCrLf & "You will need to restart SysLabel with a proper command-line.", vbExclamation, "Error": End
  'must know profile name
End If
If ReadValue("ImageCount", "") = "" Then
  MsgBox "The settings file may be corrupted." & vbCrLf & "Current profile could not be loaded.", vbExclamation, FullPath(App.Path, "display.inf"): End
  'invalid info. I don't like someone messing with INFs.
End If
Session.sProfileName = Command$
Load frmLbl
Session.iImageCount = ReadValue("ImageCount", 0) 'any images?
If Session.iImageCount = 0 Then
  frmLbl.pDisp(0).Visible = False 'hide pDisp
Else
  ReDim Session.iImagePath(Session.iImageCount - 1) 'paths
  For i = 0 To Session.iImageCount - 1
    Load frmLbl.pDisp(i) 'load a new picBox
    frmLbl.pDisp(i).Visible = True
    Session.iImagePath(i) = ReadValue("ImageFile" & i + 1, "")
    frmLbl.pDisp(i).Picture = LoadPicture(Session.iImagePath(i))
    
    frmLbl.pDisp(i).Move ReadValue("ImageLeft" & i + 1), _
    ReadValue("ImageTop" & i + 1), ReadValue _
    ("ImageWidth" & i + 1), ReadValue("ImageHeight" & i + 1) ';
    
    frmLbl.pDisp(i).AutoSize = ReadValue("ImageAutoSize" & i + 1)
    
  Next i
End If
Session.bBackTransparent = ReadValue("BackTransparent")
Session.sBackgroundImage = ReadValue("BackgroundImage")
Session.bTiledPattern = ReadValue("TiledBackground")
If Session.bBackTransparent = False Then
  frmLbl.AutoRedraw = True
  If Session.bTiledPattern = True Then
    DrawTiledPic LoadPicture(Session.sBackgroundImage), frmLbl
  Else
    DrawPict LoadPicture(Session.sBackgroundImage), frmLbl
  End If
Else
  frmLbl.AutoRedraw = False
End If
If ReadValue("ShowAbout") = False Then
  frmLbl.lbM.Top = 45
  frmLbl.lbTitle.Visible = False
  frmLbl.imgClose.Visible = False
End If

frmLbl.Move ReadValue("Left"), ReadValue("Top"), ReadValue("Width"), ReadValue("Height")

Session.sDisplayFontName = ReadValue("FontName")
Session.sDisplayFontSize = ReadValue("FontSize")

frmLbl.BackColor = ReadValue("BackgroundColor")
frmLbl.lbM.ForeColor = ReadValue("FontColor")

frmLbl.lblStatus.ForeColor = frmLbl.lbM.ForeColor

frmLbl.lbM.FontName = Session.sDisplayFontName
frmLbl.lbM.FontSize = Session.sDisplayFontSize
frmLbl.lbM.Alignment = ReadValue("TextJustify")

Session.sDisplayText = ReadValue("DisplayText")
frmLbl.lblStatus.Visible = (InStr(Session.sDisplayText, "cpu-usage-graphical") > 0)
frmLbl.picStatus.Visible = frmLbl.lblStatus.Visible

If frmLbl.lblStatus.Visible Then frmLbl.InitCPU

frmLbl.Tim.Interval = ReadValue("UpdateInterval")
frmLbl.Tim_Timer
frmLbl.Show
End Sub

Function AMPM(ByVal Hour As Integer) As String
If Hour = 0 Then Hour = 24
If Hour > 12 Then AMPM = "PM" Else AMPM = "AM"
End Function

Function NoAMPM(ByVal Hour As Integer) As Integer
'e.g. 23:00 returns 11:00 PM
If Hour = 0 Then Hour = 24
If Hour > 12 Then NoAMPM = Hour - 12 Else NoAMPM = Hour
End Function

Function MMSS(whole) As String
Dim mm, ss, sInt As Integer
mm = whole \ 60
ss = ModDecimal(whole, 60)
ss = Round(ss, 0)
If Len(mm) < 2 Then mm = "0" & mm
If Len(ss) < 2 Then ss = "0" & ss
MMSS = mm & ":" & ss
End Function

Function ModDecimal(What, Divider) As Single
Dim l As Single
l = Divider * (What \ Divider)
ModDecimal = What - l
End Function

Public Function ReadValue(Key As String, Optional Default As String, Optional Section, Optional File)
    ' Read from INI file
    Dim sReturn As String
    If IsMissing(File) Then File = FullPath(App.Path, "display.inf")
    If IsMissing(Section) Then Section = Command$
    sReturn = String(255, Chr(0))
    ReadValue = Left(sReturn, GetPrivateProfileString(Section, Key, Default, sReturn, Len(sReturn), File))
End Function

Public Sub SaveValue(Key As String, Value As String, Optional Section, Optional File)
    ' Write to INI file
    If IsMissing(File) Then File = FullPath(App.Path, "display.inf")
    If IsMissing(Section) Then Section = Command$
    WritePrivateProfileString Section, Key, Value, File
End Sub

Function FullPath(lpPath As String, lpFile As String) As String
If Right(lpPath, 1) <> "\" Then lpPath = lpPath & "\"
FullPath = lpPath & lpFile
'fullpath, after resolving the "\" problems
End Function

Function Memory(ReturnedTotal As Long, ReturnedAvailable As Long)
Dim memoryInfo As MEMORYSTATUS
GlobalMemoryStatus memoryInfo
ReturnedTotal = memoryInfo.dwTotalPhys
ReturnedAvailable = memoryInfo.dwAvailPhys
End Function

Function Processor() As String
GetSystemInfo SystemInfo
Select Case SystemInfo.dwProcessorType
Case PROCESSOR_ALPHA_21064 = 21064
Processor = "Alpha"
Case PROCESSOR_INTEL_386
Processor = "Intel 80386"
Case PROCESSOR_INTEL_486
Processor = "Intel 80486"
Case PROCESSOR_INTEL_PENTIUM
Processor = "Intel Pentium"
Case PROCESSOR_MIPS_R4000
Processor = "MIPS"
End Select
End Function

Public Function WindowsVer()
Dim infoStruct As OSVERSIONINFO
infoStruct.dwOSVersionInfoSize = Len(infoStruct)
GetVersionEx infoStruct
If infoStruct.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    WindowsVer = "Windows 95/98"
Else
    WindowsVer = "Windows NT"
End If
End Function

Function Uptime(which As String) As String
Dim gTick As Long
Dim days, mins, hours, secs
gTick = GetTickCount()
gTick = gTick / 1000
days = gTick \ 86400
hours = gTick \ 3600 - (days * 24)
mins = (gTick \ 60) Mod 60
secs = gTick Mod 60
Select Case which
Case "m"
Uptime = mins
Case "mm"
Uptime = mins
If Len(Uptime) = 1 Then Uptime = "0" & Uptime
Case "h"
Uptime = hours
Case "hh"
Uptime = hours
If Len(Uptime) = 1 Then Uptime = "0" & Uptime
Case "d"
Uptime = days
Case "dd"
Uptime = days
If Len(Uptime) = 1 Then Uptime = "0" & Uptime
Case "s"
Uptime = secs
Case "ss"
Uptime = secs
If Len(Uptime) = 1 Then Uptime = "0" & Uptime
End Select
End Function

Function GetActiveTask(hWnd As Long) As String
On Error Resume Next
Dim f As String, ln As Long
f = Space(260)
ln = GetWindowText(hWnd, f, 260)
GetActiveTask = Left(f, ln)
End Function

Function GetActiveWindow() As String
Dim lpP As POINTAPI, hw As Long
GetCursorPos lpP
hw = WindowFromPoint(lpP.X, lpP.Y)
GetActiveWindow = GetActiveTask(hw)
End Function

Function MemoryAvailable() As String
Dim lr As Long, lrs As String, lrd As Double
Memory 0, lr
lrd = lr: lrs = " bytes"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " KB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " MB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " GB"
lrs = lrd & lrs
MemoryAvailable = lrs
End Function

Function MemoryUsed() As String
Dim lr As Long, lrs As String, lrd As Double, lrtmp As Long
Memory lrtmp, lr
lr = lrtmp - lr
lrd = lr: lrs = " bytes"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " KB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " MB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " GB"
lrs = lrd & lrs
MemoryUsed = lrs
End Function

Function MemoryTotal() As String
Dim lr As Long, lrs As String, lrd As Double
Memory lr, 0
lrd = lr: lrs = " bytes"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " KB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " MB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " GB"
lrs = lrd & lrs
MemoryTotal = lrs
End Function

Function GetCPUUsage() As String
    Dim lData As Long
    Dim lType As Long
    Dim lSize As Long
    Dim hKey As Long
    Dim Qry As String
    Dim Status As Long
                  
    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
                
    If Qry <> 0 Then Exit Function
                
    lType = REG_DWORD
    lSize = 4
                
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
    
    Status = lData

GetCPUUsage = Status & "%"
End Function
