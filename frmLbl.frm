VERSION 5.00
Begin VB.Form frmLbl 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "SysLabel v1.2"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   ControlBox      =   0   'False
   Icon            =   "frmLbl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCPUStatus 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1890
      Top             =   2205
   End
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      Height          =   195
      Left            =   360
      ScaleHeight     =   135
      ScaleWidth      =   1440
      TabIndex        =   5
      Top             =   4455
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Timer Tim 
      Interval        =   1000
      Left            =   1890
      Top             =   2700
   End
   Begin VB.PictureBox pTmp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   1935
      ScaleHeight     =   450
      ScaleWidth      =   1170
      TabIndex        =   1
      Top             =   4095
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox pDisp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   0
      Left            =   675
      ScaleHeight     =   510
      ScaleWidth      =   1230
      TabIndex        =   2
      Top             =   2385
      Width           =   1230
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   450
      TabIndex        =   4
      Top             =   3690
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SysLabel v1.2 [http://sushantshome.tripod.com]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   15
      Width           =   4215
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   855
      Picture         =   "frmLbl.frx":0E42
      Top             =   0
      Width           =   270
   End
   Begin VB.Image imCloseDn 
      Height          =   225
      Left            =   1890
      Picture         =   "frmLbl.frx":0EAE
      Top             =   3555
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imCloseN 
      Height          =   225
      Left            =   1890
      Picture         =   "frmLbl.frx":1233
      Top             =   3240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbM 
      BackStyle       =   0  'Transparent
      Caption         =   "SysLabel v1.2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2805
      Left            =   45
      TabIndex        =   0
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmLbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function PaintDesktop Lib "user32" (ByVal hDC As Long) As Long

Private Sub Form_Paint()
If Session.bBackTransparent Then PaintDesktop hDC
End Sub

Private Sub Form_Resize()
imgClose.Left = ScaleWidth - imgClose.Width
lbM.Width = Width - 30
lbM.Height = Height - 30
End Sub

Private Sub imgClose_Click()
End
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClose.Picture = imCloseDn.Picture
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClose.Picture = imCloseN.Picture
End Sub

Sub Tim_Timer()
Dim st As String
st = Session.sDisplayText
st = Replace(st, "\n", vbCrLf)

st = Replace(st, "cpu-usage-text", GetCPUUsage())

st = Replace(st, "comp-name", CompName)
st = Replace(st, "windows-ver", WindowsVer)
st = Replace(st, "processor-info", Processor)

st = Replace(st, "mem-a", MemoryAvailable)
st = Replace(st, "mem-u", MemoryUsed)
st = Replace(st, "mem-t", MemoryTotal)

st = Replace(st, "active-task", GetActiveTask(GetFocus()))
st = Replace(st, "active-window", GetActiveWindow)

st = Replace(st, "uptime-dd", Uptime("dd"))
st = Replace(st, "uptime-d", Uptime("d"))

st = Replace(st, "uptime-hh", Uptime("hh"))
st = Replace(st, "uptime-h", Uptime("h"))

st = Replace(st, "uptime-mm", Uptime("mm"))
st = Replace(st, "uptime-m", Uptime("m"))

st = Replace(st, "uptime-ss", Uptime("ss"))
st = Replace(st, "uptime-s", Uptime("s"))

st = Replace(st, "time-hh", Format(Time, "hh"))
st = Replace(st, "time-h", Format(Time, "h"))

st = Replace(st, "time-nn", Format(Time, "nn"))
st = Replace(st, "time-n", Format(Time, "n"))

st = Replace(st, "time-ss", Format(Time, "ss"))
st = Replace(st, "time-s", Format(Time, "s"))

st = Replace(st, "date-dddd", Format(Date, "dddd"))
st = Replace(st, "date-ddd", Format(Date, "ddd"))
st = Replace(st, "date-dd", Format(Date, "dd"))
st = Replace(st, "date-d", Format(Date, "d"))

st = Replace(st, "date-mmmm", Format(Date, "mmmm"))
st = Replace(st, "date-mmm", Format(Date, "mmm"))
st = Replace(st, "date-mm", Format(Date, "mm"))
st = Replace(st, "date-m", Format(Date, "m"))

st = Replace(st, "date-yyyy", Format(Date, "yyyy"))
st = Replace(st, "date-yy", Format(Date, "yy"))

lbM.Caption = st
End Sub

Function CompName() As String
Dim st As String, ln As Long
st = Space(260)
ln = GetComputerName(st, 260)
CompName = Left(st, ln)
End Function

Function InitCPU()
On Error Resume Next
picStatus.Left = ReadValue("CPUGraphicLeft")
picStatus.Top = ReadValue("CPUGraphicTop")
lblStatus.Left = picStatus.Left + picStatus.Width + 45
lblStatus.Top = picStatus.Top + ((picStatus.Height - lblStatus.Height) / 2)

Dim lData As Long
Dim lType As Long
Dim lSize As Long
Dim hKey As Long
Dim Qry As String

Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)

If Qry <> 0 Then
    MsgBox "Could not open registry!", vbExclamation, "Error"
    Exit Function
End If

lType = REG_DWORD
lSize = 4

Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
Qry = RegCloseKey(hKey)

Session.sDisplayText = Replace(Session.sDisplayText, "cpu-usage-graphical", "")

tmrCPUStatus.Enabled = True
tmrCPUStatus.Interval = Tim.Interval

picStatus.BackColor = RGB(130, 130, 170)
End Function

Private Sub tmrCpuStatus_Timer()

    Dim lData As Long
    Dim lType As Long
    Dim lSize As Long
    Dim hKey As Long
    Dim Qry As String
    Dim Status As Long
                  
    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
                
    If Qry <> 0 Then Exit Sub
                
    lType = REG_DWORD
    lSize = 4
                
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
    
    Status = lData

    ' show CPU usage in Label
    lblStatus.Caption = Status & "%"
    
    ' show CPU usage in our selfmade progressbar
    ' when CPU usage is over 80% then color the status red
    If Status < 80 Then
    picStatus.Line (Status * 15, 0)-(0, 135), RGB(255, 245, 85), BF
    Else
    picStatus.Line (Status * 15, 0)-(0, 135), RGB(245, 10, 0), BF
    End If
    picStatus.Line (Status * 15, 0)-(1470, 135), RGB(130, 130, 170), BF
    
    Qry = RegCloseKey(hKey)

End Sub

