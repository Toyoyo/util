VERSION 4.00
Begin VB.Form MainForm 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Util"
   ClientHeight    =   1785
   ClientLeft      =   9000
   ClientTop       =   3630
   ClientWidth     =   2370
   Height          =   2190
   Icon            =   "MAINFORM.frx":0000
   Left            =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   158
   Top             =   3285
   Width           =   2490
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin GaugeLib.Gauge ggUserLoad 
      Height          =   270
      Left            =   1260
      TabIndex        =   0
      Top             =   330
      Width           =   1155
      _version        =   65536
      _extentx        =   2037
      _extenty        =   476
      _stockprops     =   73
      forecolor       =   32768
      backcolor       =   16777152
      innertop        =   2
      innerleft       =   2
      innerright      =   2
      innerbottom     =   2
      needlewidth     =   1
   End
   Begin GaugeLib.Gauge ggSysLoad 
      Height          =   270
      Left            =   1260
      TabIndex        =   1
      Top             =   60
      Width           =   1155
      _version        =   65536
      _extentx        =   2037
      _extenty        =   476
      _stockprops     =   73
      forecolor       =   32768
      backcolor       =   16777152
      innertop        =   2
      innerleft       =   2
      innerright      =   2
      innerbottom     =   2
      needlewidth     =   1
   End
   Begin VB.Label LblWText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/dev/null"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   1110
      Width           =   2415
   End
   Begin VB.Label LblPID 
      BackStyle       =   0  'Transparent
      Caption         =   "PID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   1590
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   150
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Label LblAct 
      BackStyle       =   0  'Transparent
      Caption         =   "WClass:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1350
      Width           =   2295
   End
   Begin VB.Label LblUserLoad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Load (USER)"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   300
      Width           =   945
   End
   Begin VB.Label LblSystemLoad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Load (SYSTEM)"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   30
      Width           =   1155
   End
   Begin VB.Label LblProc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processus count"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label LblRamFree 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreeSpace"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim bChrono As Boolean
Dim nDuree, nTmpDuree As Long

Private Sub CmdPause_Click()
    bChrono = False
End Sub

Private Sub CmdStart_Click()
    Timer.Enabled = False
    Timer.Enabled = True
    bChrono = True
End Sub

Private Sub CmdStop_Click()
    bChrono = False
    nDuree = 0
End Sub

Private Sub Form_Load()
    Timer_Timer
End Sub


Private Sub Timer_Timer()
    Dim hwnd, hwnd2 As Long
    Dim buflen As Long
    Dim buf As String
    
    'Clock
    Caption = right$(left$(Date$, 5), 2) + "/" + left$(Date$, 2) + "/" + right$(Date$, 4) _
            + " " + time$
            
        hwnd = GetActiveWindow
        
        buflen = 256
        buf = Space$(buflen - 1)
        buflen = GetClassName(hwnd, buf, buflen)
        buf = left$(buf, buflen)
        lblact = "WClass: " + buf
        
        buf = Space$(buflen - 1)
        lblpid = "PID:" + Str$(GetWindowTask(hwnd))
        
        buflen = GetWindowTextLength(hwnd)
        buf = Space$(buflen)
        buflen = GetWindowText(hwnd, buf, buflen + 1)
        buf = left$(buf, buflen)
        LblWText = buf

   'osview
    Dim nSysLoad As Integer
    Dim nUserLoad As Integer
    
    nSysLoad = 100 - GetFreeSystemResources(GFSR_SYSTEMRESOURCES)
    nUserLoad = 100 - GetFreeSystemResources(GFSR_USERRESOURCES)
    
    Select Case nSysLoad
        Case Is > 90
            ggSysLoad.ForeColor = QBColor(12)
        Case Is > 70
            ggSysLoad.ForeColor = QBColor(14)
        Case Is <= 70
            ggSysLoad.ForeColor = QBColor(2)
    End Select
    
    Select Case nUserLoad
        Case Is > 90
            ggUserLoad.ForeColor = QBColor(12)
        Case Is > 70
            ggUserLoad.ForeColor = QBColor(14)
        Case Is <= 70
            ggUserLoad.ForeColor = QBColor(2)
    End Select
    
    ggSysLoad.value = nSysLoad
    ggUserLoad.value = nUserLoad
    
    LblProc = "Active tasks: " + Str$(GetNumTasks())
    On Error GoTo subend
    LblRamFree = "Free system memory: " + Str$(Int(GetFreeSpace(0) / 1024)) + "KB"
subend:
End Sub
