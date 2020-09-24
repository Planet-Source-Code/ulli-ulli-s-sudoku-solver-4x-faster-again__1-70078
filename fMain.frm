VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fMain 
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11355
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   355
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   757
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog CDl 
      Left            =   -15
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txPuz 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   0
      Left            =   345
      MaxLength       =   1
      MousePointer    =   1  'Pfeil
      TabIndex        =   4
      Top             =   570
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txSol 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Index           =   0
      Left            =   6570
      Locked          =   -1  'True
      MaxLength       =   1
      MousePointer    =   1  'Pfeil
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton btSolve 
      Caption         =   ">> Solve >>"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5055
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Solve Puzzle"
      Top             =   2025
      Width           =   1245
   End
   Begin VB.CommandButton btExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5070
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4665
      Width           =   1245
   End
   Begin VB.CommandButton btClear 
      Caption         =   "<< Clear >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5055
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Clear Puzzle and Solution"
      Top             =   945
      Width           =   1245
   End
   Begin VB.Label lbAnimate 
      BackStyle       =   0  'Transparent
      Caption         =   "  Animate  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5205
      TabIndex        =   13
      Top             =   1590
      Width           =   915
   End
   Begin VB.Label lbSolved 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Solved"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   5055
      TabIndex        =   12
      Top             =   2655
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lbNoSol 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "No Solution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   5055
      TabIndex        =   11
      Top             =   2655
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lbHide 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   " Hide"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   10575
      TabIndex        =   10
      ToolTipText     =   "Hide Solution"
      Top             =   210
      Width           =   480
   End
   Begin VB.Image imgMe 
      Height          =   630
      Left            =   5340
      Picture         =   "fMain.frx":08CA
      ToolTipText     =   "Left: Send Mail to Author  Right: Show About"
      Top             =   90
      Width           =   675
   End
   Begin VB.Label lbTime 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   885
      Left            =   4965
      TabIndex        =   7
      Top             =   3195
      Width           =   1410
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Puzzle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   0
      Left            =   2085
      TabIndex        =   6
      Top             =   135
      Width           =   945
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Solution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Index           =   1
      Left            =   8205
      TabIndex        =   5
      Top             =   135
      Width           =   1155
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   1500
      Index           =   0
      Left            =   330
      Top             =   555
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   1500
      Index           =   1
      Left            =   330
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   1500
      Index           =   2
      Left            =   330
      Top             =   3525
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   1500
      Index           =   3
      Left            =   1815
      Top             =   3525
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   1500
      Index           =   4
      Left            =   1815
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   1500
      Index           =   5
      Left            =   1815
      Top             =   555
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   1500
      Index           =   6
      Left            =   3300
      Top             =   3525
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   1500
      Index           =   7
      Left            =   3300
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   1500
      Index           =   8
      Left            =   3300
      Top             =   555
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   4515
      Index           =   9
      Left            =   315
      Top             =   540
      Width           =   4515
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   1500
      Index           =   11
      Left            =   9525
      Top             =   555
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   1500
      Index           =   12
      Left            =   9525
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   1500
      Index           =   13
      Left            =   9525
      Top             =   3525
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   1500
      Index           =   14
      Left            =   8040
      Top             =   555
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   1500
      Index           =   15
      Left            =   8040
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   1500
      Index           =   16
      Left            =   8040
      Top             =   3525
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   1500
      Index           =   17
      Left            =   6555
      Top             =   3525
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   1500
      Index           =   18
      Left            =   6555
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   1500
      Index           =   19
      Left            =   6555
      Top             =   555
      Width           =   1500
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Puzzle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   2070
      TabIndex        =   8
      Top             =   120
      Width           =   945
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Solution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   3
      Left            =   8190
      TabIndex        =   9
      Top             =   120
      Width           =   1155
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   4515
      Index           =   10
      Left            =   6540
      Top             =   540
      Width           =   4515
   End
   Begin VB.Shape shBg 
      BorderColor     =   &H00808080&
      FillColor       =   &H00E8E8D0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   2460
      Left            =   4965
      Top             =   1920
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load Puzzle..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSavePuz 
         Caption         =   "Save &Puzzle As..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSaveSol 
         Caption         =   "Save &Solution As..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintPuz 
         Caption         =   "Print P&uzzle"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPrintSol 
         Caption         =   "Print S&olution"
         Shortcut        =   ^S
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuVoice 
         Caption         =   "Voice"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnimate 
         Caption         =   "&Animate"
         Begin VB.Menu mnuAnimSpeed 
            Caption         =   "Off"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuAnimSpeed 
            Caption         =   "Slow"
            Index           =   1
         End
         Begin VB.Menu mnuAnimSpeed 
            Caption         =   "Medium"
            Index           =   2
         End
         Begin VB.Menu mnuAnimSpeed 
            Caption         =   "Fast"
            Index           =   3
         End
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide"
         Shortcut        =   {F9}
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSolve 
         Caption         =   "&Solve"
      End
   End
   Begin VB.Menu mnuQ 
      Caption         =   "&?"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSendMail 
         Caption         =   "&Send Mail"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" () ':) Line inserted by Formatter
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function Beeper Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qRC As tRect, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal ByPos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long

Private HscFrequ    As Currency 'high speed counter frequency - using currency type as 64bit-doublelong
Private StartTick   As Currency
Private EndTick     As Currency
Private Correction  As Currency
Private DelayStart  As Currency
Private DelayEnds   As Currency

Private Type MENUITEMINFO
    cbSize          As Long
    fMask           As Long
    fType           As Long
    fState          As Long
    wID             As Long
    hSubMenu        As Long
    hbmpChecked     As Long
    hbmpUnchecked   As Long
    dwItemData      As Long
    dwTypeData      As String
    cch             As Long
End Type
Private MII As MENUITEMINFO

Private Type tRect
    L       As Long
    t       As Long
    r       As Long
    b       As Long
End Type
Private RECT        As tRect

Private Enum ApiConstants
    BDR_RAISEDOUTER = 1
    BDR_SUNKENOUTER = 2
    BDR_RAISEDINNER = 4
    BDR_SUNKENINNER = 8
    BDR_FILLET = BDR_RAISEDOUTER Or BDR_SUNKENINNER
    BDR_RIDGE = BDR_SUNKENOUTER Or BDR_RAISEDINNER
    BDR_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
    BDR_SUNKEN = BDR_SUNKENOUTER Or BDR_SUNKENINNER
    BF_RECT = 15
    BF_MONO = &H8000
    SW_SHOWNORMAL = 1
    SE_NO_ERROR = 33  'Values below 33 are error returns
    CS_DROPSHADOW = &H20000
    GCL_STYLE = -26
    REALTIME_PRIORITY_CLASS = &H100
    MFS_DEFAULT = &H1000
    MIIM_STATE = 1
End Enum

Private Vox             As SpVoice
Attribute Vox.VB_VarHelpID = -1
Private UserName        As String
Private DecSep          As String
Private Done            As Boolean
Private Timeout         As Boolean
Private Animate         As Boolean
Private Internal        As Boolean
Private LastFocus       As Integer
Private PrioClass       As Long
Private Bits(1 To 9)    As Long
Private FirstFree       As Long
Private LastFree        As Long
Private LastHit         As Long
Private PermissionBits  As Long
Private Success         As Long
Private Const Limit     As Long = 999999
Private Const Interval  As Long = -1 + 2 ^ 13 'must be of this form:  -1 + 2 ^ n
Private AnimDiv         As Long
Private i               As Long
Private j               As Long
Private k               As Long
Private Log2            As Double
Private Const Title     As String = "Ulli's Sudoku Solver"

Private Function AllAgree(ByVal Cellnumber As Long, ByVal Value As Long) As Boolean

  'cross hatching
  'returns true if row, column and block agree with Value to be put into Cell(CellNumber)

    If Groups(Cells(Cellnumber).RowNumber).Agree(Value) Then
        If Groups(Cells(Cellnumber).ColumnNumber).Agree(Value) Then
            AllAgree = Groups(Cells(Cellnumber).BlockNumber).Agree(Value)
        End If
    End If

End Function

Private Sub btClear_Click()

  'resets all values

    For j = 1 To 81
        With Cells(j - 1)
            .Value = 0
            .Fixed = False
        End With 'CELLS(j
        With txPuz(j)
            .Text = vbNullString
            .ForeColor = txPuz(0).ForeColor
            .BackColor = txPuz(0).BackColor
            .Tag = vbNullString
        End With 'TXPUZ(j)
        With txSol(j)
            .Text = vbNullString
            .ForeColor = txSol(0).ForeColor
            .BackColor = txSol(0).BackColor
        End With 'TXSOL(j)
    Next j
    For j = 0 To 26
        Groups(j).Reset
    Next j
    SetVisibleTo False
    txPuz(1).SetFocus
    Caption = Title

End Sub

Private Sub btExit_Click()

  'good bye

    mnuExit_Click

End Sub

Private Sub btSolve_Click()

  'prepares for finding the solution

  Dim Timing    As Single
  Dim Res       As String
  Dim Solved    As Boolean

    Enabled = False
    SetVisibleTo False
    For i = 1 To 81 'transfer unsolved puzzle to solution
        If Val(txPuz(i)) Then
            txSol(i) = txPuz(i)
            txSol(i).ForeColor = txPuz(i).ForeColor
            txSol(i).BackColor = txPuz(i).BackColor
          Else 'VAL(TXPUZ(I)) = FALSE/0
            txSol(i) = vbNullString
            txSol(i).ForeColor = IIf(lbHide.BorderStyle, txSol(0).BackColor, txSol(0).ForeColor)
            txSol(i).BackColor = txSol(0).BackColor
        End If
    Next i
    lbTime = vbNullString
    Screen.MousePointer = vbHourglass
    DoEvents

    'okay, let's try and find a solution...

    Rem Indent Begin
        If Complete(, False) Then 'no contradictions and all unique values completed

            'clear the tags
            For i = 0 To 80
                txPuz(i + 1).Tag = vbNullString
            Next i

            QueryPerformanceCounter StartTick 'for timimg

            'prepare for backtracking
            Done = False
            Timeout = False
            Success = 0
            LastHit = 0
            FirstFree = 0
            LastFree = 80

            'recursive backtracking
            Solve FindLeast

            QueryPerformanceCounter EndTick 'for timing

          Else 'NOT COMPLETE(,...
            EndTick = StartTick + Correction
        End If
    Rem Indent End

    '...and now display the results

    If Success <= Limit Then
        SetVisibleTo True
        Timing = (EndTick - StartTick - Correction) / HscFrequ * 1000000

        If Timeout Then
            lbSolved.Visible = False
            MsgBox "Could not find a solution within reasonable" & vbCrLf & "time -- and I'm pretty sure there is none.", vbQuestion, "Sorry..."
          Else 'Timeout = FALSE/0
            For i = 1 To 81
                k = Cells(i - 1).Value
                If k Then 'cell has a value
                    txSol(i) = k 'show solution
                  Else 'cell has no value; exit early 'K = FALSE/0
                    Exit For 'loop varying i
                End If
            Next i
            Solved = (i > 80)
            Vox.Skip "Sentence", 10
            If Solved Then
                lbNoSol.Visible = False
                Say lbSolved
              Else 'SOLVED = FALSE/0
                lbSolved.Visible = False
                Say lbNoSol
            End If
        End If
        DoEvents

        'calc timing
        If Animate Then
            Res = "N/A"
          Else 'ANIMATE = FALSE/0
            Select Case Timing
              Case Is >= 1000000
                Res = Round(Timing / 1000000, 2) & " Sec"
              Case Is >= 1000
                Res = Round(Timing / 1000, 2) & " mSec"
              Case Else
                Res = Int(Timing) & " µSec"
            End Select
        End If
        Res = Replace$(Res, DecSep, ".") & ". "

        'display timing
        lbTime = "Time" & IIf(Solved, " to solve:", ":") & vbCrLf & _
                 Res & vbCrLf & vbCrLf & _
                 Replace$(Format$(Success, "#,0 Steps"), ".", DecSep)
        lbTime_Click

        'reset puzzle
        For i = 1 To 81
            Cells(i - 1).Value = Val(txPuz(i))
        Next i
        Success = 0
        Enabled = True
        txPuz(1).SetFocus
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Complain(ByVal Index As Long, Text As String)

    Beeper 440, 30
    Say Text
    txPuz(Index) = vbNullString

End Sub

Private Function Complete(Optional ByVal Dependency As Long = 0, Optional ByVal Alarm As Boolean = True) As Boolean

  'complete any row, column or block with only one missing number
  'and look for contradictions

  Dim Missing   As Double

    Complete = True
    i = 0
    Success = 0
    Do
        With Cells(i)
            If .Value = 0 Then
                j = Groups(.ColumnNumber).PermitPattern And _
                    Groups(.RowNumber).PermitPattern And _
                    Groups(.BlockNumber).PermitPattern
                If j Then 'has at least one and possibly more possibilities for cells(i)
                    Missing = Log(j) / Log2 'which are they
                    If Missing = Int(Missing) Then 'just one value possible
                        .Value = Missing 'put that in cell
                        .Fixed = .Value 'and make cell fixed
                        With txPuz(i + 1)
                            Internal = True 'inhibit txPuz_Change processing
                            .Text = Missing 'put value into txPuz also
                            .ForeColor = txSol(0).ForeColor 'and color it as solved
                            .BackColor = txSol(0).BackColor
                            Internal = False
                            If Alarm Then
                                Beeper 2000, 10 'alert user
                                Alarm = False
                                .SetFocus
                                SetCursorPos Left / 15 + .Left + 31, Top / 15 + .Top + 75
                            End If
                            If Dependency Then 'this is a dependency from another cell during user input
                                .Tag = Dependency 'save that to undo if necessary
                            End If
                        End With 'TXPUZ(I
                        i = 0 'and restart scan
                        Success = 0
                    End If
                  Else 'no value possible for an empty cell - that's a contradiction 'J = FALSE/0
                    Complete = False 'so set to false
                    i = 80 'and exit
                End If
              Else 'NOT .VALUE...
                Success = Success + 1 'count solved cells
            End If
        End With 'CELLS(I)
        i = i + 1
    Loop Until i > 80

End Function

Private Function Contradiction(ByVal Dependency As Long) As Boolean

    Contradiction = Not Complete(Dependency)

End Function

Private Function ConvertForSpeech(ByVal Cellnumber As Long) As String

    ConvertForSpeech = ": " & Chr$((Cellnumber - 1) \ 9 + Asc("A")) & ", " & (Cellnumber + 8) Mod 9 + 1 & """"

End Function

Private Function FindLeast() As Long

  'returns the cell with the least number of possible values

  'this algorithm was inspired by Derio's sudoku solver. it is slower on easy puzzles but much faster on
  'harder ones. for example the "beast" has unbelievaby improved (from 70 Sec to 122 mSec), more than 600 times.
  'Thanks a lot, Derio!

  Dim Pattern   As Long
  Dim CurPermit As Long
  Dim MinPermit As Long
  Dim NewFirst  As Long
  Dim NewLast   As Long

    MinPermit = 10  'preset to max + 1

    Rem Indent Begin
        'equivalent to a killer heuristic (see wikipedia); the motivation is that a hit last time may still be good
        With Cells(LastHit)
            If .Value = 0 Then 'free cell

                'combined permit pattern
                Pattern = Groups(.ColumnNumber).PermitPattern And _
                          Groups(.RowNumber).PermitPattern And _
                          Groups(.BlockNumber).PermitPattern

                'count the permission bits
                CurPermit = 0
                For k = 1 To 9
                    If Pattern And Bits(k) Then
                        CurPermit = CurPermit + 1
                    End If
                Next k

                'save if less
                If CurPermit < MinPermit Then
                    MinPermit = CurPermit
                    PermissionBits = Pattern
                End If
            End If
        End With 'CELLS(LASTHIT)
    Rem Indent End

    '###########################################################################################################
    'un-commenting next line will nullify the effect of the above killer heuristic; try it to see the difference
    'MinPermit = 10
    '###########################################################################################################

    'preset range
    NewFirst = 81
    NewLast = -1

    'check all free cells to find the one with the least possibilities
    For j = FirstFree To LastFree
        With Cells(j)
            If .Value = 0 Then 'free cell

                'lower and upper limit for next time
                If j < NewFirst Then
                    NewFirst = j
                End If
                NewLast = j

                If MinPermit Then '... else skip this: MinPermit is already zero - it cannot get any lower
                    'combined permit pattern
                    Pattern = Groups(.ColumnNumber).PermitPattern And _
                              Groups(.RowNumber).PermitPattern And _
                              Groups(.BlockNumber).PermitPattern

                    'count the permission bits
                    CurPermit = 0
                    If Pattern Then '... else nothing to count
                        For k = 1 To 9
                            If Pattern And Bits(k) Then
                                CurPermit = CurPermit + 1
                                If CurPermit >= MinPermit Then 'early out; this will not be considered as good anyway
                                    Exit For 'loop varying k
                                End If
                            End If
                        Next k
                    End If

                    'save if less
                    If CurPermit < MinPermit Then
                        MinPermit = CurPermit
                        LastHit = j
                        PermissionBits = Pattern
                    End If
                End If
            End If

        End With 'CELLS(j)
    Next j

    'epilog
    FirstFree = NewFirst
    LastFree = NewLast
    FindLeast = LastHit
    Done = (NewLast = -1) 'ie no free cells

End Function

Private Sub Form_Initialize() ':) Line inserted by Formatter

    InitCommonControls ':) Line inserted by Formatter

End Sub ':) Line inserted by Formatter

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
      Case vbKeyEscape
        btClear.Value = True
        KeyCode = 0
      Case vbKeyF3
        btExit.Value = True
        KeyCode = 0
    End Select

End Sub

Private Sub Form_Load()

  Dim CmdParam As String

    Caption = Title

    'get user's name
    i = 128
    UserName = String$(i, 0)
    GetUserName UserName, i
    UserName = Left$(UserName, i + (Asc(Mid$(UserName, i, 1)) = 0))
    btExit.ToolTipText = "Good Bye, " & UserName

    'drop a form shadow
    SetClassLong hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW

    'prepare for timing measurements
    QueryPerformanceFrequency HscFrequ
    QueryPerformanceCounter StartTick
    QueryPerformanceCounter EndTick
    Correction = EndTick - StartTick

    'create all textboxes
    For i = 1 To 9
        For j = 1 To 9
            k = k + 1
            Load txPuz(k)
            With txPuz(k)
                .Move j * 33 - 10, i * 33 + 5, 32, 32
                .Visible = True
            End With 'TXPUZ(K)
            Load txSol(k)
            With txSol(k)
                .Move j * 33 + 405, i * 33 + 5, 32, 32
                .Visible = True
            End With 'txSol(K)
    Next j, i

    'instantiate and set up the classes
    For i = 0 To 26
        Set Groups(i) = New cGroup
    Next i

    For i = 0 To 80
        Set Cells(i) = New cCell
        With Cells(i)
            .Cellnumber = i
        End With 'CELLS(I)
    Next i

    With MII 'menu iten info
        .cbSize = Len(MII)
        .fMask = MIIM_STATE
        .fState = MFS_DEFAULT
        k = GetMenu(hWnd)
        For i = 0 To 2
            SetMenuItemInfo k, i, True, MII 'set to bold typeface
        Next i
    End With 'MII

    DecSep = Format$(0, "#.")
    Log2 = Log(2)
    For i = 1 To 9
        Bits(i) = 2 ^ i
    Next i
    mnuAnimSpeed_Click 0

    Set Vox = New SpVoice

    Select Case True
      Case InIDE
        If MsgBox("Compiled Code is a lot faster." & vbCrLf & vbCrLf & "Do you want to run me in the IDE anyway?", vbQuestion Or vbYesNo, Title & " [IDE]") = vbNo Then
            Unload Me
        End If
      Case App.PrevInstance
        MsgBox Title & " is already loaded.", vbExclamation, "Oops..."
        Unload Me
      Case Else
        Say "Hi " & UserName
        CmdParam = Replace$(Command$, """", "")
        If Len(CmdParam) Then
            LoadPuzzle CmdParam
            Caption = Title & " [" & Right$(CmdParam, Len(CmdParam) - InStrRev(CmdParam, "\")) & "]"
        End If
    End Select
    mnuVoice_Click

End Sub

Private Sub Form_Paint()

  'draws the frames

    With RECT
        For i = 9 To 10
            .L = sh(i).Left - 7
            .t = sh(i).Top - 30
            .r = sh(i).Left + sh(i).Width + 6
            .b = sh(i).Top + sh(i).Height + 6
            DrawEdge hDC, RECT, BDR_RAISED, BF_RECT
            .L = .L + 2
            .t = .t + 2
            .r = .r - 2
            .b = .b - 2
            DrawEdge hDC, RECT, BDR_RAISEDINNER, BF_RECT
        Next i
    End With 'RECT

End Sub

Private Sub Form_Unload(Cancel As Integer)

  'tidy up

    Success = Limit + 1
    For i = 0 To 80
        Set Cells(i) = Nothing
    Next i
    For i = 0 To 26
        Set Groups(i) = Nothing
    Next i
    If Not InIDE Then
        mnuVoice.Checked = True
        Say "Good bye " & UserName
        Say vbNullString, True
    End If
    Set Vox = Nothing
    Rem Mark Off Silent
    End
    Rem Mark On

End Sub

Private Sub imgMe_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    imgMe.BorderStyle = 1

End Sub

Private Sub imgMe_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)

    imgMe.BorderStyle = 0
    With App
        Select Case Button
          Case Is = vbLeftButton
            If ShellExecute(hWnd, vbNullString, "mailto:UMGEDV@Yahoo.com?subject=" & .ProductName & " V" & .Major & "." & .Minor & "." & .Revision & " &body=Hi Ulli,<br><br>[your message]<br><br>Best regards from " & UserName, vbNullString, .Path, SW_SHOWNORMAL) < SE_NO_ERROR Then
                MsgBox "Cannot send Mail from this System.", vbCritical, "Mail disabled/not installed"
            End If
          Case vbRightButton
            Load frmAbout
            With frmAbout
                .AppIcon(&HFFE0C0) = Icon
                .Title(&HD0D0FF) = Title
                .Version(&HF0F0A0) = "Version " & App.Major & "." & App.Minor & "." & App.Revision
                .Copyright(vbYellow) = App.LegalCopyright
                .Otherstuff1(&HF0FFF0) = "Sudoko Solver using Backtrace and Crosshatching"
                .Otherstuff2(&HE0D0D0) = "Least Possibility Cell Selection Algorithm" & vbCrLf & "by Derio"
                .Show vbModal, Me
            End With 'FRMABOUT
        End Select
    End With 'APP

End Sub

Private Function InIDE(Optional c As Boolean = False) As Boolean

  Static b  As Boolean

    b = c
    If b = False Then
        Debug.Assert InIDE(True)
    End If
    InIDE = b

End Function

Private Sub lbAnimate_Click()

    With lbAnimate
        .BorderStyle = 1 - .BorderStyle
        Animate = .BorderStyle
        mnuAnimSpeed_Click .BorderStyle * 2
    End With 'LBANIMATE
    SetVisibleTo False

End Sub

Private Sub lbHide_Click()

  'using a label to avoid difficulties with current focus

    With lbHide
        .BorderStyle = 1 - .BorderStyle
        mnuHide.Checked = .BorderStyle
        For i = 1 To 81
            If txSol(i).BackColor = txSol(0).BackColor Then
                If .BorderStyle Then
                    txSol(i).ForeColor = txSol(0).BackColor
                  Else '.BORDERSTYLE = FALSE/0
                    txSol(i).ForeColor = txSol(0).ForeColor
                End If
            End If
        Next i
        If .BorderStyle Then
            .ToolTipText = "Show Solution"
          Else '.BORDERSTYLE = FALSE/0
            .ToolTipText = "Hide Solution"
        End If
    End With 'LBHIDE

End Sub

Private Sub lbTime_Click()

  'needs some modifications to make it understandable

    Say Replace$(Replace$(lbTime, "µ", "micro "), "N/A", "not applicable."), True

End Sub

Private Sub LoadPuzzle(Filename As String)

  Dim hFile As Long
  Dim Inp   As String

    hFile = FreeFile
    Open Filename For Input As hFile
    Input #hFile, Inp
    Close hFile
    Internal = True
    For i = 1 To Len(Inp)
        If i > 81 Then
            Say "File layout is invalid"
            Exit For 'loop varying i
          Else 'NOT I...
            If IsNumeric(Mid$(Inp, i, 1)) Then
                With Cells(i - 1)
                    .Value = Mid$(Inp, i, 1)
                    .Fixed = True
                    txPuz(i) = .Value
                End With 'CELLS(I
              Else 'NOT ISNUMERIC(MID$(INP,...
                txPuz(i) = vbNullString
            End If
        End If
    Next i
    Internal = False

End Sub

Private Sub mnuAbout_Click()

    imgMe_Mouseup vbRightButton, 0, 0, 0

End Sub

Private Sub mnuAnimSpeed_Click(Index As Integer)

    For k = 0 To 3
        mnuAnimSpeed(k).Checked = (k = Index)
    Next k
    Select Case Index
      Case 1
        AnimDiv = 3
      Case 0, 2
        AnimDiv = 10
      Case 3
        AnimDiv = 30
    End Select
    lbAnimate.BorderStyle = Sgn(Index)
    Animate = Sgn(Index)
    lbAnimate.ToolTipText = AnimDiv & " Steps per Second"

End Sub

Private Sub mnuClear_Click()

    btClear.Value = True

End Sub

Private Sub mnuExit_Click()

    Visible = False
    DoEvents
    Unload Me

End Sub

Private Sub mnuHide_Click()

    mnuHide.Checked = Not mnuHide.Checked
    If mnuHide.Checked Then
        lbHide.BorderStyle = 0
      Else 'MNUHIDE.CHECKED = FALSE/0
        lbHide.BorderStyle = 1
    End If
    lbHide_Click

End Sub

Private Sub mnuLoad_Click()

    With CDl
        .InitDir = App.Path
        .DialogTitle = "Enter/Select file to load..."
        .Filename = vbNullString
        .DefaultExt = ".SKP"
        .Filter = "Sudoku Puzzle (*.SKP)|*.SKP|Sudoku Solution (*.SKS)|*.SKS"
        .Flags = cdlOFNPathMustExist Or cdlOFNLongNames
        On Error Resume Next
            .ShowOpen
            If Err = 0 Then
                btClear.Value = True
                LoadPuzzle .Filename
                Caption = Title & " [" & CDl.FileTitle & "]"
            End If
        On Error GoTo 0
    End With 'CDL

End Sub

Private Sub mnuPrintPuz_Click()

    PrintIt 1

End Sub

Private Sub mnuPrintSol_Click()

    If lbSolved.Visible = False And lbNoSol.Visible = False Then
        btSolve_Click
    End If
    PrintIt 2

End Sub

Private Sub mnuSavePuz_Click()

    SavePuzzle 0

End Sub

Private Sub mnuSaveSol_Click()

    SavePuzzle 1

End Sub

Private Sub mnuSendMail_Click()

    imgMe_Mouseup vbLeftButton, 0, 0, 0

End Sub

Private Sub mnuSolve_Click()

    btSolve.Value = True

End Sub

Private Sub mnuVoice_Click()

    mnuVoice.Checked = Not mnuVoice.Checked

End Sub

Private Sub Printer_DrawBox(Optional ByVal Small As Boolean = True)

    With RECT
        If Small Then
            .L = Printer.CurrentX - Printer.ScaleLeft - 60
            .t = Printer.CurrentY - 15
            .b = .t + 225
            .r = .L + 225
          Else 'SMALL = FALSE/0
            .L = Printer.CurrentX - Printer.ScaleLeft - 45
            .t = Printer.CurrentY + 240
            .b = .t + 2430
            .r = .L + 2445
            DrawEdge Printer.hDC, RECT, BDR_SUNKENOUTER, BF_RECT Or BF_MONO
            .L = .L - 15
            .t = .t - 15
            .b = .b + 15
            .r = .r + 15
        End If
    End With 'RECT
    DrawEdge Printer.hDC, RECT, BDR_SUNKENOUTER, BF_RECT Or BF_MONO

End Sub

Private Sub PrintIt(ByVal Num As Long)

  Dim t As Boolean

    With Printer
        .FontName = txSol(0).FontName
        .FontSize = txSol(0).FontSize
        .FontBold = True
        .ForeColor = vbBlack
        .ScaleMode = vbPixels
        .ScaleLeft = 0
        Printer.Print
        .CurrentX = (.ScaleWidth - .TextWidth(Caption)) / 2
        Printer.Print Caption

        .ScaleLeft = -1200

        For j = 1 To Num
            .FontSize = txSol(0).FontSize
            Printer.Print vbCrLf
            .ForeColor = vbBlack
            .FontBold = True
            If j = 1 Then
                Printer.Print ; "Puzzle";
              Else 'NOT J...
                Printer.Print vbCrLf
                If lbNoSol.Visible Then
                    Printer.Print ; "No Solution"
                    Exit For 'loop varying j
                End If
                Printer.Print ; "Solution";
            End If
            Printer.Print
            Printer_DrawBox False
            .FontSize = 22

            For i = 1 To 81
                If (i Mod 9) = 1 Then
                    Printer.Print
                    .CurrentY = .CurrentY + .TextWidth("1") / 2.4
                End If
                If (i Mod 27) = 1 Then
                    .CurrentY = .CurrentY + .TextWidth(" ")
                End If
                If (i Mod 3) = 1 Then
                    Printer.Print " ";
                End If
                Printer_DrawBox
                .FontBold = (txSol(i).BackColor = txSol(0).BackColor)
                .ForeColor = IIf(.FontBold, vbBlack, &H808080)
                If j = 1 Then
                    .FontBold = True
                    If Num = 1 Then
                        t = (txPuz(i).BackColor <> txPuz(0).BackColor) Or txPuz(i) = vbNullString
                      Else 'NOT NUM...
                        t = (txSol(i).BackColor = txSol(0).BackColor)
                    End If
                    If t Then
                        .ForeColor = vbWhite
                        Printer.Print "0   ";
                      Else 'T = FALSE/0
                        .ForeColor = vbBlack
                        Printer.Print txPuz(i); "   ";
                    End If
                  Else 'NOT J...
                    Printer.Print txSol(i); "   ";
                End If
            Next i
        Next j
        .ScaleLeft = 0
        .EndDoc
    End With 'PRINTER

End Sub

Private Sub SavePuzzle(ByVal Which As Long)

  Dim hFile As Long
  Dim Box   As TextBox

    With CDl
        .InitDir = App.Path
        .DialogTitle = "Enter/Select file to save..."
        If Which = 1 Then
            .DefaultExt = ".SKS"
            .Filter = "Sudoku Solution (*.SKS)|*.SKS"
          Else 'NOT WHICH...
            .DefaultExt = ".SKP"
            .Filter = "Sudoku Puzzle (*.SKP)|*.SKP"
        End If
        .Flags = cdlOFNPathMustExist Or cdlOFNLongNames Or cdlOFNOverwritePrompt
        On Error Resume Next
            CDl.ShowSave
            If Err = 0 Then
                hFile = FreeFile
                Open CDl.Filename For Output As hFile
                For k = 1 To 81
                    If Which = 1 Then
                        Set Box = txSol(k)
                      Else 'NOT WHICH...
                        Set Box = txPuz(k)
                    End If
                    If Len(Box) Then
                        Print #hFile, Box;
                      Else 'LEN(BOX) = FALSE/0
                        Print #hFile, "x";
                    End If
                Next k
                Close hFile
            End If
        On Error GoTo 0
        .Filename = vbNullString
    End With 'CDL

End Sub

Private Sub Say(Text As String, Optional Wait As Boolean = False)

    If mnuVoice.Checked Then
        If Wait Then
            Vox.WaitUntilDone 99999
        End If
        Vox.Speak Text & ".", SVSFlagsAsync
    End If

End Sub

Private Sub SetVisibleTo(Vis As Boolean)

    lbTime.Visible = Vis
    lbSolved.Visible = Vis
    lbNoSol.Visible = Vis
    shBg.Visible = Vis

End Sub

Private Sub ShowCell(ByVal Cellnumber As Long, ByVal Value As Long)

    QueryPerformanceCounter DelayStart
    With txSol(Cellnumber + 1)
        If Value Then
            .Text = Value
            DelayStart = DelayStart + HscFrequ / AnimDiv
          Else 'VALUE = FALSE/0
            .Text = vbNullString
            DelayStart = DelayStart + HscFrequ / AnimDiv / 8 'rollback is faster
        End If
        .Refresh
    End With 'TXSOL(CELLNUMBER
    Do
        DoEvents
        QueryPerformanceCounter DelayEnds
    Loop Until DelayEnds > DelayStart

End Sub

Private Sub Solve(ByVal Cellnumber As Long)

  'uses recursion / backtrack to solve the puzzle (max recursion depth will be 81)

  Dim Value     As Long 'local variables are recursive
  Dim PermBits  As Long

    If Not Done And PermissionBits Then
        PermBits = PermissionBits
        For Value = 1 To 9
            If PermBits And Bits(Value) Then 'possible value
                Cells(Cellnumber).Value = Value 'so fill it in
                If Animate Then
                    ShowCell Cellnumber, Value
                End If
                Success = Success + 1
                If (Success And Interval) = 0 Then 'at ~ 80000 steps/sec thats about 10 times per second
                    DoEvents
                End If
                If Success >= Limit Then
                    Done = True 'quit trying
                    Timeout = True 'and set timeout
                  Else 'NOT SUCCESS...
                    Solve FindLeast 'the current cell was successfully solved - recursion to next cell
                End If
                If Done Then 'no need to try further values
                    Exit For 'loop varying value
                End If
            End If
        Next Value 'try next value
        If Not Done Then
            Cells(Cellnumber).Value = 0    'no value for this cell was possible - reset cell
            If Cellnumber < FirstFree Then 'and adjust search range...
                FirstFree = Cellnumber
              ElseIf Cellnumber > LastFree Then 'NOT CELLNUMBER...
                LastFree = Cellnumber
            End If
            If Animate Then
                ShowCell Cellnumber, 0 'display empty cell
            End If
        End If
    End If

End Sub

Private Sub txPuz_Change(Index As Integer)

  'user has typed a number

  Dim iSav  As Long

    If Not Internal Then
        SetVisibleTo False
        k = Val(txPuz(Index))
        If Cells(Index - 1).Value <> k Then
            txPuz(Index).BackColor = txPuz(0).BackColor
            txPuz(Index).ForeColor = txPuz(0).ForeColor
            With Cells(Index - 1)
                If AllAgree(Index - 1, k) Then
                    .Value = k
                    If k Then
                        txPuz((Index Mod 81) + 1).SetFocus 'set focus on next textbox
                        SetCursorPos Left / 15 + txPuz((Index Mod 81) + 1).Left + 31, Top / 15 + txPuz((Index Mod 81) + 1).Top + 75
                        If Contradiction(Index) Then 'has put a value into at least one cell which has to be removed
                            For i = 0 To 80
                                If Val(txPuz(i + 1).Tag) = Index Then 'found cell
                                    txPuz(i + 1).Tag = vbNullString
                                    Cells(i).Value = 0
                                    Cells(i).Fixed = False
                                    iSav = i
                                End If
                            Next i
                            Complain Index, "Contradiction between cells" & ConvertForSpeech(Index) & " and " & ConvertForSpeech(iSav + 1)
                          Else 'CONTRADICTION(INDEX) = FALSE/0
                            .Fixed = True
                        End If
                      Else 'k = FALSE/0
                        .Fixed = False
                    End If
                  Else 'illegal input 'NOT ALLAGREE(INDEX...
                    Complain Index, "Duplicate value in cell" & ConvertForSpeech(Index)
                End If
            End With 'CELLS(INDEX
        End If
        btSolve.Value = (Success > 79) '80 solved cells is good enough to consider this puzzle as completely solved
    End If

End Sub

Private Sub txPuz_GotFocus(Index As Integer)

    txPuz(Index).SelStart = 0
    txPuz(Index).SelLength = 1

End Sub

Private Sub txPuz_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

  'keyboard navigation

    If KeyCode <> vbKeyDelete Then
        Select Case KeyCode
          Case vbKeyLeft
            txPuz((Index + 79) Mod 81 + 1).SetFocus
          Case vbKeyRight
            txPuz((Index Mod 81) + 1).SetFocus
          Case vbKeyUp
            txPuz((Index + 71) Mod 81 + 1).SetFocus
          Case vbKeyDown
            txPuz((Index + 8) Mod 81 + 1).SetFocus
        End Select
        KeyCode = 0
    End If

End Sub

Private Sub txPuz_KeyPress(Index As Integer, KeyAscii As Integer)

  'key press filtering

    If KeyAscii Then
        Select Case Chr$(KeyAscii)
          Case "1" To "9", "" 'Chr$(27) = Escape
          Case Else
            Complain Index, """" & Chr$(KeyAscii) & """ is illegal for cell" & ConvertForSpeech(Index)
            KeyAscii = 0 'kill key press
        End Select
    End If

End Sub

Private Sub txPuz_LostFocus(Index As Integer)

    LastFocus = Index

End Sub

Private Sub txSol_GotFocus(Index As Integer)

    txPuz(LastFocus).SetFocus 'put focus back to where it was

End Sub

Private Sub txSol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

  'hide and show

    If txSol(Index).BackColor = txSol(0).BackColor Then
        If txSol(Index).ForeColor = txSol(0).ForeColor Then
            If lbHide.BorderStyle Then
                txSol(Index).ForeColor = txSol(0).BackColor
            End If
          Else 'NOT TXSOL(INDEX).FORECOLOR...
            txSol(Index).ForeColor = txSol(0).ForeColor
        End If
    End If

End Sub

Private Sub txSol_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    With txSol(Index)
        If .ForeColor = .BackColor And .Text <> vbNullString Then
            .ToolTipText = " " & .Text & " "
          Else 'NOT .FORECOLOR...
            .ToolTipText = vbNullString
        End If
    End With 'TXSOL(INDEX)

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Feb-25 09:14)  Decl: 90  Code: 1084  Total: 1174 Lines
':) CommentOnly: 56 (4,8%)  Commented: 110 (9,4%)  Empty: 220 (18,7%)  Max Logic Depth: 9
':) Magic Number: 442372841
