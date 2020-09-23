VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Call Initialize_Region"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   599
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All coded by myself (thank you, thank you)"
      ForeColor       =   &H00C0C0FF&
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   3150
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "form (not on text) to drag form."
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   4320
      TabIndex        =   8
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*************************************"
      ForeColor       =   &H00C0C0FF&
      Height          =   180
      Left            =   4320
      TabIndex        =   7
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*************************************"
      ForeColor       =   &H00C0C0FF&
      Height          =   180
      Left            =   4320
      TabIndex        =   6
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hold Leftmousebutton somewhere in the "
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   2850
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ohh, me is EvDeursen by the way."
      ForeColor       =   &H00C0C0FF&
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check my other projects, if you like!!!"
      ForeColor       =   &H00C0C0FF&
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   2925
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "It's a combo of SetTimer and SetWindowRegion"
      ForeColor       =   &H00C0C0FF&
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   3300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yepp , here is some new amazing stuff."
      ForeColor       =   &H00C0C0FF&
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2850
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call SetTimer(Me.hwnd, bEventID, 100, AddressOf TimerProc)
End Sub

Private Sub Form_Paint()
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    Me.Move 0, Me.Top, Screen.Width, Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call KillTimer(Me.hwnd, bEventID)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    Call ReleaseCapture
    lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

