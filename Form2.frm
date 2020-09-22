VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Settings"
   ClientHeight    =   4920
   ClientLeft      =   3840
   ClientTop       =   2205
   ClientWidth     =   4455
   LinkTopic       =   "Form2"
   ScaleHeight     =   4920
   ScaleWidth      =   4455
   Begin VB.Frame Frame3 
      Caption         =   "Advanced"
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   4215
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2220
         TabIndex        =   20
         Text            =   "0"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2220
         TabIndex        =   19
         Text            =   "0"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2220
         TabIndex        =   18
         Text            =   "0"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Second Turn Rate:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1140
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Minute Turn Rate:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   780
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Hour Turn Rate:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   420
         Width           =   2415
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Auto-Redraw"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   4620
      Width           =   3435
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show clock in center"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   3435
   End
   Begin MSComDlg.CommonDialog Com1 
      Left            =   1080
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3540
      TabIndex        =   8
      Top             =   4380
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Layout"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   1500
      Width           =   4215
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Left            =   1080
         Max             =   300
         Min             =   2
         TabIndex        =   12
         Top             =   480
         Value           =   2
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pulse"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   780
         Width           =   3675
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Left            =   1080
         Max             =   300
         Min             =   2
         TabIndex        =   7
         Top             =   300
         Value           =   2
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Center:"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Spacing:"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4215
      Begin VB.PictureBox Picture2 
         Height          =   255
         Left            =   1740
         ScaleHeight     =   195
         ScaleWidth      =   2055
         TabIndex        =   4
         Top             =   840
         Width           =   2115
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   1740
         ScaleHeight     =   195
         ScaleWidth      =   2055
         TabIndex        =   3
         Top             =   420
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Active Numbers:"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Normal Numbers:"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   420
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
PULSE = IIf(Check1.Value = 1, True, False)
PM = SPACING - 2
PMM = SPACING + 10
End Sub

Private Sub Check2_Click()
SC = Check2.Value

End Sub

Private Sub Check3_Click()
Form1.AutoRedraw = IIf(Check3.Value = 1, True, False)
End Sub

Private Sub Command1_Click()
HS = Text1.Text
MS = Text2.Text
SS = Text3.Text
Unload Me

End Sub

Private Sub Form_Load()
Form1.Timer1.Enabled = False

Picture1.BackColor = F_NCOL
Picture2.BackColor = F_HCOL

Check1.Value = IIf(PULSE = True, 1, 0)
Text1.Text = HS
Text2.Text = MS
Text3.Text = SS
Check3.Value = IIf(Form1.AutoRedraw = True, 1, 0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Timer1.Enabled = True
End Sub

Private Sub HScroll1_Scroll()
SPACING = HScroll1.Value
Form1.DrawClock
End Sub

Private Sub HScroll2_Change()
CNTR = HScroll2.Value
Form1.DrawClock
End Sub

Private Sub HScroll2_Scroll()
CNTR = HScroll2.Value
Form1.DrawClock
End Sub

Private Sub Picture1_Click()
Com1.ShowColor
Picture1.BackColor = Com1.Color
F_NCOL = Com1.Color
Form1.DrawClock
End Sub

Private Sub Picture2_Click()
Com1.ShowColor
Picture2.BackColor = Com1.Color
F_HCOL = Com1.Color
Form1.DrawClock
End Sub

