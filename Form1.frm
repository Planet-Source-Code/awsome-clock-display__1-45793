VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   1650
   ClientTop       =   2055
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   960
      Top             =   1500
   End
   Begin VB.Timer t22 
      Interval        =   10
      Left            =   2880
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1320
      Top             =   1920
   End
   Begin VB.Label ttt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2400
      TabIndex        =   0
      Top             =   2760
      Width           =   60
   End
   Begin VB.Menu mc 
      Caption         =   "&Clock"
      Begin VB.Menu dfdf 
         Caption         =   "&Settings..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const pi As Double = 3.14
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private PX As Double, PY As Double, RX As Double, RY As Double
Private CX As Long
Private CY As Long
Private o As Double
Private oo As Double
Private ooo As Double
Private COSX((360)) As Double
Private SINX((360)) As Double
Private ff As Boolean


Private Sub dfdf_Click()
Form2.Show
End Sub

Private Sub Form_Load()
SPACING = 10
F_NCOL = RGB(100, 100, 100)
F_HCOL = RGB(255, 0, 0)
CNTR = 70
Me.ForeColor = RGB(100, 100, 100)
For i = 0 To (360)
    SINX(i) = Sin(i / (360) * 2 * pi)
    COSX(i) = Cos(i / (360) * 2 * pi)
Next
o = 0
oo = 0
CX = Me.ScaleWidth / 2
CY = Me.ScaleHeight / 2

SS = 1
MS = 0.5
HS = 0.1

DrawClock
'set second
DrawClock
End Sub

Public Sub DrawClock()
Me.Show
Me.Cls

PX = CNTR
PY = CNTR
'calc radians


 'is 12 oclock
'RX = PX * SINX(o) - PY * COSX(o)
'RY = PY * SINX(o) + PX * COSX(o)
'Me.CurrentX = RX + CX
'Me.CurrentY = RY + CY
'Me.Print 12

'first ring
For i = 0 To 59
o = o - (((360) / 60)) 'next hour
If o < 0 Then o = (360 - Abs(o))
RX = PX * SINX(o) - PY * COSX(o)
RY = PY * SINX(o) + PX * COSX(o)

If i = DateTime.Second(DateTime.Now) Then
Me.ForeColor = F_HCOL
Me.FontBold = True

TextOut Me.hdc, RX + CX, RY + CY, Str(i), Len(Str(i))
Me.FontBold = False

Else
Me.ForeColor = F_NCOL
TextOut Me.hdc, RX + CX, RY + CY, Str(i), Len(Str(i))
End If
Next

PX = CNTR + SPACING
PY = CNTR + SPACING
For i = 0 To 59
oo = oo - (((360) / 60)) 'next hour
If oo < 0 Then oo = (360 - Abs(oo))
RX = PX * SINX(oo) - PY * COSX(oo)
RY = PY * SINX(oo) + PX * COSX(oo)

If i = DateTime.Minute(DateTime.Now) Then
Me.ForeColor = F_HCOL
Me.FontBold = True

TextOut Me.hdc, RX + CX, RY + CY, Str(i), Len(Str(i))
Me.FontBold = False

Else
Dim css As Long
css = (DateTime.Second(DateTime.Now) * 2) + 50
Me.ForeColor = F_NCOL
TextOut Me.hdc, RX + CX, RY + CY, Str(i), Len(Str(i))
End If
Next


PX = CNTR + SPACING * 2
PY = CNTR + SPACING * 2
For i = 0 To 59
ooo = ooo - (((360) / 60)) 'next hour
If ooo < 0 Then ooo = (360 - Abs(ooo))
RX = PX * SINX(ooo) - PY * COSX(ooo)
RY = PY * SINX(ooo) + PX * COSX(ooo)

If i = DateTime.Hour(DateTime.Now) Then
Me.ForeColor = F_HCOL
Me.FontBold = True

TextOut Me.hdc, RX + CX, RY + CY, Str(i), Len(Str(i))
Me.FontBold = False

Else
Me.ForeColor = F_NCOL
TextOut Me.hdc, RX + CX, RY + CY, Str(i), Len(Str(i))
End If
Next

'Me.Line (CX, 0)-(CX, Me.ScaleHeight)
'Me.Line (0, CY)-(Me.ScaleWidth, CY)
End Sub

Private Sub Form_Resize()
CX = Me.ScaleWidth / 2
CY = Me.ScaleHeight / 2
DrawClock
End Sub

Private Sub HScroll1_Scroll()
o = HScroll1.Value
DrawClock
Me.Caption = HScroll1.Value

End Sub

Private Sub t22_Timer()
If PULSE Then
If ff Then
    If SPACING < PMM Then
        SPACING = SPACING + 0.1
        Exit Sub
    Else
        SPACING = PMM
        ff = False
    End If
Else
    If SPACING > PM Then
        SPACING = SPACING - 0.1
        Exit Sub
    Else
        SPACING = PM
        ff = True
    End If
End If
End If
End Sub

Private Sub Timer1_Timer()
o = o + SS
oo = oo + MS
ooo = ooo + HS
DrawClock
'Me.Caption = oo & " " & o & " " & ooo
End Sub

Private Sub Timer2_Timer()
If SC = 1 Then
    ttt.Visible = True
    
    ttt.Caption = DateTime.Time
    ttt.Move (Me.ScaleWidth / 2) - (ttt.Width / 3), (Me.ScaleHeight / 2)
Else
    ttt.Visible = False
    
End If


'DrawClock
End Sub
