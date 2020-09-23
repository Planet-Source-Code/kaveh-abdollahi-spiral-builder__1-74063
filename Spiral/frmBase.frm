VERSION 5.00
Begin VB.Form frmBase 
   Caption         =   "Spiral"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   16950
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLo 
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   34
      Top             =   4320
      Width           =   495
   End
   Begin VB.CheckBox chkAutoP 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Auto +"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Auto Draw Next"
      Top             =   2115
      Width           =   585
   End
   Begin VB.ComboBox cmbSteps 
      Height          =   315
      ItemData        =   "frmBase.frx":0000
      Left            =   1320
      List            =   "frmBase.frx":0010
      TabIndex        =   32
      Text            =   "44"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2400
      Top             =   480
   End
   Begin VB.CheckBox chkPp 
      Caption         =   "++"
      Height          =   195
      Left            =   280
      TabIndex        =   31
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtCircleSize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   29
      Text            =   "1"
      ToolTipText     =   "double click > + 1"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdPrimes 
      Caption         =   "Primes"
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAddD 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   25
      Text            =   "U"
      Top             =   8040
      Width           =   735
   End
   Begin VB.CommandButton cmdMakeDir 
      Caption         =   "New Directoy"
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CheckBox chkAutoPn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "--"
      Height          =   225
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   435
   End
   Begin VB.CommandButton cmdReset_1 
      BackColor       =   &H0000FF00&
      Caption         =   "R"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Reset to one"
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdShot 
      BackColor       =   &H0080FF80&
      Caption         =   "Shot"
      Height          =   375
      Left            =   1365
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2820
      Width           =   975
   End
   Begin VB.CommandButton cmdBdraw 
      BackColor       =   &H0080FFFF&
      Caption         =   "Draw"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Draw Next"
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox chkCls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkCircle 
      Caption         =   "Circle"
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   360
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox chkLine 
      Caption         =   "Line"
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   120
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txtMaxShot 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Text            =   "3000"
      ToolTipText     =   "Maximum Shot"
      Top             =   3225
      Width           =   495
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H000000FF&
      Caption         =   "Delete All"
      Height          =   285
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H0000FFFF&
      Caption         =   "R"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Reset to last"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtStep 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "710"
      ToolTipText     =   "double click > x 2"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "D:\9v_Visu\Sp\"
      Top             =   8400
      Width           =   4935
   End
   Begin VB.CheckBox chkAutoShot 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Auto Shot"
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Auto Save Image"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "0.1"
      ToolTipText     =   "double click > + 1"
      Top             =   2115
      Width           =   650
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   120
   End
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   280
      TabIndex        =   3
      Text            =   "1"
      ToolTipText     =   "double click > +1"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtPoints 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   280
      TabIndex        =   2
      Text            =   "100"
      ToolTipText     =   "double click > + 10"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Draw"
      Top             =   240
      Width           =   520
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   2400
      ScaleHeight     =   7290
      ScaleMode       =   0  'User
      ScaleWidth      =   14400
      TabIndex        =   0
      Top             =   120
      Width           =   14400
   End
   Begin VB.PictureBox picKaveh 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   13800
      Picture         =   "frmBase.frx":002A
      ScaleHeight     =   285
      ScaleWidth      =   2970
      TabIndex        =   15
      Top             =   7920
      Width           =   2970
   End
   Begin VB.PictureBox picPatt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   2400
      Picture         =   "frmBase.frx":397D
      ScaleHeight     =   4500
      ScaleWidth      =   1125
      TabIndex        =   18
      Top             =   3600
      Width           =   1155
   End
   Begin VB.TextBox txtPathOr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Text            =   "D:\9v_Visu\Sp\"
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "On All Boxes Double Click && Ctrl + Double Click Available"
      Height          =   375
      Left            =   5280
      TabIndex        =   33
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Circle Size ÷"
      Height          =   255
      Left            =   720
      TabIndex        =   30
      Top             =   3735
      Width           =   1095
   End
   Begin VB.Label lblPrCount 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Interval Sec"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      Height          =   255
      Left            =   1170
      TabIndex        =   10
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label lblShot 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
End
Attribute VB_Name = "frmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaveCount As Long, patt(1 To 2, 1 To 300) As Long

Private Sub chkAutoP_Click()
    If chkAutoP Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
End Sub


Private Sub chkCircle_Click()
    cmdDraw_Click
End Sub

Private Sub chkLine_Click()
    cmdDraw_Click
End Sub

Private Sub chkPp_Click()
    If chkPp Then
        Timer2.Enabled = True
    Else
        Timer2.Enabled = False
    End If
End Sub


Private Sub cmbSteps_Click()
    cmbSteps_Validate True
End Sub

Private Sub cmbSteps_Validate(Cancel As Boolean)
    txtStep = cmbSteps.Text
End Sub

Private Sub cmdBdraw_Click()
    If chkAutoPn Then
        txtNumber = Val(txtNumber) - Val(txtStep)
    Else
        txtNumber = Val(txtNumber) + Val(txtStep)
    End If
    
    cmdDraw_Click
    If Val(txtNumber) * Val(txtPoints) > 2000000000 Then cmdReset_Click: txtNumber = txtNumber + 1
    
    If chkAutoShot Then cmdShot_Click

End Sub

Private Sub cmdDel_Click()
On Error Resume Next
    Kill txtPath & "\Spp*" & ".*"
    lblShot.Caption = "0"
End Sub

Private Sub cmdDraw_Click()
Dim x As Single, y As Single, N As Long, z2 As Single, z3 As Single, T As Long, Tn As Long
Dim oX As Single, oY As Single, tm As Long, s As Long, c As Long, Si As Single
Dim x2 As Single, y2 As Single, Dis As Long, PO As Long
Dim L25(1 To 300) As Long
On Error Resume Next

pic1.ForeColor = vbBlack
pic1.DrawWidth = 1
If chkCls Then pic1.Cls
Si = Val(txtCircleSize)
N = Val(txtNumber)
PO = Val(txtPoints)
z2 = 240 / PO
z3 = Log(txtPoints)
MoveToEx pic1.hdc, 480, 270, pot
oX = 480: oY = 27
    For T = 1 To Val(txtPoints)
        pic1.ForeColor = vbBlack ' patt(2, T * z2)
        pic1.DrawWidth = 1
        x = Sin(T * N) * T * z2 + 480
        y = Cos(T * N) * T * z2 + 270
        If chkLine Then LineTo pic1.hdc, x, y
        pic1.ForeColor = patt(2, T * z2) ' vbRed
        If chkLo Then
            pic1.DrawWidth = 2 + T ^ 0.5 * Log(T) * z2 * Log(PO - T + 1) / z3 / Si
        Else
            pic1.DrawWidth = 2 + T ^ 0.5 * Log(T) * z2 * 3 / z3 / Si
        End If
        MoveToEx pic1.hdc, x, y, pot
        If chkCircle Then LineTo pic1.hdc, x, y
        oX = x: oY = y
        
    Next T
    
    If chkCircle Then pic1.ForeColor = vbRed:  MoveToEx pic1.hdc, x, y, pot:  LineTo pic1.hdc, x, y
    pic1.ForeColor = vbBlue
    If chkCircle Then pic1.DrawWidth = pic1.DrawWidth / 2: MoveToEx pic1.hdc, x, y, pot:     LineTo pic1.hdc, x, y
    pic1.DrawWidth = 5: pic1.ForeColor = vbRed: MoveToEx pic1.hdc, x, y, pot:      LineTo pic1.hdc, x, y

    pic1.ForeColor = vbRed
    pic1.FontBold = True
    pic1.FontSize = 16
    pic1.Print " Spiral"
    pic1.FontSize = 13
    pic1.ForeColor = vbBlack
    pic1.Print " Start Number: " & txtNumber Mod txtStep & "  Points: " & txtPoints
    pic1.Print " "
    pic1.Print " N=" & Format$(txtNumber, "###,###,###,###,###,###,###0") & "   N+=" & txtStep
    
    pic1.FontSize = 11
    pic1.Print "  T = 1 To " & txtPoints & " (Points)"
    pic1.Print "      X = Sin(T x N) x T"
    pic1.Print "      Y = Cos(T x " & N & ") x T"
    pic1.Print
    pic1.Print " Last 20 Distance"
    For T = Val(txtPoints) To Val(txtPoints) - 20 Step -1
        x = Sin(T * N) * T * z2 + 480
        y = Cos(T * N) * T * z2 + 270
        If T <> txtPoints Then pic1.Print GetDistance(oX, oY, x, y)
        oX = x: oY = y
     Next

    BitBlt pic1.hdc, 960 - 205, 540 - 22, 198, 19, picKaveh.hdc, 0, 0, vbSrcCopy
    BitBlt pic1.hdc, 890, 30, 50, 300, picPatt.hdc, 0, 0, vbSrcCopy
    
    If Val(lblShot) >= Val(txtMaxShot) Then chkAutoShot.Value = 0: chkAutoP.Value = 0: lblShot = 0
    
End Sub

Private Sub cmdMakeDir_Click()
Dim s As String
On Error Resume Next
    s = "D:\9v_Visu\"
    MkDir s
    s = "D:\9v_Visu\Sp\"
    MkDir s
    s = txtPathOr & "Sp_" & txtAddD & "_stN-" & txtNumber & "_sumP-" & txtPoints & "_NPlus-" & txtStep & "\"
    MkDir s
    txtPath = s
End Sub

Private Sub cmdPrimes_Click()
    cmdPrimes.Enabled = False: DoEvents
    PrimeBase
    cmdPrimes.Enabled = True
End Sub

Private Sub cmdReset_1_Click()
    txtNumber = 1
End Sub

Private Sub cmdReset_Click()
    txtNumber = txtNumber Mod txtStep
End Sub

Private Sub cmdShot_Click()
Dim s As String
    s = txtPath & "Spp-" & Format$(SaveCount, "0#######0") & ".jpg"
    SaveJpeg s, 80, pic1
    lblShot = Val(lblShot) + 1: SaveCount = SaveCount + 1
    SaveSetting "spi", "kspi", "SaveCount", SaveCount

End Sub

Private Sub Form_Activate()
Dim y As Long
picPatt.Refresh
    For y = 1 To 300
        patt(1, y) = GetPixel(picPatt.hdc, 10, y)
        patt(2, y) = GetPixel(picPatt.hdc, 10, 300 - y)
    Next y
cmdDraw_Click
End Sub

Private Sub Form_Load()
Dim y As Integer
    txtPoints = GetSetting("spi", "kspi", "txtPoints", "100")
    txtNumber = GetSetting("spi", "kspi", "txtNumber", "1")
    txtInterval = GetSetting("spi", "kspi", "txtInterval", "1")
    txtStep = GetSetting("spi", "kspi", "txtStep", "1")
    txtPath = GetSetting("spi", "kspi", "txtPath", "D:\")
    txtAddD = GetSetting("spi", "kspi", "txtAddD", "U")
    cmbSteps = GetSetting("spi", "kspi", "cmbSteps", "44")
    
    SaveCount = GetSetting("spi", "kspi", "SaveCount", "1")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "spi", "kspi", "txtPoints", txtPoints
    SaveSetting "spi", "kspi", "txtNumber", txtNumber
    SaveSetting "spi", "kspi", "txtInterval", txtInterval
    SaveSetting "spi", "kspi", "txtStep", txtStep
    SaveSetting "spi", "kspi", "txtPath", txtPath
    SaveSetting "spi", "kspi", "txtAddD", txtAddD
    SaveSetting "spi", "kspi", "cmbSteps", cmbSteps
    
    SaveSetting "spi", "kspi", "SaveCount", SaveCount
    
    End
End Sub

Private Sub Timer1_Timer()
    If chkAutoPn Then
        txtNumber = Val(txtNumber) - Val(txtStep)
    Else
        txtNumber = Val(txtNumber) + Val(txtStep)
    End If
    cmdDraw_Click
    If Val(txtNumber) * Val(txtPoints) > 2000000000 Then cmdReset_Click: txtNumber = txtNumber + 1
    
    If chkAutoShot Then cmdShot_Click
End Sub

Private Sub Timer2_Timer()
    txtPoints = Val(txtPoints) + 1
    cmdDraw_Click
End Sub

Private Sub txtCircleSize_DblClick()
    If CTRLDown Then
        If Val(txtCircleSize) > 1 Then txtCircleSize = Val(txtCircleSize) - 1
    Else
        txtCircleSize = Val(txtCircleSize) + 1
    End If
End Sub

Private Sub txtPoints_Change()
    cmdDraw_Click
End Sub

Private Sub txtPoints_DblClick()
    If CTRLDown Then
        If Val(txtPoints) > 10 Then txtPoints = Val(txtPoints) - 10
    Else
        txtPoints = Val(txtPoints) + 10
    End If
End Sub

Private Sub txtNumber_Change()
    cmdDraw_Click
End Sub

Private Sub txtNumber_DblClick()
    If CTRLDown Then
        If Val(txtNumber) > 1 Then txtNumber = Val(txtNumber) - 1
    Else
        txtNumber = Val(txtNumber) + 1
    End If
End Sub

Private Sub txtInterval_Change()
    If Val(txtInterval) > 60 Then txtInterval = 1
    Timer1.Interval = Val(txtInterval) * 1000
End Sub

Private Sub txtInterval_DblClick()
    If CTRLDown Then
        If Val(txtInterval) > 0.005 Then txtInterval = Val(txtInterval) / 2
    Else
        If Val(txtInterval) < 30 Then txtInterval = Val(txtInterval) * 2
    End If
    
'    txtInterval = Val(txtInterval) + 1
    If Val(txtInterval) > 60 Then txtInterval = 1
    Timer1.Interval = Val(txtInterval) * 1000
End Sub

Private Sub txtStep_DblClick()
    If CTRLDown Then
        If Val(txtStep) > 1 Then txtStep = Val(txtStep) \ 2
    Else
        txtStep = Val(txtStep) * 2
    End If
End Sub

Private Function ShiftDown() As Boolean
Dim RetVal As Long
    RetVal = GetAsyncKeyState(16) 'SHIFT key
    If (RetVal And 32768) <> 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
End Function
Private Function CTRLDown() As Boolean
Dim RetVal As Long
    RetVal = GetAsyncKeyState(17) 'CTRL key
    If (RetVal And 32768) <> 0 Then
        CTRLDown = True
    Else
        CTRLDown = False
    End If
End Function
Private Function ALTDown() As Boolean
Dim RetVal As Long
    RetVal = GetAsyncKeyState(18) 'ALT key
    If (RetVal And 32768) <> 0 Then
        ALTDown = True
    Else
        ALTDown = False
    End If
End Function


