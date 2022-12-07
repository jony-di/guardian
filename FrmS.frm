VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                                                  Pc"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9780
   Icon            =   "FrmS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cmb1 
      Height          =   315
      ItemData        =   "FrmS.frx":000C
      Left            =   2280
      List            =   "FrmS.frx":003D
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6588
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Preducto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Precio"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avizar a "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   11
      Top             =   2160
      Width           =   885
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   240
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver maquina"
      Height          =   195
      Index           =   2
      Left            =   4920
      TabIndex        =   9
      Top             =   2520
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cambio hora"
      Height          =   195
      Index           =   1
      Left            =   4920
      TabIndex        =   8
      Top             =   1680
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cambio de Pc"
      Height          =   195
      Index           =   0
      Left            =   4800
      TabIndex        =   7
      Top             =   840
      Width           =   990
   End
   Begin VB.Label LblTp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ 00.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total a pagar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   1620
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   5160
      Picture         =   "FrmS.frx":00A5
      ToolTipText     =   "Ver Pc"
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label LblHa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label LblHe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   690
   End
   Begin VB.Label LblT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label LblPc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   5160
      Picture         =   "FrmS.frx":0B0F
      ToolTipText     =   "Cambiar Hra de Entrada"
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image Image10 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   5160
      Picture         =   "FrmS.frx":1181
      Stretch         =   -1  'True
      ToolTipText     =   "Mover a..."
      Top             =   360
      Width           =   225
   End
   Begin VB.Image Image11 
      Height          =   720
      Left            =   5160
      Picture         =   "FrmS.frx":1B83
      ToolTipText     =   "Mover a..."
      Top             =   315
      Width           =   720
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   3015
      Left            =   4680
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Change()

End Sub

Private Sub Cmb1_Click()
    Dim ha As String
    Dim Nu As Integer
    Dim Tabla As New ADODB.Recordset
    ha = Cmb1.Text
    If ha = "" Then Exit Sub
    On Error Resume Next
    ha = TimeValue(ha)
    If Err = 13 Then MsgBox "Formato de hora no valido", vbCritical: Exit Sub
    FrmMain.Pcs.SelectedItem.SubItems(3) = Format(Cmb1.Text, "hh:mm")
    Nu = Val(Right(FrmMain.Pcs.SelectedItem.Key, Len(FrmMain.Pcs.SelectedItem.Key) - 3))
    Tabla.Open "select * from recupera where pc=" & Nu, Base, adOpenDynamic, adLockOptimistic
    Tabla(3).Value = Format(Cmb1.Text, "hh:mm")
    Tabla.Update
End Sub

Private Sub Cmb1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
    If KeyAscii = 13 Then
        Cmb1_Click
    ElseIf KeyAscii = 58 Then     ' permite introducir : para la hora
    ElseIf KeyAscii <> 8 Then   ' El 8 es la tecla de borrar (backspace)
    If Not IsNumeric(Chr(KeyAscii)) Then
        ' ... se desecha esa tecla y se avisa de que no es correcta
            Beep
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim He, ha, T As String
    
    He = FrmMain.Pcs.SelectedItem.SubItems(2)
    ha = Time
    T = TimeValue(ha) - TimeValue(He)
    LblPc.Caption = FrmMain.Pcs.SelectedItem.Key
    LblHe.Caption = "Hora de entrada: " & He
    LblHa = "Hora de salida:" & Time
    LblT = "Total de tiempo: " & Format(T, "hh:mm")
    Cmb1.Text = FrmMain.Pcs.SelectedItem.SubItems(3)
On Error Resume Next
Dim ctl As Control
    For Each ctl In Controls
        If TypeOf ctl Is Label Then _
            ctl.ForeColor = ColorF
    Next ctl
    Me.BackColor = Color
    LblTp.ForeColor = vbRed
LblTp = Format(Cobrar, "$ #,##0.00")
End Sub

Private Sub Image1_Click()
    Dim NHe As String
    Dim He, ha, T As String
    Dim Tabla As New ADODB.Recordset
    Dim Nu As Integer
    Nu = Val(Right(FrmMain.Pcs.SelectedItem.Key, Len(FrmMain.Pcs.SelectedItem.Key) - 3))
    NHe = InputBox("Escriva la nueva hora", "Cambiar Hora de entrada", FrmMain.Pcs.SelectedItem.SubItems(2))
    If NHe = "" Then Exit Sub
    On Error Resume Next
    NHe = TimeValue(NHe)
    If Err = 13 Then MsgBox "Formato de hora no valido", vbCritical: Exit Sub
    Tabla.Open "select * from recupera where pc=" & Nu, Base, adOpenDynamic, adLockOptimistic
    Tabla(2).Value = NHe
    Tabla.Update
    
    
    FrmMain.Pcs.SelectedItem.SubItems(2) = NHe
    He = NHe
    ha = Time
    T = TimeValue(ha) - TimeValue(He)
    LblPc.Caption = FrmMain.Pcs.SelectedItem.Key
    LblHe.Caption = "Hora de entrada: " & He
    LblHa = "Hora de salida:" & Time
    LblT = "Total de tiempo: " & Format(T, "hh:mm")

End Sub

Private Sub Image10_Click()
    FrmM.Show vbModal
End Sub

Private Sub Image11_Click()
    FrmM.Show vbModal
End Sub

Function Cobrar() As Currency


Dim HH, MM As Integer
Dim Ph, Pm As Currency
Dim T As String
Dim Tabla As New ADODB.Recordset
T = Right(LblT.Caption, 5)
MM = Val(Right(T, 2))
HH = Val(Left(T, 2))
Ph = FrmPre.Txt(12).Text
Unload FrmPro
Tabla.Open "Select * From cxm", Base, adOpenDynamic, adLockBatchOptimistic
Do While Not Tabla.EOF
    If Tabla(0).Value > MM Then Exit Do
    Pm = Tabla(1).Value
    Tabla.MoveNext
Loop

If Red = 1 Then
    If MM >= 30 Then HH = HH + 1
    Cobrar = HH * Ph
End If
Cobrar = HH * Ph
Cobrar = Cobrar + Pm
End Function

Private Sub Image2_Click()
    MsgBox "Solo es pocible si la maquina tiene" & vbCrLf & "GuardianCliente By JDS" & vbCrLf & "¿Lo tiene?", vbInformation + vbYesNo
End Sub

