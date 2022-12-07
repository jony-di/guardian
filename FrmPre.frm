VERSION 5.00
Begin VB.Form FrmPre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                          Precion x min"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   Icon            =   "FrmPre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Productos"
      Height          =   375
      Left            =   600
      TabIndex        =   30
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Redondiar a la hora"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   12
      Left            =   720
      TabIndex        =   12
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   11
      Left            =   2880
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   10
      Left            =   2880
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   9
      Left            =   2880
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   8
      Left            =   2880
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   7
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   6
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriva la cantidad que el Guardian cobrara por cada Hora y minuto Transcurriodos"
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
      Height          =   615
      Left            =   360
      TabIndex        =   29
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   12
      Left            =   360
      TabIndex        =   28
      Top             =   3840
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "55"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   11
      Left            =   2520
      TabIndex        =   27
      Top             =   3240
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   10
      Left            =   2520
      TabIndex        =   26
      Top             =   2760
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   9
      Left            =   2520
      TabIndex        =   25
      Top             =   2280
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   8
      Left            =   2520
      TabIndex        =   24
      Top             =   1800
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   7
      Left            =   2520
      TabIndex        =   23
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   6
      Left            =   2520
      TabIndex        =   22
      Top             =   840
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   5
      Left            =   360
      TabIndex        =   21
      Top             =   3240
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   4
      Left            =   360
      TabIndex        =   20
      Top             =   2760
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   360
      TabIndex        =   19
      Top             =   2280
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   360
      TabIndex        =   18
      Top             =   1800
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Top             =   1320
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   840
      Width           =   150
   End
End
Attribute VB_Name = "FrmPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Dim X As Byte
    For X = 0 To Txt.Count - 2
        If Check1.Value = 1 Then
            Txt(X).Enabled = False
        Else
            Txt(X).Enabled = True
        End If
    Next
End Sub

Private Sub Command1_Click()
    Dim Tabla As New ADODB.Recordset
    Tabla.Open "Select *  From cxm", Base, adOpenDynamic, adLockOptimistic
    Dim X As Byte
    For X = 0 To Txt.Count - 1
        Tabla(1).Value = Txt(X)
        Tabla.Update
        Tabla.MoveNext
    Next
    Red = Check1.Value

    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    FrmPro.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim ctl As Control
    For Each ctl In Controls
        If TypeOf ctl Is Label Then _
            ctl.ForeColor = ColorF
    Next ctl
    Check1.ForeColor = ColorF
    Me.BackColor = Color
    Check1.BackColor = Me.BackColor
    Dim Tabla As New ADODB.Recordset
    Tabla.Open "Select *  From cxm", Base, adOpenDynamic, adLockOptimistic
    Dim X As Byte
    For X = 0 To Txt.Count - 1
        Txt(X) = Format(Tabla(1).Value, "$ #,##0.00")
        Tabla.MoveNext
    Next
    Check1.Value = Red
    DoEvents
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Txt(Index).SelStart = 0
    Txt(Index).SelLength = Len(Txt(Index).Text)
End Sub

Private Sub Txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 46 Then
    ElseIf KeyAscii <> 8 Then
        If Not IsNumeric(Chr(KeyAscii)) Then
            Beep
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Txt(Index).Text = Format(Txt(Index).Text, "$ #,##0.00")
End Sub
