VERSION 5.00
Begin VB.Form FrmM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mover Pc"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2790
   Icon            =   "FrmM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd 
      Caption         =   "Mover"
      Height          =   735
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmM.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mover la "
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "FrmM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Click()
    Dim Tabla As New ADODB.Recordset, Nu As Integer
    Dim C, H
    Nu = Val(Right(FrmMain.Pcs.SelectedItem.Key, Len(FrmMain.Pcs.SelectedItem.Key) - 3))
    Tabla.Open "select * from recupera where pc=" & Nu, Base, adOpenDynamic, adLockOptimistic
    C = Tabla(1)
    H = Tabla(2)
    Tabla.Close
    Tabla.Open "delete from recupera where pc=" & Nu, Base, adOpenDynamic, adLockOptimistic
    Nu = Val(Right(Combo1.Text, Len(Combo1.Text) - 3))
    Tabla.Open "select * from recupera", Base, adOpenDynamic, adLockOptimistic
    Tabla.AddNew
    Tabla(0) = Nu
    Tabla(1) = C
    Tabla(2) = H
    Tabla.Update
    FrmMain.Form_Load
    Me.Hide
    Unload FrmS
    Unload Me
End Sub

Private Sub Form_Load()
    Me.BackColor = Color
    Dim x As Integer
    Dim ctl As Control
    For Each ctl In Controls
        If TypeOf ctl Is Label Then _
            ctl.ForeColor = ColorF
    Next ctl
    Cmd.BackColor = ColorP
    Label1.Caption = Label1.Caption & FrmMain.Pcs.SelectedItem.Key
    Combo1.Clear
    For x = 1 To FrmMain.Pcs.ListItems.Count
        If FrmMain.Pcs.ListItems(x).Icon = 1 Then
            Combo1.AddItem FrmMain.Pcs.ListItems(x).Key
        End If
    Next

End Sub
