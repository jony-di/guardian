VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00C17C20&
   Caption         =   "OLD - TIMES -  JDSystem [ Guardian ]"
   ClientHeight    =   9540
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15165
   ClipControls    =   0   'False
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   9540
   ScaleWidth      =   15165
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   9960
      Top             =   6480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ListView Pcs 
      Height          =   3720
      Left            =   6360
      TabIndex        =   0
      Top             =   2160
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6562
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   -2147483643
      Appearance      =   0
      OLEDragMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Hora de entrada"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Avisar en"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3120
      Top             =   6600
   End
   Begin MSComctlLib.ImageCombo IC 
      Height          =   330
      Left            =   11160
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "ImageList1"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11640
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IL2 
      Left            =   4560
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1276
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   5400
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   7920
      Picture         =   "FrmMain.frx":20C8
      ToolTipText     =   "Cerar Caja"
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   8760
      Picture         =   "FrmMain.frx":23D2
      Stretch         =   -1  'True
      ToolTipText     =   "Lista de trabajos"
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   465
      Left            =   9960
      Picture         =   "FrmMain.frx":40CC
      Stretch         =   -1  'True
      ToolTipText     =   "Unir cuentas"
      Top             =   75
      Width           =   435
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   10560
      Picture         =   "FrmMain.frx":49D1
      Stretch         =   -1  'True
      ToolTipText     =   "Lista de Precios"
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   6840
      Picture         =   "FrmMain.frx":BEC3
      Stretch         =   -1  'True
      ToolTipText     =   "Estilo de Guardian"
      Top             =   45
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   7440
      Picture         =   "FrmMain.frx":1214D
      Stretch         =   -1  'True
      ToolTipText     =   "Vistas"
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3255
   End
   Begin VB.Menu MnuSol 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu MnuI 
         Caption         =   "Iconos"
      End
      Begin VB.Menu MnuL 
         Caption         =   "Lista"
      End
      Begin VB.Menu MnuD 
         Caption         =   "Detalles"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim It As Boolean
Dim BPcs As Byte
Dim i  As Integer
Sub CPcs(): Dim Pcss
    Pcss = 150
    Pcs.ListItems.Clear
    For BPcs = 1 To Pcss
        If BPcs < 10 Then
            Pcs.ListItems.Add , "Pc-00" & BPcs, "Pc-00" & BPcs, 1, 1
        ElseIf BPcs < 100 Then
            Pcs.ListItems.Add , "Pc-" & BPcs, "Pc-0" & BPcs, 1, 1
        Else
            Pcs.ListItems.Add , "Pc-" & BPcs, "Pc-" & BPcs, 1, 1
        End If
        Pcs.ListItems(BPcs).SubItems(4) = "Disponible"
        Pcs.ListItems(BPcs).ForeColor = ColorFD
    Next BPcs
    
    Recupera
End Sub
Private Sub Recupera()
On Error Resume Next
    Dim X As Byte
    Dim Tabla As New ADODB.Recordset
    Tabla.Open "Select * From Recupera", base, adOpenDynamic, adLockBatchOptimistic
    Do While Not Tabla.EOF
        X = Tabla(0).Value
        Pcs.ListItems(X).Icon = 2
        Pcs.ListItems(X).SmallIcon = 2
        Pcs.ListItems(X).SubItems(1) = Tabla(1).Value
        Pcs.ListItems(X).SubItems(2) = Tabla(2).Value
        Pcs.ListItems(X).SubItems(3) = Tabla(3).Value
        Pcs.ListItems(X).SubItems(4) = "Ocupada"
        Pcs.ListItems(X).ForeColor = ColorFO
        If Pcs.View = lvwIcon Then Pcs.ListItems(X).Text = Pcs.ListItems(X).Key & vbCrLf & Pcs.ListItems(X).SubItems(1)
        Tabla.MoveNext
    Loop
End Sub

Private Sub Combo1_Click()
    Dim Tabla As New ADODB.Recordset
    Tabla.Open
End Sub

Sub Form_Load()
Attribute Form_Load.VB_UserMemId = -552
    Dim SEmp As String
     
    On Error Resume Next
    Color = Val(GetSetting(App.Title, "Estilo", "Color"))
    ColorF = Val(GetSetting(App.Title, "Estilo", "ColorF"))
    ColorP = Val((GetSetting(App.Title, "Estilo", "ColorP")))
    ColorFD = Val(GetSetting(App.Title, "Estilo", "ColorD"))
    ColorFO = Val(GetSetting(App.Title, "Estilo", "ColorFO"))
    ColorFF = Val(GetSetting(App.Title, "Estilo", "ColorFF"))
    
    Sbg = GetSetting(App.Title, "Estilo", "Fondo")
    Sbgb = GetSetting(App.Title, "Estilo", "Barra")
    SEmp = GetSetting(App.Title, "Estilo", "Emp")
    Red = Val(GetSetting(App.Title, "Precios", "Red"))
    SFunt = GetSetting(App.Title, "Estilo", "Font")
    Cuad = Val(GetSetting(App.Title, "Estilo", "Cuad"))
    Fonts = Val(GetSetting(App.Title, "Estilo", "FontS"))
    Vistas = Val(GetSetting(App.Title, "Estilo", "Vistas"))
    IcoSize = Val(GetSetting(App.Title, "Estilo", "IcoSize"))
    
    If IcoSize = 0 Then IcoSize = 32
    
    If SEmp = "" Then
        SEmp = InputBox("El nombre de su negosio", App.Title)
        SaveSetting App.Title, "Estilo", "Emp", SEmp
    End If
    
    Me.Caption = SEmp & " usa JD-System [Guardian]"
    Me.BackColor = Color
    
    Dim ctl As Control
    
    For Each ctl In Controls
        If TypeOf ctl Is Label Then ctl.ForeColor = ColorF
    Next ctl
     
    ImageList1.ListImages.Add , , LoadPicture(App.Path & "\01.ico")
    ImageList1.ListImages.Add , , LoadPicture(App.Path & "\02.ico")
    ImageList1.ListImages.Add , , LoadPicture(App.Path & "\03.ico")
    IC.ComboItems.Clear
    IC.ComboItems.Add , , "Todas", 1, 1
    IC.ComboItems.Add , , "Solo diponibles", 2, 2
    IC.ComboItems.Add , , "Solo Ocupadas", 3, 3
    IC.ComboItems.Add , , "Solo Fuera de servicio", 4, 4
    IC.ComboItems(1).Selected = True
    
    Set Pcs.Icons = Nothing
    Set Pcs.SmallIcons = Nothing
    IL1.ListImages.Clear
    
    IL1.ImageHeight = IcoSize
    IL1.ImageWidth = IcoSize
    IL1.ListImages.Add , , LoadPicture(App.Path & "\01.ico")
    IL1.ListImages.Add , , LoadPicture(App.Path & "\02.ico")
    IL1.ListImages.Add , , LoadPicture(App.Path & "\03.ico")
    Set Pcs.Icons = IL1
    Set Pcs.SmallIcons = IL1
    
    Pcs.ColumnHeaders(3).Width = Len(Pcs.ColumnHeaders(3).Text) * 120
    Image2.Picture = LoadPicture(App.Path & "\Estilos\" & Sbgb)
    Pcs.Picture = LoadPicture(App.Path & "\Estilos\" & Sbg)
    Pcs.GridLines = CBool(Cuad)
    Pcs.BackColor = ColorP
    Pcs.Font.Name = SFunt
    Pcs.Font.Size = Fonts
    Pcs.View = Vistas
    
    
    If Err Then FrmE.Show vbModal
    base.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Recupera.mdb;Persist Security Info=False"
    If Err.Number = -2147467259 Then MsgBox Err.Number & vbCrLf & Err.Description, vbCritical: End
    DoEvents
    CPcs
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    Pcs.Width = (Me.Width - Pcs.Left) - 100
    Pcs.Height = (Me.Height - Pcs.Top) - 1000
    
    DoEvents
    Image2.Width = Me.Width - 100
    IC.Left = Me.Width - IC.Width - 200
    
    Image8.Left = Me.Width - IC.Width - Image3.Width - 1000
    Image5.Left = Image8.Left - Image8.Width - 200
    Image3.Left = Image5.Left - Image5.Width - 1000
    Image4.Left = Image3.Left - Image3.Width - 200
    Image6.Left = Image4.Left - Image4.Width - 1000
    Image7.Left = Image6.Left - Image6.Width - 200
    

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Precios", "Red", Red
    SaveSetting App.Title, "Estilo", "Vistas", Vistas
    End
End Sub

Private Sub IC_Click()
    Dim X As Byte
    CPcs
    If IC.Text = "Todas" Then Exit Sub
    
    If IC.Text = "Solo diponibles" Then
        X = 1
        Do While Not X = BPcs
            If Pcs.ListItems(X).Icon = 2 Then
                Pcs.ListItems.Remove X
                 BPcs = BPcs - 1
                 X = X - 1
                 
            ElseIf Pcs.ListItems(X).Icon = 3 Then
                Pcs.ListItems.Remove X
                BPcs = BPcs - 1
                X = X - 1
            End If
        X = X + 1
        Loop
        
    ElseIf IC.Text = "Solo Ocupadas" Then
        X = 1
        Do While Not X = BPcs
            If Pcs.ListItems(X).Icon = 1 Then
                Pcs.ListItems.Remove X
                 BPcs = BPcs - 1
                 X = X - 1
            ElseIf Pcs.ListItems(X).Icon = 3 Then
                Pcs.ListItems.Remove X
                BPcs = BPcs - 1
                X = X - 1
            End If
        X = X + 1
        Loop
        
    ElseIf IC.Text = "Solo Fuera de servicio" Then
        X = 1
        Do While Not X = BPcs
            If Pcs.ListItems(X).Icon = 1 Then
                Pcs.ListItems.Remove X
                 BPcs = BPcs - 1
                 X = X - 1
            ElseIf Pcs.ListItems(X).Icon = 2 Then
                Pcs.ListItems.Remove X
                BPcs = BPcs - 1
                X = X - 1
            End If
        X = X + 1
        Loop
    End If
    Pcs.SetFocus
End Sub


Private Sub Image3_Click()
    MnuSol.Enabled = True
    PopupMenu MnuSol
    MnuSol.Enabled = False
End Sub

Private Sub Image4_Click()
    FrmE.Show vbModal
End Sub

Private Sub Image5_Click()
    FrmPre.Show vbModal
End Sub

Private Sub Image8_Click()
    Dim Tabla As New ADODB.Recordset
    Tabla.Open "select * from Recupera", base, adOpenDynamic, adLockBatchOptimistic
    If Tabla.EOF Then
        Frmc.Show vbModal
    Else
        MsgBox "Hay maquina(s) en renta", vbCritical
    End If
End Sub


Private Sub MnuD_Click()
    Dim X As Byte
    Pcs.View = lvwReport
    For X = 1 To BPcs - 1
        Pcs.ListItems(X).Text = Pcs.ListItems(X).Key
    Next
    Vistas = Pcs.View
End Sub

Private Sub MnuI_Click()
    Dim X As Byte
    Pcs.View = lvwIcon
    For X = 1 To BPcs - 1
        Pcs.ListItems(X).Text = Pcs.ListItems(X).Key & vbCrLf & Pcs.ListItems(X).SubItems(1)
    Next
    Vistas = Pcs.View
End Sub

Private Sub MnuL_Click()
    Dim X As Byte
    Pcs.View = lvwList
    For X = 1 To BPcs - 1
        Pcs.ListItems(X).Text = Pcs.ListItems(X).Key
    Next
    Vistas = Pcs.View
End Sub

Private Sub Pcs_ColumnClick(ByVal Colum As MSComctlLib.ColumnHeader)
        Pcs.SortKey = Colum.Index - 1
    If Pcs.SortOrder = lvwDescending Then
        Pcs.SortOrder = lvwAscending
    Else
        Pcs.SortOrder = lvwDescending
    End If
End Sub

Private Sub Pcs_ItemClick(ByVal Item As MSComctlLib.ListItem)
It = True
End Sub

Private Sub Pcs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ESE
End Sub

Private Sub Pcs_LostFocus()
    i = 0
End Sub

Private Sub Pcs_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If It = True Then
        ESE
        It = False
    End If
End Sub

Private Sub ESE()
    If Pcs.SelectedItem.Icon = 1 Then
        Dim Us As String
        Dim Tabla As New ADODB.Recordset, Nu As Integer
        Us = InputBox("Nombre del Cliente", "Guardia ", "Cliente Conocido")
        If Us = "" Then Exit Sub
        If Us = "Cliente Conocido" Then
            Us = "CC"
        End If
        Nu = Val(Right(Pcs.SelectedItem.Key, Len(Pcs.SelectedItem.Key) - 3))
        Pcs.SelectedItem.Icon = 2
        Pcs.SelectedItem.SmallIcon = 2
        Pcs.SelectedItem.SubItems(1) = UCase(Us)
        Pcs.SelectedItem.SubItems(2) = Time
        Pcs.SelectedItem.SubItems(4) = "Ocupada"
        Pcs.SelectedItem.ForeColor = ColorFO
        If Pcs.View = lvwIcon Then Pcs.ListItems(Nu).Text = Pcs.ListItems(Nu).Key & vbCrLf & Pcs.ListItems(Nu).SubItems(1)
        Tabla.Open "select * from recupera where pc=" & Nu, base, adOpenDynamic, adLockOptimistic
        If Tabla.EOF = False Then
            Tabla.Close
            MsgBox "Error de almasenado Pc ocupada", vbCritical
            CPcs
            Exit Sub
        End If
        Tabla.AddNew
        Tabla(0).Value = Nu
        Tabla(1).Value = Pcs.SelectedItem.SubItems(1)
        Tabla(2).Value = Pcs.SelectedItem.SubItems(2)
        Tabla.Update
        IC_Click
    ElseIf Pcs.SelectedItem.Icon = 2 Then
        FrmS.Show vbModal
    End If
    
End Sub

Private Sub Timer1_Timer()
    Dim X As Integer
    Dim Tabla As New ADODB.Recordset
    Dim He, H, ht, av As String
    For X = 1 To Pcs.ListItems.Count
        If Pcs.ListItems(X).SubItems(3) <> "" Then
                    He = Format(Pcs.ListItems(X).SubItems(2), "hh:mm")
                    H = Format(Time, "hh:mm")
                    ht = TimeValue(He) - TimeValue(H)
                    ht = Format(ht, "hh:mm")
                    H = Format(Pcs.ListItems(X).SubItems(3), "hh:mm")
                    
                    If TimeValue(ht) >= TimeValue(H) Then
                        av = MsgBox("La " & Pcs.ListItems(X).Key & " a terminado su tiempo " & vbCrLf & "¿Desea darle salida?", vbInformation + vbYesNo)
                        If av = vbNo Then
                            MsgBox "La alarma a sido quitada", vbExclamation
                            Pcs.ListItems(X).SubItems(3) = ""
                            Dim Nu As Integer
                            
                                Nu = X
                                Tabla.Open "select * from recupera where pc=" & Nu, base, adOpenDynamic, adLockOptimistic
                                Tabla(3).Value = ""
                                Tabla.Update
                        Else
                            Pcs.ListItems(3).SubItems(3) = ""
                            Pcs.ListItems(X).Selected = True
                            FrmS.Show vbModal
                        End If
                    End If
        
        End If
    Next X
End Sub
