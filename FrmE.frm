VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmE 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                                          Apariencia"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   Icon            =   "FrmE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4815
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   6120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontStrikeThru  =   -1  'True
      FontUnderLine   =   -1  'True
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   5655
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   600
         Picture         =   "FrmE.frx":000C
         ScaleHeight     =   2595
         ScaleWidth      =   3810
         TabIndex        =   15
         Top             =   1440
         Width           =   3810
         Begin VB.Image Image2 
            Height          =   200
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   5055
         End
         Begin VB.Image Image1 
            Height          =   2415
            Left            =   1080
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2775
         End
         Begin VB.Image Image4 
            Height          =   300
            Left            =   80
            Picture         =   "FrmE.frx":2049A
            Top             =   735
            Width           =   930
         End
         Begin VB.Label Lbl 
            Height          =   2475
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label Lbl 
            Height          =   2565
            Index           =   1
            Left            =   960
            TabIndex        =   16
            Top             =   120
            Width           =   2835
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmE.frx":2138C
         Left            =   360
         List            =   "FrmE.frx":213CF
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   2775
      End
      Begin VB.PictureBox Picture4 
         Height          =   2775
         Left            =   480
         ScaleHeight     =   2715
         ScaleWidth      =   4035
         TabIndex        =   30
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Un tema es un estilo de configuracion visual que hace mas agradable la utilizacion de Gurardian"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5175
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5790
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   10213
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Temas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Escritorio"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Apariencia"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "Cuadricula en modo Detalles"
         Height          =   615
         Left            =   240
         TabIndex        =   31
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Color de la letra en general"
         Height          =   555
         Left            =   3240
         TabIndex        =   27
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Color de la letra del elemeto"
         Height          =   555
         Left            =   3240
         TabIndex        =   26
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FrmE.frx":214D8
         Left            =   3240
         List            =   "FrmE.frx":214E5
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3120
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3720
         Width           =   1830
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Restaurar"
         Height          =   375
         Left            =   3720
         TabIndex        =   21
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cambiar Icono..."
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   2400
         Width           =   1335
      End
      Begin MSComctlLib.ListView Pcs 
         Height          =   1560
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2752
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hora de entrada"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Estado"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList IL1 
         Left            =   720
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmE.frx":21504
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmE.frx":28A06
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmE.frx":2ECA0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   5
         Left            =   2520
         TabIndex        =   28
         Top             =   4440
         Width           =   585
      End
      Begin VB.Label Lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   4
         Left            =   2520
         TabIndex        =   25
         Top             =   3600
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   5655
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2880
         Left            =   480
         Picture         =   "FrmE.frx":30232
         ScaleHeight     =   2880
         ScaleWidth      =   3840
         TabIndex        =   12
         Top             =   720
         Width           =   3840
         Begin VB.Image Image6 
            Height          =   225
            Left            =   0
            Stretch         =   -1  'True
            Top             =   195
            Width           =   3855
         End
         Begin VB.Image Image7 
            Height          =   120
            Left            =   0
            Picture         =   "FrmE.frx":54274
            Top             =   2760
            Width           =   3840
         End
         Begin VB.Image Image3 
            Height          =   300
            Left            =   0
            Picture         =   "FrmE.frx":55AB6
            Top             =   840
            Width           =   930
         End
         Begin VB.Image Image5 
            Height          =   2415
            Left            =   960
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Lbl 
            Height          =   2365
            Index           =   2
            Left            =   960
            TabIndex        =   14
            Top             =   360
            Width           =   2955
         End
         Begin VB.Label Lbl 
            Height          =   2385
            Index           =   3
            Left            =   0
            TabIndex        =   13
            Top             =   360
            Width           =   1035
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cambio el color del fondo..."
         Height          =   495
         Left            =   2280
         TabIndex        =   11
         Top             =   4320
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cambio de imagen..."
         Height          =   495
         Left            =   2280
         TabIndex        =   10
         Top             =   3840
         Width           =   2175
      End
      Begin VB.FileListBox File1 
         Height          =   1260
         Left            =   120
         Pattern         =   "*.jpg"
         TabIndex        =   5
         Top             =   3720
         Width           =   2055
      End
      Begin VB.FileListBox File2 
         Height          =   1260
         Left            =   3480
         Pattern         =   "*.bmp"
         TabIndex        =   4
         Top             =   3720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox Picture3 
         Height          =   3015
         Left            =   360
         ScaleHeight     =   2955
         ScaleWidth      =   4035
         TabIndex        =   29
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Personalisa aun mas guardian elige que estilo con que imagen te gustaria tener solo da clic, selecciona y listo"
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   120
         Width           =   4455
      End
   End
End
Attribute VB_Name = "FrmE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Imagen()
On Error Resume Next
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
    Pcs.ListItems.Clear
    Pcs.ListItems.Add , , " Pc-01" & vbCrLf & "Disponible", 1, 1
    Pcs.ListItems.Add , , " Pc-02" & vbCrLf & "CC", 2, 2
    Pcs.ListItems.Add , , " Pc-03" & vbCrLf & "Fuera de servicio", 3, 3
    Pcs.ListItems(1).ForeColor = ColorFD
    Pcs.ListItems(2).ForeColor = ColorFO
    Pcs.ListItems(3).ForeColor = ColorFF
End Sub

Private Sub Cmd1_Click()
On Error Resume Next
    ColorF = Pcs.ForeColor
    ColorFD = Pcs.ListItems(1).ForeColor
    ColorFO = Pcs.ListItems(2).ForeColor
    ColorFF = Pcs.ListItems(3).ForeColor
    Cuad = Check1.Value
    SFunt = Pcs.Font.Name
    Fonts = Pcs.Font.Size
    SaveSetting App.Title, "Estilo", "Barra", Sbgb
    SaveSetting App.Title, "Estilo", "Fondo", Sbg
    SaveSetting App.Title, "Estilo", "Color", Color
    SaveSetting App.Title, "Estilo", "ColorP", ColorP
    SaveSetting App.Title, "Estilo", "ColorF", ColorF
    SaveSetting App.Title, "Estilo", "ColorD", ColorFD
    SaveSetting App.Title, "Estilo", "ColorFO", ColorFO
    SaveSetting App.Title, "Estilo", "ColorFF", ColorFF
    SaveSetting App.Title, "Estilo", "Font", SFunt
    SaveSetting App.Title, "Estilo", "FontS", Fonts
    SaveSetting App.Title, "Estilo", "IcoSize", IcoSize
    SaveSetting App.Title, "Estilo", "Cuad", Cuad
    SaveSetting App.Title, "Estilo", "Vistas", Vistas
    FrmMain.Form_Load
    DoEvents
    Unload Me
End Sub

Private Sub Combo2_Click()
    Pcs.Font.Name = Combo2.Text
End Sub

Private Sub Combo3_Click()
    Pcs.Font.Size = Val(Combo3.Text)
End Sub

Private Sub Combo4_Click()
    If Combo4.Text = "16 X 16" Then
        IcoSize = 16
    ElseIf Combo4.Text = "32 X 32" Then
        IcoSize = 32
    Else
        IcoSize = 64
    End If
Imagen
End Sub


Private Sub Command1_Click()
    On Error Resume Next
        CommonDialog1.ShowOpen
        If CommonDialog1.FileName <> "" Then
            FileCopy CommonDialog1.FileName, App.Path & "\estilos\img\" & CommonDialog1.FileTitle
            Sbg = "\img\" & CommonDialog1.FileTitle
            Image1.Stretch = False
            Image1.Picture = LoadPicture(App.Path & "\estilos\" & Sbg)
            Image1.Stretch = True
            Image1.Width = Image1.Width / 2.8
            Image1.Height = Image1.Height / 2.8
            Image5.Stretch = False
            Image5.Picture = LoadPicture(App.Path & "\estilos\" & Sbg)
            Image5.Stretch = True
            Image5.Width = Image5.Width / 2.9
            Image5.Height = Image5.Height / 2.9
        End If
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    CommonDialog1.ShowColor
    If Err Then Exit Sub
    ColorP = CommonDialog1.Color
    Lbl(1).BackColor = ColorP
    Lbl(2).BackColor = ColorP
End Sub

Private Sub Combo1_Click()
    'File1.ListIndex = Combo1.ListIndex
    'File2.ListIndex = File1.ListIndex
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    On Error Resume Next
        CommonDialog1.ShowOpen
    If Err Then Exit Sub
    
    Select Case Pcs.SelectedItem.Index
    Case 1
        FileCopy CommonDialog1.FileName, App.Path & "\01.ico"
    Case 2
        FileCopy CommonDialog1.FileName, App.Path & "\02.ico"
    Case 3
        FileCopy CommonDialog1.FileName, App.Path & "\03.ico"
    End Select
    Imagen
End Sub

Private Sub Command5_Click()
    Select Case Pcs.SelectedItem.Index
    Case 1
        FileCopy App.Path & "\Estilos\01.ico", App.Path & "\01.ico"
    Case 2
        FileCopy App.Path & "\Estilos\02.ico", App.Path & "\02.ico"
    Case 3
        FileCopy App.Path & "\Estilos\03.ico", App.Path & "\03.ico"
    End Select
    Imagen
End Sub


Private Sub Command6_Click()
    On Error Resume Next
    CommonDialog1.ShowColor
    Pcs.SelectedItem.ForeColor = CommonDialog1.Color
    Lbl(4).BackColor = CommonDialog1.Color
End Sub

Private Sub Command7_Click()
    On Error Resume Next
    CommonDialog1.ShowColor
    If Err Then Exit Sub
    Pcs.ForeColor = CommonDialog1.Color
    Lbl(5).BackColor = CommonDialog1.Color
End Sub

Private Sub File1_Click(): Dim Fil As String
    Sbg = File1.FileName
    Image1.Stretch = False
    Image1.Picture = LoadPicture(App.Path & "\estilos" & "\" & Sbg)
    Image1.Stretch = True
    Image1.Width = Image1.Width / 2.8
    Image1.Height = Image1.Height / 2.8
    Image5.Stretch = False
    Image5.Picture = LoadPicture(App.Path & "\estilos" & "\" & Sbg)
    Image5.Stretch = True
    Image5.Width = Image5.Width / 2.9
    Image5.Height = Image5.Height / 2.9
    If File1.FileName = "bkg01.jpg" Then ColorP = RGB(152, 164, 206)
    If File1.FileName = "bkg02.jpg" Then ColorP = RGB(165, 150, 157)
    If File1.FileName = "bkg03.jpg" Then ColorP = RGB(223, 182, 154)
    If File1.FileName = "bkg04.jpg" Then ColorP = RGB(179, 214, 52)
    If File1.FileName = "bkg05.jpg" Then ColorP = RGB(255, 237, 104)
    If File1.FileName = "bkg06.jpg" Then ColorP = RGB(166, 161, 132)
    If File1.FileName = "bkg07.jpg" Then ColorP = RGB(108, 153, 218)
    If File1.FileName = "bkg08.jpg" Then ColorP = RGB(45, 104, 174)
    If File1.FileName = "bkg09.jpg" Then ColorP = RGB(184, 111, 9)
    If File1.FileName = "bkg10.jpg" Then ColorP = RGB(49, 66, 30)
    If File1.FileName = "bkg11.jpg" Then ColorP = RGB(254, 247, 166)
    If File1.FileName = "bkg12.jpg" Then ColorP = RGB(215, 223, 244)
    If File1.FileName = "bkg13.jpg" Then ColorP = RGB(48, 160, 0)
    If File1.FileName = "bkg14.jpg" Then ColorP = RGB(252, 88, 1)
    If File1.FileName = "bkg15.jpg" Then ColorP = RGB(11, 154, 248)
    If File1.FileName = "bkg16.jpg" Then ColorP = RGB(0, 0, 0)
    If File1.FileName = "bkg17.jpg" Then ColorP = RGB(251, 219, 222)
    If File1.FileName = "bkg18.jpg" Then ColorP = RGB(251, 228, 182)
    If File1.FileName = "bkg19.jpg" Then ColorP = RGB(243, 248, 252)
    If File1.FileName = "bkg20.jpg" Then ColorP = RGB(248, 230, 241)
    If File1.FileName = "bkg21.jpg" Then ColorP = RGB(255, 255, 255)
    Sbg = File1.FileName
    Lbl(1).BackColor = ColorP
    Lbl(2).BackColor = ColorP
End Sub

Private Sub File2_Click()
        Image2.Picture = LoadPicture(App.Path & "\estilos" & "\" & File2.FileName)
        Image6.Picture = LoadPicture(App.Path & "\estilos" & "\" & File2.FileName)
        If File2.FileName = "bkg01.bmp" Then Color = RGB(64, 58, 97)
        If File2.FileName = "bkg02.bmp" Then Color = RGB(71, 69, 72)
        If File2.FileName = "bkg03.bmp" Then Color = RGB(85, 38, 39)
        If File2.FileName = "bkg04.bmp" Then Color = RGB(0, 0, 0)
        If File2.FileName = "bkg05.bmp" Then Color = RGB(0, 0, 0)
        If File2.FileName = "bkg06.bmp" Then Color = RGB(230, 227, 210)
        If File2.FileName = "bkg07.bmp" Then Color = RGB(118, 160, 212)
        If File2.FileName = "bkg08.bmp" Then Color = RGB(37, 187, 235)
        If File2.FileName = "bkg09.bmp" Then Color = RGB(55, 29, 5)
        If File2.FileName = "bkg10.bmp" Then Color = RGB(86, 169, 39)
        If File2.FileName = "bkg11.bmp" Then Color = RGB(255, 144, 0)
        If File2.FileName = "bkg12.bmp" Then Color = RGB(155, 181, 208)
        If File2.FileName = "bkg13.bmp" Then Color = RGB(245, 150, 38)
        If File2.FileName = "bkg14.bmp" Then Color = RGB(58, 113, 202)
        If File2.FileName = "bkg15.bmp" Then Color = RGB(55, 113, 203)
        If File2.FileName = "bkg16.bmp" Then Color = RGB(251, 197, 56)
        If File2.FileName = "bkg17.bmp" Then Color = RGB(250, 208, 207)
        If File2.FileName = "bkg18.bmp" Then Color = RGB(253, 187, 147)
        If File2.FileName = "bkg19.bmp" Then Color = RGB(170, 203, 228)
        If File2.FileName = "bkg20.bmp" Then Color = RGB(234, 171, 214)
        If File2.FileName = "bkg21.bmp" Then Color = RGB(183, 200, 238)
        Sbgb = File2.FileName

        Lbl(0).BackColor = Color
        Lbl(3).BackColor = Color
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim X As Integer
    File1.Path = App.Path & "\estilos"
    File2.Path = App.Path & "\estilos"
    Image1.Stretch = False
    Image1.Picture = LoadPicture(App.Path & "\estilos" & "\" & Sbg)
    Image1.Stretch = True
    Image1.Width = Image1.Width / 2.8
    Image1.Height = Image1.Height / 2.8
    DoEvents
    Image5.Stretch = False
    Image5.Picture = LoadPicture(App.Path & "\estilos" & "\" & Sbg)
    Image5.Stretch = True
    Image5.Width = Image5.Width / 2.9
    Image5.Height = Image5.Height / 2.9
    Image2.Picture = LoadPicture(App.Path & "\estilos" & "\" & Sbgb)
    Image6.Picture = LoadPicture(App.Path & "\estilos" & "\" & Sbgb)
    Lbl(4).BackColor = Pcs.SelectedItem.ForeColor
    Lbl(0).BackColor = Color
    Lbl(3).BackColor = Color
    Lbl(1).BackColor = ColorP
    Lbl(2).BackColor = ColorP
    Lbl(5).BackColor = ColorF
    Pcs.Picture = LoadPicture(Sbg)
    Pcs.BackColor = ColorP
    Me.BackColor = Color
    Frame1.BackColor = Color
    Frame2.BackColor = Color
    Frame3.BackColor = Color
    Check1.BackColor = Color
    Check1.ForeColor = ColorF
    Check1.Value = Cuad
    Dim ctl As Control
    For Each ctl In Controls
        If TypeOf ctl Is Label Then _
            ctl.ForeColor = ColorF
    Next ctl
    Imagen
    For X = 0 To Screen.FontCount - 1
        Combo2.AddItem Screen.Fonts(X)
        Combo3.AddItem X + 8
    Next
 
    If IcoSize = 16 Then
        Combo4.ListIndex = 0
    ElseIf IcoSize = 32 Then
        Combo4.ListIndex = 1
    Else
        Combo4.ListIndex = 2
    End If
    Combo2.Text = SFunt
    Combo3.Text = Round(Fonts)
End Sub
Private Sub Pcs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Lbl(4).BackColor = Item.ForeColor
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = 1 Then
        Frame1.ZOrder (0)
    ElseIf TabStrip1.SelectedItem.Index = 2 Then
        Frame2.ZOrder (0)
    ElseIf TabStrip1.SelectedItem.Index = 3 Then
        Frame3.ZOrder (0)

    Else
    End If
End Sub
