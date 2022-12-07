VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "FrmPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   1080
      ScaleHeight     =   2235
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton Command4 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   3600
         TabIndex        =   8
         Top             =   1560
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2520
         TabIndex        =   7
         Top             =   1560
         Width           =   1035
      End
      Begin VB.TextBox Txt 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Txt 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   570
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1005
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "ImageList2"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drecripcion"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   120
         Width           =   555
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar Gurpo"
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar Nuevo grupo"
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   15130843
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   68
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":000C
            Key             =   "Unidad Zip"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":0C5E
            Key             =   "Botella"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":18B0
            Key             =   "Cerveza"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2502
            Key             =   "Camara Digital"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":3154
            Key             =   "Camara Fotografica"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":3DA6
            Key             =   "Cd con caja"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":49F8
            Key             =   "Drive CD-R"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":564A
            Key             =   "CD-R"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":629C
            Key             =   "CD-ROM"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":6EEE
            Key             =   "CD-RW"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":7B40
            Key             =   "Coca-cola"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":8792
            Key             =   "Coca-cola lata"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":93E4
            Key             =   "Cafe chico"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":A036
            Key             =   "Cafe"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":AC88
            Key             =   "Cafe con galletas"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":B8DA
            Key             =   "Monitor"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":BDDC
            Key             =   "nidad Zip"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":CA2E
            Key             =   "CPU"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":D680
            Key             =   "CPU2"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":E2D2
            Key             =   "Galleta"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":EF24
            Key             =   "Galletas"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":FB76
            Key             =   "Galletas2"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":107C8
            Key             =   "Targeta Expandible"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1141A
            Key             =   "Disquette"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1206C
            Key             =   "Disquette2"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":12CBE
            Key             =   "Cafe grande"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":13210
            Key             =   "Refresco"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":13E62
            Key             =   "Baso"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":14AB4
            Key             =   "DVD-R"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":15706
            Key             =   "DVD-Rom"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":16358
            Key             =   "DVD-RW"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":16FAA
            Key             =   "Amburgesa"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":17BFC
            Key             =   "Audifonos"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1884E
            Key             =   "Hot-Dog"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":194A0
            Key             =   "Helado"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1A0F2
            Key             =   "Helado2"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1AD44
            Key             =   "Helado3"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1B996
            Key             =   "Paleta"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1C5E8
            Key             =   "Bevida"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1D23A
            Key             =   "Plato"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1DE8C
            Key             =   "Teclado"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1EADE
            Key             =   "Copa de helado"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":1F730
            Key             =   "Monitor2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":20382
            Key             =   "Teclado2"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":20FD4
            Key             =   "Cel3"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":21C26
            Key             =   "Camara"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":22878
            Key             =   "Palm"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":234CA
            Key             =   "Pluma"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2411C
            Key             =   "Pepsi"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":24D6E
            Key             =   "Celular1"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":259C0
            Key             =   "Telefono"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":26612
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":27264
            Key             =   "Pizza grande"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":27EB6
            Key             =   "Pizza Mediana"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":28B08
            Key             =   "Revanada1"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2975A
            Key             =   "Pizza chica"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2A3AC
            Key             =   "Revanada"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2AFFE
            Key             =   "Impreciones a color"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2BC50
            Key             =   "Impreciones B/N"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2C8A2
            Key             =   "Zanwich"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2D4F4
            Key             =   "Sumarino"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2E146
            Key             =   "Escaner"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2ED98
            Key             =   "Escaner2"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":2F9EA
            Key             =   "Cel"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":3063C
            Key             =   "Tacos"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":3128E
            Key             =   "Taza de cafe"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":31EE0
            Key             =   "Fax"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPro.frx":32B32
            Key             =   "Copa"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Picture1.Visible = True
    Txt(0).SetFocus
End Sub

Private Sub Command3_Click()
    Picture1.Visible = False
End Sub

Private Sub Command4_Click()
        Picture1.Visible = False
End Sub

Private Sub Form_Load()
Dim X As Integer
        Me.BackColor = Color
        Picture1.BackColor = Color
        For X = 1 To ImageList2.ListImages.Count
            ImageCombo1.ComboItems.Add , , ImageList2.ListImages(X).Key, X, X
        Next
End Sub
