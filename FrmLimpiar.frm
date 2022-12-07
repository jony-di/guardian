VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1725
   Icon            =   "FrmLimpiar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    DeleteSetting App.Title, "Precios", "Red"
    DeleteSetting App.Title, "Estilo", "ColorP"
    DeleteSetting App.Title, "Estilo", "ColorF"
    DeleteSetting App.Title, "Estilo", "ColorD"
    DeleteSetting App.Title, "Estilo", "ColorFO"
    DeleteSetting App.Title, "Estilo", "ColorFF"
    DeleteSetting App.Title, "Estilo", "Fondo"
    DeleteSetting App.Title, "Estilo", "Barra"
    DeleteSetting App.Title, "Estilo", "Emp"
    DeleteSetting App.Title, "Estilo", "Font"
    DeleteSetting App.Title, "Estilo", "Cuad"
    DeleteSetting App.Title, "Estilo", "Color"
    DeleteSetting App.Title, "Estilo", "FontS"
    DeleteSetting App.Title, "Estilo", "Vistas"
    DeleteSetting App.Title, "Estilo", "IcoSize"
    End
End Sub
