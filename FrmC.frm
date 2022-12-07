VERSION 5.00
Begin VB.Form FrmC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cerrara Caja"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2445
   Icon            =   "FrmC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Terminar dia"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Imprimir "
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Cerrar Caja"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Frmc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNewXlsFile  As String
Dim sXlsTemplate As String

Private Const SW_SHOWNORMAL = 1

Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const SE_ERR_ACCESSDENIED = 5
Private Const SE_ERR_OOM = 8
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_SHARE = 26
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_DLLNOTFOUND = 32

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Private Sub cmdExport_Click()
Dim Tabla As New ADODB.Recordset
Dim i           As Long
Dim j           As Long
Dim lRowCount   As Long
Dim lPasteCount As Long
Dim sLtr        As String
Dim sStart      As String
Dim sEnd        As String
Dim sRowData    As String
Dim sSelData    As String
Dim Emp         As String
Dim oExcelApp   As Excel.Application
Dim oWs         As Excel.Worksheet
Dim oWb         As Excel.Workbook
Dim cNUMROWS As Integer
Emp = GetSetting(App.Title, "Estilo", "Emp")
Tabla.Open "Select * From Registro", Base, adOpenDynamic, adLockBatchOptimistic
If Tabla.EOF Then MsgBox "No hay Registros", vbExclamation: Exit Sub
 Do While Not Tabla.EOF
    cNUMROWS = cNUMROWS + 1
    Tabla.MoveNext
 Loop
 Tabla.MoveFirst
Const cNUMCOLS = 7

Const cFIXEDROWS = 5
Const cCLIPROWS = 500

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass
If Dir(sNewXlsFile) <> "" Then Kill sNewXlsFile
'
' Create an invisible Excel instance.
'
' Open a previously created worksheet that has most
' of the desired formatting already. Save this template
' as a new file so as not to destroy it.
'
Set oExcelApp = CreateObject("EXCEL.APPLICATION")
oExcelApp.Visible = False
oExcelApp.Workbooks.Open FileName:=sXlsTemplate, ReadOnly:=True, ignoreReadOnlyRecommended:=True
Set oWs = oExcelApp.ActiveSheet
Set oWb = oExcelApp.ActiveWorkbook
oWs.SaveAs FileName:=sNewXlsFile
'
' Populate the header information by writting
' directly to specific cells.
'
' Note:
' Strings are prefixed with a quote mark.
'
With oWs
    .Cells(1, 2).Value = Emp
    .Cells(2, 2).Value = Date
    .Cells(3, 2).Value = Time
    .Cells(4, 1).Value = "Folio"
    .Cells(4, 2).Value = "Pc"
    .Cells(4, 3).Value = "Cliente"
    .Cells(4, 4).Value = "Hora de entrada"
    .Cells(4, 5).Value = "Hora de salida"
    .Cells(4, 6).Value = "Horas acomuladas"
    .Cells(4, 7).Value = "Otros"
    .Cells(4, 8).Value = "Total"
End With
'
' Now lets populate the "body" of the spreadsheet.
' Determine the range of cells to be populated
' and change their format to numeric.
'
sStart = "A" & CStr(cFIXEDROWS + 1)
sLtr = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", cNUMCOLS + 1, 1)
sEnd = sLtr & CStr(cFIXEDROWS + cNUMROWS + 1)

oWs.Range(sStart, sEnd).Select
oWs.Range(sStart, sEnd).Activate
'
' Populate the body of the spreadsheet.
'
sSelData = ""
lRowCount = 0
lPasteCount = 0

For i = 0 To cNUMROWS - 1
    sRowData = ""
    '
    ' Create the rows to send to Excel. Each row
    ' is a tab delimited string of values terminated
    ' by a carriage return and line feed. Data can
    ' come from a grid or other source.
    '
    For j = 0 To cNUMCOLS
        sRowData = sRowData & Tabla(j).Value & vbTab
    Next
    sRowData = Left$(sRowData, Len(sRowData) - 1)
    '
    ' Rows are accumulated into blocks then stored in
    ' the clipboard and pasted into Excel in one shot.
    '
    ' They can be written one at a time but this is
    ' faster since the data is kept in memory and
    ' there are fewer calls to Excel.
    '
    sSelData = sSelData + sRowData + vbCrLf
    lRowCount = lRowCount + 1
    
    If lRowCount = cCLIPROWS Then
        Clipboard.Clear
        Clipboard.SetText sSelData
        sSelData = ""
        With oWs
            .Range("A" & CStr(lPasteCount * cCLIPROWS + cFIXEDROWS)).Select
            .Paste
            .Range("A1").Select
        End With
        lRowCount = 0
        lPasteCount = lPasteCount + 1
    End If
    Tabla.MoveNext
Next
Clipboard.Clear
Clipboard.SetText sSelData
With oWs
    .Range("A" & CStr(lPasteCount * cCLIPROWS + cFIXEDROWS)).Select
    .Paste
    .Range("A1").Select
End With
j = (lPasteCount * cCLIPROWS) + cFIXEDROWS + lRowCount
For i = 1 To cNUMCOLS + 1
    With oWs.Cells(j, i)
        .Borders(xlTop).LineStyle = xlDouble
        .Font.Bold = True
        .Font.ColorIndex = 3
    End With
Next
oWs.Cells(j, 8).Value = "=SUM(H" & CStr(cFIXEDROWS) & ":H" & CStr(j - 1) & ")"

'
' Save the changed worksheet.
'
oWb.Save
oWb.Saved = True
'
' Terminate and release the Excel objects.
'
oExcelApp.Quit
Set oWs = Nothing
Set oWb = Nothing
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
MsgBox "Cierre de caja terminado", vbInformation, "Excel Export Example"
Tabla.Close
Tabla.Open "Delete from Registro", Base, adOpenDynamic, adLockBatchOptimistic
cmdExport.Enabled = False
Exit Sub

ErrorHandler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " (" & CStr(Err.Number) & ")", vbExclamation, "Excel Export Example"
    On Error Resume Next
    oExcelApp.Quit
    Set oWs = Nothing
    Set oWb = Nothing
    Set oExcelApp = Nothing
End Sub
Private Sub cmdView_Click()
Dim lresult As Long
'
' ShellExecute opens or prints a specified file.
' For a complete description see my "ShellExecute" example.
'
On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass
lresult = ShellExecute(Me.hwnd, "open", sNewXlsFile & vbNullChar, "", 0, SW_SHOWNORMAL)
Call pDisplayError(lresult)
Exit Sub

ErrorHandler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " (" & CStr(Err.Number) & ")", vbExclamation, "Excel Export Example"
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
On Error Resume Next
sXlsTemplate = App.Path & "\Registros\dia.xls"
sNewXlsFile = App.Path & "\Registros\" & Format(Now, "DDDD dd-MMMM-YYYY") & ".xls"
Me.BackColor = Color
End Sub

Public Sub pDisplayError(ByVal lresult As Long)
Dim sError As String
'
' ShellExecute returns a value greater than 32 if there were no errors.
' Check its return code and display the associated error message.
'
Select Case lresult
    Case 0
        sError = "The operating system is out of memory or resources."
    Case ERROR_FILE_NOT_FOUND
        sError = "The specified file was not found."
    Case ERROR_PATH_NOT_FOUND
        sError = "The specified path was not found."
    Case ERROR_BAD_FORMAT
        sError = "The .exe file is invalid (non-Win32® .exe or error in .exe image)."
    Case SE_ERR_ACCESSDENIED
        sError = "The operating system denied access to the specified file."
    Case SE_ERR_ASSOCINCOMPLETE
        sError = "The file name association is incomplete or invalid."
    Case SE_ERR_DDEBUSY
        sError = "The DDE transaction could not be completed because other DDE transactions were being processed."
    Case SE_ERR_DDEFAIL
        sError = "The DDE transaction failed."
    Case SE_ERR_DDETIMEOUT
        sError = "The DDE transaction could not be completed because the request timed out."
    Case SE_ERR_DLLNOTFOUND
        sError = "The specified dynamic-link library was not found."
    Case SE_ERR_NOASSOC
        sError = "There is no application associated with the given file name extension."
    Case SE_ERR_OOM
        sError = "There was not enough memory to complete the operation."
    Case SE_ERR_SHARE
        sError = "A sharing violation occurred."
    Case Else
        sError = ""
End Select

Screen.MousePointer = vbDefault
If lresult <= 32 Then
    MsgBox sError, vbCritical, "Excel/ShellExecute Example"
End If

End Sub
