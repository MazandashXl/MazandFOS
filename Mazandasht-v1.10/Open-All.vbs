Option Explicit
Dim app, fso, folder, names, item, book, ws
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
names = Array("Info", "Weigh", "Txn", "Basket", "Buy", "Sale", "Staff", "Cost", "Cheq", "Fin", "Partner", "Dash", "Invoice", "Help")
For Each item In names
  If Not fso.FileExists(fso.BuildPath(folder, item & ".xlsx")) Then
    MsgBox "Missing workbook: " & item & ".xlsx" & vbCrLf & "Keep every workbook together in the same folder.", 16, "Mazandasht"
    WScript.Quit 1
  End If
Next
Set app = CreateObject("Excel.Application")
app.Visible = False
app.AskToUpdateLinks = False
app.ScreenUpdating = False
app.DisplayAlerts = False
app.MapPaperSize = False
On Error Resume Next
For Each item In names
  Set book = app.Workbooks.Open(fso.BuildPath(folder, item & ".xlsx"), 0, False)
  If Err.Number <> 0 Then
    MsgBox "Could not open " & item & ".xlsx: " & Err.Description, 16, "Mazandasht"
    app.Visible = True
    app.DisplayAlerts = True
    app.ScreenUpdating = True
    WScript.Quit 1
  End If
  app.Calculation = -4135
Next
For Each book In app.Workbooks
  For Each ws In book.Worksheets
    If book.Name = "Invoice.xlsx" And ws.Name = "A5" Then
      ws.PageSetup.PaperSize = 11
    Else
      ws.PageSetup.PaperSize = 9
    End If
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    If book.Name = "Invoice.xlsx" And (ws.Name = "A4" Or ws.Name = "A4Flex" Or ws.Name = "A5") Then
      ws.PageSetup.FitToPagesTall = 1
    Else
      ws.PageSetup.FitToPagesTall = False
    End If
  Next
Next
app.Calculation = -4105
app.CalculateFullRebuild
app.Workbooks("Dash.xlsx").Worksheets("Main").Activate
app.DisplayAlerts = True
app.ScreenUpdating = True
app.Visible = True
On Error GoTo 0

