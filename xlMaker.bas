Attribute VB_Name = "Module1"
' Version 2024-04-02
Global wb0, wb2 As Workbook
Global ws As Worksheet
Global sPathIn, sPathTemplate As String
Global srcDta, buffer As String
Global templateName As String
Global parm() As String
Sub main()
Set wb0 = ActiveWorkbook
sPathIn = "c:\xlMaker\"
xCmd = sPathIn + "xlmaker.xCmd"

Open xCmd For Input As #1
Do Until EOF(1)
    Line Input #1, srcDta
    Call loadBuffer
Loop

Close #1

End Sub
Sub loadBuffer()
srcDta = Trim(srcDta)
If (srcDta = "" Or Left(srcDta, 2) = "//") Then Exit Sub
buffer = buffer + srcDta
Call cmdExec
End Sub
Sub cmdExec()
Dim p1, p2 As Integer
Dim command, parms As String

If Right(buffer, 1) <> ";" Then Exit Sub

p1 = InStr(buffer, "(")
command = Left(buffer, p1 - 1)

p2 = InStrRev(buffer, ")")
parms = Mid(buffer, p1 + 1, p2 - p1 - 1)
parm = Split(parms, ",")
For i = 0 To UBound(parm)
    parm(i) = Trim(parm(i))
Next i

Run command

buffer = ""
ReDim parm(0)
End Sub
Public Sub DLTSET()
Dim setName As String
On Error Resume Next
setName = parm(0)
KILL setName
End Sub
Public Sub DLTXL()
Dim pathName As String
pathName = parm(0)
If Dir(pathName) <> "" Then
    KILL pathName
End If
End Sub
Public Sub OVRXLT()
Dim fileID As String

fileID = parm(0)

Set wb0 = Workbooks.Open(fileID)
End Sub
Public Sub NEWXL()

Set wb2 = Workbooks.Add

End Sub
Public Sub ADDSHEET()
Dim sTemplate, sNewSheet, fileName As String

sTemplate = parm(0)
sNewSheet = parm(1)

If UBound(parm) < 2 Then
  fileName = parm(1)
Else
  fileName = parm(2)
End If

wb0.Sheets(sTemplate).Visible = True
wb0.Sheets(sTemplate).Copy wb2.Sheets(wb2.Sheets.Count)

Set ws = ActiveSheet

ws.Name = sNewSheet

For Each nX In ws.Names
    Call loadData(sNewSheet, nX, fileName)
Next nX

End Sub
Public Sub loadData(ByVal sFile As String, ByVal nX As Name, ByVal fileName As String)

Dim pathdata As String
p = InStrRev(nX.Name, "!")

sName = Mid(nX.Name, p + 1)

If sName = "_FilterDatabase" Then
    Exit Sub
End If

dataPath = sPathIn & "in\" & fileName & "." & sName
sRange = nX.Value

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & dataPath, Destination:=Range(sRange))
    .Name = "n_" + sName
    .AdjustColumnWidth = False
    .FieldNames = False
    .RowNumbers = False
    .FillAdjacentFormulas = True
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlOverwriteCells
    .SavePassword = False
    .SaveData = True
    .RefreshPeriod = False
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 1252
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlNone
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery = False
End With

ActiveWorkbook.Connections(1).Delete

End Sub
Public Sub RUNVBA()

' Run "'" & WorkbookName & "'!" & MacroName, argument1, argument2
' Application.Run "AnotherWorkbook.xlsm!NameOfMacro"

Run wb0.Name & "!specif", ws

End Sub
Public Sub SAVXL()
Dim pathName As String
pathName = parm(0)

Application.DisplayAlerts = False
For Each Sheet In wb2.Sheets
    If Left(Sheet.Name, 5) = "Sheet" Then Sheet.Delete
Next

wb2.Sheets(1).Select
wb2.SaveAs pathName
wb2.Close
End Sub
Public Sub RLSXLT()
wb0.Close
End Sub



