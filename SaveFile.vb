Sub SaveFile()
'
' SaveFile Macro
'

'
    Workbooks.Open Filename:= _
        "C:\Users\tjmacdonald\Documents\DM Tool Kit\toolNames.xlsx"
    Workbooks.Open Filename:= _
        "C:\Users\tjmacdonald\Documents\DM Tool Kit\toolTemplate.xlsx"
    Dim counter As Integer
    counter = 1
    For counter = 1 To 56 Step 1
        Windows("tool.xlsb").Activate
        Sheets("Dashboard 2").Select
        Sheets("Dashboard 2").Copy Before:=Workbooks("toolTemplate.xlsx").Sheets(1)
        Application.WindowState = xlMaximized
        Windows("toolNames").Activate
        Range("A" & counter).Select
        Dim name As String
        name = Selection.Text
        Windows("toolTemplate").Activate
        Sheets("Dashboard 2").Select
        Range("C8").Select
        ActiveCell.FormulaR1C1 = name
        Application.CutCopyMode = False
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "C:\Users\tjmacdonald\Documents\DM Tool Kit\Dashboards\" + name + ".pdf", _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=False, OpenAfterPublish:=False
    Next counter
    Windows("toolNames").Close
    Windows("toolTemplate").Close
End Sub