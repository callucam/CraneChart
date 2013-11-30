Imports Microsoft.Office.Interop

Public Class Form2

    Private Sub ModifyH_Click(sender As Object, e As EventArgs) Handles ModifyH.Click

        Dim oXL As Excel.Application = Nothing
        Dim oWBs As Excel.Workbooks = Nothing
        Dim oWB As Excel.Workbook = Nothing
        Dim DebuggingWorksheet As Excel.Worksheet = Nothing
        Dim ResultsWorksheet As Excel.Worksheet = Nothing
        Dim ContentWorksheet As Excel.Worksheet = Nothing

        Dim oCells As Excel.Range = Nothing

        oXL = New Excel.Application
        oXL.Visible = True
        oWBs = oXL.Workbooks
        oWB = oWBs.Open("C:\CraneChartCalc\Input\Hydrostatics\AddHydrostatics.xlsx")
        'oWB = oWBs.Open(My.Settings.XLSXTemplateSetting)
        'DebuggingWorksheet = oWB.Worksheets(1)
        'DebuggingWorksheet.Range("g2:s1000").Font.ColorIndex = 3

    End Sub
End Class