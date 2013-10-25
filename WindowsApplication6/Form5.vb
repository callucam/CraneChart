Public Class Form5

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox(Form1.ConfigIndex)
        For i = 0 To DataGridViewListChart.RowCount - 1

            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<Config>(Form1.ConfigIndex).<ListChart>(Form1.MachineListIndex).<Row1>(i).<OppRad>.Value = DataGridViewListChart.Item(0, i).Value
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<Config>(Form1.ConfigIndex).<ListChart>(Form1.MachineListIndex).<Row1>(i).<BoomAngle>.Value = DataGridViewListChart.Item(1, i).Value
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<Config>(Form1.ConfigIndex).<ListChart>(Form1.MachineListIndex).<Row1>(i).<BoomPoint>.Value = DataGridViewListChart.Item(2, i).Value
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<Config>(Form1.ConfigIndex).<ListChart>(Form1.MachineListIndex).<Row1>(i).<Capacity>.Value = DataGridViewListChart.Item(3, i).Value
        Next
        Form1.xCrane.Save(My.Settings.CraneListFileSetting)

        Dim row As DataRow

        row = Form1.DataSetEventLog.Tables(0).NewRow
        row.Item(0) = Now.ToOADate
        row.Item(1) = Form1.CraneNameTextbox.Text & "." & "MachineListTable" & "Change(" & Form1.CurrentRowCrane & "." & Form1.ConfigIndex & "." & Form1.MachineListIndex & ")"
        Form1.DataSetEventLog.Tables(0).Rows.Add(row)
        Form1.DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
        Form1.DataSetEventLog.Clear()
        Form1.DataSetEventLog.ReadXml(My.Settings.EventLogSetting)
        Form1.DataGridViewEventLog.DataSource = Form1.DataSetEventLog
        Form1.DataGridViewEventLog.DataMember = "Row"
    End Sub

    Private Sub PasteChart_Click(sender As Object, e As EventArgs) Handles PasteChart.Click
        Dim s As String
        Try
            s = Clipboard.GetText()
            Dim i, ii As Integer

            Dim tArr() As String = s.Split(ControlChars.NewLine)
            Dim arT() As String
            Dim cc, iRow, iCol As Integer

            iRow = DataGridViewListChart.SelectedCells(0).RowIndex
            iCol = DataGridViewListChart.SelectedCells(0).ColumnIndex
            For i = 0 To tArr.Length - 1
                If tArr(i) <> "" Then
                    arT = tArr(i).Split(vbTab)
                    cc = iCol
                    For ii = 0 To arT.Length - 1
                        If cc > DataGridViewListChart.ColumnCount - 1 Then Exit For
                        If iRow > DataGridViewListChart.Rows.Count - 1 Then Exit Sub
                        With DataGridViewListChart.Item(cc, iRow)
                            .Value = arT(ii).TrimStart

                        End With
                        cc = cc + 1
                    Next
                    iRow = iRow + 1
                End If

            Next

        Catch ex As Exception
            MsgBox("Please redo Copy and Click on cell")
        End Try
    End Sub
End Class