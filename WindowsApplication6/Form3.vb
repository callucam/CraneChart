Public Class Form3

    Private Sub CopyC_Click(sender As Object, e As EventArgs) Handles CopyC.Click
        Dim s As String
        Try
            s = Clipboard.GetText()
            Dim i, ii As Integer

            Dim tArr() As String = s.Split(ControlChars.NewLine)
            Dim arT() As String
            Dim cc, iRow, iCol As Integer

            iRow = DataGridView1.SelectedCells(0).RowIndex
            iCol = DataGridView1.SelectedCells(0).ColumnIndex
            For i = 0 To tArr.Length - 1
                If tArr(i) <> "" Then
                    arT = tArr(i).Split(vbTab)
                    cc = iCol
                    For ii = 0 To arT.Length - 1
                        If cc > DataGridView1.ColumnCount - 1 Then Exit For
                        If iRow > DataGridView1.Rows.Count - 1 Then Exit Sub
                        With DataGridView1.Item(cc, iRow)
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

    Private Sub SaveC_Click(sender As Object, e As EventArgs) Handles SaveC.Click
        '    For i = 0 To DataGridView1.RowCount - 1
        '        Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<Name>.Value = DataGridViewWAC.Item(0, i).Value
        '        Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<Mass>.Value = DataGridViewWAC.Item(1, i).Value
        '        Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<LCG>.Value = DataGridViewWAC.Item(2, i).Value
        '        Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<TCG>.Value = DataGridViewWAC.Item(3, i).Value
        '        Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<VCG>.Value = DataGridViewWAC.Item(4, i).Value
        '    Next
        '    Form1.xCrane.Save(My.Settings.CraneListFileSetting)
    End Sub

End Class