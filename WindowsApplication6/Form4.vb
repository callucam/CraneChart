Public Class Form4

    Private Sub SaveWACButton_Click(sender As Object, e As EventArgs) Handles SaveWACButton.Click
        For i = 0 To DataGridViewWAC.RowCount - 1
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<Name>.Value = DataGridViewWAC.Item(0, i).Value
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<Mass>.Value = DataGridViewWAC.Item(1, i).Value
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<LCG>.Value = DataGridViewWAC.Item(2, i).Value
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<TCG>.Value = DataGridViewWAC.Item(3, i).Value
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<WAC>.<Row3>(i).<VCG>.Value = DataGridViewWAC.Item(4, i).Value
        Next
        Form1.xCrane.Save(My.Settings.CraneListFileSetting)
    End Sub

    Private Sub SaveBoomWeightButton_Click(sender As Object, e As EventArgs) Handles SaveBoomWeightButton.Click
        For i = 0 To DataGridViewBoomLengthWeight.RowCount - 1
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<BoomLengthWeight>.<Row2>(i).<BoomLength>.Value = DataGridViewBoomLengthWeight.Item(0, i).Value
            Form1.xCrane...<Crane>(Form1.CurrentRowCrane).<BoomLengthWeight>.<Row2>(i).<BoomWeight>.Value = DataGridViewBoomLengthWeight.Item(1, i).Value
        Next
        Form1.xCrane.Save(My.Settings.CraneListFileSetting)
    End Sub

    Private Sub PasteWACButton_Click(sender As Object, e As EventArgs) Handles PasteWACButton.Click
        Dim s As String
        Try
            s = Clipboard.GetText()
            Dim i, ii As Integer

            Dim tArr() As String = s.Split(ControlChars.NewLine)
            Dim arT() As String
            Dim cc, iRow, iCol As Integer

            iRow = DataGridViewWAC.SelectedCells(0).RowIndex
            iCol = DataGridViewWAC.SelectedCells(0).ColumnIndex
            For i = 0 To tArr.Length - 1
                If tArr(i) <> "" Then
                    arT = tArr(i).Split(vbTab)
                    cc = iCol
                    For ii = 0 To arT.Length - 1
                        If cc > DataGridViewWAC.ColumnCount - 1 Then Exit For
                        If iRow > DataGridViewWAC.Rows.Count - 1 Then Exit Sub
                        With DataGridViewWAC.Item(cc, iRow)
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

    Private Sub PasteBoomWeightButton_Click(sender As Object, e As EventArgs) Handles PasteBoomWeightButton.Click
        Dim s As String
        Try
            s = Clipboard.GetText()
            Dim i, ii As Integer

            Dim tArr() As String = s.Split(ControlChars.NewLine)
            Dim arT() As String
            Dim cc, iRow, iCol As Integer

            iRow = DataGridViewBoomLengthWeight.SelectedCells(0).RowIndex
            iCol = DataGridViewBoomLengthWeight.SelectedCells(0).ColumnIndex
            For i = 0 To tArr.Length - 1
                If tArr(i) <> "" Then
                    arT = tArr(i).Split(vbTab)
                    cc = iCol
                    For ii = 0 To arT.Length - 1
                        If cc > DataGridViewBoomLengthWeight.ColumnCount - 1 Then Exit For
                        If iRow > DataGridViewBoomLengthWeight.Rows.Count - 1 Then Exit Sub
                        With DataGridViewBoomLengthWeight.Item(cc, iRow)
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