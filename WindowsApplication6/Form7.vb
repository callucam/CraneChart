Public Class Form7

    Public DevelopmentFilePath As String = "C:\CraneChartCalc\Input\"
    Public j As Integer = 0
    Public DataTableBarge As DataTable
    Public CurrentRowBarge As Integer
    Public CurrentRowCrane As Integer
    Public xCrane As XElement
    Public ConfigIndex As Integer
    Public MachineListIndex As Integer
    Public AllAroundRotation As Boolean
    Public Const HydrostaticsDataSize As Integer = 50 '50 or 200
    Public Const CrossCurveDataSize As Integer = 50
    Public Const RotationArraySize As Integer = 11
    Public DisplacementArray(0 To HydrostaticsDataSize - 1) As Double
    Public LCFDraftArray(0 To HydrostaticsDataSize - 1) As Double
    Public LCFArray(0 To HydrostaticsDataSize - 1) As Double
    Public MTCArray(0 To HydrostaticsDataSize - 1) As Double
    Public LCBArray(0 To HydrostaticsDataSize - 1) As Double
    Public FPDraftArray(0 To HydrostaticsDataSize - 1) As Double
    Public APDraftArray(0 To HydrostaticsDataSize - 1) As Double
    Public RMArray(0 To HydrostaticsDataSize - 1) As Double
    Public GMTArray(0 To HydrostaticsDataSize - 1) As Double
    Public GMLArray(0 To HydrostaticsDataSize - 1) As Double
    Public DataTableCCondensed As New DataTable

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GeneralReset()

    End Sub

    Private Sub FillComboBoxes()

        ChooseBargeComboBox.Items.Clear()
        ChooseCraneComboBox.Items.Clear()

        Dim i As Integer

        DataSetBarge.Clear()
        DataSetBarge.ReadXml(My.Settings.BargeListFileSetting)
        DataTableBarge = DataSetBarge.Tables(0)
        For i = 0 To DataSetBarge.Tables(0).Rows.Count - 1
            ChooseBargeComboBox.Items.Add(DataSetBarge.Tables(0).Rows(i).Item(0).ToString)
        Next
        'ChooseBargeComboBox.Items.Add("Create New ...")
        xCrane = XElement.Load(My.Settings.CraneListFileSetting)
        For i = 0 To xCrane...<Crane>.Count - 1
            ChooseCraneComboBox.Items.Add(xCrane...<Crane>(i).<ProposedCrane>.Value)
        Next
        'ChooseCraneComboBox.Items.Add("Create New ...")
    End Sub

    Private Sub GeneralReset()

        DataSetBarge.Clear()
        DataSetBarge.ReadXml(My.Settings.BargeListFileSetting)
        FillComboBoxes()

    End Sub

    Private Sub ChooseBargeComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ChooseBargeComboBox.SelectedIndexChanged

        Form1.LoadBarge(ChooseBargeComboBox.SelectedIndex)
        'Form1.IllustrateSetup()

    End Sub


    Private Sub ConfigComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ConfigComboBox.SelectedIndexChanged

        Form1.LoadConfig(ConfigComboBox.SelectedIndex)

    End Sub

    Private Sub ChooseCraneComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ChooseCraneComboBox.SelectedIndexChanged
        Form1.LoadCrane(ChooseCraneComboBox.SelectedIndex)
    End Sub
End Class