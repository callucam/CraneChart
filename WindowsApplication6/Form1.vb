#Region "Imports directives"

Imports System.Reflection
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Scripting
Imports Microsoft.VisualBasic

#End Region

Public Class Form1

    Public DevelopmentFilePath As String = "C:\CraneChartCalc\Input\"
    Public j As Integer = 0
    Public DataTableBarge As DataTable
    Public DataTableC As DataTable
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

    Public TotalMass As Double
    Public TotalLCG As Double
    Public TotalTCG As Double
    Public TotalVCG As Double
    Public ww As Integer

    Public LCFDraft As Double
    Public LCF As Double
    Public MTC As Double
    Public LCB As Double
    Public FPDraft As Double
    Public APDraft As Double
    Public RM As Double
    Public TrimInCm As Double
    Public TrimInM As Double
    Public ProportionAft As Double
    Public ProportionFwd As Double
    Public GMt As Double
    Public RM2 As Double
    Public GMl As Double
    Public TrimAngle As Double
    Public Density As Double = 1.0
    Public BargeLengthVar As Double
    Public BargeBreadthVar As Double
    Public BargeDepthVar As Double
    Public aa As Double
    Public gg As Double
    Public cc(0 To 1) As Double
    Public dd(0 To 1) As Double
    Public ee(0 To 1) As Double
    Public ff(0 To 1) As Double

    Public GZMax As Double
    Public GZArea As Double
    Public IndexAtMaxGZ As Integer
    Public IndexAtVanishingGZ As Integer
    Public KNArray(0 To 900) As Double
    Public GZArray(0 To 900) As Double
    Public HeelArray(0 To 900) As Double
    Public SM(0 To 900) As Double
    Public Residual(0 To 900) As Double
    Public CCArray(0 To CrossCurveDataSize - 1, 0 To 900) As Double
    Public DisplacementCCArray(0 To CrossCurveDataSize - 1) As Double
    Public MinimumResidual As Double
    Public IndexAtMinimumResidual As Integer

    Public HeelAtMinimumResidual As Double
    Public HeelAtMaxGZ As Double
    Public HeelAtVanishingGZ As Double
    Public HeelAtVanishingGZFromEQ As Double
    Public LeastFreeboard As Double
    Public DebugIndex As Integer
    Public DebuggingWorksheet As Excel.Worksheet = Nothing

    Public GMLLimit As Double
    Public Test1 As Double
    Public Test2 As Double
    Public Test3 As Double
    Public Test4 As Double
    Public Test5 As Double
    'Dim Test6 As String
    Public Test7 As Double
    Public Test8 As Double
    Public Test9 As Double
    Public Test10 As Double
    Public Test11 As Double
    Public Test12 As Double
    Public Test13 As Double
    Public ListLimit As Double
    Public ResultsWorksheet As Excel.Worksheet = Nothing
    Public HydrostaticDataTab As String
    Public CrossCurveDataTab As String
    Public BoomHingetoCraneCL As Double
    Public CrawlerBasetoBoomHinge As Double
    Public PositionofCraneCentretoFrontofTrack As Double
    Public ModelAbbreviation As String
    Public SheetName As String
    Public MachineList As String
    Public EffectiveBoomLength As Double
    Public PositionofCraneCentre As Double
    Public PositionofCraneCentreOffCentreline As Double
    Public NominalBoomLength As Double
    Public BoomExtension As Double
    Public BoomHeight As Double

    Public CraneSerialNumber As String
    Public BargeNameVar As String
    Public ProposedCraneVar As String

    Public LightshipWAC(0 To 0, 0 To 3) As Double
    Public BallastWAC(0 To 0, 0 To 3) As Double
    Public CraneWAC(0 To 4, 0 To 3) As Double
    Public CraneWACPreLift(0 To 4, 0 To 3) As Double
    Public TotalWAC(0 To 2, 0 To 3) As Double
    Public TotalWACPreLift(0 To 2, 0 To 3) As Double

    Public DeckTimberHeight As Double

    Public CraneTrim As Double
    'Dim CraneTrim2 As Double
    Public MaxLoadAllowedByBargeVizStability As Double
    Public HalfMaxMomentAllowed As Double
    Public FullMaxMomentAllowed As Double

    Public ModifiedOperatingRadius As Double
    Public ModifiedCounterweightRadius As Double
    Public HalfMaxLoadAllowedByBargeVizTipping As Double
    Public FullMaxLoadAllowedByBargeVizTipping As Double
    Public MaxLoadAllowedbyBarge As Double

    Public ModifiedUpperworksRadius As Double
    Public ModifiedCrawlerRadius As Double
    Public ModifiedBoomRadius As Double

    Public CounterweightMass As Double
    Public CounterweightLCG As Double
    Public CounterweightVCG As Double
    Public UpperworksMass As Double
    Public UpperworksLCG As Double
    Public UpperworksVCG As Double
    Public CrawlersMass As Double
    Public CrawlersLCG As Double
    Public CrawlersVCG As Double
    Public BoomMass As Double
    Public BoomLCG As Double
    Public BoomVCG As Double

    Public CraneLCG As Double
    Public CraneTCG As Double
    Public CraneVCG As Double

    Public Mass As Double

    Public CraneLoad As Double
    Public DeckInclinationAngle As Double
    Public InitialCraneLoad As Double

    Public OperatingRadius As Double

    Public BoomAngle As Double
    Public BoomPoint As Double


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Label1.Text = "Version: " & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString

        DataTableCCondensed.Columns.Add("Disp. kg")
        DataTableCCondensed.Columns.Add("LCG [m]")
        DataTableCCondensed.Columns.Add("KN @ 0 deg")
        DataTableCCondensed.Columns.Add("KN @ 10 deg")
        DataTableCCondensed.Columns.Add("KN @ 20 deg")
        DataTableCCondensed.Columns.Add("KN @ 30 deg")
        DataTableCCondensed.Columns.Add("KN @ 40 deg")
        DataTableCCondensed.Columns.Add("KN @ 50 deg")
        DataTableCCondensed.Columns.Add("KN @ 60 deg")
        DataTableCCondensed.Columns.Add("KN @ 70 deg")
        DataTableCCondensed.Columns.Add("KN @ 80 deg")
        DataTableCCondensed.Columns.Add("KN @ 90 deg")

        GeneralReset()


    End Sub

    Private Sub LoadCraneButton_Click(sender As Object, e As EventArgs) Handles LoadCraneTabButton.Click

        Me.TabControl2.SelectedIndex = 2

    End Sub

    Private Sub SaveButton_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub LoadBargeButton_Click(sender As Object, e As EventArgs) Handles LoadBargeTabButton.Click

        Me.TabControl2.SelectedIndex = 1

    End Sub

    Private Sub SaveBargeButton_Click(sender As Object, e As EventArgs)
        DataSetBarge.Tables(0).AcceptChanges()
        DataSetBarge.AcceptChanges()
        DataSetBarge.WriteXml(My.Settings.BargeListFileSetting)
        MsgBox("Change Recorded")
    End Sub

    Private Sub LoadCriteria_Click(sender As Object, e As EventArgs) Handles LoadCriteriaTabButton.Click
        Me.TabControl2.SelectedIndex = 5
    End Sub

    Private Sub RunButton_Click(sender As Object, e As EventArgs) Handles RunButton.Click

        ' All data in m, kg, s, deg

        Dim CranePosition As Double

        Dim CraneName As String
        Dim RotationArray()
        Dim ActiveSet As Integer
        Dim Rotation1 As Integer
        Dim AbsoluteRow As Integer
        

        Dim Truth As Integer


        Dim i As Integer
        Dim k As Integer
        Dim m As Integer
        Dim j As Integer
        Dim xx As Integer
        Dim yy As Integer
        Dim q As Integer
        Dim r As Integer

        Dim FirstRotation As Double
        Dim Reason(0 To 3) As String
        Dim ReasonID As Integer

        Dim MaxLoadAllowedByBargeArray(0 To RotationArraySize - 1) As Double
        Dim HeelAtMinimumResidualArray(0 To RotationArraySize - 1) As Double
        Dim LeastFreeboardArray(0 To RotationArraySize - 1) As Double
        Dim CraneListArray(0 To RotationArraySize - 1) As Double
        Dim CraneTrimArray(0 To RotationArraySize - 1) As Double

        Dim IntervalStart As Object
        Dim CalculationAllowance As Double
        Dim WorksheetExists As Integer

        Dim IncludeDebugging As Boolean
        Dim DerateIncrement As Double

        Dim oXL As Excel.Application = Nothing
        Dim oWBs As Excel.Workbooks = Nothing
        Dim oWB As Excel.Workbook = Nothing

        Dim ContentWorksheet As Excel.Worksheet = Nothing

        Dim oCells As Excel.Range = Nothing

        Dim CraneList As Double
        Dim StabilityCheck As Integer
        Dim CraneCheck As Integer
        Dim Couple As Integer

        

        'Me.TabControl2.SelectedIndex = 2

        'StartProgressBar()

        oXL = New Excel.Application
        oXL.Visible = True
        oWBs = oXL.Workbooks
        oWB = oWBs.Open(My.Settings.XLSXTemplateSetting)
        DebuggingWorksheet = oWB.Worksheets(1)

        DebuggingWorksheet.Range("g2:s1000").Font.ColorIndex = 3

        'RenameSheetIfDuplicate()

        DebugIndex = 0

        'Barge Input

        CurrentRowBarge = 0
        'CurrentRowCrane = 0

        BargeNameVar = BargeNameTextbox.Text
        BargeLengthVar = Me.BargeLength.Text
        BargeBreadthVar = Me.BargeBreadth.Text
        BargeDepthVar = Me.BargeDepth.Text

        HydrostaticDataTab = Me.HydrostaticsDataTag.Text
        CrossCurveDataTab = Me.CrossCurvesDataTag.Text

        'Crane Input

        ProposedCraneVar = xCrane...<ProposedCrane>(CurrentRowCrane).Value
        CraneSerialNumber = xCrane...<CraneSerialNumber>(CurrentRowCrane).Value
        BoomHingetoCraneCL = CDbl(xCrane...<BoomHingetoCraneCL>(CurrentRowCrane).Value)
        CrawlerBasetoBoomHinge = xCrane...<CrawlerBasetoBoomHinge>(CurrentRowCrane).Value
        PositionofCraneCentretoFrontofTrack = xCrane...<PositionofCraneCentretoFrontofTrack>(CurrentRowCrane).Value
        PositionofCraneCentreOffCentreline = PositionofCraneCentreOffCentrelineTextbox.Text

        If AllAroundRadio.Checked = True Then ModelAbbreviation = "360"
        If OverEndRadio.Checked = True Then ModelAbbreviation = "000"

        MachineList = ListChartComboBox.SelectedItem.ToString

        NominalBoomLength = BoomLengthTextbox.Text

        'CranePosition = DataGridViewPosition.Item(0, 0).Value '#### FIX THIS
        CranePosition = DistanceFromEndTextbox.Text

        ReDim RotationArray(RotationArraySize)

        For j = 0 To RotationArraySize - 1
            RotationArray(j) = DataSetRotation.Tables(0).Rows(j).Item(0).ToString
        Next

        PositionofCraneCentre = BargeLengthVar - PositionofCraneCentretoFrontofTrack - CranePosition
        EffectiveBoomLength = NominalBoomLength

        ListLimit = xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<MaxMachineList>.Value

        If ListLimit = 0 Then
            ListLimit = 0.44
        End If


        ActiveSet = 0
        SheetName = NominalBoomLength / 0.3048 & "-" & ModelAbbreviation & "-" & ListChartComboBox.SelectedItem.ToString & "-" & Math.Round(CranePosition / 0.3048, 0)

        oWB.Worksheets(2).Copy(After:=oWB.Worksheets(2))
        oWB.Worksheets(3).Name = SheetName
        ResultsWorksheet = oWB.Worksheets(3)

        WriteInputValues()

        For xx = 0 To CrossCurveDataSize - 1
            For yy = 0 To 900
                CCArray(xx, yy) = DataSetC.Tables(0).Rows(xx).Item(yy + 2).ToString
            Next
            If DataSetC.Tables(0).Rows(xx).Item(0).ToString <> "" Then
                DisplacementCCArray(xx) = DataSetC.Tables(0).Rows(xx).Item(0).ToString
            Else
                DisplacementCCArray(xx) = 999999999
            End If
        Next

        LightshipWAC(0, 0) = LightshipMass.Text
        LightshipWAC(0, 1) = LightshipLCG.Text
        LightshipWAC(0, 2) = LightshipTCG.Text
        LightshipWAC(0, 3) = LightshipVCG.Text

        BallastWAC(0, 0) = BallastMass.Text
        BallastWAC(0, 1) = BallastLCG.Text
        BallastWAC(0, 2) = BallastTCG.Text
        BallastWAC(0, 3) = BallastVCG.Text

        DataSetSM.ReadXml(My.Settings.SimpsonMultipliersSetting)

        For yy = 0 To 900
            SM(yy) = DataSetSM.Tables(0).Rows(0).Item(yy).ToString
            HeelArray(yy) = DataSetSM.Tables(0).Rows(1).Item(yy).ToString
        Next

        ReadCriteria()

        DeckTimberHeight = CranePadThickness.Text
        IncludeDebugging = IncludeDebuggingCheckBox.Checked
        DerateIncrement = DerateIncrementTextbox.Text

        ''frm.ProgressBar1.Value = 2
        ''frm.ProgressBar1.Refresh

        CalculationAllowance = CalculationAllowanceTextbox.Text

        Dim row As DataRow
        row = DataSetEventLog.Tables(0).NewRow
        row.Item(0) = Now.ToOADate
        row.Item(1) = "RUN: " & BargeName.ToString & "/ " & ProposedCraneVar.ToString & "/ " & SheetName
        DataSetEventLog.Tables(0).Rows.Add(row)
        DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
        FillEventLog()

        'For i = 0 To DataGridViewListChart.RowCount - 1
        For i = 0 To xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1>.Count - 1

            '    frm.ProgressBar1.Value = 2 + i
            '    frm.ProgressBar1.Refresh

            'MsgBox(xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1>(i).<OppRad>.Value)

            OperatingRadius = xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1>(i).<OppRad>.Value
            BoomAngle = xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1>(i).<BoomAngle>.Value
            BoomPoint = xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1>(i).<BoomPoint>.Value
            CraneLoad = xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1>(i).<Capacity>.Value

            'OperatingRadius = DataGridViewListChart.Item(0, i).Value
            'BoomAngle = DataGridViewListChart.Item(1, i).Value
            'BoomPoint = DataGridViewListChart.Item(2, i).Value
            'CraneLoad = DataGridViewListChart.Item(3, i).Value

            InitialCraneLoad = CraneLoad

            FirstRotation = RotationArray(0)

            Truth = 0

            IntervalStart = DateTime.Now.ToOADate()

            Do While Truth < UBound(RotationArray) And DateTime.Now.ToOADate() < IntervalStart + 1 / 24 / 60 / 60 * CalculationAllowance

                '        'MsgBox (Time & " " & IntervalStart + 1 / 24 / 60 / 60 * 2)

                If RotationArray(0) = FirstRotation Then
                    Truth = 0
                End If

                'MsgBox(UBound(RotationArray))

                For k = 0 To UBound(RotationArray) - 1

                    AbsoluteRow = ActiveSet * UBound(RotationArray) + k - 1
                    Rotation1 = RotationArray(k)

                    BoomExtension = OperatingRadius - BoomHingetoCraneCL
                    BoomHeight = Math.Sin(BoomAngle * 3.1415 / 180) * EffectiveBoomLength

                    SetupWeight(CraneLoad, Rotation1)
                    EquilibriumHydrostatics()
                    LargeAngleStability()
                    CraneList = Math.Sin((Rotation1) * 3.1415 / 180) * (TrimAngle) + Math.Cos((Rotation1) * 3.1415 / 180) * (HeelAtMinimumResidual)

                    GMLLimit = 0.02 * BargeLengthVar ^ 2 / LCFDraft

                    If LCFDraft >= Test1 And FPDraft >= Test2 And APDraft >= Test3 And Math.Abs(HeelAtMinimumResidual) <= Test4 And GMt >= BargeBreadthVar * Test5 And GMl >= GMLLimit And Math.Abs(TrimAngle) <= Test7 And LeastFreeboard >= Test8 And HeelAtMaxGZ > Test9 And GZMax >= Test10 And GZArea >= Test11 And HeelAtVanishingGZ >= Test12 Then

                        StabilityCheck = 1
                    Else
                        StabilityCheck = 0
                        ReasonID = 1
                    End If

                    If Math.Abs(CraneList) <= ListLimit + Test13 Then
                        CraneCheck = 1
                    Else
                        CraneCheck = 0
                        ReasonID = 3
                    End If

                    Couple = StabilityCheck * CraneCheck

                    Reason(0) = "Not derated."
                    Reason(1) = "Limited by barge stability."
                    Reason(2) = "Limited by crane overturning."
                    Reason(3) = "Limited by machine list."

                    CraneTrim = Math.Cos((Rotation1 * 3.1415 / 180)) * (TrimAngle) + Math.Sin((Rotation1) * 3.1415 / 180) * (HeelAtMinimumResidual)

                    CalculateMaxMoment()

                    If Couple = 1 Then

                        ReasonID = 1

                        If HalfMaxLoadAllowedByBargeVizTipping < MaxLoadAllowedByBargeVizStability Then
                            ReasonID = 2
                        End If

                        MaxLoadAllowedByBargeArray(k) = MaxLoadAllowedbyBarge
                        HeelAtMinimumResidualArray(k) = HeelAtMinimumResidual
                        LeastFreeboardArray(k) = LeastFreeboard
                        CraneListArray(k) = CraneList
                        CraneTrimArray(k) = CraneTrim

                    Else

                    End If

                    'Creating of Loading Condition Data

                    If IncludeDebugging Then
                        PrintDebugResultsInput(0, Rotation1, OperatingRadius, BoomAngle, CraneList)
                        PrintDebugResultsAlt1(4, Rotation1, OperatingRadius, BoomAngle, CraneList)

                        CraneWACPreLift = CraneWAC
                        CraneWACPreLift(4, 0) = 0
                        SetupWeight(0, Rotation1)
                        EquilibriumHydrostatics()
                        LargeAngleStability()
                        CraneList = Math.Sin((Rotation1) * 3.1415 / 180) * (TrimAngle) + Math.Cos((Rotation1) * 3.1415 / 180) * (HeelAtMinimumResidual)
                        CraneTrim = Math.Cos((Rotation1 * 3.1415 / 180)) * (TrimAngle) + Math.Sin((Rotation1) * 3.1415 / 180) * (HeelAtMinimumResidual)
                        ModifiedOperatingRadius = EffectiveBoomLength * Math.Cos((BoomAngle + CraneTrim) * 3.1415 / 180) + BoomHingetoCraneCL

                        PrintDebugResultsAlt1(24, Rotation1, OperatingRadius, BoomAngle, CraneList)

                        'DebuggingWorksheet.Range("ax2").Offset(DebugIndex, 0).Value = BoomAngle + CraneTrim
                        'DebuggingWorksheet.Range("ay2").Offset(DebugIndex, 0).Value = EffectiveBoomLength * (Math.Cos(3.1415 / 180 * (BoomAngle + CraneTrim))) + BoomHingetoCraneCL

                        ''Other

                        ''            Worksheets("Debugging").Range("j20").Resize(1, 7).Offset(DebugIndex, 0) = CraneTrimArray(k)
                        ''            Worksheets("Debugging").Range("j20").Resize(1, 1).Offset(DebugIndex, 0) = StabilityCheck
                        ''            Worksheets("Debugging").Range("k20").Resize(1, 1).Offset(DebugIndex, 0) = CraneCheck
                        ''            Worksheets("Debugging").Range("l20").Resize(1, 1).Offset(DebugIndex, 0) = Couple
                        ''            Worksheets("Debugging").Range("m20").Resize(1, 1).Offset(DebugIndex, 0) = Truth
                        '' Worksheets("Debugging").Range("n20").Resize(1, 1).Offset(DebugIndex, 0) = Reason(ReasonID)

                        ''Worksheets("Debugging").Range("h2").Offset(DebugIndex, 0) = TotalWAC(2, 1)
                        ''Worksheets("Debugging").Range("i2").Offset(DebugIndex, 0) = TotalWAC(3, 1)
                        ''Worksheets("Debugging").Range("k2").Offset(DebugIndex, 0) = EffectiveBoomLength
                        ''Worksheets("Debugging").Range("k2").Offset(DebugIndex, 0) = NominalBoomLength



                    End If

                    Truth = Truth + Couple

                    If IncludeDebugging Then

                        'DebuggingWorksheet.Range("t2").Offset(DebugIndex, 0).Value = Now()
                        'DebuggingWorksheet.Range("u2").Offset(DebugIndex, 0).Value = Truth
                        'DebuggingWorksheet.Range("v2").Offset(DebugIndex, 0).Value = Couple
                        'DebuggingWorksheet.Range("w2").Offset(DebugIndex, 0).Value = Reason(ReasonID)

                    End If

                    DebugIndex = DebugIndex + 1


                Next

                CraneLoad = CraneLoad * DerateIncrement

                MaxLoadAllowedbyBarge = 999999999999.0#
                HeelAtMinimumResidual = 0
                LeastFreeboard = 999999999999.0#
                CraneList = 0
                CraneTrim = 0

                For w = 0 To RotationArraySize - 1

                    If MaxLoadAllowedbyBarge > MaxLoadAllowedByBargeArray(w) Then
                        MaxLoadAllowedbyBarge = MaxLoadAllowedByBargeArray(w)
                    Else
                    End If

                    If Math.Abs(HeelAtMinimumResidual) < Math.Abs(HeelAtMinimumResidualArray(w)) Then
                        HeelAtMinimumResidual = HeelAtMinimumResidualArray(w)
                    Else
                    End If

                    If LeastFreeboard > LeastFreeboardArray(w) Then
                        LeastFreeboard = LeastFreeboardArray(w)
                    Else
                    End If

                    If Math.Abs(CraneList) < Math.Abs(CraneListArray(w)) Then
                        CraneList = CraneListArray(w)
                    Else
                    End If

                    If Math.Abs(CraneTrim) < Math.Abs(CraneTrimArray(w)) Then
                        CraneTrim = CraneTrimArray(w)
                    Else
                    End If

                Next

                ResultsWorksheet.Range("a13").Offset(i, 0).Value = OperatingRadius / 0.3048
                ResultsWorksheet.Range("b13").Offset(i, 0).Value = BoomAngle
                ResultsWorksheet.Range("c13").Offset(i, 0).Value = BoomPoint / 0.3048
                ResultsWorksheet.Range("d13").Offset(i, 0).Value = Math.Round(MaxLoadAllowedbyBarge * 2.20462262 / 100 - 1, 0) * 100
                ResultsWorksheet.Range("l13").Offset(i, 0).Value = InitialCraneLoad * 2.20462262
                ResultsWorksheet.Range("e13").Offset(i, 0).Value = HeelAtMinimumResidual
                ResultsWorksheet.Range("f13").Offset(i, 0).Value = LeastFreeboard / 0.3048
                ResultsWorksheet.Range("g13").Offset(i, 0).Value = CraneList
                ResultsWorksheet.Range("h13").Offset(i, 0).Value = CraneTrim
                'Worksheets(SheetName).Range("k13").Offset(i, 0) = Now

                '        If Time > IntervalStart + 1 / 24 / 60 / 60 * CalculationAllowance Then
                '        Worksheets(SheetName).Range("d13").Offset(i, 0) = "Timeout"
                '        End If

                If MaxLoadAllowedbyBarge = InitialCraneLoad Then
                    ReasonID = 0
                End If

                ResultsWorksheet.Range("d13").Offset(i, 0).NumberFormat = "0,000 """""

                If ReasonID = 1 Then
                    ResultsWorksheet.Range("d13").Offset(i, 0).NumberFormat = "0,000 ""*"""
                ElseIf ReasonID = 2 Then
                    ResultsWorksheet.Range("d13").Offset(i, 0).NumberFormat = "0,000 ""**"""
                ElseIf ReasonID = 3 Then
                    ResultsWorksheet.Range("d13").Offset(i, 0).NumberFormat = "0,000 ""***"""
                End If

            Loop


            ActiveSet = ActiveSet + 1

        Next

        'ContentWorksheet = oWB.Worksheets(2)
        'ContentWorksheet.Range("A1").Value = BargeName.ToString
        'ContentWorksheet.Range("A2").Value = CraneNameTextbox.Text
        'ContentWorksheet.Range("A3").Value = A3.Text

        'oWB.Worksheets(4).PageSetup.CenterFooter = A3.Text
        'oWB.Worksheets(5).PageSetup.CenterFooter = A3.Text
        'oWB.Worksheets(6).PageSetup.CenterFooter = A3.Text
        'oWB.Worksheets(7).PageSetup.CenterFooter = A3.Text
        'oWB.Worksheets(8).PageSetup.CenterFooter = A3.Text
        'oWB.Worksheets(9).PageSetup.CenterFooter = A3.Text
        'oWB.Worksheets(10).PageSetup.CenterFooter = A3.Text

        'ContentWorksheet.Range("A4").Value = A4.Text
        'ContentWorksheet.Range("A5").Value = A5.text
        'ContentWorksheet.Range("A6").Value = Now.ToShortTimeString
        'ContentWorksheet.Range("A7").Value = A7.Text
        'ContentWorksheet.Range("A8").Value = A8.Text
        'ContentWorksheet.Range("A9").Value = A9.Text
        'ContentWorksheet.Range("A10").Value = A10.Text
        'ContentWorksheet.Range("A10").Value = A10.Text
        'ContentWorksheet.Range("A15").Value = A15.Text
        'ContentWorksheet.Range("A16").Value = A16.Text
        'ContentWorksheet.Range("A17").Value = A17.Text
        'ContentWorksheet.Range("A18").Value = A18.Text
        'ContentWorksheet.Range("A19").Value = CounterweightTextbox.Text
        'ContentWorksheet.Range("A20").Value = BoomTypeTextbox.Text
        'ContentWorksheet.Range("A21").Value = BoomNameTextbox.Text
        'ContentWorksheet.Range("A22").Value = DeckTimberHeight.ToString
        ''ContentWorksheet.Range("A23").Value = CalibrationID.ToString
        ''ContentWorksheet.Range("A24").Value = BargeSerialNumber.ToString
        ''ContentWorksheet.Range("A25").Value = PortOfRegistry.ToString
        'ContentWorksheet.Range("A26").Value = Neat(BargeLength.ToString / 0.3048)
        'ContentWorksheet.Range("A27").Value = Neat(BargeBreadth.ToString / 0.3048)
        'ContentWorksheet.Range("A28").Value = Neat(BargeDepth.ToString / 0.3048)
        ''ContentWorksheet.Range("A29").Value = BallastName.ToString
        'ContentWorksheet.Range("A30").Value = Neat(LightshipMass.Text * 0.0009842)
        'ContentWorksheet.Range("A31").Value = Neat(LightshipLCG.Text / 0.3048)
        'ContentWorksheet.Range("A32").Value = Neat(LightshipTCG.Text / 0.3048)
        'ContentWorksheet.Range("A33").Value = Neat(LightshipVCG.Text / 0.3048)
        ''ContentWorksheet.Range("A34").Value = OffsetTopAngle.ToString
        'ContentWorksheet.Range("A35").Value = Neat(BoomHingetoCraneCL.ToString / 0.3048)
        'ContentWorksheet.Range("A36").Value = Neat(CrawlerBasetoBoomHinge.ToString / 0.3048)
        ''ContentWorksheet.Range("A37").Value = SideFramesExtended?
        'ContentWorksheet.Range("A38").Value = Neat(PositionofCraneCentretoFrontofTrack.ToString / 0.3048)
        'ContentWorksheet.Range("A39").Value = ListChartComboBox.Items(0).ToString
        'ContentWorksheet.Range("A40").Value = Neat(DistanceFromEndTextbox.Text / 0.3048)
        ''ContentWorksheet.Range("A41").Value = Neat(DistanceFromEndTextbox.Text / 0.3048)
        'ContentWorksheet.Range("A41").Value = Neat(CraneWAC(0, 0) * 0.0009842)
        'ContentWorksheet.Range("A42").Value = Neat(CraneWAC(0, 1) / 0.3048)
        'ContentWorksheet.Range("A43").Value = Neat(CraneWAC(0, 2) / 0.3048)
        'ContentWorksheet.Range("A44").Value = Neat(CraneWAC(0, 3) / 0.3048)

        'ContentWorksheet.Range("A45").Value = Neat(CraneWAC(1, 0) * 0.0009842)
        'ContentWorksheet.Range("A46").Value = Neat(CraneWAC(1, 1) / 0.3048)
        'ContentWorksheet.Range("A47").Value = Neat(CraneWAC(1, 2) / 0.3048)
        'ContentWorksheet.Range("A48").Value = Neat(CraneWAC(1, 3) / 0.3048)

        'ContentWorksheet.Range("A49").Value = Neat(CraneWAC(2, 0) * 0.0009842)
        'ContentWorksheet.Range("A50").Value = Neat(CraneWAC(2, 1) / 0.3048)
        'ContentWorksheet.Range("A51").Value = Neat(CraneWAC(2, 2) / 0.3048)
        'ContentWorksheet.Range("A52").Value = Neat(CraneWAC(2, 3) / 0.3048)


        ResultsWorksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible
        ResultsWorksheet.Activate()



        ''Unload frm

        'Worksheets(SheetName).Shapes(1).Delete()

        'With Worksheets(SheetName)
        '    .Cells.Copy()
        '    .Cells.PasteSpecial(xlPasteValues)
        'End With

        'ActiveWorkbook.Save()

    End Sub

    Private Sub FillPositionArray()
        DataSetPosition.Clear()
        DataSetPosition.ReadXml(My.Settings.StandardCranePositionsSetting)
        DataGridViewPosition.DataSource = DataSetPosition
        DataGridViewPosition.DataMember = "Row"
        PositionsComboBox.Items.Clear()
        Dim i As Integer
        For i = 0 To DataSetPosition.Tables(0).Rows.Count - 1
            PositionsComboBox.Items.Add(DataSetPosition.Tables(0).Rows(i).Item(0).ToString)
        Next

        PositionsComboBox.Text = DataSetPosition.Tables(0).Rows(0).Item(0).ToString
        DistanceFromEndTextbox.Text = DataSetPosition.Tables(0).Rows(0).Item(1).ToString
        PositionofCraneCentreOffCentrelineTextbox.Text = DataSetPosition.Tables(0).Rows(0).Item(2).ToString

    End Sub

    Private Sub FillCriteriaArray()
        DataSetCriteria.Clear()
        CriteriaComboBox.Items.Add("Criteria.xml")
        CriteriaComboBox.Text = "Criteria.xml"


    End Sub

    Private Function Forecast(TotalMass As Double, yy As Double(), xx As Double()) As Object
        Dim y As Double
        Dim m As Double
        Dim b As Double
        m = (yy(1) - yy(0)) / (xx(1) - xx(0))
        b = yy(1) - m * xx(1)
        y = m * TotalMass + b
        Return y
    End Function

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles AllAroundRadio.CheckedChanged

        FillRotationArray(My.Settings.AllAroundRotationArraySetting)
        'DrawPositions(BargeLength.Text, DataSetPosition, DataSetRotation, DistanceFromEndTextbox.Text)
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles OverEndRadio.CheckedChanged

        FillRotationArray(My.Settings.OverEndRotationArraySetting)
        'DrawPositions(BargeLength.Text, DataSetPosition, DataSetRotation, DistanceFromEndTextbox.Text)
    End Sub

    Private Sub LoadPositionTabButton_Click(sender As Object, e As EventArgs) Handles LoadPositionTabButton.Click
        'FillPositionArray()
        Me.TabControl2.SelectedIndex = 4
    End Sub

    Private Sub FillRotationArray(p1 As String)
        DataSetRotation.Clear()
        DataSetRotation.ReadXml(p1)
        DataGridViewRotation.DataSource = DataSetRotation
        DataGridViewRotation.DataMember = "Row"
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

    Private Sub ChooseBargeComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ChooseBargeComboBox.SelectedIndexChanged

        LoadBarge(ChooseBargeComboBox.SelectedIndex)
        DisplayBargeCharts()
        IllustrateSetup()

    End Sub

    Private Sub ChooseCraneComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ChooseCraneComboBox.SelectedIndexChanged

        LoadCrane(ChooseCraneComboBox.SelectedIndex)
        DrawBarge(BargeLength.Text, BargeBreadth.Text, BargeDepth.Text)
        DrawCrane(BargeLength.Text, DistanceFromEndTextbox.Text)

    End Sub

    Private Sub LoadRotationTabButton_Click(sender As Object, e As EventArgs) Handles LoadRotationTabButton.Click
        Me.TabControl2.SelectedIndex = 3
    End Sub

    Private Sub StepBargeBack_Click(sender As Object, e As EventArgs) Handles StepBargeBack.Click
        If ChooseBargeComboBox.SelectedIndex > 0 Then ChooseBargeComboBox.SelectedIndex = ChooseBargeComboBox.SelectedIndex - 1
    End Sub

    Private Sub StepBargeAhead_Click(sender As Object, e As EventArgs) Handles StepBargeAhead.Click
        If ChooseBargeComboBox.SelectedIndex + 1 < ChooseBargeComboBox.Items.Count Then ChooseBargeComboBox.SelectedIndex = ChooseBargeComboBox.SelectedIndex + 1
    End Sub

    Private Sub HydrostaticsChart_Click(sender As Object, e As EventArgs) Handles HydrostaticsChart.Click
        Form2.Show()
        If ChooseBargeComboBox.SelectedIndex >= 0 Then
            Form2.DataGridView1.DataSource = DataSetH
            Form2.DataGridView1.DataMember = "Hydrostatics"
            Form2.DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
            Form2.DataGridView1.AllowUserToAddRows = False
            Form2.DataGridView1.AllowUserToDeleteRows = False
            Form2.DataGridView1.AllowUserToResizeColumns = False
        Else
            MsgBox("Error: Please select barge.")
        End If
    End Sub

    Private Sub CrossCurvesChart_Click(sender As Object, e As EventArgs) Handles CrossCurvesChart.Click
        Form3.Show()
        If ChooseBargeComboBox.SelectedIndex >= 0 Then
            Form3.DataGridView1.DataSource = DataTableCCondensed
            Form3.DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
            Form3.DataGridView1.AllowUserToAddRows = False
            Form3.DataGridView1.AllowUserToDeleteRows = False
            Form3.DataGridView1.AllowUserToResizeColumns = False
        Else
            MsgBox("Error: Please select barge.")
        End If
    End Sub

    Private Sub StepCraneBack_Click(sender As Object, e As EventArgs) Handles StepCraneBack.Click
        If ChooseCraneComboBox.SelectedIndex > 0 Then ChooseCraneComboBox.SelectedIndex = ChooseCraneComboBox.SelectedIndex - 1
    End Sub

    Private Sub StepCraneAhead_Click(sender As Object, e As EventArgs) Handles StepCraneAhead.Click
        If ChooseCraneComboBox.SelectedIndex + 1 < ChooseCraneComboBox.Items.Count Then ChooseCraneComboBox.SelectedIndex = ChooseCraneComboBox.SelectedIndex + 1
    End Sub

    Private Function PlaceCranePoint(MassName As String, Mass As String, LCG As String, TCG As String, VCG As String, BoomHingetoCraneCL As String, CrawlerBasetoBoomHinge As String, PositionofCraneCentretoFrontofTrack As String) As Point
        Dim CentrePoint As Point

        Dim DeltaX As Integer = 1092 - 989
        Dim DeltaY As Integer = 153 - 210

        Dim RatioX As Double = DeltaX / PositionofCraneCentretoFrontofTrack
        Dim RatioY As Double = DeltaY / PositionofCraneCentretoFrontofTrack

        CentrePoint.X = 989 + RatioX * CDbl(LCG)
        CentrePoint.Y = 210 + RatioY * CDbl(VCG)

        Return CentrePoint

    End Function

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        If ConfigComboBox.SelectedIndex + 1 < ConfigComboBox.Items.Count Then ConfigComboBox.SelectedIndex = ConfigComboBox.SelectedIndex + 1
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        If ConfigComboBox.SelectedIndex > 0 Then ConfigComboBox.SelectedIndex = ConfigComboBox.SelectedIndex - 1
    End Sub

    Private Sub ConfigComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ConfigComboBox.SelectedIndexChanged

        LoadConfig(ConfigComboBox.SelectedIndex)

    End Sub

    Private Sub ListChartComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListChartComboBox.SelectedIndexChanged

        MachineListIndex = ListChartComboBox.SelectedIndex


        'If MachineListIndex = ListChartComboBox.Items.Count - 1 Then
        '    CreateNewListRoutine()
        'Else

        Dim t As Integer
        Dim LandChartSeries As New DataVisualization.Charting.Series

        If MachineListIndex >= 0 Then
            Chart1.Series.Clear()
            LandChartSeries.Name = "LandChartSeries"
            LandChartSeries.ChartType = DataVisualization.Charting.SeriesChartType.Point
            For t = 0 To xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1>.Count - 1
                LandChartSeries.Points.AddXY(xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1>(t).<OppRad>.Value, xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1>(t).<Capacity>.Value)
            Next
            Chart1.Series.Add(LandChartSeries)



            MachineListTextbox.Text = "Chart" & xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<MaxMachineList>.Value

        End If
        'End If
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        If ListChartComboBox.SelectedIndex + 1 < ListChartComboBox.Items.Count Then ListChartComboBox.SelectedIndex = ListChartComboBox.SelectedIndex + 1
    End Sub

    Private Sub Label58_Click(sender As Object, e As EventArgs) Handles Label58.Click
        If ListChartComboBox.SelectedIndex > 0 Then ListChartComboBox.SelectedIndex = ListChartComboBox.SelectedIndex - 1
    End Sub

    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click
        Form5.Show()

        If ListChartComboBox.SelectedIndex >= 0 Then
            Dim ListChart = From st In xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(MachineListIndex).<Row1> _
    Select New With { _
    .OppRad = st.<OppRad>.Value, _
    .BoomAngle = st.<BoomAngle>.Value, _
    .BoomPoint = st.<BoomPoint>.Value, _
    .Capacity = st.<Capacity>.Value
    }

            Form5.DataGridViewListChart.DataSource = ListChart.ToList
        Else
            MsgBox("Error: Please select list chart.")
        End If

    End Sub

    Private Sub ShowWeights_Click(sender As Object, e As EventArgs) Handles UpperworksCentreShape.Click, CrawlersCentreShape.Click, CTWTCentreShape.Click
        Form4.Show()

        If ChooseCraneComboBox.SelectedIndex >= 0 Then

            Dim WAC = From st In xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3> _
    Select New With { _
    .Name = st.<Name>.Value, _
    .Mass = st.<Mass>.Value, _
    .LCG = st.<LCG>.Value, _
    .TCG = st.<TCG>.Value, _
    .VCG = st.<VCG>.Value
    }

            Dim BoomLengthWeight = From st In xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2> _
                Select New With { _
                .SampleBoomLength = st.<BoomLength>.Value, _
                .SampleBoomWeight = st.<BoomWeight>.Value
                }

            Form4.DataGridViewWAC.DataSource = WAC.ToList
            Form4.DataGridViewBoomLengthWeight.DataSource = BoomLengthWeight.ToList

        Else
            MsgBox("Error: Please select crane.")
        End If
    End Sub

    Private Function Neat(p1 As String) As Object
        Return Math.Round(CDbl(p1), 2)
    End Function

    Private Sub PositionsComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PositionsComboBox.SelectedIndexChanged
        If PositionsComboBox.SelectedIndex >= 0 Then
            DistanceFromEndTextbox.Text = DataSetPosition.Tables(0).Rows(PositionsComboBox.SelectedIndex).Item(1).ToString
            PositionofCraneCentreOffCentrelineTextbox.Text = DataSetPosition.Tables(0).Rows(PositionsComboBox.SelectedIndex).Item(2).ToString
        End If
    End Sub

    Private Sub CriteriaComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CriteriaComboBox.SelectedIndexChanged
        DataSetCriteria.ReadXml(My.Settings.CriteriaFolderSetting & CriteriaComboBox.SelectedItem.ToString)
        DataGridViewCriteria.DataSource = DataSetCriteria
        DataGridViewCriteria.DataMember = "Row"
    End Sub

    Private Sub FillEventLog()
        DataSetEventLog.Clear()
        DataSetEventLog.ReadXml(My.Settings.EventLogSetting)
        DataGridViewEventLog.DataSource = DataSetEventLog
        DataGridViewEventLog.DataMember = "Row"
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataSetRotation.Tables(0).AcceptChanges()
        DataSetRotation.AcceptChanges()

        If AllAroundRadio.Checked = True Then
            DataSetRotation.WriteXml(My.Settings.AllAroundRotationArraySetting)
        Else
            DataSetRotation.WriteXml(My.Settings.OverEndRotationArraySetting)
        End If

        Dim row As DataRow

        row = DataSetEventLog.Tables(0).NewRow
        row.Item(0) = Now.ToOADate
        row.Item(1) = "RotationSetChanged"
        DataSetEventLog.Tables(0).Rows.Add(row)
        DataSetEventLog.WriteXml(My.Settings.EventLogSetting)


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        DataSetPosition.Tables(0).AcceptChanges()
        DataSetPosition.AcceptChanges()

        DataSetPosition.WriteXml(My.Settings.StandardCranePositionsSetting)

        Dim row As DataRow

        row = DataSetEventLog.Tables(0).NewRow
        row.Item(0) = Now.ToOADate
        row.Item(1) = "PositionSetChanged"
        DataSetEventLog.Tables(0).Rows.Add(row)
        DataSetEventLog.WriteXml(My.Settings.EventLogSetting)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        DataSetCriteria.Tables(0).AcceptChanges()
        DataSetCriteria.AcceptChanges()

        DataSetCriteria.WriteXml(My.Settings.CriteriaFolderSetting & "Criteria.xml")

        Dim row As DataRow

        row = DataSetEventLog.Tables(0).NewRow
        row.Item(0) = Now.ToOADate
        row.Item(1) = "CriteriaSetChanged"
        DataSetEventLog.Tables(0).Rows.Add(row)
        DataSetEventLog.WriteXml(My.Settings.EventLogSetting)

    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        End
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click

        SaveFileDialog1.CreatePrompt = True
        SaveFileDialog1.OverwritePrompt = True
        SaveFileDialog1.FileName = "Setup" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "-" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now)
        SaveFileDialog1.DefaultExt = "xml"
        SaveFileDialog1.AddExtension = True
        SaveFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        SaveFileDialog1.ShowDialog()
        Dim xSetup As XElement
        xSetup = XElement.Load(My.Settings.SetupSetting)
        xSetup...<CurrentRowBarge>.Value = CurrentRowBarge
        xSetup...<CurrentRowCrane>.Value = CurrentRowCrane
        xSetup...<ConfigIndex>.Value = ConfigIndex
        xSetup...<MachineListIndex>.Value = MachineListIndex
        xSetup.Save(SaveFileDialog1.FileName)


    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        OpenFileDialog1.FileName = ""
        OpenFileDialog1.DefaultExt = "xml"
        OpenFileDialog1.AddExtension = True
        OpenFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()

        Dim xSetup As XElement
        xSetup = XElement.Load(OpenFileDialog1.FileName)
        ChooseBargeComboBox.SelectedIndex = xSetup...<CurrentRowBarge>.Value
        ChooseCraneComboBox.SelectedIndex = xSetup...<CurrentRowCrane>.Value
        ConfigComboBox.SelectedIndex = xSetup...<ConfigIndex>.Value
        ListChartComboBox.SelectedIndex = xSetup...<MachineListIndex>.Value
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        MsgBox("Not yet implemented.")
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        MsgBox("Not yet implemented.")
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        MsgBox("Not yet implemented.")
    End Sub

    Private Sub CopyToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem1.Click
        MsgBox("Not yet implemented.")
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        MsgBox("Not yet implemented.")
    End Sub

    Private Sub LoadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadToolStripMenuItem.Click
        Form6.Show()
        Form6.CraneListFileSettingtb.Text = My.Settings.CraneListFileSetting
        Form6.BargeListFileSettingtb.Text = My.Settings.BargeListFileSetting
        Form6.AllAroundRotationArraySettingtb.Text = My.Settings.AllAroundRotationArraySetting
        Form6.EventLogSettingtb.Text = My.Settings.EventLogSetting
        Form6.StandardCranePositionsSettingtb.Text = My.Settings.StandardCranePositionsSetting
        Form6.CriteriaFolderSettingtb.Text = My.Settings.CriteriaFolderSetting
        Form6.OverEndRotationArraySettingtb.Text = My.Settings.OverEndRotationArraySetting
        Form6.SimpsonMultipliersSettingtb.Text = My.Settings.SimpsonMultipliersSetting
        Form6.SetupSettingtb.Text = My.Settings.SetupSetting
        Form6.XLSXTemplateSettingtb.Text = My.Settings.XLSXTemplateSetting
        Form6.HydrostaticsFolderSettingtb.Text = My.Settings.HydrostaticsFolderSetting
        Form6.CrossCurveFolderSettingtb.Text = My.Settings.CrossCurveFolderSetting
    End Sub

    Private Sub ViewHelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewHelpToolStripMenuItem.Click
        MsgBox("Not yet impletemented.")
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("Not yet impletemented.")
    End Sub

    Public Sub DrawBarge(BLength As String, BBreadth As String, BDepth As String)

        Dim BargeSize As Size
        Dim PointShift As Point

        If ChooseBargeComboBox.SelectedIndex >= 0 Then

            BargeSize.Height = 10 * BargeDepth.Text
            BargeSize.Width = 10 * BargeLength.Text
            BargeShape.Size = BargeSize

            BargeSize.Height = 10 * BargeBreadth.Text
            BargeSize.Width = 10 * BargeLength.Text
            BargePlanShape.Size = BargeSize

            PointShift.X = 335 - 300 + 10 * BargeLength.Text
            PointShift.Y = 140
            FwdRakeShape.EndPoint = PointShift

            PointShift.X = 385 - 300 + 10 * BargeLength.Text
            PointShift.Y = 110
            FwdRakeShape.StartPoint = PointShift
        End If
    End Sub

    Public Sub DrawCrane(BLength As String, DistanceFromEnd As String)
        Dim PointShift As Point
        If ChooseBargeComboBox.SelectedIndex >= 0 Then
            PointShift.X = 294 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = CrawlerAftShape.Location.Y
            CrawlerAftShape.Location = PointShift

            PointShift.X = 332 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = CrawlerFwdShape.Location.Y
            CrawlerFwdShape.Location = PointShift

            PointShift.X = 304 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = CrawlerCentreShape.Location.Y
            CrawlerCentreShape.Location = PointShift

            PointShift.X = 284 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = UpperworksShape.Location.Y
            UpperworksShape.Location = PointShift

            PointShift.X = 327 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = UpperworksCabShape.Location.Y
            UpperworksCabShape.Location = PointShift

            PointShift.X = 329 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = UpperworksCabWindowShape.Location.Y
            UpperworksCabWindowShape.Location = PointShift

            PointShift.X = 275 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = CounterweightShape.Location.Y
            CounterweightShape.Location = PointShift

            PointShift.X = 266 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = CounterweightAftShape.Location.Y
            CounterweightAftShape.Location = PointShift

            PointShift.X = 300 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = 184 + (BargeBreadth.Text * 10 - 100) / 2
            CrawlerPlanPortShape.Location = PointShift

            PointShift.X = 300 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = 206 + (BargeBreadth.Text * 10 - 100) / 2
            CrawlerPlanStbdShape.Location = PointShift

            PointShift.X = 266 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = CounterweightAftShape.Location.Y
            CounterweightAftShape.Location = PointShift

            PointShift.X = 409 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = BoomShape.StartPoint.Y
            BoomShape.StartPoint = PointShift

            PointShift.X = 347 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = BoomShape.EndPoint.Y
            BoomShape.EndPoint = PointShift
        End If
    End Sub

    Public Sub DrawPositions(BLength As String, dataSet As DataSet, dataSetRot As DataSet, DistanceFromEnd As String)
        Dim PointShift As Point


        If BargeLength.Text > 1 And DistanceFromEndTextbox.Text >= 0 Then

            If dataSet.Tables(0).Rows.Count >= 1 Then
                PointShift.X = 370 + (10 * BLength - 300) - dataSet.Tables(0).Rows(0).Item(1).ToString * 10
                PointShift.Y = Position1Shape.StartPoint.Y
                Position1Shape.StartPoint = PointShift
                PointShift.Y = Position1Shape.EndPoint.Y
                Position1Shape.EndPoint = PointShift
            End If
            If dataSet.Tables(0).Rows.Count >= 2 Then
                PointShift.X = 370 + (10 * BLength - 300) - dataSet.Tables(0).Rows(1).Item(1).ToString * 10
                PointShift.Y = Position1Shape.StartPoint.Y
                Position2Shape.StartPoint = PointShift
                PointShift.Y = Position1Shape.EndPoint.Y
                Position2Shape.EndPoint = PointShift
            End If
            If dataSet.Tables(0).Rows.Count >= 3 Then
                PointShift.X = 370 + (10 * BLength - 300) - dataSet.Tables(0).Rows(2).Item(1).ToString * 10
                PointShift.Y = Position1Shape.StartPoint.Y
                Position3Shape.StartPoint = PointShift
                PointShift.Y = Position1Shape.EndPoint.Y
                Position3Shape.EndPoint = PointShift
            End If

            PointShift.X = 345 + (10 * BLength - 300) - DistanceFromEnd * 10 + 20
            PointShift.Y = 200 + (BargeBreadth.Text * 10 - 100) / 2

            BoomPlan1Shape.StartPoint = PointShift
            BoomPlan2Shape.StartPoint = PointShift
            BoomPlan3Shape.StartPoint = PointShift
            BoomPlan4Shape.StartPoint = PointShift
            BoomPlan5Shape.StartPoint = PointShift
            BoomPlan6Shape.StartPoint = PointShift
            BoomPlan7Shape.StartPoint = PointShift
            BoomPlan8Shape.StartPoint = PointShift
            BoomPlan9Shape.StartPoint = PointShift
            BoomPlan10Shape.StartPoint = PointShift
            BoomPlan11Shape.StartPoint = PointShift

            BoomPlan1Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(0).Item(0).ToString)
            BoomPlan2Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(1).Item(0).ToString)
            BoomPlan3Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(2).Item(0).ToString)
            BoomPlan4Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(3).Item(0).ToString)
            BoomPlan5Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(4).Item(0).ToString)
            BoomPlan6Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(5).Item(0).ToString)
            BoomPlan7Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(6).Item(0).ToString)
            BoomPlan8Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(7).Item(0).ToString)
            BoomPlan9Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(8).Item(0).ToString)
            BoomPlan10Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(9).Item(0).ToString)
            BoomPlan11Shape.EndPoint = PointShiftFunction(dataSetRot.Tables(0).Rows(10).Item(0).ToString)
        End If
    End Sub

    Public Function PointShiftFunction(MaxAngle As Double) As Point
        Dim PointShift As Point
        PointShift.X = 345 + (10 * BargeLength.Text - 300) - DistanceFromEndTextbox.Text * 10 + 20 + 90 * Math.Cos(MaxAngle * 3.1415 / 180)
        PointShift.Y = 200 + (BargeBreadth.Text * 10 - 100) / 2 + 90 * Math.Sin(MaxAngle * 3.1415 / 180)
        Return PointShift
    End Function

    Private Sub CreateNewBargeRoutine()

        Dim NewBarge As XElement
        Dim BargeList As XElement
        Dim placeholder As Integer = 0

        BargeList = XElement.Load(My.Settings.BargeListFileSetting)

        NewBarge = <Barge>
                       <BargeName><%= placeholder %></BargeName>
                       <BargeLength><%= placeholder %></BargeLength>
                       <BargeBreadth><%= placeholder %></BargeBreadth>
                       <BargeDepth><%= placeholder %></BargeDepth>
                       <Calibration><%= placeholder %></Calibration>
                       <BallastDescription><%= placeholder %></BallastDescription>
                       <HydrostaticDataTab><%= "9999H.xml" %></HydrostaticDataTab>
                       <CrossCurveDataTab><%= "9999C.xml" %></CrossCurveDataTab>
                       <LightshipMass><%= placeholder %></LightshipMass>
                       <LightshipLCG><%= placeholder %></LightshipLCG>
                       <LightshipTCG><%= placeholder %></LightshipTCG>
                       <LightshipVCG><%= placeholder %></LightshipVCG>
                       <BargeOwner><%= placeholder %></BargeOwner>
                       <CranePadThickness><%= placeholder %></CranePadThickness>
                       <BallastMass><%= placeholder %></BallastMass>
                       <BallastLCG><%= placeholder %></BallastLCG>
                       <BallastTCG><%= placeholder %></BallastTCG>
                       <BallastVCG><%= placeholder %></BallastVCG>
                       <AftRake><%= False %></AftRake>
                       <FwdRake><%= False %></FwdRake>
                       <HouseworksAft><%= False %></HouseworksAft>
                       <PlanView><%= "9999plan.bmp" %></PlanView>
                       <ProfileView><%= "9999profile.bmp" %></ProfileView>
                   </Barge>

        BargeList.Add(NewBarge)
        BargeList.Save(My.Settings.BargeListFileSetting)

        GeneralReset()

        'ChooseBargeComboBox.SelectedIndex = ChooseBargeComboBox.Items.Count - 1

    End Sub

    Private Sub CreateNewCraneRoutine()
        Dim NewCrane As XElement

        Dim placeholder As Integer = 99999

        NewCrane =
        <Crane>
            <ProposedCrane><%= placeholder %></ProposedCrane>
            <CraneSerialNumber><%= placeholder %></CraneSerialNumber>
            <BoomHingetoCraneCL><%= placeholder %></BoomHingetoCraneCL>
            <CrawlerBasetoBoomHinge><%= placeholder %></CrawlerBasetoBoomHinge>
            <PositionofCraneCentretoFrontofTrack><%= placeholder %></PositionofCraneCentretoFrontofTrack>
            <CraneOwner><%= placeholder %></CraneOwner>
            <WAC>
                <Row3>
                    <Name><%= "CounterweightWAC" %></Name>
                    <Mass><%= placeholder %></Mass>
                    <LCG><%= placeholder %></LCG>
                    <TCG><%= placeholder %></TCG>
                    <VCG><%= placeholder %></VCG>
                </Row3>
                <Row3>
                    <Name><%= "UpperworksWAC" %></Name>
                    <Mass><%= placeholder %></Mass>
                    <LCG><%= placeholder %></LCG>
                    <TCG><%= placeholder %></TCG>
                    <VCG><%= placeholder %></VCG>
                </Row3>
                <Row3>
                    <Name><%= "CrawlersWAC" %></Name>
                    <Mass><%= placeholder %></Mass>
                    <LCG><%= placeholder %></LCG>
                    <TCG><%= placeholder %></TCG>
                    <VCG><%= placeholder %></VCG>
                </Row3>
            </WAC>
            <BoomLengthWeight>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
                <Row2>
                    <BoomLength><%= placeholder %></BoomLength>
                    <BoomWeight><%= placeholder %></BoomWeight>
                </Row2>
            </BoomLengthWeight>
            <Config>
                <BoomType><%= placeholder %></BoomType>
                <Counterweight><%= placeholder %></Counterweight>
                <BoomLength><%= placeholder %></BoomLength>
                <BoomName><%= placeholder %></BoomName>
                <ListChart>
                    <MaxMachineList><%= placeholder %></MaxMachineList>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                    <Row1>
                        <OppRad><%= placeholder %></OppRad>
                        <BoomAngle><%= placeholder %></BoomAngle>
                        <BoomPoint><%= placeholder %></BoomPoint>
                        <Capacity><%= placeholder %></Capacity>
                    </Row1>
                </ListChart>
            </Config>
        </Crane>

        xCrane.Add(NewCrane)
        xCrane.Save(My.Settings.CraneListFileSetting)

        'xCrane = XElement.Load(My.Settings.CraneListFileSetting)

        GeneralReset()

        ChooseCraneComboBox.SelectedIndex = ChooseCraneComboBox.Items.Count - 1
        ConfigComboBox.SelectedIndex = 0
        ListChartComboBox.SelectedIndex = 0

    End Sub

    Private Sub CreateNewConfigRoutine()
        MsgBox("createconfig")
    End Sub

    Private Sub CreateNewListRoutine()
        MsgBox("createlist")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CreateNewBargeRoutine()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        CreateNewCraneRoutine()
    End Sub

    Private Sub GeneralReset()
        'TextBox1.Text = My.Settings.CraneListFileSetting
        'TextBox2.Text = My.Settings.BargeListFileSetting

        DataSetBarge.Clear()
        DataSetBarge.ReadXml(My.Settings.BargeListFileSetting)


        'ConfigComboBox.SelectedItem = "Config1"
        'ListChartComboBox.SelectedItem = "Chart0"

        'AllAroundRotation = True
        FillRotationArray(My.Settings.AllAroundRotationArraySetting)
        FillPositionArray()
        FillCriteriaArray()
        FillComboBoxes()
        FillEventLog()

    End Sub

    Private Sub SaveBargeChangesButton_Click(sender As Object, e As EventArgs) Handles SaveBargeChangesButton.Click

        ImplementBargeSave(BargeName.Text, 0)
        ImplementBargeSave(BargeLength.Text, 1)
        ImplementBargeSave(BargeBreadth.Text, 2)
        ImplementBargeSave(BargeDepth.Text, 3)
        ImplementBargeSave(Calibration.Text, 4)
        ImplementBargeSave(BallastDescription.Text, 5)
        ImplementBargeSave(HydrostaticsDataTag.Text, 6)
        ImplementBargeSave(CrossCurvesDataTag.Text, 7)
        ImplementBargeSave(LightshipMass.Text, 8)
        ImplementBargeSave(LightshipLCG.Text, 9)
        ImplementBargeSave(LightshipTCG.Text, 10)
        ImplementBargeSave(LightshipVCG.Text, 11)
        'ImplementBargeSave(BargeOwner.Text, 12)
        ImplementBargeSave(CranePadThickness.Text, 13)
        ImplementBargeSave(BallastMass.Text, 14)
        ImplementBargeSave(BallastLCG.Text, 15)
        ImplementBargeSave(BallastTCG.Text, 16)
        ImplementBargeSave(BallastVCG.Text, 17)
        ImplementBargeSave(AftRakeCheckBox.Checked, 18)
        ImplementBargeSave(FwdRakeCheckBox.Checked, 19)
        ImplementBargeSave(HouseAftCheckBox.Checked, 20)
        ImplementBargeSave(PlanLabel.Text, 21)
        ImplementBargeSave(ProfileLabel.Text, 22)

        GeneralReset()
        ChooseBargeComboBox.SelectedIndex = CurrentRowBarge

    End Sub

    Private Sub ImplementBargeSave(NewValue As String, i As Integer)
        Dim XMLString(22) As String
        XMLString = {"BargeName", "BargeLength", "BargeBreadth", "BargeDepth", "Calibration", "BallastDescription", "HydrostaticDataTab", "CrossCurveDataTab", "LightshipMass", "LightshipLCG", "LightshipTCG", "LightshipVCG", "BargeOwner", "CranePadThickness", "BallastMass", "BallastLCG", "BallastTCG", "BallastVCG", "AftRake", "FwdRake", "HouseworksAft", "PlanView", "ProfileView"}
        Dim row As DataRow
        If DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(i).ToString <> NewValue And NewValue <> "" Then
            MsgBox(XMLString(i) & " has changed from " & DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(i).ToString & " to " & NewValue)
            row = DataSetEventLog.Tables(0).NewRow
            row.Item(0) = Now.ToOADate
            row.Item(1) = BargeNameTextbox.Text & "." & XMLString(i) & "Change(" & DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(i).ToString & "," & NewValue & ")"
            DataSetEventLog.Tables(0).Rows.Add(row)
            DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
            FillEventLog()
            DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(XMLString(i)) = NewValue
            DataSetBarge.WriteXml(My.Settings.BargeListFileSetting)
        Else
        End If

    End Sub

    Private Sub AftRakeCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles AftRakeCheckBox.CheckedChanged

        If AftRakeCheckBox.Checked Then
            AftRakeShape.Visible = True
        Else
            AftRakeShape.Visible = False
        End If

    End Sub

    Private Sub FwdRakeCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles FwdRakeCheckBox.CheckedChanged
        If FwdRakeCheckBox.Checked Then
            FwdRakeShape.Visible = True
        Else
            FwdRakeShape.Visible = False
        End If
    End Sub

    Private Sub HouseAftCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles HouseAftCheckBox.CheckedChanged
        If HouseAftCheckBox.Checked Then
            HouseShape.Visible = True
        Else
            HouseShape.Visible = False
        End If
    End Sub

    Private Sub SaveCraneChangesButton_Click(sender As Object, e As EventArgs) Handles SaveCraneChangesButton.Click
        ImplementCraneSave()



        GeneralReset()
        ChooseCraneComboBox.SelectedIndex = CurrentRowCrane
    End Sub

    Private Sub ImplementCraneSave()
        Dim NewValue As String
        Dim ChangedParameter As String = "Counterweight"
        Dim row As DataRow

        NewValue = CounterweightTextbox.Text
        ChangedParameter = "Counterweight"

        If xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<Counterweight>.Value <> NewValue And NewValue <> "" Then
            MsgBox(ChangedParameter & " has changed from " & xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<Counterweight>.Value & " to " & NewValue)
            row = DataSetEventLog.Tables(0).NewRow
            row.Item(0) = Now.ToOADate
            row.Item(1) = CraneNameTextbox.Text & "." & ChangedParameter & "Change(" & xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<Counterweight>.Value & "," & NewValue & ")"
            DataSetEventLog.Tables(0).Rows.Add(row)
            DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
            FillEventLog()
            xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<Counterweight>.Value = NewValue
            xCrane.Save(My.Settings.CraneListFileSetting)
        Else
        End If

        ChangedParameter = "BoomLength"
        NewValue = BoomLengthTextbox.Text

        If xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomLength>.Value <> NewValue And NewValue <> "" Then
            MsgBox(ChangedParameter & " has changed from " & xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomLength>.Value & " to " & NewValue)

            row = DataSetEventLog.Tables(0).NewRow
            row.Item(0) = Now.ToOADate
            row.Item(1) = CraneNameTextbox.Text & "." & ChangedParameter & "Change(" & xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomLength>.Value & "," & NewValue & ")"
            DataSetEventLog.Tables(0).Rows.Add(row)
            DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
            FillEventLog()
            xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomLength>.Value = NewValue
            xCrane.Save(My.Settings.CraneListFileSetting)

        Else
        End If

        ChangedParameter = "BoomType"
        NewValue = BoomTypeTextbox.Text

        If xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomType>.Value <> NewValue And NewValue <> "" Then
            MsgBox(ChangedParameter & " has changed from " & xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomType>.Value & " to " & NewValue)
            row = DataSetEventLog.Tables(0).NewRow
            row.Item(0) = Now.ToOADate
            row.Item(1) = CraneNameTextbox.Text & "." & ChangedParameter & "Change(" & xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomType>.Value & "," & NewValue & ")"
            DataSetEventLog.Tables(0).Rows.Add(row)
            DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
            FillEventLog()
            xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomType>.Value = NewValue
            xCrane.Save(My.Settings.CraneListFileSetting)

        Else
        End If

        ChangedParameter = "CraneName"
        NewValue = ProposedCraneTextbox.Text


        If xCrane...<Crane>(CurrentRowCrane).<ProposedCrane>.Value <> NewValue And NewValue <> "" Then
            MsgBox(ChangedParameter & " has changed from " & xCrane...<Crane>(CurrentRowCrane).<ProposedCrane>.Value & " to " & NewValue)
            row = DataSetEventLog.Tables(0).NewRow
            row.Item(0) = Now.ToOADate
            row.Item(1) = CraneNameTextbox.Text & "." & ChangedParameter & "Change(" & xCrane...<Crane>(CurrentRowCrane).<ProposedCrane>.Value & "," & NewValue & ")"
            DataSetEventLog.Tables(0).Rows.Add(row)
            DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
            FillEventLog()
            xCrane...<Crane>(CurrentRowCrane).<ProposedCrane>.Value = NewValue
            xCrane.Save(My.Settings.CraneListFileSetting)

        Else
        End If
        ChangedParameter = "CraneSerialNumber"
        NewValue = CraneSerialNumberTextbox.Text

        If xCrane...<Crane>(CurrentRowCrane).<CraneSerialNumber>.Value <> NewValue And NewValue <> "" Then
            MsgBox(ChangedParameter & " has changed from " & xCrane...<Crane>(CurrentRowCrane).<CraneSerialNumber>.Value & " to " & NewValue)
            row = DataSetEventLog.Tables(0).NewRow
            row.Item(0) = Now.ToOADate
            row.Item(1) = CraneNameTextbox.Text & "." & ChangedParameter & "Change(" & xCrane...<Crane>(CurrentRowCrane).<CraneSerialNumber>.Value & "," & NewValue & ")"
            DataSetEventLog.Tables(0).Rows.Add(row)
            DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
            FillEventLog()
            xCrane...<Crane>(CurrentRowCrane).<CraneSerialNumber>.Value = NewValue
            xCrane.Save(My.Settings.CraneListFileSetting)

        Else
        End If


        ChangedParameter = "BoomHingetoCraneCL"
        NewValue = BoomHingetoCraneCLTextbox.Text


        If xCrane...<Crane>(CurrentRowCrane).<BoomHingetoCraneCL>.Value <> NewValue And NewValue <> "" Then
            MsgBox(ChangedParameter & " has changed from " & xCrane...<Crane>(CurrentRowCrane).<BoomHingetoCraneCL>.Value & " to " & NewValue)
            row = DataSetEventLog.Tables(0).NewRow
            row.Item(0) = Now.ToOADate
            row.Item(1) = CraneNameTextbox.Text & "." & ChangedParameter & "Change(" & xCrane...<Crane>(CurrentRowCrane).<BoomHingetoCraneCL>.Value & "," & NewValue & ")"
            DataSetEventLog.Tables(0).Rows.Add(row)
            DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
            FillEventLog()
            xCrane...<Crane>(CurrentRowCrane).<BoomHingetoCraneCL>.Value = NewValue
            xCrane.Save(My.Settings.CraneListFileSetting)

        Else
        End If


        ChangedParameter = "CrawlerBasetoBoomHinge"
        NewValue = CrawlerBasetoBoomHingeTextbox.Text


        If xCrane...<Crane>(CurrentRowCrane).<CrawlerBasetoBoomHinge>.Value <> NewValue And NewValue <> "" Then
            MsgBox(ChangedParameter & " has changed from " & xCrane...<Crane>(CurrentRowCrane).<CrawlerBasetoBoomHinge>.Value & " to " & NewValue)
            row = DataSetEventLog.Tables(0).NewRow
            row.Item(0) = Now.ToOADate
            row.Item(1) = CraneNameTextbox.Text & "." & ChangedParameter & "Change(" & xCrane...<Crane>(CurrentRowCrane).<CrawlerBasetoBoomHinge>.Value & "," & NewValue & ")"
            DataSetEventLog.Tables(0).Rows.Add(row)
            DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
            FillEventLog()
            xCrane...<Crane>(CurrentRowCrane).<CrawlerBasetoBoomHinge>.Value = NewValue
            xCrane.Save(My.Settings.CraneListFileSetting)

        Else
        End If


        ChangedParameter = "PositionofCraneCentretoFrontofTrack"
        NewValue = PositionofCraneCentretoFrontofTrackTextbox.Text


        If xCrane...<Crane>(CurrentRowCrane).<PositionofCraneCentretoFrontofTrack>.Value <> NewValue And NewValue <> "" Then
            MsgBox(ChangedParameter & " has changed from " & xCrane...<Crane>(CurrentRowCrane).<PositionofCraneCentretoFrontofTrack>.Value & " to " & NewValue)
            row = DataSetEventLog.Tables(0).NewRow
            row.Item(0) = Now.ToOADate
            row.Item(1) = CraneNameTextbox.Text & "." & ChangedParameter & "Change(" & xCrane...<Crane>(CurrentRowCrane).<PositionofCraneCentretoFrontofTrack>.Value & "," & NewValue & ")"
            DataSetEventLog.Tables(0).Rows.Add(row)
            DataSetEventLog.WriteXml(My.Settings.EventLogSetting)
            FillEventLog()
            xCrane...<Crane>(CurrentRowCrane).<PositionofCraneCentretoFrontofTrack>.Value = NewValue
            xCrane.Save(My.Settings.CraneListFileSetting)

        Else
        End If

    End Sub

    Private Sub ChooseBargeCalibrationComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ChooseBargeCalibrationComboBox.SelectedIndexChanged

        'Enter the survey date (attached the survey report?)
        'Enter the four freeboards
        'Identify the barge and its configuration
        'Identify its position on the barge
        'Identify the rotation of the crane
        'Identify the load (if any)
        'Identify the boom angle
        'Identify any ballast weights on or off

        'Calculate the heel
        'Calculate the fwd and aft draft

        'Calculate the equilibrium hydrostatics

        'Calculate the weights and centres of the crane
        'Calculate the weights and centres of the ballast

        'Deduct these from the "as surveyed condition"

        'Create a new lightship displacement and LCG/ TCG
        'calculate a VCG

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form7.Show()
    End Sub

    Public Sub LoadBarge(ChooseBargeComboBoxSelectedIndex)
        DataSetH.Clear()
        DataSetC.Clear()

        CurrentRowBarge = ChooseBargeComboBoxSelectedIndex

        'If CurrentRowBarge = ChooseBargeComboBox.Items.Count - 1 Then
        '    CreateNewBargeRoutine()

        'Else

        BargeLength.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(1).ToString
        BargeBreadth.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(2).ToString
        BargeDepth.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(3).ToString

        Calibration.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(4).ToString
        BallastDescription.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(5).ToString

        HydrostaticsDataTag.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(6).ToString
        CrossCurvesDataTag.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(7).ToString

        LightshipMass.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(8).ToString
        LightshipLCG.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(9).ToString
        LightshipTCG.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(10).ToString
        LightshipVCG.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(11).ToString
        CranePadThickness.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(13).ToString

        BallastMass.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(14).ToString
        BallastLCG.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(15).ToString
        BallastTCG.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(16).ToString
        BallastVCG.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(17).ToString


        AftRakeCheckBox.Checked = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(18).ToString
        FwdRakeCheckBox.Checked = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(19).ToString
        HouseAftCheckBox.Checked = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(20).ToString

        'PictureBox7.Load(DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(21).ToString)
        'PictureBox5.Load(DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(22).ToString)

        ProfileLabel.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(22).ToString
        PlanLabel.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(21).ToString


        DataSetH.ReadXml(My.Settings.HydrostaticsFolderSetting & DataTableBarge.Rows(CurrentRowBarge).Item(6).ToString)

        Dim t As Integer


        For t = 0 To HydrostaticsDataSize - 1

            DisplacementArray(t) = DataSetH.Tables(0).Rows(t).Item(0).ToString
            LCFDraftArray(t) = DataSetH.Tables(0).Rows(t).Item(4).ToString
            LCFArray(t) = DataSetH.Tables(0).Rows(t).Item(15).ToString
            MTCArray(t) = DataSetH.Tables(0).Rows(t).Item(25).ToString
            LCBArray(t) = DataSetH.Tables(0).Rows(t).Item(14).ToString
            FPDraftArray(t) = DataSetH.Tables(0).Rows(t).Item(2).ToString
            APDraftArray(t) = DataSetH.Tables(0).Rows(t).Item(3).ToString
            RMArray(t) = DataSetH.Tables(0).Rows(t).Item(26).ToString
            'Changed GMTArray to reference KMt in hydrostatics b21 changed to b23 and GML B22 --> B24 if this works eventually change variable name to KM
            GMTArray(t) = DataSetH.Tables(0).Rows(t).Item(22).ToString
            GMLArray(t) = DataSetH.Tables(0).Rows(t).Item(23).ToString

        Next


        DataSetC.ReadXml(My.Settings.CrossCurveFolderSetting & DataTableBarge.Rows(CurrentRowBarge).Item(7).ToString)
        DataTableC = DataSetC.Tables(0)

        BargeNameTextbox.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(0).ToString
        BargeName.Text = DataSetBarge.Tables(0).Rows(CurrentRowBarge).Item(0).ToString
        Label25.Text = CurrentRowBarge + 1 & " / " & ChooseBargeComboBox.Items.Count

        'End If
    End Sub

    Public Sub DisplayBargeCharts()
        HydrostaticsChart.Series.Clear()
        CrossCurvesChart.Series.Clear()

        Dim DraftSeries As New DataVisualization.Charting.Series
        Dim LCBSeries As New DataVisualization.Charting.Series

        Dim KN00Series As New DataVisualization.Charting.Series
        Dim KN10Series As New DataVisualization.Charting.Series
        Dim KN20Series As New DataVisualization.Charting.Series
        Dim KN30Series As New DataVisualization.Charting.Series
        Dim KN40Series As New DataVisualization.Charting.Series
        Dim KN50Series As New DataVisualization.Charting.Series
        Dim KN60Series As New DataVisualization.Charting.Series
        Dim KN70Series As New DataVisualization.Charting.Series
        Dim KN80Series As New DataVisualization.Charting.Series
        Dim KN90Series As New DataVisualization.Charting.Series

        Dim column As New DataColumn

        Dim t As Integer


        For t = 0 To HydrostaticsDataSize - 1


            DraftSeries.Points.AddXY(DisplacementArray(t), LCFDraftArray(t))
            LCBSeries.Points.AddXY(DisplacementArray(t), LCBArray(t))

        Next

        DraftSeries.Name = "LCF Draft"
        LCBSeries.Name = "LCB"
        KN00Series.Name = "KN @ 0 deg"
        KN10Series.Name = "KN @ 10 deg"
        KN20Series.Name = "KN @ 20 deg"
        KN30Series.Name = "KN @ 30 deg"
        KN40Series.Name = "KN @ 40 deg"
        KN50Series.Name = "KN @ 50 deg"
        KN60Series.Name = "KN @ 60 deg"
        KN70Series.Name = "KN @ 70 deg"
        KN80Series.Name = "KN @ 80 deg"
        KN90Series.Name = "KN @ 90 deg"

        DraftSeries.ChartType = DataVisualization.Charting.SeriesChartType.Point
        LCBSeries.ChartType = DataVisualization.Charting.SeriesChartType.Point

        HydrostaticsChart.Series.Add(DraftSeries)
        HydrostaticsChart.Series.Add(LCBSeries)


        KN00Series.ChartType = DataVisualization.Charting.SeriesChartType.Point
        KN10Series.ChartType = DataVisualization.Charting.SeriesChartType.Point
        KN20Series.ChartType = DataVisualization.Charting.SeriesChartType.Point
        KN30Series.ChartType = DataVisualization.Charting.SeriesChartType.Point
        KN40Series.ChartType = DataVisualization.Charting.SeriesChartType.Point
        KN50Series.ChartType = DataVisualization.Charting.SeriesChartType.Point
        KN60Series.ChartType = DataVisualization.Charting.SeriesChartType.Point
        KN70Series.ChartType = DataVisualization.Charting.SeriesChartType.Point
        KN80Series.ChartType = DataVisualization.Charting.SeriesChartType.Point
        KN90Series.ChartType = DataVisualization.Charting.SeriesChartType.Point

        For i = 0 To DataTableC.Rows.Count - 1
            DataTableCCondensed.Rows.Add()
            DataTableCCondensed.Rows(i).Item(0) = DataTableC.Rows(i).Item(0).ToString
            DataTableCCondensed.Rows(i).Item(1) = DataTableC.Rows(i).Item(1).ToString
            DataTableCCondensed.Rows(i).Item(2) = DataTableC.Rows(i).Item(2).ToString
            DataTableCCondensed.Rows(i).Item(3) = DataTableC.Rows(i).Item(102).ToString
            DataTableCCondensed.Rows(i).Item(4) = DataTableC.Rows(i).Item(202).ToString
            DataTableCCondensed.Rows(i).Item(5) = DataTableC.Rows(i).Item(302).ToString
            DataTableCCondensed.Rows(i).Item(6) = DataTableC.Rows(i).Item(402).ToString
            DataTableCCondensed.Rows(i).Item(7) = DataTableC.Rows(i).Item(502).ToString
            DataTableCCondensed.Rows(i).Item(8) = DataTableC.Rows(i).Item(602).ToString
            DataTableCCondensed.Rows(i).Item(9) = DataTableC.Rows(i).Item(702).ToString
            DataTableCCondensed.Rows(i).Item(10) = DataTableC.Rows(i).Item(802).ToString
            DataTableCCondensed.Rows(i).Item(11) = DataTableC.Rows(i).Item(902).ToString
            KN00Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(2).ToString)
            KN10Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(102).ToString)
            KN20Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(202).ToString)
            KN30Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(302).ToString)
            KN40Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(402).ToString)
            KN50Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(502).ToString)
            KN60Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(602).ToString)
            KN70Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(702).ToString)
            KN80Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(802).ToString)
            KN90Series.Points.AddXY(DataTableC.Rows(i).Item(0).ToString, DataTableC.Rows(i).Item(902).ToString)

        Next

        CrossCurvesChart.Series.Add(KN00Series)
        CrossCurvesChart.Series.Add(KN10Series)
        CrossCurvesChart.Series.Add(KN20Series)
        CrossCurvesChart.Series.Add(KN30Series)
        CrossCurvesChart.Series.Add(KN40Series)
        CrossCurvesChart.Series.Add(KN50Series)
        CrossCurvesChart.Series.Add(KN60Series)
        CrossCurvesChart.Series.Add(KN70Series)
        CrossCurvesChart.Series.Add(KN80Series)
        CrossCurvesChart.Series.Add(KN90Series)

    End Sub

    Public Sub IllustrateSetup()
        DrawBarge(BargeLength.Text, BargeDepth.Text, BargeDepth.Text)
        DrawCrane(BargeLength.Text, DistanceFromEndTextbox.Text)
        DrawPositions(BargeLength.Text, DataSetPosition, DataSetRotation, DistanceFromEndTextbox.Text)
    End Sub

    Public Sub LoadConfig(ConfigComboBoxSelectedIndex)

        ConfigIndex = ConfigComboBoxSelectedIndex

        'If ConfigIndex = ConfigComboBox.Items.Count - 1 Then
        '    CreateNewConfigRoutine()
        'Else

        If ConfigIndex >= 0 Then

            BoomTypeTextbox.Text = xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomType>.Value
            CounterweightTextbox.Text = xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<Counterweight>.Value
            BoomLengthTextbox.Text = xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomLength>.Value
            BoomNameTextbox.Text = xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<BoomName>.Value

            Dim q As Integer

            ListChartComboBox.Items.Clear()
            ListChartComboBox.Text = " "

            For q = 0 To xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>.Count - 1
                ListChartComboBox.Items.Add("Chart" & xCrane...<Crane>(CurrentRowCrane).<Config>(ConfigIndex).<ListChart>(q).<MaxMachineList>.Value)
            Next

            'ListChartComboBox.Items.Add("Create New ...")

        End If
        'End If
    End Sub

    Public Sub LoadCrane(ChooseCraneComboBoxSelectedIndex)
        Dim CentrePoint As Point

        CurrentRowCrane = ChooseCraneComboBoxSelectedIndex

        'If CurrentRowCrane = ChooseCraneComboBox.Items.Count - 1 Then
        '    CreateNewCraneRoutine()
        'Else

        ProposedCraneTextbox.Text = xCrane...<Crane>(CurrentRowCrane).<ProposedCrane>.Value
        CraneSerialNumberTextbox.Text = xCrane...<Crane>(CurrentRowCrane).<CraneSerialNumber>.Value
        BoomHingetoCraneCLTextbox.Text = xCrane...<Crane>(CurrentRowCrane).<BoomHingetoCraneCL>.Value
        CrawlerBasetoBoomHingeTextbox.Text = xCrane...<Crane>(CurrentRowCrane).<CrawlerBasetoBoomHinge>.Value
        PositionofCraneCentretoFrontofTrackTextbox.Text = xCrane...<Crane>(CurrentRowCrane).<PositionofCraneCentretoFrontofTrack>.Value

        CraneNameTextbox.Text = ProposedCraneTextbox.Text
        Label55.Text = CurrentRowCrane + 1 & " / " & ChooseCraneComboBox.Items.Count

        ConfigComboBox.Items.Clear()
        ListChartComboBox.Items.Clear()
        ConfigComboBox.Text = ""
        ListChartComboBox.Text = ""

        For k = 0 To xCrane...<Crane>(CurrentRowCrane).<Config>.Count - 1
            ConfigComboBox.Items.Add(xCrane...<Crane>(CurrentRowCrane).<Config>(k).<BoomName>.Value)
        Next

        'ConfigComboBox.Items.Add("Create New ...")

        CentrePoint = PlaceCranePoint(xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(0).<Name>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(0).<Mass>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(0).<LCG>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(0).<TCG>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(0).<VCG>.Value, xCrane...<Crane>(CurrentRowCrane).<BoomHingetoCraneCL>.Value, xCrane...<Crane>(CurrentRowCrane).<CrawlerBasetoBoomHinge>.Value, xCrane...<Crane>(CurrentRowCrane).<PositionofCraneCentretoFrontofTrack>.Value)
        CTWTCentreShape.Location = CentrePoint
        CentrePoint = PlaceCranePoint(xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(1).<Name>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(1).<Mass>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(1).<LCG>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(1).<TCG>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(1).<VCG>.Value, xCrane...<Crane>(CurrentRowCrane).<BoomHingetoCraneCL>.Value, xCrane...<Crane>(CurrentRowCrane).<CrawlerBasetoBoomHinge>.Value, xCrane...<Crane>(CurrentRowCrane).<PositionofCraneCentretoFrontofTrack>.Value)
        UpperworksCentreShape.Location = CentrePoint
        CentrePoint = PlaceCranePoint(xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(2).<Name>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(2).<Mass>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(2).<LCG>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(2).<TCG>.Value, xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(2).<VCG>.Value, xCrane...<Crane>(CurrentRowCrane).<BoomHingetoCraneCL>.Value, xCrane...<Crane>(CurrentRowCrane).<CrawlerBasetoBoomHinge>.Value, xCrane...<Crane>(CurrentRowCrane).<PositionofCraneCentretoFrontofTrack>.Value)
        CrawlersCentreShape.Location = CentrePoint

        BoomTypeTextbox.Text = ""
        CounterweightTextbox.Text = ""
        BoomLengthTextbox.Text = ""
        BoomNameTextbox.Text = ""

        Chart1.Series.Clear()

        MachineListTextbox.Text = ""


        'End If
    End Sub

    Private Sub SetupWeight(CraneLoad As Double, Rotation1 As Integer)

        Dim LCG As Double
        Dim TCG As Double
        Dim VCG As Double

        Dim LCGMoment As Double
        Dim TCGMoment As Double
        Dim VCGMoment As Double

        Dim TotalLCGMoment As Double
        Dim TotalTCGMoment As Double
        Dim TotalVCGMoment As Double

        For m = 0 To 2
            'For j = 0 To 3
            'CraneWAC(m, j) = DataGridViewWAC.Item(j + 1, m).Value
            CraneWAC(m, 0) = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(m).<Mass>.Value
            CraneWAC(m, 1) = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(m).<LCG>.Value
            CraneWAC(m, 2) = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(m).<TCG>.Value
            CraneWAC(m, 3) = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(m).<VCG>.Value
            'Next
        Next

        ww = 0

        'Do While DataGridViewBoomLengthWeight.Item(0, ww).Value <= EffectiveBoomLength
        Do While xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2>(ww).<BoomLength>.Value <= EffectiveBoomLength
            ww = ww + 1
        Loop

        CraneWAC(3, 0) = Forecast(EffectiveBoomLength, {xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2>(ww - 1).<BoomWeight>.Value, xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2>(ww).<BoomWeight>.Value}, {xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2>(ww - 1).<BoomLength>.Value, xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2>(ww).<BoomLength>.Value})
        CraneWAC(3, 1) = BoomExtension / 2 + BoomHingetoCraneCL
        CraneWAC(3, 2) = 0
        CraneWAC(3, 3) = CrawlerBasetoBoomHinge + BoomHeight / 2

        CraneWAC(4, 0) = CraneLoad
        CraneWAC(4, 1) = BoomExtension + BoomHingetoCraneCL
        CraneWAC(4, 2) = 0
        CraneWAC(4, 3) = CrawlerBasetoBoomHinge + BoomHeight

        For m = 0 To 4

            LCG = CraneWAC(m, 1)
            TCG = CraneWAC(m, 2)
            CraneWAC(m, 1) = Math.Cos(Rotation1 * 3.1415 / 180) * LCG + Math.Sin(Rotation1 * 3.1415 / 180) * TCG
            'CraneWAC(m, 2) = Math.Cos(Rotation1 * 3.1415 / 180) * TCG + Math.Sin(Rotation1 * 3.1415 / 180) * LCG  ERROR
            CraneWAC(m, 2) = Math.Cos(Rotation1 * 3.1415 / 180) * TCG - Math.Sin(Rotation1 * 3.1415 / 180) * LCG

        Next

        Mass = 0
        LCGMoment = 0
        TCGMoment = 0
        VCGMoment = 0

        For m = 0 To 4

            Mass = Mass + CraneWAC(m, 0)
            LCGMoment = LCGMoment + CraneWAC(m, 0) * CraneWAC(m, 1)
            TCGMoment = TCGMoment + CraneWAC(m, 0) * CraneWAC(m, 2)
            VCGMoment = VCGMoment + CraneWAC(m, 0) * CraneWAC(m, 3)

        Next

        CraneLCG = LCGMoment / Mass
        CraneTCG = TCGMoment / Mass
        CraneVCG = VCGMoment / Mass

        TotalWAC(0, 0) = LightshipWAC(0, 0)
        TotalWAC(0, 1) = LightshipWAC(0, 1)
        TotalWAC(0, 2) = LightshipWAC(0, 2)
        TotalWAC(0, 3) = LightshipWAC(0, 3)

        TotalWAC(1, 0) = BallastWAC(0, 0)
        TotalWAC(1, 1) = BallastWAC(0, 1)
        TotalWAC(1, 2) = BallastWAC(0, 2)
        TotalWAC(1, 3) = BallastWAC(0, 3)

        TotalWAC(2, 0) = Mass
        TotalWAC(2, 1) = CraneLCG + PositionofCraneCentre
        TotalWAC(2, 2) = CraneTCG + PositionofCraneCentreOffCentreline
        TotalWAC(2, 3) = CraneVCG + BargeDepthVar + DeckTimberHeight

        TotalMass = 0
        TotalLCGMoment = 0
        TotalTCGMoment = 0
        TotalVCGMoment = 0

        For m = 0 To 2

            TotalMass = TotalMass + TotalWAC(m, 0)
            TotalLCGMoment = TotalLCGMoment + TotalWAC(m, 0) * TotalWAC(m, 1)
            TotalTCGMoment = TotalTCGMoment + TotalWAC(m, 0) * TotalWAC(m, 2)
            TotalVCGMoment = TotalVCGMoment + TotalWAC(m, 0) * TotalWAC(m, 3)

        Next

        TotalLCG = TotalLCGMoment / TotalMass
        TotalTCG = TotalTCGMoment / TotalMass
        TotalVCG = TotalVCGMoment / TotalMass

    End Sub

    Public Sub EquilibriumHydrostatics()

        aa = 0

        Do While DisplacementArray(aa) <= TotalMass
            aa = aa + 1
        Loop

        If aa > HydrostaticsDataSize Then
            MsgBox("Routine Fails.  The test load exceeds the max displacement of the displacement array.")
            'Unload frm
            'Worksheets(SheetName).Shapes(1).Delete()
            'Cells.Select()
            'With Selection.Font
            '    .Color = -16776961
            '    .TintAndShade = 0
            'End With
            'Set a = Worksheets(SheetName).Shapes.AddTextEffect(PresetTextEffect:=msoTextEffect2, Text:="FAIL", FontName:="Arial Black", FontSize:=100, FontBold:=msoFalse, FontItalic:=msoFalse, Left:=50, Top:=50)
        End If

        cc = {LCFDraftArray(aa - 1), LCFDraftArray(aa)}
        dd = {DisplacementArray(aa - 1), DisplacementArray(aa)}
        LCFDraft = Forecast(TotalMass, cc, dd)

        cc = {LCFArray(aa - 1), LCFArray(aa)}
        LCF = Forecast(TotalMass, cc, dd)

        cc = {MTCArray(aa - 1), MTCArray(aa)}
        MTC = Forecast(TotalMass, cc, dd)

        cc = {LCBArray(aa - 1), LCBArray(aa)}
        LCB = Forecast(TotalMass, cc, dd)

        cc = {FPDraftArray(aa - 1), FPDraftArray(aa)}

        FPDraft = Forecast(TotalMass, cc, dd)

        cc = {APDraftArray(aa - 1), APDraftArray(aa)}
        APDraft = Forecast(TotalMass, cc, dd)

        cc = {RMArray(aa - 1), RMArray(aa)}
        RM = Forecast(TotalMass, cc, dd)

        TrimInCm = (LCB - TotalLCG) * TotalMass / (Density * 1000) / MTC

        TrimInM = TrimInCm / 100

        ProportionAft = LCF / BargeLengthVar

        ProportionFwd = 1 - ProportionAft

        FPDraft = FPDraft - ProportionFwd * TrimInM
        APDraft = APDraft + ProportionAft * TrimInM

        cc = {GMTArray(aa - 1), GMTArray(aa)}
        GMt = Forecast(TotalMass, cc, dd) - TotalVCG

        RM2 = TotalMass * GMt * Math.Sin((1) * 3.1415 / 180)

        cc = {GMLArray(aa - 1), GMLArray(aa)}
        GMl = Forecast(TotalMass, cc, dd) - TotalVCG

        TrimAngle = 180 / 3.1415 * Math.Asin((APDraft - FPDraft) / BargeLengthVar)
    End Sub

    Private Sub LargeAngleStability()


        GZMax = 0
        GZArea = 0
        IndexAtMaxGZ = 0
        IndexAtVanishingGZ = 0

        For q = 0 To 900

            gg = 0

            Do While DisplacementCCArray(gg) <= TotalMass
                gg = gg + 1
            Loop

            ee = {CCArray(gg - 1, q), CCArray(gg, q)}
            ff = {DisplacementCCArray(gg - 1), DisplacementCCArray(gg)}
            KNArray(q) = Forecast(TotalMass, ee, ff)
            GZArray(q) = KNArray(q) - TotalVCG * Math.Sin(HeelArray(q) * 3.1415 / 180)

            Residual(q) = Math.Abs(GZArray(q) - Math.Abs(TotalTCG * Math.Cos(HeelArray(q) * 3.1415 / 180)))
            'Residual(q) = GZArray(q) - TotalTCG * Cos(HeelArray(q) * 3.1415 / 180)
            'Residual(q) = Abs(GZArray(q) - TotalTCG * Cos(HeelArray(q) * 3.1415 / 180))

            If GZArray(q) >= GZMax Then
                GZMax = GZArray(q)
                IndexAtMaxGZ = IndexAtMaxGZ + 1
                GZArea = GZArea + SM(q) * GZArray(q)
            Else
            End If
            If GZArray(q) > 0 Then
                IndexAtVanishingGZ = IndexAtVanishingGZ + 1
            Else
            End If
        Next

        MinimumResidual = 999999999
        IndexAtMinimumResidual = 0



        For r = 0 To IndexAtMaxGZ
            If Residual(r) < MinimumResidual Then
                MinimumResidual = Residual(r)
                IndexAtMinimumResidual = IndexAtMinimumResidual + 1
            Else
            End If
        Next

        MinimumResidual = Residual(IndexAtMinimumResidual - 1)
        If HeelArray(IndexAtMinimumResidual - 1) = 0 Then
            HeelAtMinimumResidual = HeelArray(IndexAtMinimumResidual - 1)
        Else
            HeelAtMinimumResidual = HeelArray(IndexAtMinimumResidual - 1) * TotalTCG / Math.Abs(TotalTCG)
        End If


        GZMax = GZArray(IndexAtMaxGZ)
        GZMax = GZMax - Math.Abs(TotalTCG)
        HeelAtMaxGZ = HeelArray(IndexAtMaxGZ)
        HeelAtVanishingGZ = HeelArray(IndexAtVanishingGZ)
        HeelAtVanishingGZFromEQ = HeelAtVanishingGZ - HeelAtMinimumResidual

        'GZArea = GZArea / 3 * 3.1415 / 180 * 0.1 - TotalTCG * 3.1415 / 180 * HeelAtMaxGZ
        GZArea = GZArea / 3 * 3.1415 / 180 * 0.1 - Math.Abs(TotalTCG * 3.1415 / 180 * HeelAtMaxGZ) + Math.Abs(TotalTCG * 3.1415 / 180 * HeelAtMinimumResidual / 2)
        LeastFreeboard = BargeDepthVar - Math.Max(APDraft, FPDraft) + BargeBreadthVar * Math.Tan(Math.Abs(HeelAtMinimumResidual) * -3.1415 / 180) / 2 - 0.075

        'CraneList = Sin((Rotation1) * 3.1415 / 180) * (TrimAngle) + Cos((Rotation1) * 3.1415 / 180) * (HeelAtMinimumResidual)

    End Sub

    Private Sub PrintDebugResults(p1 As Integer, Rotation1 As Double, OperatingRadius As Double, BoomAngle As Double, CraneList As Double)
        DebuggingWorksheet.Activate()

        DebuggingWorksheet.Range("a2").Offset(DebugIndex, p1).Value = Rotation1
        DebuggingWorksheet.Range("b2").Offset(DebugIndex, p1).Value = OperatingRadius
        DebuggingWorksheet.Range("c2").Offset(DebugIndex, p1).Value = TotalMass
        DebuggingWorksheet.Range("d2").Offset(DebugIndex, p1).Value = TotalLCG
        DebuggingWorksheet.Range("e2").Offset(DebugIndex, p1).Value = TotalTCG
        DebuggingWorksheet.Range("f2").Offset(DebugIndex, p1).Value = TotalVCG

        ''Hydrostatic Results

        DebuggingWorksheet.Range("g2").Offset(DebugIndex, p1).Value = LCFDraft
        DebuggingWorksheet.Range("h2").Offset(DebugIndex, p1).Value = FPDraft
        DebuggingWorksheet.Range("i2").Offset(DebugIndex, p1).Value = APDraft
        DebuggingWorksheet.Range("j2").Offset(DebugIndex, p1).Value = HeelAtMinimumResidual
        DebuggingWorksheet.Range("k2").Offset(DebugIndex, p1).Value = GMt
        DebuggingWorksheet.Range("l2").Offset(DebugIndex, p1).Value = GMl
        DebuggingWorksheet.Range("m2").Offset(DebugIndex, p1).Value = TrimAngle
        DebuggingWorksheet.Range("n2").Offset(DebugIndex, p1).Value = LeastFreeboard

        ''Large Angle Stability Results

        DebuggingWorksheet.Range("o2").Offset(DebugIndex, p1).Value = HeelAtMaxGZ
        DebuggingWorksheet.Range("p2").Offset(DebugIndex, p1).Value = GZMax
        DebuggingWorksheet.Range("q2").Offset(DebugIndex, p1).Value = GZArea
        DebuggingWorksheet.Range("r2").Offset(DebugIndex, p1).Value = HeelAtVanishingGZ
        DebuggingWorksheet.Range("s2").Offset(DebugIndex, p1).Value = CraneList

        DebuggingWorksheet.Range("x2").Offset(DebugIndex, p1).Value = MaxLoadAllowedbyBarge
        DebuggingWorksheet.Range("y2").Offset(DebugIndex, p1).Value = BoomAngle + CraneTrim

        DebuggingWorksheet.Range("z2").Offset(DebugIndex, p1).Value = Mass
        DebuggingWorksheet.Range("aa2").Offset(DebugIndex, p1).Value = CraneLCG
        DebuggingWorksheet.Range("ab2").Offset(DebugIndex, p1).Value = CraneTCG
        DebuggingWorksheet.Range("ac2").Offset(DebugIndex, p1).Value = CraneVCG

        DebuggingWorksheet.Range("ad2").Offset(DebugIndex, p1).Value = TotalWAC(0, 0)
        DebuggingWorksheet.Range("ae2").Offset(DebugIndex, p1).Value = TotalWAC(0, 1)
        DebuggingWorksheet.Range("af2").Offset(DebugIndex, p1).Value = TotalWAC(0, 2)
        DebuggingWorksheet.Range("ag2").Offset(DebugIndex, p1).Value = TotalWAC(0, 3)

        DebuggingWorksheet.Range("ah2").Offset(DebugIndex, p1).Value = TotalWAC(1, 0)
        DebuggingWorksheet.Range("ai2").Offset(DebugIndex, p1).Value = TotalWAC(1, 1)
        DebuggingWorksheet.Range("aj2").Offset(DebugIndex, p1).Value = TotalWAC(1, 2)
        DebuggingWorksheet.Range("ak2").Offset(DebugIndex, p1).Value = TotalWAC(1, 3)

        DebuggingWorksheet.Range("al2").Offset(DebugIndex, p1).Value = TotalWAC(2, 0)
        DebuggingWorksheet.Range("am2").Offset(DebugIndex, p1).Value = TotalWAC(2, 1)
        DebuggingWorksheet.Range("an2").Offset(DebugIndex, p1).Value = TotalWAC(2, 2)
        DebuggingWorksheet.Range("ao2").Offset(DebugIndex, p1).Value = TotalWAC(2, 3)

        DebuggingWorksheet.Range("ap2").Offset(DebugIndex, p1).Value = CraneLoad

        DebuggingWorksheet.Range("aq2").Offset(DebugIndex, p1).Value = CraneTrim
        'DebuggingWorksheet.Range("ar2").Offset(DebugIndex, p1).Value = BoomAngle + CraneTrim2
        DebuggingWorksheet.Range("as2").Offset(DebugIndex, p1).Value = EffectiveBoomLength * (Math.Cos(3.1415 / 180 * (BoomAngle + CraneTrim))) + BoomHingetoCraneCL


        DebuggingWorksheet.Range("az2").Offset(DebugIndex, 0).Value = HalfMaxLoadAllowedByBargeVizTipping
        DebuggingWorksheet.Range("ba2").Offset(DebugIndex, 0).Value = FullMaxLoadAllowedByBargeVizTipping



        DeckInclinationAngle = 180 / 3.1415 * Math.Atan(TrimAngle / HeelAtMinimumResidual)

        DebuggingWorksheet.Range("bb2").Offset(DebugIndex, 0).Value = DeckInclinationAngle 'angle for max or min deck inclination

        DebuggingWorksheet.Range("bc2").Offset(DebugIndex, 0).Value = Math.Sin(DeckInclinationAngle * 3.1415 / 180) * HeelAtMinimumResidual + Math.Cos(DeckInclinationAngle * 3.1415 / 180) * TrimAngle

        DebuggingWorksheet.Range("bd2").Offset(DebugIndex, 0).Value = Math.Sin((DeckInclinationAngle + 180) * 3.1415 / 180) * HeelAtMinimumResidual + Math.Cos((DeckInclinationAngle + 180) * 3.1415 / 180) * TrimAngle


        DebuggingWorksheet.Range("ax2").Offset(DebugIndex, 0).Value = ModifiedOperatingRadius


        DebuggingWorksheet.Range("aw2").Offset(DebugIndex, 0).Value = Math.Cos(Math.Acos(((OperatingRadius / 0.3048) / (EffectiveBoomLength / 0.3048))) + CraneTrim * 3.1415 / 180) * EffectiveBoomLength



        ''Hydrostatic Results

        If LCFDraft >= Test1 Then DebuggingWorksheet.Range("g2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If FPDraft >= Test2 Then DebuggingWorksheet.Range("h2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If APDraft >= Test3 Then DebuggingWorksheet.Range("i2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If Math.Abs(HeelAtMinimumResidual) <= Test4 Then DebuggingWorksheet.Range("j2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If GMt >= BargeBreadthVar * Test5 Then DebuggingWorksheet.Range("k2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If GMl >= GMLLimit Then DebuggingWorksheet.Range("l2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If Math.Abs(TrimAngle) <= Test7 Then DebuggingWorksheet.Range("m2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If LeastFreeboard >= Test8 Then DebuggingWorksheet.Range("n2").Offset(DebugIndex, p1).Font.ColorIndex = 4

        ''Large Angle Stability Results

        If HeelAtMaxGZ > Test9 Then DebuggingWorksheet.Range("o2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If GZMax >= Test10 Then DebuggingWorksheet.Range("p2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If GZArea >= Test11 Then DebuggingWorksheet.Range("q2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If HeelAtVanishingGZ >= Test12 Then DebuggingWorksheet.Range("r2").Offset(DebugIndex, p1).Font.ColorIndex = 4
        If Math.Abs(CraneList) <= ListLimit + Test13 Then DebuggingWorksheet.Range("s2").Offset(DebugIndex, p1).Font.ColorIndex = 4
    End Sub

    Private Sub ReadCriteria()
        Test1 = DataSetCriteria.Tables(0).Rows(0).Item(1).ToString
        Test2 = DataSetCriteria.Tables(0).Rows(1).Item(1).ToString
        Test3 = DataSetCriteria.Tables(0).Rows(2).Item(1).ToString
        Test4 = DataSetCriteria.Tables(0).Rows(3).Item(1).ToString
        Test5 = DataSetCriteria.Tables(0).Rows(4).Item(1).ToString
        'Test6 = Worksheets(SheetName).Range("p67").Value
        Test7 = DataSetCriteria.Tables(0).Rows(6).Item(1).ToString
        Test8 = DataSetCriteria.Tables(0).Rows(7).Item(1).ToString
        Test9 = DataSetCriteria.Tables(0).Rows(8).Item(1).ToString
        Test10 = DataSetCriteria.Tables(0).Rows(9).Item(1).ToString
        Test11 = DataSetCriteria.Tables(0).Rows(10).Item(1).ToString
        Test12 = DataSetCriteria.Tables(0).Rows(11).Item(1).ToString
        Test13 = DataSetCriteria.Tables(0).Rows(12).Item(1).ToString
    End Sub

    Private Sub StartProgressBar()
        'frm.ProgressBar1.Max = Worksheets(SheetName).Range("p20") + 2 - 1
        'frm.ProgressBar1.Value = 1
        'frm.ProgressBar1.Refresh
    End Sub

    Private Sub RenameSheetIfDuplicate()
        'Dim Sht As Excel.Worksheet
        'WorksheetExists = 0
        'For Each Sht In oSheet2.Worksheets
        '    If Microsoft.VisualBasic.Left(Sht.Name, Len(SheetName)) = SheetName Then WorksheetExists = WorksheetExists + 1
        'Next Sht
        'If WorksheetExists > 0 Then
        '    SheetName = SheetName & " (" & WorksheetExists + 1 & ")"
        'End If
    End Sub

    Private Sub WriteInputValues()
        ResultsWorksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible
        ResultsWorksheet.Activate()

        ResultsWorksheet.Range("p2").Value = BargeNameVar.ToString
        ResultsWorksheet.Range("p3").Value = BargeLengthVar.ToString
        ResultsWorksheet.Range("p4").Value = BargeBreadthVar.ToString
        ResultsWorksheet.Range("p5").Value = BargeDepthVar.ToString
        ResultsWorksheet.Range("p6").Value = HydrostaticDataTab.ToString
        ResultsWorksheet.Range("p7").Value = CrossCurveDataTab.ToString

        ResultsWorksheet.Range("p10").Value = BoomHingetoCraneCL.ToString()
        ResultsWorksheet.Range("p11").Value = CrawlerBasetoBoomHinge.ToString()
        ResultsWorksheet.Range("p12").Value = PositionofCraneCentretoFrontofTrack.ToString()
        ResultsWorksheet.Range("p13").Value = CraneSerialNumber.ToString()
        ResultsWorksheet.Range("p14").Value = Neat(EffectiveBoomLength.ToString / 0.3048)
        ResultsWorksheet.Range("p15").Value = BoomNameTextbox.Text
        ResultsWorksheet.Range("p16").Value = CounterweightTextbox.Text
        ResultsWorksheet.Range("p17").Value = ProposedCraneVar.ToString
        ResultsWorksheet.Range("p18").Value = MachineList.ToString()
        ResultsWorksheet.Range("p19").Value = MachineList.ToString()
        ResultsWorksheet.Range("p24").Value = Neat(DistanceFromEndTextbox.Text / 0.3048)
        ResultsWorksheet.Range("p25").Value = PositionsComboBox.SelectedItem.ToString
        ResultsWorksheet.Range("p26").Value = DataSetRotation.Tables(0).Rows(0).Item(0)
        ResultsWorksheet.Range("p27").Value = DataSetRotation.Tables(0).Rows(DataSetRotation.Tables(0).Rows.Count - 1).Item(0)
        ResultsWorksheet.Range("p28").Value = "OVER BOW"
        If AllAroundRadio.Checked Then ResultsWorksheet.Range("p28").Value = "ALL AROUND"
        ResultsWorksheet.Range("p29").Value = ModelAbbreviation.ToString
        ResultsWorksheet.Range("p30").Value = SheetName.ToString
        ResultsWorksheet.Range("p32").Value = PositionofCraneCentreOffCentreline.ToString
        If NotesRichTextBox.Text <> "" Then
            ResultsWorksheet.Range("b49").Value = "e) "
            ResultsWorksheet.Range("c49").Value = "Run-time notes:" & NotesRichTextBox.Text
        End If

        ResultsWorksheet.Range("p33").Value = Calibration.Text
        ResultsWorksheet.Range("p34").Value = BallastDescription.Text
    End Sub


    Private Sub PrintDebugResultsInput(p1 As Integer, Rotation1 As Double, OperatingRadius As Double, BoomAngle As Double, CraneList As Double)

        DebuggingWorksheet.Activate()

        DebuggingWorksheet.Range("a2").Offset(DebugIndex, p1).Value = Rotation1
        DebuggingWorksheet.Range("b2").Offset(DebugIndex, p1).Value = OperatingRadius
        DebuggingWorksheet.Range("c2").Offset(DebugIndex, p1).Value = BoomAngle
        DebuggingWorksheet.Range("d2").Offset(DebugIndex, p1).Value = InitialCraneLoad
        DebuggingWorksheet.Range("e2").Offset(DebugIndex, p1).Value = HalfMaxMomentAllowed
        DebuggingWorksheet.Range("f2").Offset(DebugIndex, p1).Value = HalfMaxLoadAllowedByBargeVizTipping
        DebuggingWorksheet.Range("g2").Offset(DebugIndex, p1).Value = FullMaxMomentAllowed
        DebuggingWorksheet.Range("h2").Offset(DebugIndex, p1).Value = FullMaxLoadAllowedByBargeVizTipping
        DebuggingWorksheet.Range("i2").Offset(DebugIndex, p1).Value = MaxLoadAllowedbyBarge



    End Sub

    Private Sub PrintDebugResultsAlt1(p1 As Integer, Rotation1 As Double, OperatingRadius As Double, BoomAngle As Double, CraneList As Double)

        'Load Case stats

        DebuggingWorksheet.Range("h2").Offset(DebugIndex, p1).Value = TotalMass
        DebuggingWorksheet.Range("i2").Offset(DebugIndex, p1).Value = TotalLCG
        DebuggingWorksheet.Range("j2").Offset(DebugIndex, p1).Value = TotalTCG
        DebuggingWorksheet.Range("k2").Offset(DebugIndex, p1).Value = TotalVCG

        'Hydrostatic Results

        DebuggingWorksheet.Range("l2").Offset(DebugIndex, p1).Value = LCFDraft
        DebuggingWorksheet.Range("m2").Offset(DebugIndex, p1).Value = FPDraft
        DebuggingWorksheet.Range("n2").Offset(DebugIndex, p1).Value = APDraft
        DebuggingWorksheet.Range("o2").Offset(DebugIndex, p1).Value = HeelAtMinimumResidual
        DebuggingWorksheet.Range("p2").Offset(DebugIndex, p1).Value = GMt
        DebuggingWorksheet.Range("q2").Offset(DebugIndex, p1).Value = GMl
        DebuggingWorksheet.Range("r2").Offset(DebugIndex, p1).Value = TrimAngle
        DebuggingWorksheet.Range("s2").Offset(DebugIndex, p1).Value = LeastFreeboard

        'Large Angle Stability Results

        DebuggingWorksheet.Range("t2").Offset(DebugIndex, p1).Value = HeelAtMaxGZ
        DebuggingWorksheet.Range("u2").Offset(DebugIndex, p1).Value = GZMax
        DebuggingWorksheet.Range("v2").Offset(DebugIndex, p1).Value = GZArea
        DebuggingWorksheet.Range("w2").Offset(DebugIndex, p1).Value = HeelAtVanishingGZ

        'Machine Stats

        DebuggingWorksheet.Range("x2").Offset(DebugIndex, p1).Value = CraneTrim
        DebuggingWorksheet.Range("y2").Offset(DebugIndex, p1).Value = BoomAngle + CraneTrim
        DebuggingWorksheet.Range("z2").Offset(DebugIndex, p1).Value = CraneList

        DebuggingWorksheet.Range("aa2").Offset(DebugIndex, p1).Value = ModifiedOperatingRadius

    End Sub

    Private Sub CalculateMaxMoment()
        CounterweightMass = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(0).<Mass>.Value
        CounterweightLCG = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(0).<LCG>.Value
        CounterweightVCG = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(0).<VCG>.Value
        UpperworksMass = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(1).<Mass>.Value
        UpperworksLCG = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(1).<LCG>.Value
        UpperworksVCG = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(1).<VCG>.Value
        CrawlersMass = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(2).<Mass>.Value
        CrawlersLCG = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(2).<LCG>.Value
        CrawlersVCG = xCrane...<Crane>(CurrentRowCrane).<WAC>.<Row3>(2).<VCG>.Value
        BoomMass = Forecast(EffectiveBoomLength, {xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2>(ww - 1).<BoomWeight>.Value, xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2>(ww).<BoomWeight>.Value}, {xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2>(ww - 1).<BoomLength>.Value, xCrane...<Crane>(CurrentRowCrane).<BoomLengthWeight>.<Row2>(ww).<BoomLength>.Value})
        BoomLCG = (BoomExtension / 2 + BoomHingetoCraneCL)
        BoomVCG = CrawlerBasetoBoomHinge + BoomHeight / 2
        HalfMaxMomentAllowed = OperatingRadius * InitialCraneLoad
        FullMaxMomentAllowed = OperatingRadius * InitialCraneLoad + CounterweightMass * CounterweightLCG + UpperworksMass * UpperworksLCG + CrawlersMass * CrawlersLCG + BoomMass * BoomLCG


        MaxLoadAllowedByBargeVizStability = CraneLoad

        ModifiedOperatingRadius = EffectiveBoomLength * Math.Cos((BoomAngle + CraneTrim) * 3.1415 / 180) + BoomHingetoCraneCL

        ModifiedCounterweightRadius = CounterweightLCG * Math.Cos(3.1415 / 180 * (CraneTrim)) - CounterweightVCG * Math.Sin(3.1415 / 180 * (CraneTrim))
        ModifiedUpperworksRadius = UpperworksLCG * (Math.Cos(3.1415 / 180 * (CraneTrim))) - UpperworksVCG * Math.Sin(3.1415 / 180 * (CraneTrim))
        ModifiedCrawlerRadius = CrawlersLCG * (Math.Cos(3.1415 / 180 * (CraneTrim))) - CrawlersVCG * Math.Sin(3.1415 / 180 * (CraneTrim))
        ModifiedBoomRadius = BoomLCG * (Math.Cos(3.1415 / 180 * (CraneTrim))) - BoomVCG * Math.Sin(3.1415 / 180 * (CraneTrim))

        HalfMaxLoadAllowedByBargeVizTipping = HalfMaxMomentAllowed / ModifiedOperatingRadius

        FullMaxLoadAllowedByBargeVizTipping = (FullMaxMomentAllowed - (CounterweightMass * ModifiedCounterweightRadius + UpperworksMass * ModifiedUpperworksRadius + CrawlersMass * ModifiedCrawlerRadius + BoomMass * ModifiedBoomRadius)) / ModifiedOperatingRadius

        MaxLoadAllowedbyBarge = Math.Min(HalfMaxLoadAllowedByBargeVizTipping, MaxLoadAllowedByBargeVizStability)

    End Sub

End Class



