Imports OmicronLab.VectorNetworkAnalysis.AutomationInterface
Imports OmicronLab.VectorNetworkAnalysis.AutomationInterface.Interfaces
Imports OmicronLab.VectorNetworkAnalysis.AutomationInterface.Enumerations
Imports OmicronLab.VectorNetworkAnalysis.AutomationInterface.Interfaces.Measurements
Imports OmicronLab.VectorNetworkAnalysis.AutomationInterface.DataTypes

Imports System.IO
Imports System.IO.Ports
Imports System.Threading.Thread
Imports System.Threading
Imports System.Windows.Forms.Application
Imports System.Timers
Imports System.Web
Imports System.Environment
Imports System.Globalization
Imports System.Runtime.CompilerServices
Imports Microsoft.ApplicationInsights.Extensibility
Imports System.ServiceModel
Imports Microsoft.AspNetCore.Mvc.Filters
Imports VisaComLib

Imports System.DateTime
Imports System.Management

Public Class Measure

    Private myOPM As OnePortMeasurement
    Private myPWRSPLY As FormattedIO488
    Private TwoStageIsNeeded As Boolean = False

    Private myCH As Integer
    Private myStartV As Double
    Private myStopV As Double
    Private myCount As Integer
    Private myCC As Double

    Private mySavePath As String

    Private measureContinue As Boolean = False


    Private phase As Integer = 0
    ''' <summary>
    ''' phase 0 - warning
    ''' phase 1 - connection confirm
    ''' phase 2 - status panel for measurement progress
    ''' phase 3 - connection confirm
    ''' phase 4 - status panel for measurement progress
    ''' phase 5 - done
    ''' </summary>

    Private biasStep0 As New List(Of Double)
    Private biasStep1 As New List(Of Double)

    Private biasStep0IsPositive As Boolean = True
    Private biasStep1IsPositive As Boolean = True
    Private biasPolaritySwitching As Boolean = False

    Private WithEvents Blinker As New System.Timers.Timer



    Sub New(ByRef opm As OnePortMeasurement,
            ByRef dimm_pwr As FormattedIO488,
            ByVal biasCH As Integer,
            ByVal biasStartV As Double,
            ByVal biasStopV As Double,
            ByVal biasCount As Integer,
            ByVal biasCC As Double,
            ByVal savePath As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        myOPM = opm
        myPWRSPLY = dimm_pwr

        myCH = biasCH
        myStartV = biasStartV
        myStopV = biasStopV
        myCount = biasCount
        myCC = biasCC

        mySavePath = Path.Combine(savePath, String.Format(
                                  "{0}_{1}To{2}.csv",
                                  Today.ToShortDateString,
                                    If(biasStartV > 0, "P", "N") & biasStartV.ToString("F1"),
                                    If(biasStopV > 0, "P", "N") & biasStopV.ToString("F1")))
        If File.Exists(mySavePath) Then
            mySavePath &= "_" & Now.ToShortTimeString
        End If


        Blinker.Interval = 500

        biasStep0IsPositive = biasStartV > 0
        biasStep1IsPositive = biasStopV > 0
        biasPolaritySwitching = biasStartV * biasStopV > 0

        Dim biasStep As Double = (biasStopV - biasStartV) / biasCount

        For i As Integer = 0 To biasStep - 1
            If biasStartV * (biasStartV + i * biasStep) >= 0 Then
                If i = biasStep - 1 Then
                    biasStep0.Add(biasStopV)
                Else
                    biasStep0.Add(biasStartV + i * biasStep)
                End If

            Else
                If i = biasStep - 1 Then
                    biasStep1.Add(biasStopV)
                Else
                    biasStep1.Add(biasStartV + i * biasStep)
                End If
            End If
        Next


        If biasStartV * biasStopV < 0 Then
            phase = 0
        Else
            phase = 1
        End If

        Label_Progress0.Visible = True
        Label_Progress1.Visible = False
        Label_Progress2.Visible = False
        Label_Progress3.Visible = False
    End Sub

    Private Sub Measure_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        PanelChange(phase)
    End Sub

    Private Sub PanelChange(ByVal thisPhase)
        Select Case phase
            Case 0  'warning
                With TableLayoutPanel_Main
                    With .RowStyles(1)
                        .SizeType = SizeType.Percent
                        .Height = 100
                    End With
                    With .RowStyles(2)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                    With .RowStyles(3)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                End With

                Button_Go.Text = "NEXT"
            Case 1  'connection confirm
                With TableLayoutPanel_Main
                    With .RowStyles(1)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                    With .RowStyles(2)
                        .SizeType = SizeType.Percent
                        .Height = 100
                    End With
                    With .RowStyles(3)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                End With

                Button_Go.Text = "NEXT"

                If biasStep0(0) > 0 Then
                    Label_PWRP.Text = "+"
                    Label_PWRN.Text = "-"
                Else
                    Label_PWRP.Text = "-"
                    Label_PWRN.Text = "+"
                End If
                Label_PWRNotice.Text = String.Format("Start Bias Voltage : {0}V" &
                                                     vbCrLf &
                                                     "Check up power cable connections!",
                                                     biasStep0(0).ToString("F3"))

            Case 2 'Mearure progress 1
                With TableLayoutPanel_Main
                    With .RowStyles(1)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                    With .RowStyles(2)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                    With .RowStyles(3)
                        .SizeType = SizeType.Percent
                        .Height = 100
                    End With
                End With

                Button_Go.Text = "GO"

                Label_MainStatus.Text = "Press Button GO" & vbCrLf & "to" & vbCrLf & "Start Measurement"

            Case 3  'connection confirm
                With TableLayoutPanel_Main
                    With .RowStyles(1)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                    With .RowStyles(2)
                        .SizeType = SizeType.Percent
                        .Height = 100
                    End With
                    With .RowStyles(3)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                End With

                Button_Go.Text = "NEXT"

                If biasStep1(biasStep1.Count - 1) > 0 Then
                    Label_PWRP.Text = "+"
                    Label_PWRN.Text = "-"
                Else
                    Label_PWRP.Text = "-"
                    Label_PWRN.Text = "+"
                End If
                Label_PWRNotice.Text = String.Format("Stop Bias Voltage : {0}V" &
                                                     vbCrLf &
                                                     "Check up power cable connections!",
                                                     biasStep1(biasStep1.Count - 1).ToString("F3"))
            Case 4
                With TableLayoutPanel_Main
                    With .RowStyles(1)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                    With .RowStyles(2)
                        .SizeType = SizeType.Absolute
                        .Height = 0
                    End With
                    With .RowStyles(3)
                        .SizeType = SizeType.Percent
                        .Height = 100
                    End With
                End With

                Button_Go.Text = "GO"

                Label_MainStatus.Text = "Press Button GO" & vbCrLf & "to" & vbCrLf & "Start Measurement"
            Case 5
                Me.Close()
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button_Go.Click
        Button_Go.Enabled = False
        Select Case phase
            Case 0
                phase = 1   'connection confirm
                PanelChange(phase)
            Case 1
                phase = 2   'status panel for measurement progress
                PanelChange(phase)
            Case 2

                Button_Stop.Enabled = True
                phase = DoMeasure()
                Button_Stop.Enabled = False
                PanelChange(phase)
            Case 3
                phase = 4
                PanelChange(phase)
            Case 4
                Button_Stop.Enabled = True
                phase = DoMeasure()
                Button_Stop.Enabled = False
                PanelChange(phase)
            Case 5
                PanelChange(phase)

        End Select
        Button_Go.Enabled = True
    End Sub

    Private Function DoMeasure() As Integer
        myPWRSPLY.WriteString(String.Format("CURR {0}, (@{1})", myCC.ToString, myCH.ToString))
        ProgressBar_Main.Maximum = biasStep0.Count + biasStep1.Count

        If phase = 2 Then

            Using writer As New StreamWriter(mySavePath, False)
                Dim title As String = "Time,Bias Voltage, CC, Current, category,"
                Dim measuredFrequency As Double() = myOPM.MeasurementFrequencies
                For i As Integer = 0 To measuredFrequency.Length - 1
                    title &= String.Format("{0}Hz{1}",
                                           measuredFrequency(i).ToString,
                                           If(i = measuredFrequency.Length - 1, Nothing, ","))
                    DoEvents()
                Next

                writer.WriteLine(title)

                myPWRSPLY.WriteString(String.Format("OUTP 1, (@{0})", myCH.ToString))

                For i As Integer = 0 To biasStep0.Count - 1

                    If measureContinue Then
                        myPWRSPLY.WriteString(
                        String.Format(
                            "VOLT {0},(@{1})", biasStep0(i).ToString, myCH.ToString))

                        Sleep(100)

                        myPWRSPLY.WriteString(
                            String.Format("MEAS:CURR? (@{0})",
                                    myCH.ToString))
                        Dim measuredCurrent As String = Nothing
                        Try
                            measuredCurrent = myPWRSPLY.ReadString
                        Catch ex As Exception

                        End Try

                        With Label_MainStatus
                            .Text = String.Format(
                                "Step : {0}/{1}" &
                                vbCrLf & "Bias Voltage : {2}  Measured Current : {3}" &
                                vbCrLf & "Frequency Sweeping from {4} to {5}",
                                (i + 1).ToString, (biasStep0.Count + biasStep1.Count).ToString,
                                GetVoltage(biasStep0(i)), GetCurrent(measuredCurrent),
                                GetFrequency(measuredFrequency(0)), GetFrequency(measuredFrequency(measuredFrequency.Length - 1)))
                        End With


                        Dim measureTime As DateTime = Now
                        Dim state As ExecutionState = myOPM.ExecuteMeasurement()

                        While Not state = ExecutionState.Ok
                            With Label_MainStatus
                                .Invoke(
                                    Sub()
                                        .BackColor = If(.BackColor = Color.White, Color.LightYellow, Color.White)
                                    End Sub)
                            End With

                            Sleep(250)
                            DoEvents()
                        End While

                        Dim magnitude As Double() = myOPM.MeasurementResults.magnitude
                        Dim real As Double() = myOPM.MeasurementResults.Real
                        Dim imag As Double() = myOPM.MeasurementResults.Imaginary
                        Dim cs As Double() = myOPM.Measurementresults.cs
                        Dim ls As Double() = myOPM.Measurementresults.ls
                        Dim rs As Double() = myOPM.Measurementresults.rs


                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep0(i).ToString,
                                myCC,
                                measuredCurrent,
                                "MAG"))
                        For k As Integer = 0 To magnitude.Length - 1
                            If k = magnitude.Length - 1 Then
                                writer.WriteLine(magnitude(k).ToString)
                            Else
                                writer.WriteLine(magnitude(k).ToString & ",")
                            End If
                            DoEvents()
                        Next


                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep0(i).ToString,
                                myCC,
                                measuredCurrent,
                                "REAL"))
                        For k As Integer = 0 To real.Length - 1
                            If k = real.Length - 1 Then
                                writer.WriteLine(real(k).ToString)
                            Else
                                writer.WriteLine(real(k).ToString & ",")
                            End If
                            DoEvents()
                        Next

                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep0(i).ToString,
                                myCC,
                                measuredCurrent,
                                "IMAG"))
                        For k As Integer = 0 To imag.Length - 1
                            If k = imag.Length - 1 Then
                                writer.WriteLine(imag(k).ToString)
                            Else
                                writer.WriteLine(imag(k).ToString & ",")
                            End If
                            DoEvents()
                        Next

                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep0(i).ToString,
                                myCC,
                                measuredCurrent,
                                "Cs"))
                        For k As Integer = 0 To cs.Length - 1
                            If k = cs.Length - 1 Then
                                writer.WriteLine(cs(k).ToString)
                            Else
                                writer.WriteLine(cs(k).ToString & ",")
                            End If
                            DoEvents()
                        Next

                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep0(i).ToString,
                                myCC,
                                measuredCurrent,
                                "Ls"))
                        For k As Integer = 0 To ls.Length - 1
                            If k = ls.Length - 1 Then
                                writer.WriteLine(ls(k).ToString)
                            Else
                                writer.WriteLine(ls(k).ToString & ",")
                            End If
                            DoEvents()
                        Next

                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep0(i).ToString,
                                myCC,
                                measuredCurrent,
                                "Rs"))
                        For k As Integer = 0 To rs.Length - 1
                            If k = rs.Length - 1 Then
                                writer.WriteLine(rs(k).ToString)
                            Else
                                writer.WriteLine(rs(k).ToString & ",")
                            End If
                            DoEvents()
                        Next


                        Label_MainStatus.BackColor = Color.White
                    Else
                        Exit For
                    End If

                    ProgressBar_Main.Value = i
                    DoEvents()
                Next


                myPWRSPLY.WriteString(String.Format("OUTP 0, (@{0})", myCH.ToString))
                writer.Close()
            End Using

            If biasStep1.Count = 0 Then
                Return 5
            Else
                Return 3
            End If

        ElseIf phase = 4 Then
            Using writer As New StreamWriter(mySavePath, True)
                'Dim title As String = "Time,Bias Voltage, CC, Current, category,"
                Dim measuredFrequency As Double() = myOPM.MeasurementFrequencies
                'For i As Integer = 0 To measuredFrequency.Length - 1
                '    title &= String.Format("{0}Hz{1}",
                '                           measuredFrequency(i).ToString,
                '                           If(i = measuredFrequency.Length - 1, Nothing, ","))
                '    DoEvents()
                'Next

                'writer.WriteLine(title)

                myPWRSPLY.WriteString(String.Format("OUTP 1, (@{0})", myCH.ToString))

                For i As Integer = 0 To biasStep1.Count - 1

                    If measureContinue Then
                        myPWRSPLY.WriteString(
                        String.Format(
                            "VOLT {0},(@{1})", biasStep1(i).ToString, myCH.ToString))

                        Sleep(100)

                        myPWRSPLY.WriteString(
                            String.Format("MEAS:CURR? (@{0})",
                                    myCH.ToString))
                        Dim measuredCurrent As String = Nothing
                        Try
                            measuredCurrent = myPWRSPLY.ReadString
                        Catch ex As Exception

                        End Try

                        With Label_MainStatus
                            .Text = String.Format(
                                "Step : {0}/{1}" &
                                vbCrLf & "Bias Voltage : {2}  Measured Current : {3}" &
                                vbCrLf & "Frequency Sweeping from {4} to {5}",
                                (i + 1).ToString, (biasStep1.Count + biasStep1.Count).ToString,
                                GetVoltage(biasStep1(i)), GetCurrent(measuredCurrent),
                                GetFrequency(measuredFrequency(0)), GetFrequency(measuredFrequency(measuredFrequency.Length - 1)))
                        End With


                        Dim measureTime As DateTime = Now
                        Dim state As ExecutionState = myOPM.ExecuteMeasurement()

                        While Not state = ExecutionState.Ok
                            With Label_MainStatus
                                .Invoke(
                                    Sub()
                                        .BackColor = If(.BackColor = Color.White, Color.LightYellow, Color.White)
                                    End Sub)
                            End With

                            Sleep(250)
                            DoEvents()
                        End While

                        Dim magnitude As Double() = myOPM.MeasurementResults.magnitude
                        Dim real As Double() = myOPM.MeasurementResults.Real
                        Dim imag As Double() = myOPM.MeasurementResults.Imaginary
                        Dim cs As Double() = myOPM.Measurementresults.cs
                        Dim ls As Double() = myOPM.Measurementresults.ls
                        Dim rs As Double() = myOPM.Measurementresults.rs


                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep1(i).ToString,
                                myCC,
                                measuredCurrent,
                                "MAG"))
                        For k As Integer = 0 To magnitude.Length - 1
                            If k = magnitude.Length - 1 Then
                                writer.WriteLine(magnitude(k).ToString)
                            Else
                                writer.WriteLine(magnitude(k).ToString & ",")
                            End If
                            DoEvents()
                        Next


                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep1(i).ToString,
                                myCC,
                                measuredCurrent,
                                "REAL"))
                        For k As Integer = 0 To real.Length - 1
                            If k = real.Length - 1 Then
                                writer.WriteLine(real(k).ToString)
                            Else
                                writer.WriteLine(real(k).ToString & ",")
                            End If
                            DoEvents()
                        Next

                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep1(i).ToString,
                                myCC,
                                measuredCurrent,
                                "IMAG"))
                        For k As Integer = 0 To imag.Length - 1
                            If k = imag.Length - 1 Then
                                writer.WriteLine(imag(k).ToString)
                            Else
                                writer.WriteLine(imag(k).ToString & ",")
                            End If
                            DoEvents()
                        Next

                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep1(i).ToString,
                                myCC,
                                measuredCurrent,
                                "Cs"))
                        For k As Integer = 0 To cs.Length - 1
                            If k = cs.Length - 1 Then
                                writer.WriteLine(cs(k).ToString)
                            Else
                                writer.WriteLine(cs(k).ToString & ",")
                            End If
                            DoEvents()
                        Next

                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep1(i).ToString,
                                myCC,
                                measuredCurrent,
                                "Ls"))
                        For k As Integer = 0 To ls.Length - 1
                            If k = ls.Length - 1 Then
                                writer.WriteLine(ls(k).ToString)
                            Else
                                writer.WriteLine(ls(k).ToString & ",")
                            End If
                            DoEvents()
                        Next

                        writer.Write(
                            String.Format("{0},{1},{2},{3},{5},",
                                measureTime.ToShortTimeString,
                                biasStep1(i).ToString,
                                myCC,
                                measuredCurrent,
                                "Rs"))
                        For k As Integer = 0 To rs.Length - 1
                            If k = rs.Length - 1 Then
                                writer.WriteLine(rs(k).ToString)
                            Else
                                writer.WriteLine(rs(k).ToString & ",")
                            End If
                            DoEvents()
                        Next


                        Label_MainStatus.BackColor = Color.White
                    Else
                        Exit For
                    End If

                    ProgressBar_Main.Value = i
                    DoEvents()
                Next


                myPWRSPLY.WriteString(String.Format("OUTP 0, (@{0})", myCH.ToString))
                writer.Close()
            End Using



            Return 5
        End If


        Return 5
    End Function


    Private Function GetVoltage(ByVal V As Double) As String
        If V = 0 Then
            Return "0 V"
        Else
            Dim index As Integer = Math.Floor(Math.Log10(V))
            Dim unitIndex As Integer = (index \ 3) * 3

            Return (V * 10 ^ (index - unitIndex)).ToString("F3") & " " & GetVoltageUnit(V)
        End If
    End Function

    Private Function GetVoltageUnit(ByVal V As Double) As String
        Dim index As Integer = Math.Floor(Math.Log10(V))
        Dim unitIndex As Integer = (index \ 3) * 3

        Select Case unitIndex
            Case 0
                Return "V"
            Case -3
                Return "mV"
            Case -6
                Return "uV"
            Case -9
                Return "nV"
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function GetVoltage(ByVal V As String) As String
        If V Is Nothing Then
            Return Nothing
        Else
            Return GetVoltage(CDbl(V))
        End If
    End Function

    Private Function GetCurrent(ByVal C As Double) As String
        If C = 0 Then
            Return "0 A"
        Else
            Dim index As Integer = Math.Floor(Math.Log10(C))
            Dim unitIndex As Integer = (index \ 3) * 3

            Return (C * 10 ^ (index - unitIndex)).ToString("F3") & " " & GetCurrentUnit(C)
        End If
    End Function

    Private Function GetCurrent(ByVal C As String) As String
        If C Is Nothing Then
            Return Nothing
        Else
            Return GetCurrent(CDbl(C))
        End If
    End Function

    Private Function GetCurrentUnit(ByVal C As Double) As String
        Dim index As Integer = Math.Floor(Math.Log10(C))
        Dim unitIndex As Integer = (index \ 3) * 3

        Select Case unitIndex
            Case 0
                Return "A"
            Case -3
                Return "mA"
            Case -6
                Return "uA"
            Case -9
                Return "nA"
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function GetFrequency(ByVal F As Double) As String
        If F = 0 Then
            Return "0 Hz"
        Else
            Dim index As Integer = Math.Floor(Math.Log10(F))
            Dim unitIndex As Integer = (index \ 3) * 3

            Return (F * 10 ^ (index - unitIndex)).ToString("F3") & " " & GetFrequencyUnit(F)
        End If
    End Function

    Private Function GetFrequency(ByVal F As String) As String
        If F Is Nothing Then
            Return Nothing
        Else
            Return GetFrequency(CDbl(F))
        End If
    End Function

    Private Function GetFrequencyUnit(ByVal F As Double) As String
        Dim index As Integer = Math.Floor(Math.Log10(F))
        Dim unitIndex As Integer = (index \ 3) * 3

        Select Case unitIndex
            Case 0
                Return "Hz"
            Case 3
                Return "KHz"
            Case 6
                Return "MHz"
            Case 9
                Return "GHz"
            Case Else
                Return Nothing
        End Select
    End Function

    Private Sub Button_Stop_Click(sender As Object, e As EventArgs) Handles Button_Stop.Click
        measureContinue = False
    End Sub
End Class