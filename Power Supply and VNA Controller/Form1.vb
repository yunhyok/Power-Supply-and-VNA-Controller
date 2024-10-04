Imports VisaComLib

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



Public Class Form1

    Private ADDR_PWRSPL As String = Nothing
    Private ADDR_VNA As String = Nothing

    Private DIMM_PWRSPL As New FormattedIO488
    Private DIMM_VNA As New FormattedIO488

    Private RM_PWRSPL As New ResourceManager
    Private RM_VNA As New ResourceManager

    Private PWRSPLisConnected As Boolean = False
    Private VNAisConnected As Boolean = False

    Private WithEvents Timer1 As New System.Timers.Timer

    Private AutomationInterface As New BodeAutomation

    Private Bode100 As BodeDevice
    Private myOnePortMeasuremnt As OnePortMeasurement

    Private myOPMIsConfigured As Boolean = False

    Private CalFileName As String
    Private LoadCalFileUsage As Boolean = False
    Private CalOPEN As Boolean = False
    Private CalSHORT As Boolean = False
    Private CalLOAD As Boolean = False

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Timer1.Interval = 1000


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '' Get the list of available resources
        'Dim resources As String() = ResourceManager.GetLocalManager.FindResources("?*")


        '' Find the power supply
        'For Each s As String In resources
        '    If s.Contains("USB0::0x1AB1::0x0E11::DP8A") Then
        '        ADDR_PWRSPL = s
        '        Exit For
        '    End If
        'Next

        '' Find the VNA
        'For Each s As String In resources
        '    If s.Contains("USB0::0x2A8D::0x5D01::MY58180001") Then
        '        ADDR_VNA = s
        '        Exit For
        '    End If
        'Next

        'ADDR_PWRSPL = TextBox_ADDR_PWRSPL.Text

        '' If the power supply is found, connect to it
        'If ADDR_PWRSPL IsNot Nothing Then
        '    Try
        '        DIMM_PWRSPL.IO = RM_PWRSPL.Open(ADDR_PWRSPL, AccessMode.NO_LOCK, 2000, "")
        '        PWRSPLisConnected = True
        '        MsgBox("Connecting Succcessful")
        '    Catch ex As Exception
        '        MsgBox("Error connecting to power supply: " & ex.Message)
        '    End Try
        'Else
        '    MsgBox("Power supply not found")
        'End If

        '' If the VNA is found, connect to it
        'If ADDR_VNA IsNot Nothing Then
        '    Try
        '        DIMM_VNA.IO = RM_VNA.Open(ADDR_VNA, AccessMode.NO_LOCK, 2000, "")
        '        VNAisConnected = True
        '    Catch ex As Exception
        '        MsgBox("Error connecting to VNA: " & ex.Message)
        '    End Try
        'Else
        '    MsgBox("VNA not found")
        'End If

    End Sub

    Private Sub Button_CONN_PWRSPL_Click(sender As Object, e As EventArgs) Handles Button_CONN_PWRSPL.Click


        If PWRSPLisConnected Then
            Try
                Timer1.Stop()
                DIMM_PWRSPL.IO.Close()
                PWRSPLisConnected = False

                With Button_CONN_PWRSPL
                    .BackColor = Color.LightGray
                    .Font = New Font(.Font, FontStyle.Regular)
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        Else
            ADDR_PWRSPL = TextBox_ADDR_PWRSPL.Text

            Try
                DIMM_PWRSPL.IO = RM_PWRSPL.Open(ADDR_PWRSPL, AccessMode.NO_LOCK, 2000, "")
                PWRSPLisConnected = True



                With Button_CONN_PWRSPL
                    .BackColor = Color.GreenYellow
                    .Font = New Font(.Font, FontStyle.Bold)
                End With

                DIMM_PWRSPL.WriteString("OUTP 0, (@1)")
                DIMM_PWRSPL.WriteString("OUTP 0, (@2)")
                DIMM_PWRSPL.WriteString("OUTP 0, (@3)")

                'Timer1.Start()
                'GetStatusPWRSPL()
            Catch ex As Exception
                MessageBox.Show(ex.Message)

            End Try
        End If

    End Sub

    Private Sub GetStatusPWRSPL()
        If PWRSPLisConnected Then
            Try
                Dim opc As String = "0"
                While CInt(opc) <> 1
                    Try
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                    Catch ex As Exception

                    End Try
                    Sleep((New Random).Next(100))
                    DoEvents()
                End While


                DIMM_PWRSPL.WriteString("MEAS:VOLT? CH1")
                Dim v1 As String = DIMM_PWRSPL.ReadString
                Dim vv1 As String() = v1.Split("E")
                Dim index1 As Integer = CInt(vv1(1))
                Dim index11 As Integer = Math.Floor(CDec(vv1(1)) / 3) * 3
                Dim base1 As Decimal = CDec(vv1(0)) * 10 ^ (index11 - index1)
                With Label_DISP_VOLT_CH1
                    .Invoke(
                        Sub()
                            .Text = base1.ToString("F3")
                        End Sub)
                End With
                With Label_DISP_UNIT_VOLT_CH1
                    .Invoke(
                        Sub()
                            .Text = GetVoltUnit(index11)
                        End Sub)
                End With

                opc = "0"
                While CInt(opc) <> 1
                    Try
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                    Catch ex As Exception

                    End Try
                    Sleep((New Random).Next(100))
                    DoEvents()
                End While

                DIMM_PWRSPL.WriteString("MEAS:CURR? CH1")
                Dim c1 As String = DIMM_PWRSPL.ReadString
                Dim cc1 As String() = c1.Split("E")
                Dim index2 As Integer = CInt(cc1(1))
                Dim index22 As Integer = Math.Floor(CDec(cc1(1)) / 3) * 3
                Dim base2 As Decimal = CDec(cc1(0)) * 10 ^ (index22 - index2)

                With Label_DISP_CURR_CH1
                    .Invoke(
                        Sub()
                            .Text = base2.ToString("F3")
                        End Sub)
                End With
                With Label_DISP_UNIT_CURR_CH1
                    .Invoke(
                        Sub()
                            .Text = GetCurrUnit(index22)
                        End Sub)
                End With

                opc = "0"
                While CInt(opc) <> 1
                    Try
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                    Catch ex As Exception

                    End Try
                    Sleep((New Random).Next(100))
                    DoEvents()
                End While

                DIMM_PWRSPL.WriteString("MEAS:VOLT? CH2")
                Dim v2 As String = DIMM_PWRSPL.ReadString
                Dim vv2 As String() = v2.Split("E")

                Dim index3 As Integer = CInt(vv2(1))
                Dim index33 As Integer = Math.Floor(CDec(vv2(1)) / 3) * 3
                Dim base3 As Decimal = CDec(vv2(0)) * 10 ^ (index33 - index3)

                With Label_DISP_VOLT_CH2
                    .Invoke(
                        Sub()
                            .Text = base3.ToString("F3")
                        End Sub)
                End With
                With Label_DISP_UNIT_VOLT_CH2
                    .Invoke(
                        Sub()
                            .Text = GetVoltUnit(index33)
                        End Sub)
                End With

                opc = "0"
                While CInt(opc) <> 1
                    Try
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                    Catch ex As Exception

                    End Try
                    Sleep((New Random).Next(100))
                    DoEvents()
                End While

                DIMM_PWRSPL.WriteString("MEAS:CURR? CH2")
                Dim c2 As String = DIMM_PWRSPL.ReadString
                Dim cc2 As String() = c2.Split("E")

                Dim index4 As Integer = CInt(cc2(1))
                Dim index44 As Integer = Math.Floor(CDec(cc2(1)) / 3) * 3
                Dim base4 As Decimal = CDec(cc2(0)) * 10 ^ (index44 - index4)


                With Label_DISP_CURR_CH2
                    .Invoke(
                        Sub()
                            .Text = base4.ToString("F3")
                        End Sub)
                End With
                With Label_DISP_UNIT_CURR_CH2
                    .Invoke(
                        Sub()
                            .Text = GetCurrUnit(index44)
                        End Sub)
                End With


                opc = "0"
                While CInt(opc) <> 1
                    Try
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                    Catch ex As Exception

                    End Try
                    Sleep((New Random).Next(100))
                    DoEvents()
                End While

                DIMM_PWRSPL.WriteString("MEAS:VOLT? CH3")
                Dim v3 As String = DIMM_PWRSPL.ReadString
                Dim vv3 As String() = v3.Split("E")
                Dim index5 As Integer = CInt(vv3(1))
                Dim index55 As Integer = Math.Floor(CDec(vv3(1)) / 3) * 3
                Dim base5 As Decimal = CDec(vv3(0)) * 10 ^ (index55 - index5)

                With Label_DISP_VOLT_CH3
                    .Invoke(
                        Sub()
                            .Text = base5.ToString("F3")
                        End Sub)
                End With
                With Label_DISP_UNIT_VOLT_CH3
                    .Invoke(
                        Sub()
                            .Text = GetVoltUnit(index55)
                        End Sub)
                End With

                opc = "0"
                While CInt(opc) <> 1
                    Try
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                    Catch ex As Exception

                    End Try
                    Sleep((New Random).Next(100))
                    DoEvents()
                End While

                DIMM_PWRSPL.WriteString("MEAS:CURR? CH3")
                Dim c3 As String = DIMM_PWRSPL.ReadString
                Dim cc3 As String() = c3.Split("E")
                Dim index6 As Integer = CInt(cc3(1))
                Dim index66 As Integer = Math.Floor(CDec(cc3(1)) / 3) * 3
                Dim base6 As Decimal = CDec(cc3(0)) * 10 ^ (index66 - index6)

                With Label_DISP_CURR_CH3
                    .Invoke(
                        Sub()
                            .Text = base6.ToString("F3")
                        End Sub)
                End With
                With Label_DISP_UNIT_CURR_CH3
                    .Invoke(
                        Sub()
                            .Text = GetCurrUnit(index66)
                        End Sub)
                End With



            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Elapsed
        GetStatusPWRSPL()
    End Sub

    Private Sub Button_ON_Click(sender As Object, e As EventArgs) Handles Button_OUTP_CH1.Click, Button_OUTP_CH2.Click, Button_OUTP_CH3.Click

        If PWRSPLisConnected Then
            With CType(sender, Button)
                Dim name As String = .Name
                Dim ch As String = name.Substring(name.Length - 1)

                Dim querySentence As String = String.Format("STAT:QUES:INST:ISUM{0}:COND?", ch)
                Try
                    Dim opc As String = "0"
                    While CInt(opc) <> 1
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                        DoEvents()
                    End While

                    DIMM_PWRSPL.WriteString(querySentence)
                    '0 - The output is off or unregulated.
                    '1 - The output is CC(constant current) operating mode.
                    '2 - The output is CV(constant voltage) operating mode.
                    '3 - The output has a hardware failure.
                    Dim chStatus As String = DIMM_PWRSPL.ReadString

                    If CInt(chStatus) = 0 Then
                        Dim commandSentence As String = String.Format("OUTP 1, (@{0})", ch)
                        DIMM_PWRSPL.WriteString(commandSentence)
                        .BackColor = Color.Yellow
                        .Font = New Font(.Font, FontStyle.Bold)
                        .Text = "ON"
                    Else
                        Dim commandSentence As String = String.Format("OUTP 0, (@{0})", ch)
                        DIMM_PWRSPL.WriteString(commandSentence)
                        .BackColor = Color.LightGray
                        .Font = New Font(.Font, FontStyle.Regular)
                        .Text = "OFF"
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try


            End With
        End If

    End Sub

    Private Sub Button_SET_CH1_Click(sender As Object, e As EventArgs) Handles Button_SET_CH1.Click, Button_SET_CH2.Click, Button_SET_CH3.Click
        With CType(sender, Button)
            Dim name As String = .Name
            Dim ch As String = name.Substring(name.Length - 1)


            Try
                Dim opc As String = "0"
                While CInt(opc) <> 1
                    DIMM_PWRSPL.WriteString("*OPC?")
                    opc = DIMM_PWRSPL.ReadString
                    DoEvents()
                End While

                Dim VOLT_BOX As TextBox = CType(Me.Controls.Find(String.Format("Textbox_SET_VOLT_CH{0}", ch), True)(0), TextBox)

                Dim commandSentence As String = String.Format("VOLT {0}, (@{1})", VOLT_BOX.Text, ch)
                DIMM_PWRSPL.WriteString(commandSentence)

                Dim V1 As String = VOLT_BOX.Text
                VOLT_BOX.Text = CDec(V1).ToString("F3", CultureInfo.InvariantCulture)

                opc = "0"
                While CInt(opc) <> 1
                    DIMM_PWRSPL.WriteString("*OPC?")
                    opc = DIMM_PWRSPL.ReadString
                    DoEvents()
                End While

                Dim CURR_BOX As TextBox = CType(Me.Controls.Find(String.Format("Textbox_SET_CURR_CH{0}", ch), True)(0), TextBox)

                commandSentence = String.Format("CURR {0}, (@{1})", CURR_BOX.Text, ch)
                DIMM_PWRSPL.WriteString(commandSentence)

                CURR_BOX.Text = CDec(CURR_BOX.Text).ToString("F3")

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End With

    End Sub


    Private Shared Function GetVoltUnit(ByVal index As Integer) As String
        Select Case index
            Case 0
                Return "V"
            Case -3
                Return "mV"
            Case -6
                Return "uV"
            Case -9
                Return "nV"
            Case Else
                Return "err"

        End Select
    End Function

    Private Shared Function GetCurrUnit(ByVal index As Integer) As String
        Select Case index
            Case 0
                Return "A"
            Case -3
                Return "mA"
            Case -6
                Return "uA"
            Case -9
                Return "nA"
            Case Else
                Return "err"

        End Select
    End Function

    Private Sub Button_PWRSPL_Refresh_Click(sender As Object, e As EventArgs) Handles Button_PWRSPL_Refresh.Click
        GetStatusPWRSPL()
    End Sub

    Private Sub Button_VNA_AutoConnect_Click(sender As Object, e As EventArgs) Handles Button_VNA_AutoConnect.Click
        If VNAisConnected Then

            Try
                Bode100.ShutDown()

                With Button_VNA_AutoConnect
                    .BackColor = Color.LightGray
                    .Font = New Font(.Font, FontStyle.Regular)
                End With
                VNAisConnected = False
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        Else
            Try
                Dim devices As String() = AutomationInterface.ScanForFreeDevices

                If devices.Count > 0 Then
                    Dim autoConnection As BodeAutomationInterface = New BodeAutomation

                    Bode100 = autoConnection.Connect
                    Label_VNA_ID.Text = Bode100.SerialNumber


                    With Button_VNA_AutoConnect
                        .BackColor = Color.GreenYellow
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With

                    VNAisConnected = True
                Else
                    MsgBox("No VNA is found.")
                End If


            Catch ex As Exception
                MsgBox("Error connecting to VNA: " & ex.Message)

            End Try
        End If


    End Sub

    Private Sub Button_CONN_VNA_Click(sender As Object, e As EventArgs) Handles Button_CONN_VNA.Click
        If VNAisConnected Then

            Try
                Bode100.ShutDown()

                With Button_CONN_VNA
                    .BackColor = Color.LightGray
                    .Font = New Font(.Font, FontStyle.Regular)
                End With
                VNAisConnected = False
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        Else
            Try
                Dim devices As String() = AutomationInterface.ScanForFreeDevices

                If devices.Count > 0 Then
                    Dim autoConnection As BodeAutomationInterface = New BodeAutomation

                    Bode100 = autoConnection.Connect
                    Label_VNA_ID.Text = Bode100.SerialNumber


                    With Button_CONN_VNA
                        .BackColor = Color.GreenYellow
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With

                    VNAisConnected = True
                Else
                    MsgBox("No VNA is found.")
                End If


            Catch ex As Exception
                MsgBox("Error connecting to VNA: " & ex.Message)

            End Try
        End If

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim setupFileName As String = "setup.ini"
        If File.Exists(setupFileName) Then

            Using sr As New StreamReader(setupFileName)

                With sr
                    TextBox_ADDR_PWRSPL.Text = sr.ReadLine.Split(vbTab)(1)
                    TextBox_ADDR_VNA.Text = sr.ReadLine.Split(vbTab)(1)
                    ComboBox_PWRCH.SelectedIndex = CInt(sr.ReadLine.Split(vbTab)(1))
                    TextBox_PWRCC.Text = sr.ReadLine.Split(vbTab)(1)
                    TextBox_PWRBiasStart.Text = sr.ReadLine.Split(vbTab)(1)
                    TextBox_PWRBiasStop.Text = sr.ReadLine.Split(vbTab)(1)
                    TextBox_PWRCount.Text = sr.ReadLine.Split(vbTab)(1)
                    TextBox_VNAStartFrequency.Text = sr.ReadLine.Split(vbTab)(1)
                    ComboBox_VNAStartFrequencyUnit.SelectedIndex = CInt(sr.ReadLine.Split(vbTab)(1))
                    TextBox_VNAStopFrequency.Text = sr.ReadLine.Split(vbTab)(1)
                    ComboBox_VNAStopFrequencyUnit.SelectedIndex = CInt(sr.ReadLine.Split(vbTab)(1))
                    RadioButton_VNASweepLinear.Checked = CBool(sr.ReadLine.Split(vbTab)(1))
                    RadioButton_VNASweepLog.Checked = CBool(sr.ReadLine.Split(vbTab)(1))
                    ComboBox_VNANumberOfPoints.SelectedIndex = CInt(sr.ReadLine.Split(vbTab)(1))
                    Label_DefaultDIR.Text = sr.ReadLine.Split(vbTab)(1)

                    GetVStep()

                End With
                sr.Close()
            End Using
        Else
            ComboBox_PWRCH.SelectedIndex = 1
            TextBox_PWRCC.Text = "0.100"
            TextBox_PWRBiasStart.Text = "20"
            TextBox_PWRBiasStop.Text = "-20"
            TextBox_PWRCount.Text = "161"

            Label_PWRStep.Text = GetVStep()

            TextBox_VNAStartFrequency.Text = "10"
            ComboBox_VNAStartFrequencyUnit.SelectedIndex = 0
            TextBox_VNAStopFrequency.Text = "1"
            ComboBox_VNAStopFrequencyUnit.SelectedIndex = 3

            RadioButton_VNASweepLinear.Checked = True
            RadioButton_VNASweepLog.Checked = False

            ComboBox_VNANumberOfPoints.SelectedIndex = 5

            Label_DefaultDIR.Text = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), "MeasureData")

            WriteSetupFile()

        End If
    End Sub


    Private Function GetVStep() As String
        Try
            Dim startV As Decimal = CDec(TextBox_PWRBiasStart.Text)
            Dim stopV As Decimal = CDec(TextBox_PWRBiasStop.Text)
            Dim count As Integer = CInt(TextBox_PWRCount.Text)

            Return ((stopV - startV) / (count - 1)).ToString("F3")

        Catch ex As Exception
            Return "Error"
        End Try

    End Function

    Private Sub WriteSetupFile()
        Dim setupFileName As String = "setup.ini"
        Using sw As New StreamWriter(setupFileName)
            With sw
                .WriteLine("Power Supply Address" & vbTab & TextBox_ADDR_PWRSPL.Text)
                .WriteLine("VNA Address" & vbTab & TextBox_ADDR_VNA.Text)
                .WriteLine("Power Channel" & vbTab & ComboBox_PWRCH.SelectedIndex)
                .WriteLine("Power CC" & vbTab & TextBox_PWRCC.Text)
                .WriteLine("Bias Start" & vbTab & TextBox_PWRBiasStart.Text)
                .WriteLine("Bias Stop" & vbTab & TextBox_PWRBiasStop.Text)
                .WriteLine("Bias Count" & vbTab & TextBox_PWRCount.Text)
                .WriteLine("VNA Sweep Start Frequency" & vbTab & TextBox_VNAStartFrequency.Text)
                .WriteLine("VNA Sweep Start Unit" & vbTab & ComboBox_VNAStartFrequencyUnit.SelectedIndex)
                .WriteLine("VNA Sweep Stop Frequency" & vbTab & TextBox_VNAStopFrequency.Text)
                .WriteLine("VNA Sweep Stop Unit" & vbTab & ComboBox_VNAStopFrequencyUnit.SelectedIndex)
                .WriteLine("VNA Sweep Linear Scale" & vbTab & RadioButton_VNASweepLinear.Checked.ToString)
                .WriteLine("VNA Sweep Log Scale" & vbTab & RadioButton_VNASweepLog.Checked.ToString)
                .WriteLine("VNA Sweep Points" & vbTab & ComboBox_VNANumberOfPoints.SelectedIndex)
                .WriteLine("Default DIR" & vbTab & Label_DefaultDIR.Text)
            End With
            sw.Close()
        End Using
    End Sub

    Private Sub Button_DefaultDIR_Click(sender As Object, e As EventArgs) Handles Button_DefaultDIR.Click
        Dim fBD As New FolderBrowserDialog
        With fBD
            Dim defaultDIR As String = Label_DefaultDIR.Text
            If Not Path.Exists(defaultDIR) Then
                defaultDIR = GetFolderPath(SpecialFolder.Desktop)
            End If

            .InitialDirectory = defaultDIR

            If .ShowDialog = DialogResult.OK Then
                Label_DefaultDIR.Text = .SelectedPath
            End If

        End With
    End Sub

    Private Sub Button_PWRSetupCheck_Click(sender As Object, e As EventArgs) Handles Button_PWRSetupCheck.Click
        If PWRSPLisConnected Then
            Try
                With DIMM_PWRSPL

                    .WriteString("*CLS")
                    ToolStripStatusLabel_Status.Text = "INIT"

                    Dim ch As String = ComboBox_PWRCH.SelectedIndex + 1
                    Dim ccERR, startERR, stopERR As String


                    Dim opc As String = "0"
                    While CInt(opc) <> 1
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                        Sleep((New Random).Next(100))
                        DoEvents()
                    End While

                    .WriteString(String.Format("CURR {0},(@{1})", TextBox_PWRCC.Text, ch))
                    Sleep(500)
                    .WriteString("*ESR?")
                    ccERR = DIMM_PWRSPL.ReadString


                    opc = "0"
                    While CInt(opc) <> 1
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                        Sleep((New Random).Next(100))
                        DoEvents()
                    End While
                    .WriteString(String.Format("VOLT {0},(@{1})", Math.Abs(CDbl(TextBox_PWRBiasStart.Text)).ToString, ch))
                    .WriteString("*ESR?")
                    startERR = DIMM_PWRSPL.ReadString

                    opc = "0"
                    While CInt(opc) <> 1
                        DIMM_PWRSPL.WriteString("*OPC?")
                        opc = DIMM_PWRSPL.ReadString
                        Sleep((New Random).Next(100))
                        DoEvents()
                    End While
                    .WriteString(String.Format("VOLT {0},(@{1})", Math.Abs(CDbl(TextBox_PWRBiasStop.Text)).ToString, ch))
                    .WriteString("*ESR?")
                    stopERR = DIMM_PWRSPL.ReadString

                    If CInt(ccERR) + CInt(startERR) + CInt(stopERR) > 0 Then
                        Throw New Exception(String.Format("CC ERR : {0}   Start ERR : {1}   Stop ERR : {2}", ccERR, startERR, stopERR))
                    ElseIf CInt(TextBox_PWRCount.text) > 1 Then
                        Throw New Exception("Bias Count should be larger than 1.")
                    Else
                        MessageBox.Show("Power Supply Setup seems to be normal.")
                    End If

                End With
            Catch ex As Exception
                MessageBox.Show("Power Supply Setup Error!!" & vbCrLf & ex.Message)
            End Try
        Else
            MessageBox.Show("Power Supply is not connected.")
        End If
    End Sub

    Private Sub Button_VNASetupCheck_Click(sender As Object, e As EventArgs) Handles Button_VNASetupCheck.Click

        If VNAisConnected Then
            myOPMIsConfigured = False
            Try
                With Bode100
                    myOnePortMeasuremnt = .Impedance.CreateOnePortMeasurement
                    With myOnePortMeasuremnt
                        Dim startFreq As Double = CDbl(TextBox_VNAStartFrequency.Text)
                        startFreq *= 10 ^ (ComboBox_VNAStartFrequencyUnit.SelectedIndex * 3)
                        Dim stopFreq As Double = CDbl(TextBox_VNAStopFrequency.Text)
                        stopFreq *= 10 ^ (ComboBox_VNAStopFrequencyUnit.SelectedIndex * 3)
                        Dim nop As Integer = CInt(ComboBox_VNANumberOfPoints.SelectedItem.ToString)
                        Dim sweepMode As Integer = If(RadioButton_VNASweepLinear.Checked, 0, 1)
                        .ConfigureSweep(startFreq, stopFreq, nop, sweepMode)
                    End With
                    myOPMIsConfigured = True
                    MessageBox.Show("VNA setup seems to be normal.")
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Else
            MessageBox.Show("VNA is not connected.")
        End If
    End Sub

    Private Sub Button_VNACalOPEN_Click(sender As Object, e As EventArgs) Handles Button_VNACalOPEN.Click, Button_VNACalSHORT.Click, Button_VNACalLOAD.Click
        If VNAisConnected Then
            Dim thisButton As Button = CType(sender, Button)
            Dim calCategory As String = thisButton.Text
            Dim calForm As New Cal(myOnePortMeasuremnt, calCategory)
            calForm.ShowDialog()
            If calForm.ProcessCompleted Then

                With thisButton
                    .BackColor = Color.GreenYellow
                    .ForeColor = Color.Black
                    .Font = New Font(.Font, FontStyle.Bold)
                End With

                Select Case calCategory
                    Case "OPEN"
                        CalOPEN = True
                    Case "SHORT"
                        CalSHORT = True
                    Case "LOAD"
                        CalLOAD = True
                End Select

                If CalOPEN AndAlso CalSHORT AndAlso CalLOAD Then
                    Dim myPath As String = Path.Combine(GetFolderPath(Environment.SpecialFolder.Desktop), "myCalibration.mcalx")
                    If myOnePortMeasuremnt.Calibration.SaveCalibration(myPath) Then
                        MessageBox.Show("Calibration is completed and saved at DeskTop.")
                    Else
                        MessageBox.Show("Failed to save the calibration result as a file.")
                    End If

                End If

            End If
        Else
            MessageBox.Show("VNA is not connected.")
        End If



    End Sub

    Private Sub Button_VNALoadCalFile_Click(sender As Object, e As EventArgs) Handles Button_VNALoadCalFile.Click
        If VNAisConnected Then
            Dim ofD As New OpenFileDialog
            With ofD
                .DefaultExt = ".mcalx"
                .InitialDirectory = GetFolderPath(Environment.SpecialFolder.Desktop)
            End With

            If ofD.ShowDialog = DialogResult.OK Then
                CalFileName = ofD.FileName
                LoadCalFileUsage = True
                If myOnePortMeasuremnt.Calibration.UserRange.LoadCalibration(CalFileName) Then
                    CalOPEN = True
                    CalSHORT = True
                    CalLOAD = True

                    With Button_VNACalLOAD
                        .BackColor = Color.GreenYellow
                        .ForeColor = Color.Black
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    With Button_VNACalOPEN
                        .BackColor = Color.GreenYellow
                        .ForeColor = Color.Black
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    With Button_VNACalSHORT
                        .BackColor = Color.GreenYellow
                        .ForeColor = Color.Black
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With

                    MessageBox.Show("Calibration file is loaded.")
                Else
                    MessageBox.Show("Failed to load the calibration file.")
                End If
            End If

        Else
            MessageBox.Show("VNA is not connected.")
        End If
    End Sub
End Class
