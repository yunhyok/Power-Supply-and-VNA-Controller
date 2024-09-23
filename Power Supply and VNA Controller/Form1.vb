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
Imports System.Globalization
Imports System.Runtime.CompilerServices
Imports Microsoft.ApplicationInsights.Extensibility



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
End Class
