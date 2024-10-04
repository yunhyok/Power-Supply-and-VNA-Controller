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

Public Class Cal

    Private completed As Boolean = False

    Private category As String
    Private OPM As OnePortMeasurement
    Private state As ExecutionState

    Sub New(ByRef myOPM As OnePortMeasurement, ByVal CalCategory As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        category = CalCategory
        OPM = myOPM

        Label_Title.Text = String.Format("{0} Calibration", CalCategory)
        Label_Status.Text = String.Format("Ready" & vbCrLf & "for" & vbCrLf & "{0} Calibration??", CalCategory)

    End Sub


    Public ReadOnly Property ProcessCompleted As Boolean
        Get
            Return completed
        End Get
    End Property

    Private Sub Button_YES_Click(sender As Object, e As EventArgs) Handles Button_YES.Click

        Button_Cancel.Enabled = False
        Button_YES.Enabled = False


        Select Case category
            Case "OPEN"
                state = OPM.Calibration.UserRange.ExecuteOpen()
            Case "SHORT"
                state = OPM.Calibration.UserRange.ExecuteShort()
            Case "LOAD"
                state = OPM.Calibration.UserRange.ExecuteLoad()
        End Select

        Dim dots As Integer = 0
        While OPM.Calibration.UserRange.IsActive
            Label_Status.Text = "Wait "
            For i As Integer = 0 To dots
                Label_Status.Text &= "."
                If dots = 5 Then
                    dots = 0
                End If
            Next
            DoEvents()
            Sleep(500)
        End While

        completed = True
        Label_Status.Text = "Calibration Completed!"
        Sleep(1000)
        Me.Close()
    End Sub

    Private Sub Button_Cancel_Click(sender As Object, e As EventArgs) Handles Button_Cancel.Click
        completed = False
        Me.Close()
    End Sub
End Class