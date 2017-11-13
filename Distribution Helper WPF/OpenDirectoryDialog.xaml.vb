Imports System.Windows.Forms
Imports MahApps.Metro.Controls

Public Class OpenDirectoryDialog
    Inherits MetroWindow

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Left = My.Windows.MainWindow.Left - 200
        Top = My.Windows.MainWindow.Left - 300
    End Sub


    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        My.Windows.MainWindow.TurnOffBrowseBttns()
        DistributionPathACBox.AutoCompleteMode = AutoCompleteMode.Suggest
        DistributionPathACBox.AutoCompleteSource = AutoCompleteSource.AllSystemSources
        MyHost.Child = DistributionPathACBox
    End Sub


    Private Sub DistributionPathACBox_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles DistributionPathACBox.PreviewKeyDown
        If e.KeyValue.Equals(13) Then
            If System.IO.Directory.Exists(Me.DistributionPathACBox.Text) Then
                'Enter and path exists
                Dim path = Me.DistributionPathACBox.Text
                My.Windows.MainWindow.TurnOffBrowseBttns()
                Me.Close()
                My.Windows.MainWindow.FindFilesAndCreateProgramSelectWindow(path)
            Else
                Me.Title = "Invalid path"
            End If
        End If
        Console.Write("pressed " & Int(e.KeyValue) & vbCrLf)
    End Sub
End Class