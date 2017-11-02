Imports MahApps.Metro.Controls
Imports Xceed.Wpf.Toolkit

Class MainWindow
    Inherits MetroWindow



    Private Sub Distribution_Helper_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded

        Dim DistributionsDataSet As Distribution_Helper_WPF.DistributionsDataSet = CType(Me.FindResource("DistributionsDataSet"), Distribution_Helper_WPF.DistributionsDataSet)
        'Load data into the table Distributions. You can modify this code as needed.
        Dim DistributionsDataSetDistributionsTableAdapter As Distribution_Helper_WPF.DistributionsDataSetTableAdapters.DistributionsTableAdapter = New Distribution_Helper_WPF.DistributionsDataSetTableAdapters.DistributionsTableAdapter()
        DistributionsDataSetDistributionsTableAdapter.Fill(DistributionsDataSet.Distributions)
        Dim DistributionsViewSource As System.Windows.Data.CollectionViewSource = CType(Me.FindResource("DistributionsViewSource"), System.Windows.Data.CollectionViewSource)
        DistributionsViewSource.View.MoveCurrentToFirst()
    End Sub

    Private Sub DistributionsDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles DistributionsDataGrid.SelectionChanged

    End Sub

    Private Sub RefreshDB_Click(sender As Object, e As RoutedEventArgs) Handles RefreshDB.Click

    End Sub
End Class
