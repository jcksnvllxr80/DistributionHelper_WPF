Public Class DraggableListViewItem
    Inherits ListViewItem


    Private _deferSelection As Boolean = False

    Protected Overrides Sub OnPreviewMouseLeftButtonDown(e As MouseButtonEventArgs)
        If e.ClickCount = 1 And IsSelected Then
            _deferSelection = True
        Else
            MyBase.OnPreviewMouseLeftButtonDown(e)
        End If
    End Sub


    Protected Overrides Sub OnPreviewMouseLeftButtonUp(e As MouseButtonEventArgs)
        If _deferSelection Then
            _deferSelection = False
        End If
        MyBase.OnMouseLeftButtonUp(e)
    End Sub


    Protected Overrides Sub OnMouseLeave(e As MouseEventArgs)
        MyBase.OnMouseLeave(e)
    End Sub
End Class
