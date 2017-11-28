Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO

Public Class DocPrintObj
    Private FileStr As String
    Private Printer As String
    Private PrintFont As Font
    Private streamToPrint As StreamReader

    Public Sub New(str As String, prntr As String)
        Me.FileStr = str
        Me.Printer = prntr
    End Sub


    Public Sub SetFont(myFont As String, size As Double)
        Dim fontFamily As New FontFamily(myFont)
        Me.PrintFont = New Font(fontFamily, size)
    End Sub


    Public Function GetFileStr()
        Return Me.FileStr
    End Function


    Public Function GetPrinter()
        Return Me.Printer
    End Function


    Public Sub Print()
        Dim xpsDocument As New Xps.Packaging.XpsDocument(Me.FileStr, FileAccess.ReadWrite)
        Dim fixedDocSeq As FixedDocumentSequence = xpsDocument.GetFixedDocumentSequence()

        Try
            streamToPrint = New StreamReader(Me.FileStr)
            Try
                Dim pd As New PrintDocument()
                AddHandler pd.PrintPage, AddressOf Me.pd_PrintPage
                pd.Print()
            Finally
                streamToPrint.Close()
            End Try
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        fixedDocSeq = Nothing
        xpsDocument.Close()
        xpsDocument = Nothing
        System.IO.File.Delete(Me.FileStr)
    End Sub


    Public Overrides Function ToString() As String
        Return "File:" & Me.FileStr
    End Function


    Private Sub pd_PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        Dim linesPerPage As Single = 0
        Dim yPos As Single = 0
        Dim count As Integer = 0
        Dim leftMargin As Single = ev.MarginBounds.Left
        Dim topMargin As Single = ev.MarginBounds.Top
        Dim line As String = Nothing

        ' Calculate the number of lines per page.
        linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics)

        ' Print each line of the file.
        While count < linesPerPage
            line = streamToPrint.ReadLine()
            If line Is Nothing Then
                Exit While
            End If
            yPos = topMargin + count * Me.PrintFont.GetHeight(ev.Graphics)
            ev.Graphics.DrawString(line, Me.PrintFont, Brushes.Black, leftMargin, yPos, New StringFormat())
            count += 1
        End While

        ' If more lines exist, print another page.
        If (line IsNot Nothing) Then
            ev.HasMorePages = True
        Else
            ev.HasMorePages = False
        End If
    End Sub
End Class
