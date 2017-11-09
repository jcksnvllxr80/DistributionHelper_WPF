Imports MahApps.Metro.Controls
Imports System.Drawing
Imports System.IO
Imports System.ComponentModel
Imports Xceed.Wpf.Toolkit
'Imports Microsoft.Office.Interop.Outlook



Class MainWindow
    Inherits MetroWindow

    Dim tempInfoString = ""
    Dim DistributionPrograms As New LinkedList(Of Object)
    Dim DistributionDataLoaded As Boolean = False
    Dim locationInfo As LocationData
    Dim MainConnection = New SqlClient.SqlConnection(
                            "Data Source=XJALAP0569\SQLEXPRESS;Initial Catalog=Distributions;
                            Integrated Security=True;MultipleActiveResultSets=True")
    Dim AutoCompleteURLs

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Function GetConnectionOpen() As SqlClient.SqlConnection
        Try
            MainConnection.Open()
            Return MainConnection
        Catch
            MsgBox("Server conection error.")
            Close()
        End Try
    End Function


    Private Sub PrintLocationInfo()
        Dim printer As New PrintDialog
        Dim result = printer.ShowDialog()
        If result Then
            Dim CloneDoc As FlowDocument = LocationInfoViewer.Document
            CloneDoc.PageHeight = printer.PrintableAreaHeight
            CloneDoc.PageWidth = printer.PrintableAreaWidth
            CloneDoc.Foreground = System.Windows.Media.Brushes.Black
            Dim idocument As IDocumentPaginatorSource = CloneDoc
            printer.PrintDocument(idocument.DocumentPaginator, Me.DistroPathText.Text & " Location Info")
        End If
    End Sub

    Private Sub Distribution_Helper_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        FillDataGridFromDB()
    End Sub


    Private Sub FillDataGridFromDB()
        Dim DistributionsDataSet As Distribution_Helper_WPF.DistributionsDataSet = CType(Me.FindResource("DistributionsDataSet"), Distribution_Helper_WPF.DistributionsDataSet)
        'Load data into the table Distributions. You can modify this code as needed.
        Dim DistributionsDataSetDistributionsTableAdapter As Distribution_Helper_WPF.DistributionsDataSetTableAdapters.DistributionsTableAdapter =
                New Distribution_Helper_WPF.DistributionsDataSetTableAdapters.DistributionsTableAdapter()
        DistributionsDataSetDistributionsTableAdapter.Fill(DistributionsDataSet.Distributions)

        Dim progName = Trim(Me.LocationNameText.Text)
        If FilterDatabaseByLocation.IsChecked Then
            If Not progName = "" Then
                Dim dt = New DistributionsDataSet.DistributionsDataTable
                DistributionsDataSetDistributionsTableAdapter.Fill(dt)
                DistributionsDataGrid.DataContext = dt
                Dim blv As IBindingListView = dt.DefaultView
                blv.Filter = "locationName = '" & Me.LocationNameText.Text & "'"
                DistributionsDataGrid.ItemsSource = dt.DefaultView
            Else
                MsgBox("Location Name Field is empty.")
            End If
        Else
            Dim DistributionsViewSource As System.Windows.Data.CollectionViewSource = CType(Me.FindResource("DistributionsViewSource"), System.Windows.Data.CollectionViewSource)
            DistributionsViewSource.View.MoveCurrentToFirst()
            DistributionsDataGrid.ItemsSource = DistributionsDataSet.Distributions.DefaultView
        End If

        'read from database and fill in combo boxes
        LoadCustomerComboBox()
        LoadCustomerJobNumComboBox()
        LoadInternalJobNumComboBox()
    End Sub


    Private Sub Text_MouseEnter(sender As Object, e As MouseEventArgs)
        StatusLabel.Text = sender.Tag
    End Sub

    Private Sub Text_MouseLeave(sender As Object, e As MouseEventArgs)
        StatusLabel.Text = tempInfoString
    End Sub


    Private Sub CreateEmail()
        Dim subjectSubstr = "– Program Books & Chips"
        Dim bodyStr As String = vbTab & vbTab & vbTab & vbCr & locationInfo.GetLocationName & vbCr & locationInfo.GetCity & ", " &
                    locationInfo.GetState & " / MP. " & locationInfo.GetMilePost & vbCr & locationInfo.GetDivision & " Division / " &
                    locationInfo.GetSubdivision & " Subdivision" & vbCr & locationInfo.GetCustomerNumber & vbCr & locationInfo.GetInternalNumber &
                    vbCr & vbCr & "I have sent " & locationInfo.GetLocationName & " (" & locationInfo.GetMilePost &
                    ") program book(s) and EPROMs (executive and application) to the address above with FeDEx " &
                    extraText.Text & " shipping." &
                    vbCr & vbCr &
                    DistributionAddressText.Text &
                    vbCr & vbCr &
                    vbTab & "Tracking Number:" & vbTab & vbTab & extraText.Text & vbCr &
                    vbTab & "Invoice Number:" & vbTab & vbTab & extraText.Text & vbCr &
                    vbTab & "Reference:" & vbTab & vbTab & "XORAIL CORP" & vbCr &
                    vbTab & "Service type:" & vbTab & vbTab & extraText.Text & vbCr &
                    vbTab & "Packaging type:" & vbTab & vbTab & "FedEx Box"
        Dim TO_Recipients As String = ""
        Dim CC_Recipients As String = "Miller, John <J.Miller@xorail.com>; Holmes, Daryl (D.Holmes@xorail.com)"
        Dim subjectStr As String = locationInfo.GetCustomerNumber & "; " & locationInfo.GetLocationName & " " & subjectSubstr

        Dim otlApp = CreateObject("Outlook.Application")
        Dim olMailItem As Object = Nothing
        Dim otlNewMail = otlApp.CreateItem(olMailItem)
        Dim WshShell = CreateObject("WScript.Shell")
        With otlNewMail
            .Subject = subjectStr
            WshShell.AppActivate(subjectStr & " - Message (HTML)")
            .To = TO_Recipients
            .CC = CC_Recipients
            .Display
            Dim objDoc = otlApp.ActiveInspector.WordEditor
            Dim objSel = objDoc.Windows(1).Selection
            objSel.InsertBefore(bodyStr)
        End With
        WshShell = Nothing
        otlNewMail = Nothing
        otlApp = Nothing
    End Sub


    Private Sub InsertDistInfoToDatabase()
        Dim connection = GetConnectionOpen()

        Dim totalProgress = DistributionPrograms.Count
        Dim currentProgress = 1
        Me.ProgressBar.Visibility = Visibility.Visible

        For Each Prog In DistributionPrograms
            If Prog Is Nothing Then
                Exit For
            End If
            Dim nextPrimaryKey = GetNextPrimaryKey(connection)
            Dim nextRevNumber = getNextRevisionNumber(connection, Prog.GetName)
            Dim userFieldData As New MainWindowData(Me.LocationNameText.Text, Me.CustomerComboBox.Text,
                                                    Me.CustomerJobNumComboBox.Text, Me.InternalJobNumComboBox.Text,
                                                    Me.DistributionDatePicker.DisplayDate)
            Prog.InsertDistributionToDB(connection, nextPrimaryKey, nextRevNumber, userFieldData)
            'Console.WriteLine("Primary Key: " & nextPrimaryKey & vbCrLf & "Revision Number: " & nextRevNumber & vbCrLf & vbCrLf)
            currentProgress += 1
            Me.ProgressBar.Value = 100 * currentProgress / totalProgress
        Next

        Me.ProgressBar.Value = 100
        StatusLabel.Text = "Distribution information was inserted successfully to the database"

        connection.Close()

        LoadInternalJobNumComboBox()
        LoadCustomerComboBox()
        LoadCustomerJobNumComboBox()
        FillDataGridFromDB()

        Me.ProgressBar.Visibility = Visibility.Hidden
    End Sub


    Private Function GetNextPrimaryKey(con As SqlClient.SqlConnection) As Integer
        Dim cmd As New SqlClient.SqlCommand
        Dim reader As SqlClient.SqlDataReader

        cmd.CommandText = "SELECT MAX(ID) FROM Distributions"
        cmd.CommandType = CommandType.Text
        cmd.Connection = con

        reader = cmd.ExecuteReader()
        ' Data is accessible through the DataReader object here.
        reader.Read()
        Dim maxPrimaryKey = CType(reader, IDataRecord)

        If TypeOf maxPrimaryKey(0) Is DBNull Then
            Return 0
        Else
            Return Int(maxPrimaryKey(0)) + 1
        End If

    End Function


    Private Sub LoadCustomerComboBox()
        Dim cmd As New SqlClient.SqlCommand
        Dim reader As SqlClient.SqlDataReader
        Dim nextRev = 0

        cmd.CommandText = "SELECT DISTINCT customer FROM Distributions"
        cmd.CommandType = CommandType.Text
        cmd.Connection = GetConnectionOpen()

        CustomerComboBox.ItemsSource = Nothing
        reader = cmd.ExecuteReader()
        Dim candidates As New List(Of String)
        While reader.Read()
            Dim customer = CType(reader, IDataRecord)

            If Not TypeOf customer(0) Is DBNull Then
                candidates.Add(customer(0))
            End If
        End While

        CustomerComboBox.ItemsSource = candidates
        cmd.Connection.Close()
    End Sub


    Private Sub LoadInternalJobNumComboBox()
        Dim cmd As New SqlClient.SqlCommand
        Dim reader As SqlClient.SqlDataReader
        Dim nextRev = 0

        cmd.CommandText = "SELECT DISTINCT internalJobNum FROM Distributions"
        cmd.CommandType = CommandType.Text
        cmd.Connection = GetConnectionOpen()


        InternalJobNumComboBox.ItemsSource = Nothing
        reader = cmd.ExecuteReader()
        Dim candidates As New List(Of String)
        While reader.Read()
            Dim internalJobNum = CType(reader, IDataRecord)

            If Not TypeOf internalJobNum(0) Is DBNull Then
                candidates.Add(internalJobNum(0))
            End If
        End While

        InternalJobNumComboBox.ItemsSource = candidates
        cmd.Connection.Close()
    End Sub


    Private Sub LoadCustomerJobNumComboBox()
        Dim cmd As New SqlClient.SqlCommand
        Dim reader As SqlClient.SqlDataReader
        Dim nextRev = 0

        cmd.CommandText = "SELECT DISTINCT customerJobNum FROM Distributions WHERE customerJobNum <> '';"
        cmd.CommandType = CommandType.Text
        cmd.Connection = GetConnectionOpen()

        CustomerJobNumComboBox.InputBindings.Clear()
        reader = cmd.ExecuteReader()
        Dim candidates As New List(Of String)
        While reader.Read()
            Dim customerJobNum = CType(reader, IDataRecord)

            If Not TypeOf customerJobNum(0) Is DBNull Then
                candidates.Add(customerJobNum(0))
            End If
        End While

        CustomerJobNumComboBox.ItemsSource = candidates
        cmd.Connection.Close()
    End Sub


    Private Function getNextRevisionNumber(con As SqlClient.SqlConnection, programName As String) As Integer
        Dim cmd As New SqlClient.SqlCommand
        Dim reader As SqlClient.SqlDataReader
        Dim nextRev = 0

        cmd.CommandText = "SELECT MAX(revision) FROM Distributions WHERE (programName = '" & programName & "' AND internalJobNum = '" & InternalJobNumComboBox.Text & "');"
        cmd.CommandType = CommandType.Text
        cmd.Connection = con

        reader = cmd.ExecuteReader()
        ' Data is accessible through the DataReader object here.
        If reader.Read() Then
            Dim revision = CType(reader, IDataRecord)

            If Not TypeOf revision(0) Is DBNull Then
                nextRev = Int(revision(0)) + 1
            End If
        End If

        Return nextRev
    End Function


    Private Sub DistroPathText_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles DistroPathText.PreviewKeyDown
        If e.Key = Key.Enter Then
            If System.IO.Directory.Exists(Me.DistroPathText.Text) Then
                'Enter and path exists
                FindFilesAndCreateProgramSelectWindow()
                tempInfoString = ""
                StatusLabel.Text = tempInfoString
            Else
                tempInfoString = "Invalid path"
            End If
            StatusLabel.Text = tempInfoString
        Else
            PopulatePathDropBox()
        End If
        'Console.Write("pressed " & Int(e.Key) & vbCrLf)
    End Sub


    Private Sub FindFilesAndCreateProgramSelectWindow()
        StatusLabel.Text = "Looking for software to distribute..."
        Dim j = 0
        Dim filesys = CreateObject("Scripting.FileSystemObject")
        If filesys.FolderExists(Me.DistroPathText.Text) Then
            DistributionPrograms.Clear()
            ProgramWrapPanel.Children.RemoveRange(0, ProgramWrapPanel.Children.Count)

            Dim Folder = filesys.getfolder(Me.DistroPathText.Text)
            For Each File In Folder.Files
                Dim filetype = filesys.GetExtensionName(File)
                Dim filename = filesys.GetFileName(File)

                If InStr(filename, "~") = 0 And Not filetype Is Nothing Then 'dont use system files
                    Dim typeStr = filetype.ToUpper
                    If New String() {"CCF", "LOC", "ML2", "GN2", "MAP"}.Contains(typeStr) Then
                        ' Create a checkbox
                        Dim checkBox As New CheckBox()
                        ' Add checkbox to form
                        Me.ProgramWrapPanel.Children.Add(checkBox)

                        'Set size, position, ...
                        checkBox.Content = "_" & filename
                        checkBox.Tag = Me.DistroPathText.Text
                        checkBox.Width = 150
                        checkBox.FontSize = 14.0
                        checkBox.IsChecked = True
                        j = j + 1
                    End If
                End If
            Next
            Dim buttonPanel = New StackPanel
            buttonPanel.Orientation = Orientation.Horizontal
            Me.ProgramWrapPanel.Children.Add(buttonPanel)
            If j > 0 Then
                Dim OkButton As New Button()
                buttonPanel.Children.Add(OkButton)
                OkButton.Content = "OK"
                AddHandler OkButton.Click, AddressOf OkButton_Click


                Dim CancelButton As New Button()

                buttonPanel.Children.Add(CancelButton)
                CancelButton.Content = "Cancel"
                AddHandler CancelButton.Click, AddressOf CancelButton_Click

                ShowProgramSelectorPanel(True)
            Else
                MsgBox("There are no files to distribute in the selected folder.")
            End If
        Else
            Me.DistroPathText.Text = "Invalid path."
        End If
    End Sub


    Private Sub ShowProgramSelectorPanel(open As Boolean)
        If open Then
            ProgramSelectorLabel.Visibility = Visibility.Visible
            ProgramWrapPanel.Visibility = Visibility.Visible
        Else
            ProgramSelectorLabel.Visibility = Visibility.Hidden
            ProgramWrapPanel.Visibility = Visibility.Hidden
            ProgramWrapPanel.Children.RemoveRange(0, ProgramWrapPanel.Children.Count)
        End If
    End Sub


    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        Dim dirPathComponents = Split(Me.DistroPathText.Text, "\")
        If dirPathComponents.Length > 1 Then
            Me.CustomerComboBox.Text = UCase(dirPathComponents(1))
        End If

        'check for info worksheet in XRL folder
        Dim infoFile = GetInfoFile()

        If infoFile <> "" Then
            StatusLabel.Text = "Reading info worksheet..."
            locationInfo = New LocationData(infoFile)
            Me.CustomerJobNumComboBox.Text = locationInfo.GetCustomerNumber()
            Me.InternalJobNumComboBox.Text = locationInfo.GetInternalNumber()
            Me.LocationNameText.Text = locationInfo.GetLocationName()
            locationInfo.SetCustomer(Me.CustomerComboBox.Text)
        End If

        'DistributionPrograms.Clear()
        'ProgramWrapPanel.Children.RemoveRange(0, ProgramWrapPanel.Children.Count)

        For Each controlObj In Me.ProgramWrapPanel.FindChildren(Of CheckBox)
            If controlObj.GetType() Is GetType(CheckBox) Then
                If controlObj.IsChecked Then
                    controlObj.Content = controlObj.Content.Substring(1)
                    If DistributionPrograms.First Is Nothing Then
                        DistributionPrograms.AddFirst(DetermineProgramType(controlObj))
                    Else
                        DistributionPrograms.AddLast(DetermineProgramType(controlObj))
                    End If

                End If
            End If
        Next
        ShowProgramSelectorPanel(False)
        For Each Prog In DistributionPrograms
            If Not Prog Is Nothing Then
                EnableDataViewFunctions()
                Exit For
            Else
                DisableDataViewFunctions()
            End If
        Next
    End Sub


    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        ShowProgramSelectorPanel(False)
    End Sub


    Private Function DetermineProgramType(chkbox As CheckBox) As ProgramFile
        Dim filetype = System.IO.Path.GetExtension(chkbox.Content)
        Dim filename = System.IO.Path.GetFileNameWithoutExtension(chkbox.Content)
        Dim filePath = chkbox.Tag

        Dim typeStr = filetype.ToUpper
        Dim program
        'MsgBox(typeStr)
        Select Case typeStr
            Case ".CCF"
                If System.IO.File.Exists(Me.DistroPathText.Text & "\" & filename & ".H30") Then
                    program = New NonVitalHLC(filename, filePath, "ACE")
                ElseIf System.IO.File.Exists(Me.DistroPathText.Text & "\" & filename & ".H14") Then
                    program = New VitalHLC(filename, filePath, "ACE")
                Else
                    program = New ElectroLogIXS(filename, filePath)
                End If
            Case ".LOC"
                If System.IO.File.Exists(Me.DistroPathText.Text & "\" & filename & ".H30") Then
                    program = New NonVitalHLC(filename, filePath, "ALC")
                Else
                    program = New VitalHLC(filename, filePath, "ALC")
                End If
            Case ".ML2"
                program = New ML2(filename, filePath)
            Case ".GN2"
                program = New GN2(filename, filePath)
            Case ".MAP"
                program = New EC4(filename, filePath)
            Case Else
                program = Nothing
        End Select
        Return program

    End Function


    Private Sub SelectFolder()
        Dim StartDir As String
        If Not Me.DistroPathText.Text = "" Then
            StartDir = Me.DistroPathText.Text
        Else
            StartDir = "P:\"
        End If
        Dim FilesDirectory = GetDistributionFolder(StartDir, "Browse for folder containing distribution files")
        StatusLabel.Text = tempInfoString
        If Not FilesDirectory = "" Then
            DistroPathText.Text = FilesDirectory
            FindFilesAndCreateProgramSelectWindow()
        End If
    End Sub


    Private Function GetInfoFile() As String
        StatusLabel.Text = "Looking for info worksheet..."
        Dim searchStr = "XRL"
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim f = fso.GetFolder(DistroPathText.Text)
        Dim subFldrs = f.SubFolders
        For Each subF In subFldrs
            If subF.name.Length >= searchStr.Length Then
                If (UCase(subF.name.Substring(0, 3)) = searchStr) Then
                    Dim infoPathStr = fso.GetAbsolutePathName(subF)
                    Dim folder = fso.GetFolder(infoPathStr)
                    For Each f In folder.Files
                        If f.Name.Length > 19 Then
                            If (f.Name.Substring(f.Name.Length - 19) = "info worksheet.xlsm") Then
                                Return infoPathStr & "\" & f.Name
                                fso = Nothing
                                Exit Function
                            End If
                        End If
                    Next
                End If
            End If
        Next
        Return ""
    End Function


    Private Function GetDistributionFolder(currentDir As String, instructionString As String) As String
        Const msoFileDialogOpen = 4
        Dim FolderPath = Nothing
        Dim i = 0
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim objWord = CreateObject("Word.Application")
        Dim DistributionFolder As String
        objWord.ChangeFileOpenDirectory(currentDir)

        With objWord.FileDialog(msoFileDialogOpen)
            .Title = instructionString
            .AllowMultiSelect = False
            .Filters.Clear
            '.Filters.Add("Log Files", "*.LOG")

            If .Show = -1 Then  'short form
                objWord.Visible = True
                objWord.Activate
                For Each Folder In .SelectedItems  'short form
                    Dim objFile = fso.GetFolder(Folder)
                    FolderPath = objFile.Path
                    i += 1
                Next
            End If
        End With
        If i = 0 Then
            DistributionFolder = ""
        Else
            DistributionFolder = FolderPath
        End If
        objWord.Quit
        objWord = Nothing
        Return DistributionFolder
    End Function


    Private Function CreateDistInfoDocument() As FlowDocument
        Dim DocFont As New Font("Arial", 12)
        Dim ProgInfoStr = ""
        If Not locationInfo Is Nothing Then
            ProgInfoStr = locationInfo.ToString()
        End If

        For Each Prog In DistributionPrograms
            If Prog Is Nothing Then
                Exit For
            End If
            ProgInfoStr = ProgInfoStr & Prog.ToString & vbCrLf & vbCrLf
        Next
        Dim paragraph As New Paragraph
        paragraph.Inlines.Add(ProgInfoStr)
        Return New FlowDocument(paragraph)
    End Function


    Private Sub EnableDataViewFunctions()
        DistributionDataLoaded = True

        PrintMenu.IsEnabled = True
        PrintToolBttn.IsEnabled = True
        'PrintLocInfoMenuItem.IsEnabled = True

        PrintPreviewTab.Visibility = Visibility.Visible
        LocationInfoViewer.Document = CreateDistInfoDocument()

        'SaveToolBttn.Enabled = True
        'SaveMenuItem.Enabled = True
        'SaveAsMenuItem.Enabled = True

        If Me.CustomerComboBox.Text.Trim = "" Or Me.InternalJobNumComboBox.Text.Trim = "" Or Me.LocationNameText.Text.Trim = "" Then
            tempInfoString = "Fill customer, internal job number, and location name fields."
            StatusLabel.Text = tempInfoString
        Else
            EnableCreationControls()
        End If
    End Sub


    Private Sub DisableDataViewFunctions()
        DistributionDataLoaded = False

        PrintMenu.IsEnabled = False
        PrintToolBttn.IsEnabled = False
        'PrintLocInfoMenuItem.IsEnabled = False

        PrintPreviewTab.Visibility = Visibility.Hidden

        'SaveToolBttn.Enabled = False
        'SaveMenuItem.Enabled = False
        'SaveAsMenuItem.Enabled = False

        InsertToDBToolBttn.IsEnabled = False
        InsertToDBMenuItem.IsEnabled = False

        CreateLabelsToolBttn.IsEnabled = False

        CreateEmailToolBttn.IsEnabled = False

        'CreateLetterToolBttn.Enabled = False
    End Sub


    Private Sub CheckFieldsForData()
        If DistributionDataLoaded Then
            If Not (Me.CustomerComboBox.Text.Trim = "" Or Me.InternalJobNumComboBox.Text.Trim = "" Or Me.LocationNameText.Text.Trim = "") Then
                EnableCreationControls()
                tempInfoString = ""
            Else
                DisableCreationControls()
                tempInfoString = "Fill customer, internal job number, and location name fields."
            End If
            StatusLabel.Text = tempInfoString
        End If
    End Sub


    Private Sub CustomerComboBox_TextChanged(sender As Object, e As EventArgs) Handles CustomerComboBox.TextChanged
        CheckFieldsForData()
    End Sub


    Private Sub InternalJobNumComboBox_TextChanged(sender As Object, e As EventArgs) Handles InternalJobNumComboBox.TextChanged
        CheckFieldsForData()
    End Sub


    Private Sub LocationNameText_TextChanged(sender As Object, e As EventArgs) Handles LocationNameText.TextChanged
        CheckFieldsForData()
    End Sub


    Private Sub CreateLabelsToolBttn_Click(sender As Object, e As EventArgs) Handles CreateLabelsToolBttn.Click
        CreateLabels()
    End Sub


    Private Sub CreateLabels()
        Dim labelPath = FindOrCreateLabelsDirectory()
        Dim doc As XDocument = XDocument.Load("resources\Blank.label")
        Dim labelnode = doc.Descendants("String")
        For Each prog In DistributionPrograms
            If Not prog Is Nothing Then
                If prog.GetEquipType() = "EC4" Then
                    labelnode(0).Value = prog.MAPLabelStr
                    labelnode(1).Value = prog.MAPLabelStr
                    doc.Save(labelPath & "\" & prog.GetName & ".label")
                ElseIf {"VHLC", "NVHLC"}.Contains(prog.GetEquipType()) Then
                    labelnode(0).Value = prog.evenLabelStr
                    labelnode(1).Value = prog.oddLabelStr
                    doc.Save(labelPath & "\" & prog.GetName & ".label")
                End If
            End If
        Next
    End Sub


    Private Function FindOrCreateLabelsDirectory() As String
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim distributionFldrPath As String = Nothing
        Dim labelsFldrPath As String = Nothing
        Dim searchStr As String

        searchStr = "Distribution"
        Dim f = fso.GetFolder(DistroPathText.Text)
        For Each subF In f.subFolders
            If subF.Name.Length >= searchStr.Length Then
                If (UCase(subF.name.Substring(0, 12)) = searchStr) Then
                    distributionFldrPath = fso.GetAbsolutePathName(subF)
                    Exit For
                End If
            End If
        Next

        If distributionFldrPath Is Nothing Then
            distributionFldrPath = DistroPathText.Text & "\" & searchStr
            System.IO.Directory.CreateDirectory(distributionFldrPath)
        End If

        searchStr = "Labels"
        f = fso.GetFolder(distributionFldrPath)
        For Each subF In f.subFolders
            If subF.Name.Length >= searchStr.Length Then
                If (UCase(subF.name.Substring(0, 6)) = searchStr) Then
                    labelsFldrPath = fso.GetAbsolutePathName(subF)
                    Exit For
                End If
            End If
        Next

        If labelsFldrPath Is Nothing Then
            labelsFldrPath = distributionFldrPath & "\" & searchStr
            System.IO.Directory.CreateDirectory(labelsFldrPath)
        End If

        Return labelsFldrPath
    End Function


    Private Sub EnableCreationControls()
        FillDataGridFromDB()

        FilterDatabaseByLocation.IsEnabled = True
        InsertToDBToolBttn.IsEnabled = True
        InsertToDBMenuItem.IsEnabled = True
        RefreshDBToolBttn.IsEnabled = True
        DatabaseMenu.IsEnabled = True
        DatabaseTab.Visibility = Visibility.Visible

        CreateEmailToolBttn.IsEnabled = True

        CreateLabelsToolBttn.IsEnabled = True

        'CreateLetterToolBttn.isEnabled = True

    End Sub


    Private Sub DisableCreationControls()
        FilterDatabaseByLocation.IsEnabled = False
        InsertToDBToolBttn.IsEnabled = False
        InsertToDBMenuItem.IsEnabled = False
        RefreshDBToolBttn.IsEnabled = False
        DatabaseMenu.IsEnabled = False
        DatabaseTab.Visibility = Visibility.Hidden

        CreateEmailToolBttn.IsEnabled = False

        CreateLabelsToolBttn.IsEnabled = False

        'CreateLetterToolBttn.IsEnabled = False

    End Sub


    Private Sub Exit_Application(sender As Object, e As EventArgs)
        Close()
    End Sub

    Private Sub DistroPathText_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles DistroPathText.MouseDoubleClick
        SelectFolder()
    End Sub


    Private Sub PopulatePathDropBox()

        Dim startDir = ""
        Dim searchStr
        Dim posOfSlash = DistroPathText.Text.LastIndexOf("\")
        If posOfSlash > 0 Then
            startDir = DistroPathText.Text.Substring(0, posOfSlash + 1)
            If posOfSlash = DistroPathText.Text.Length - 1 Then
                searchStr = ""
            Else
                searchStr = DistroPathText.Text.Substring(posOfSlash + 1)
            End If
        ElseIf DistroPathText.Text.Length = 2 Then
            searchStr = DistroPathText.Text & "\"
        ElseIf DistroPathText.Text.Length = 1 Then
            searchStr = DistroPathText.Text & ":\"
        Else
            searchStr = ""
        End If

        Try
            Dim contents = Directory.GetDirectories(startDir, searchStr & "*", SearchOption.TopDirectoryOnly)
            Dim contentsList = contents.ToList
            If contentsList IsNot Nothing Then
                DistroPathText.IsDropDownOpen = True
                DistroPathText.ItemsSource = contents.ToList
            End If
        Catch e As Exception
            Console.WriteLine("cant find " & startDir & searchStr & " : {0}", e.ToString())
        End Try
    End Sub
End Class
