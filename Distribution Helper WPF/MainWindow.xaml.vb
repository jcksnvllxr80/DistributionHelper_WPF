Imports MahApps.Metro.Controls
'Imports Xceed.Wpf.Toolkit
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO

Class MainWindow
    Inherits MetroWindow

    Dim tempInfoString = ""
    Dim DistributionPrograms As New LinkedList(Of Object)
    Dim DistributionDataLoaded As Boolean = False
    Dim locationInfo As LocationData
    Dim MainConnection = New SqlClient.SqlConnection("Data Source=XJALAP0569\SQLEXPRESS;Initial Catalog=Distributions;Integrated Security=True;MultipleActiveResultSets=True")

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub Distribution_Helper_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        fillDataGridFromDB()
    End Sub

    Private Sub fillDataGridFromDB()
        Dim DistributionsDataSet As Distribution_Helper_WPF.DistributionsDataSet = CType(Me.FindResource("DistributionsDataSet"), Distribution_Helper_WPF.DistributionsDataSet)
        'Load data into the table Distributions. You can modify this code as needed.
        Dim DistributionsDataSetDistributionsTableAdapter As Distribution_Helper_WPF.DistributionsDataSetTableAdapters.DistributionsTableAdapter =
            New Distribution_Helper_WPF.DistributionsDataSetTableAdapters.DistributionsTableAdapter()
        DistributionsDataSetDistributionsTableAdapter.Fill(DistributionsDataSet.Distributions)
        Dim DistributionsViewSource As System.Windows.Data.CollectionViewSource = CType(Me.FindResource("DistributionsViewSource"), System.Windows.Data.CollectionViewSource)
        DistributionsViewSource.View.MoveCurrentToFirst()

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


    Private Sub InsertDistInfoToDatabase()
        Dim connection = MainConnection
        connection.Open()

        Dim totalProgress = DistributionPrograms.Count
        Dim currentProgress = 1
        Me.ProgressBar.Visibility = Visibility.Visible

        For Each Prog In DistributionPrograms
            If Prog Is Nothing Then
                Exit For
            End If
            Dim nextPrimaryKey = GetNextPrimaryKey(connection)
            Dim nextRevNumber = getNextRevisionNumber(connection, Prog.GetName)
            Prog.InsertDistributionToDB(connection, nextPrimaryKey, nextRevNumber)
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
        BindGridAndFilterByProgName(Me.LocationNameText.Text)

        Me.ProgressBar.Visibility = Visibility.Hidden
    End Sub


    Private Sub BindGridAndFilterByProgName(progName As String)
        'DataGridView1.DataSource = DistributionsBindingSource

        progName = Trim(progName)
        If progName = "" Then Exit Sub

        fillDataGridFromDB()

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
        cmd.Connection = MainConnection
        cmd.Connection.Open()

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
        cmd.Connection = MainConnection
        cmd.Connection.Open()

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
        cmd.Connection = MainConnection
        cmd.Connection.Open()

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


    Private Sub DistroPathText_KeyDown(sender As Object, e As KeyEventArgs) Handles DistroPathText.KeyDown
        If Int(e.Key) = 6 Then
            If System.IO.Directory.Exists(Me.DistroPathText.Text) Then
                'Enter or TAB and path exists
                FindFilesAndCreateProgramSelectWindow()
                tempInfoString = ""
                StatusLabel.Text = tempInfoString
            Else
                tempInfoString = "Invalid path"
                StatusLabel.Text = tempInfoString
            End If
        End If
        'Console.Write("pressed " & Int(e.Key) & vbCrLf)
    End Sub


    Private Sub FindFilesAndCreateProgramSelectWindow()

        StatusLabel.Text = "Looking for software to distribute..."
        Dim j = 0
        Dim filesys = CreateObject("Scripting.FileSystemObject")
        If filesys.FolderExists(Me.DistroPathText.Text) Then
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
                        checkBox.Content = filename
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

        DistributionPrograms.Clear()
        For Each controlObj In Me.ProgramWrapPanel.FindChildren(Of CheckBox)
            If controlObj.GetType() Is GetType(CheckBox) Then
                If controlObj.IsChecked Then
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
        objWord.ChangeFileOpenDirectory(currentDir)

        With objWord.FileDialog(msoFileDialogOpen)
            .Title = instructionString
            .AllowMultiSelect = False
            .Filters.Clear
            .AllowMultiSelect = False
            .Filters.Clear
            '.Filters.Add("Log Files", "*.LOG")

            If .Show = -1 Then  'short form
                For Each Folder In .SelectedItems  'short form
                    Dim objFile = fso.GetFolder(Folder)
                    FolderPath = objFile.Path
                    i += 1
                Next
            End If
        End With
        If i = 0 Then
            GetDistributionFolder = ""
        Else
            GetDistributionFolder = FolderPath
        End If
        objWord.Quit
        objWord = Nothing
    End Function


    Private Function CreateDistributionDocument() As FlowDocument
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

        'PrintToolBttn.Enabled = True
        'PrintMenuItem.Enabled = True

        'PrintPrevToolBttn.Enabled = True
        'PrintPreviewMenuItem.Enabled = True
        PrintPreviewTab.Visibility = Visibility.Visible
        LocationInfoViewer.Document = CreateDistributionDocument()

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

        'PrintToolBttn.Enabled = False
        'PrintMenuItem.Enabled = False

        'PrintPrevToolBttn.Enabled = False
        'PrintPreviewMenuItem.Enabled = False
        PrintPreviewTab.Visibility = Visibility.Hidden

        'SaveToolBttn.Enabled = False
        'SaveMenuItem.Enabled = False
        'SaveAsMenuItem.Enabled = False

        InsertToDBToolBttn.IsEnabled = False
        InsertToDBMenuItem.IsEnabled = False

        CreateLabelsToolBttn.IsEnabled = False

        'DistEmailToolBttn.Enabled = False

        'CreateLetterToolBttn.Enabled = False

    End Sub


    Private Sub CustomerComboBox_TextChanged(sender As Object, e As EventArgs) Handles CustomerComboBox.TextChanged
        If Not (Me.CustomerComboBox.Text.Trim.Equals("") Or Me.InternalJobNumComboBox.Text.Trim.Equals("") Or Me.LocationNameText.Text.Trim.Equals("")) And DistributionDataLoaded Then
            EnableCreationControls()
            tempInfoString = ""
        Else
            DisableCreationControls()
            tempInfoString = "Fill customer, internal job number, and location name fields."
        End If
        StatusLabel.Text = tempInfoString
    End Sub


    Private Sub InternalJobNumComboBox_TextChanged(sender As Object, e As EventArgs) Handles InternalJobNumComboBox.TextChanged
        If Not (Me.CustomerComboBox.Text.Trim = "" Or Me.InternalJobNumComboBox.Text.Trim = "" Or Me.LocationNameText.Text.Trim = "") And DistributionDataLoaded Then
            EnableCreationControls()
            tempInfoString = ""
        Else
            DisableCreationControls()
            tempInfoString = "Fill customer, internal job number, and location name fields."
        End If
        StatusLabel.Text = tempInfoString
    End Sub


    Private Sub LocationNameText_TextChanged(sender As Object, e As EventArgs) Handles LocationNameText.TextChanged
        If Not (Me.CustomerComboBox.Text.Trim = "" Or Me.InternalJobNumComboBox.Text.Trim = "" Or Me.LocationNameText.Text.Trim = "") And DistributionDataLoaded Then
            EnableCreationControls()
            tempInfoString = ""
        Else
            DisableCreationControls()
            tempInfoString = "Fill customer, internal job number, and location name fields."
        End If
        StatusLabel.Text = tempInfoString
    End Sub


    Private Sub CreateLabelsToolBttn_Click(sender As Object, e As EventArgs) Handles CreateLabelsToolBttn.Click
        CreateLabels()
    End Sub


    Private Sub CreateLabels()
        Dim labelPath = FindOrCreateLabelsDirectory()

        Dim doc As XDocument = XDocument.Load("Blank.label")
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
        BindGridAndFilterByProgName(Me.LocationNameText.Text)

        InsertToDBToolBttn.IsEnabled = True
        InsertToDBMenuItem.IsEnabled = True
        RefreshDBToolBttn.IsEnabled = True
        DatabaseTab.Visibility = Visibility.Visible

        CreateLabelsToolBttn.IsEnabled = True

        'DistEmailToolBttn.isEnabled = True

        'CreateLetterToolBttn.isEnabled = True

    End Sub


    Private Sub DisableCreationControls()
        InsertToDBToolBttn.IsEnabled = False
        InsertToDBMenuItem.IsEnabled = False
        RefreshDBToolBttn.IsEnabled = False
        DatabaseTab.Visibility = Visibility.Hidden

        CreateLabelsToolBttn.IsEnabled = False

        'DistEmailToolBttn.isEnabled = False

        'CreateLetterToolBttn.IsEnabled = False

    End Sub


    Private Sub Exit_Application(sender As Object, e As EventArgs)
        Close()
    End Sub

End Class
