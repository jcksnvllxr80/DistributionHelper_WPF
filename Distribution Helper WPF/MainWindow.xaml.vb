Imports MahApps.Metro.Controls
Imports System.Drawing
Imports System.IO
Imports System.ComponentModel
Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Printing
Imports System.Drawing.Printing
Imports PdfSharp.Pdf.Printing
'Imports Xceed.Wpf.Toolkit

Class MainWindow
    Inherits MetroWindow

    Private InsertToDatabseBGWorker As New BackgroundWorker()
    Private MineLocationDataBGWorker As New BackgroundWorker()
    Private createLetterBGWorker As New BackgroundWorker()
    Private printFilesBGWorker As New BackgroundWorker()
    Private tempInfoString = ""
    Private DistributionPrograms As New LinkedList(Of Object)
    Private DistributionDataLoaded As Boolean
    Private locationInfo As LocationData
    Private user As UserObject
    Private LocationDataFlowDoc As FlowDocument = Nothing
    Private LinkCompareFlowDoc As FlowDocument = Nothing
    Private RtvPrinter As String = "\\XJASRV0001\XJAPRT0019 - 4th Floor East - Xerox WorkCentre 7970 PCL6"

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub GetUser()
        user = New UserObject()
        Distribution_Helper.Title &= " - (" & user.FullName & ")"
        System.Console.WriteLine(vbCrLf & "Welcome to the Distribution Helper, " & user.FirstName &
                                 "." & vbCrLf & "--------------------------------------------" &
                                 vbCrLf & "Your credentials should be as shown below:" & vbCrLf & user.ToString)
    End Sub


    Private Function GetConnectionOpen() As SqlClient.SqlConnection
        Dim MainConnection = New SqlClient.SqlConnection(
                            "Data Source=XJALAP0569\SQLEXPRESS;Initial Catalog=Distributions;
                            Integrated Security=True;MultipleActiveResultSets=True")
        Try
            MainConnection.Open()
            Return MainConnection
        Catch
            MsgBox("Server conection error.")
            Close()
            Return Nothing
        End Try
    End Function


    Private Sub PrintLocationInfo(sender As Object, e As RoutedEventArgs)
        Dim printer As New PrintDialog
        Dim result = printer.ShowDialog()
        If result Then
            Dim CloneDoc As FlowDocument = LocationInfoViewer.Document
            CloneDoc.PageHeight = printer.PrintableAreaHeight
            CloneDoc.PageWidth = printer.PrintableAreaWidth
            CloneDoc.Foreground = System.Windows.Media.Brushes.Black
            Dim idocument As IDocumentPaginatorSource = CloneDoc
            Dim buttonPressed As Button = sender
            If buttonPressed.Name.Equals("PrintLinkCompMenuItem") Then
                If PrintLinkCompMenuItem.IsEnabled Then
                    printer.PrintDocument(idocument.DocumentPaginator, locationInfo.GetLocationName & " Remote Link Comparison")
                End If
            Else
                printer.PrintDocument(idocument.DocumentPaginator, Me.DistroPathText.Text & " Location Info")
            End If
        End If
    End Sub

    Private Sub Distribution_Helper_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        GetUser()

        DistroPathText.AutoCompleteMode = Forms.AutoCompleteMode.Suggest
        DistroPathText.AutoCompleteSource = Forms.AutoCompleteSource.AllSystemSources
        MyHost.Child = DistroPathText

        ShippingMethodBox.Items.Add("Standard (3-5 Days)")
        ShippingMethodBox.Items.Add("Express (1-2 Days)")
        ShippingMethodBox.Items.Add("Overnight")
        ShippingMethodBox.SelectedItem = "Standard (3-5 Days)"

        MineLocationDataBGWorker.WorkerReportsProgress = True
        'MineLocationDataBGWorker.WorkerSupportsCancellation = True
        AddHandler MineLocationDataBGWorker.DoWork, AddressOf BackgroundWorker_MineLocationData
        AddHandler MineLocationDataBGWorker.ProgressChanged, AddressOf BackgroundWorker_MiningProgressChanged
        AddHandler MineLocationDataBGWorker.RunWorkerCompleted, AddressOf BackgroundWorker_MiningWorkerCompleted

        InsertToDatabseBGWorker.WorkerReportsProgress = True
        'InsertToDatabseBGWorker.WorkerSupportsCancellation = True
        AddHandler InsertToDatabseBGWorker.DoWork, AddressOf BackgroundWorker_InsertToDB
        AddHandler InsertToDatabseBGWorker.ProgressChanged, AddressOf BackgroundWorker_InsertionProgressChanged
        AddHandler InsertToDatabseBGWorker.RunWorkerCompleted, AddressOf BackgroundWorker_InsertionWorkerCompleted

        createLetterBGWorker.WorkerReportsProgress = True
        'InsertToDatabseBGWorker.WorkerSupportsCancellation = True
        AddHandler createLetterBGWorker.DoWork, AddressOf BackgroundWorker_CreateLetter
        AddHandler createLetterBGWorker.ProgressChanged, AddressOf BackgroundWorker_LetterCreationProgressChanged
        AddHandler createLetterBGWorker.RunWorkerCompleted, AddressOf BackgroundWorker_LetterCreationWorkerCompleted

        printFilesBGWorker.WorkerReportsProgress = True
        'InsertToDatabseBGWorker.WorkerSupportsCancellation = True
        AddHandler printFilesBGWorker.DoWork, AddressOf BackgroundWorker_PrintFiles
        AddHandler printFilesBGWorker.ProgressChanged, AddressOf BackgroundWorker_FilePrintingProgressChanged
        AddHandler printFilesBGWorker.RunWorkerCompleted, AddressOf BackgroundWorker_FilePrintWorkerCompleted

    End Sub


    Private Sub GetDistroTextWidth()
        DistroPathText.Width = Me.Width - Int(DistroPathText.Width / 20)
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
            Dim DistributionsViewSource As System.Windows.Data.CollectionViewSource =
                CType(Me.FindResource("DistributionsViewSource"), System.Windows.Data.CollectionViewSource)
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
        Dim bodyStr As String = locationInfo.GetLocationName & vbCr & locationInfo.GetCity & ", " &
                    locationInfo.GetState & " / MP. " & locationInfo.GetMilePost & vbCr & locationInfo.GetDivision & " Division / " &
                    locationInfo.GetSubdivision & " Subdivision" & vbCr & locationInfo.GetCustomerNumber & vbCr & locationInfo.GetInternalNumber &
                    vbCr & vbCr & RecipientNameText.Text & vbCr & AddressStreetText.Text & vbCr &
                    AddressCityText.Text & ", " & AddressStateBox.Text & " " & AddressZipCodeText.Text & vbCr & vbCr &
                    "I have sent " & locationInfo.GetLocationName & " (" & locationInfo.GetMilePost &
                    ") program book(s) and EPROMs (executive and application) to the address above with FeDEx " &
                    ShippingMethodBox.SelectedItem & " shipping." & vbCr & vbCr &
                    "    Tracking Number:" & vbTab & TrackingNumberText.Text & vbCr &
                    "    Invoice Number:" & vbTab & InvoiceNumText.Text & vbCr &
                    "    Reference:" & vbTab & vbTab & "XORAIL CORP" & vbCr &
                    "    Service type:" & vbTab & vbTab & "FedEx " & ShippingMethodBox.Text & vbCr &
                    "    Packaging type:" & vbTab & "FedEx Box"
        Dim TO_Recipients As String = ""
        Dim CC_Recipients As String = "Miller, John <J.Miller@xorail.com>; Holmes, Daryl (D.Holmes@xorail.com)"
        Dim subjectStr As String = locationInfo.GetCustomerNumber & "; " & locationInfo.GetLocationName & " " & subjectSubstr

        Dim otlApp = CreateObject("Outlook.Application")
        Dim olMailItem As Object = Nothing
        Dim otlNewMail = otlApp.CreateItem(olMailItem)
        Dim WshShell = CreateObject("WScript.Shell")
        With otlNewMail
            .Display
            WshShell.AppActivate(subjectStr & " - Message (HTML)")
            .Subject = subjectStr
            .To = TO_Recipients
            .CC = CC_Recipients
            Dim objDoc = otlApp.ActiveInspector.WordEditor
            Dim objSel = objDoc.Windows(1).Selection
            objSel.InsertBefore(bodyStr)
        End With
        WshShell = Nothing
        otlNewMail = Nothing
        otlApp = Nothing
    End Sub


    Private Sub InsertDistInfoToDatabase()
        If locationInfo IsNot Nothing Then
            Me.ProgressBar.Visibility = Visibility.Visible
            TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Normal

            Dim myDate As Date = Me.DistributionDatePicker.DisplayDate
            InsertToDatabseBGWorker.RunWorkerAsync(myDate) ' this starts the background worker
        Else
            MsgBox("what to do when an attempt to add to the database is made but the fields were not properly filled at some point. this is know because locationInfo is = Nothing")
        End If
    End Sub


    Private Sub BackgroundWorker_InsertToDB(sender As Object, e As DoWorkEventArgs)
        'this sub is started when the run worker async is started
        Dim connection = GetConnectionOpen()
        Dim totalProgress = DistributionPrograms.Count
        Dim currentProgress = 0
        Dim myDate As Date = e.Argument

        For Each Prog In DistributionPrograms
            If Prog Is Nothing Then
                Exit For
            End If
            Dim nextPrimaryKey = GetNextPrimaryKey(connection)
            Dim nextRevNumber = GetNextRevisionNumber(connection, Prog.GetName, locationInfo.GetInternalNumber)
            Dim userFieldData As New MainWindowData(locationInfo.GetLocationName, locationInfo.GetCustomer,
                                                    locationInfo.GetCustomerNumber, locationInfo.GetInternalNumber,
                                                    myDate)
            Prog.InsertDistributionToDB(connection, nextPrimaryKey, nextRevNumber, userFieldData)
            'Console.WriteLine("Primary Key: " & nextPrimaryKey & vbCrLf & "Revision Number: " & nextRevNumber & vbCrLf & vbCrLf)
            currentProgress += 1
            InsertToDatabseBGWorker.ReportProgress(100 * currentProgress / totalProgress)
        Next

        connection.Close()
    End Sub


    Private Sub BackgroundWorker_InsertionWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        'this runs when the background worker has completed its running thread
        StatusLabel.Text = "Distribution information was inserted successfully to the database"
        LoadInternalJobNumComboBox()
        LoadCustomerComboBox()
        LoadCustomerJobNumComboBox()
        FillDataGridFromDB()
        Me.ProgressBar.Visibility = Visibility.Hidden
        Me.ProgressBar.Value = 0
        TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.None

    End Sub


    Private Sub BackgroundWorker_InsertionProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)
        'this is called when the background worker is told to report progress
        Me.ProgressBar.Value = e.ProgressPercentage
        TaskbarItemInfo.ProgressValue = CDbl(e.ProgressPercentage) / 100
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


    Private Function GetNextRevisionNumber(con As SqlClient.SqlConnection, programName As String, internalJobNum As String) As Integer
        Dim cmd As New SqlClient.SqlCommand
        Dim reader As SqlClient.SqlDataReader
        Dim nextRev = 0

        cmd.CommandText = "SELECT MAX(revision) FROM Distributions WHERE (programName = '" & programName & "' AND internalJobNum = '" & internalJobNum & "');"
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


    Private Sub AddChassisToLinksTab()
        ChassisListView.Items.Clear()
        Dim currentChassis = DistributionPrograms.First
        Do While currentChassis IsNot Nothing
            'currentChassis.Value
            Dim chassisListItem As New ListViewItem
            chassisListItem.Content = currentChassis.Value.GetName()
            chassisListItem.Background = Media.Brushes.Transparent
            AddHandler chassisListItem.PreviewMouseMove, AddressOf ItemListPreviewMouseMove
            If currentChassis.Value.GetEquipType().Equals("ElectroLogIXS") Then
                ChassisListView.Items.Add(chassisListItem)
            Else
                ChassisListView.Items.Clear()
                LinkCompareInstructions.Text = "Link compare is only available for ElectroLogIXS locations"
            End If
            currentChassis = currentChassis.Next
        Loop
    End Sub


    Private Sub AddDropPanelsForRemoteChassis(numberOfRemotes As Short)
        Dim NumRemotesQuotient As Integer
        If numberOfRemotes > 1 Then
            NumRemotesQuotient = Int(8 / numberOfRemotes)
        Else
            NumRemotesQuotient = 5
        End If

        For i = 1 To numberOfRemotes
            Dim myRemoteImage As New System.Windows.Controls.Image()
            'myRemoteImage.Name = "RemoteHouseDropImage"
            myRemoteImage.Source = New BitmapImage(New Uri("resources\blu_house_88x.png", UriKind.Relative))
            myRemoteImage.Stretch = Stretch.Fill

            Dim remoteDropPanel As New StackPanel
            remoteDropPanel.Name = "Remote" & i & "DropPanel"
            remoteDropPanel.Tag = "Remote" & i
            remoteDropPanel.AllowDrop = True
            remoteDropPanel.Children.Add(myRemoteImage)
            AddHandler remoteDropPanel.Drop, AddressOf DragAndDropStack_Drop
            AddHandler remoteDropPanel.DragOver, AddressOf DropPanel_DragOver
            AddHandler remoteDropPanel.DragLeave, AddressOf DropPanel_DragLeave

            Dim myTextBlock As New TextBlock

            Dim remoteLabel As New Label
            remoteLabel.Tag = remoteDropPanel.Tag
            remoteLabel.Foreground = System.Windows.Media.Brushes.DarkOliveGreen
            remoteLabel.FontSize = 13
            remoteLabel.VerticalContentAlignment = VerticalAlignment.Top
            remoteLabel.HorizontalContentAlignment = HorizontalAlignment.Center
            remoteLabel.Content = myTextBlock '"Remote " & i
            remoteLabel.Content.Text = "Remote " & i

            Dim ColumnLocation As Integer = NumRemotesQuotient + 2 * (i - 1)
            If ColumnLocation < 0 Then
                ColumnLocation = 0
            End If
            RemoteLinkGrid.Children.Add(remoteDropPanel)
            Grid.SetRow(remoteDropPanel, 1)
            Grid.SetColumn(remoteDropPanel, ColumnLocation)
            RemoteLinkGrid.Children.Add(remoteLabel)
            Grid.SetRow(remoteLabel, 2)
            Grid.SetColumn(remoteLabel, ColumnLocation)
        Next
    End Sub


    Private Sub FindFilesAndCreateProgramSelectPanel()
        StatusLabel.Text = "Looking for software to distribute..."
        Dim j = 0
        If Directory.Exists(Me.DistroPathText.Text) Then
            DistributionTab.Visibility = Visibility.Visible
            DistributionTabGrid.Visibility = Visibility.Visible
            DistributionPrograms.Clear()
            ProgramWrapPanel.Children.RemoveRange(0, ProgramWrapPanel.Children.Count)
            ChassisListView.Items.Clear()
            RemoteLinkGrid.Children.RemoveRange(0, RemoteLinkGrid.Children.Count)
            PrintListView.Items.Clear()

            For Each File In Directory.GetFiles(Me.DistroPathText.Text)
                Dim filesys = CreateObject("Scripting.FileSystemObject")
                Dim filetype = filesys.GetExtensionName(File)
                Dim filename = filesys.GetFileName(File)

                If InStr(filename, "~") = 0 And Not filetype Is Nothing Then 'dont use system files
                    Dim typeStr = filetype.ToUpper
                    If New String() {"CCF", "LOC", "ML2", "GN2", "MAP"}.Contains(typeStr) Then
                        ' Create a checkbox
                        Dim checkBox As New CheckBox()
                        ' Add checkbox to form
                        Me.ProgramWrapPanel.Children.Add(checkBox)

                        Dim textBlock As New TextBlock
                        textBlock.Text = filename

                        'Set size, position, ...
                        checkBox.Content = textBlock
                        checkBox.Tag = Me.DistroPathText.Text
                        checkBox.Width = 150
                        'checkBox.Height = 20
                        checkBox.FontSize = 14.0
                        checkBox.IsChecked = True
                        Dim separator As New Separator()
                        Me.ProgramWrapPanel.Children.Add(separator)
                        separator.BorderThickness = New Thickness(3)
                        separator.BorderBrush = System.Windows.Media.Brushes.Transparent
                        j = j + 1
                    End If

                    If New String() {"ALL", "RPT", "LER", "ML2", "DOC", "DOCX"}.Contains(typeStr) And Not filename.ToString.Contains("_print_list_") Then
                        ' Create a ListViewItem
                        Dim printListItem As New ListViewItem()
                        ' Add ListViewItem to form
                        printListItem.Content = filename
                        printListItem.Tag = File
                        printListItem.Background = Media.Brushes.Transparent
                        Me.PrintListView.Items.Add(printListItem)
                    End If
                End If
            Next

            Dim buttonPanel = New StackPanel With {
                .Orientation = Orientation.Vertical
            }
            buttonPanel.Height = 200
            buttonPanel.Width = 100
            Me.ProgramWrapPanel.Children.Add(buttonPanel)
            If j > 0 Then
                Dim OkButton As New Button()
                buttonPanel.Children.Add(OkButton)
                OkButton.Content = "OK"
                OkButton.Height = 32
                OkButton.Width = 90
                AddHandler OkButton.Click, AddressOf OkButton_Click

                Dim separator As New Separator()
                buttonPanel.Children.Add(separator)
                separator.BorderThickness = New Thickness(5)
                separator.BorderBrush = System.Windows.Media.Brushes.Transparent

                Dim CancelButton As New Button()
                buttonPanel.Children.Add(CancelButton)
                CancelButton.Content = "Cancel"
                CancelButton.Height = 32
                CancelButton.Width = 90
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
            ProgramSelectBorder.Visibility = Visibility.Visible
        Else
            ProgramSelectorLabel.Visibility = Visibility.Hidden
            ProgramWrapPanel.Visibility = Visibility.Hidden
            ProgramSelectBorder.Visibility = Visibility.Hidden
            ProgramWrapPanel.Children.RemoveRange(0, ProgramWrapPanel.Children.Count)
        End If
    End Sub


    Private Sub BackgroundWorker_MiningProgressChanged(sender As Object, e As ProgressChangedEventArgs)
        Me.ProgressBar.Value = e.ProgressPercentage
        TaskbarItemInfo.ProgressValue = CDbl(e.ProgressPercentage) / 100
        If Me.ProgressBar.Value = 20 Then
            Me.StatusLabel.Text = "Looking for info worksheet..."
        ElseIf Me.ProgressBar.Value = 35 Then
            Me.StatusLabel.Text = "Reading info worksheet..."
        ElseIf Me.ProgressBar.Value = 50 Then
            Me.StatusLabel.Text = "Mining location information..."
        End If
    End Sub


    Private Sub BackgroundWorker_MiningWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        Me.ProgressBar.Visibility = Visibility.Hidden
        TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.None

        If locationInfo IsNot Nothing Then
            Me.CustomerJobNumComboBox.Text = locationInfo.GetCustomerNumber()
            Me.InternalJobNumComboBox.Text = locationInfo.GetInternalNumber()
            Me.LocationNameText.Text = locationInfo.GetLocationName()
            Me.AddressStateBox.Text = locationInfo.GetState()
            Me.AddressCityText.Text = locationInfo.GetCity()
            locationInfo.SetCustomer(Me.CustomerComboBox.Text)

            Dim reducedTestDir As String = Nothing
            If locationInfo.GetRTVPfolderNum IsNot Nothing Then
                'Console.WriteLine("RTVP folder num is """ & locationInfo.GetRTVPfolderNum & """.")

                Dim validationYearDir = "P:\Validation\" & locationInfo.GetRTVPyear & "\" & locationInfo.GetCustomer
                reducedTestDir = FindSubfolderContaining(validationYearDir, locationInfo.GetRTVPfolderNum & "_", 0)

                Dim reducedTestDistributionDir As String = Nothing
                If reducedTestDir IsNot Nothing Then
                    reducedTestDistributionDir = FindSubfolderContaining(reducedTestDir, "Distribution", 0)
                End If

                If reducedTestDistributionDir IsNot Nothing Then
                    Dim fso = CreateObject("Scripting.FileSystemObject")
                    For Each file In fso.GetFolder(reducedTestDistributionDir).Files
                        If System.IO.Path.GetExtension(file.Name).ToUpper.Equals(".PDF") Then
                            ' Create a ListViewItem
                            Dim printListItem As New ListViewItem()
                            ' Add ListViewItem to form
                            printListItem.Content = file.Name
                            printListItem.Tag = file.Path
                            printListItem.Background = Media.Brushes.Transparent
                            Me.PrintListView.Items.Add(printListItem)
                        End If
                    Next
                End If
            End If

            PrintListPanelColumn.Width = GridLength.Auto
        End If

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


    Private Function FindSubfolderContaining(folderToIterate As String, searchStr As String, startPos As Short) As String
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim match As String = Nothing
        Dim folder = fso.GetFolder(folderToIterate)
        For Each subF In folder.subFolders
            If subF.Name.Length >= searchStr.Length Then
                If UCase(subF.name.Substring(startPos, searchStr.Length).Equals(searchStr)) Then
                    match = fso.GetAbsolutePathName(subF)
                    Exit For 'found it, stop looking
                End If
            End If
        Next
        fso = Nothing
        Return match
    End Function


    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        Me.ProgressBar.Visibility = Visibility.Visible
        Me.ProgressBar.Value = 0
        TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Normal

        Dim StrArray(2) As String
        Dim dirPathComponents = Split(Me.DistroPathText.Text, "\")
        If dirPathComponents.Length > 1 Then
            Me.CustomerComboBox.Text = UCase(dirPathComponents(1))
        End If
        Dim checkboxList As New List(Of Array)
        For Each cb In Me.ProgramWrapPanel.FindChildren(Of CheckBox)
            If cb.GetType() Is GetType(CheckBox) Then
                If cb.IsChecked Then
                    StrArray = {cb.Content.Text, cb.Tag}
                    checkboxList.Add(StrArray)
                End If
            End If
        Next
        MineLocationDataBGWorker.RunWorkerAsync(checkboxList)
    End Sub


    Private Sub BackgroundWorker_MineLocationData(sender As Object, e As DoWorkEventArgs)
        Dim checkboxList = e.Argument

        MineLocationDataBGWorker.ReportProgress(15)
        'check for info worksheet in XRL folder
        Dim infoFile = GetInfoFile()


        If infoFile <> "" Then
            MineLocationDataBGWorker.ReportProgress(25)
            locationInfo = New LocationData(infoFile)

        End If

        MineLocationDataBGWorker.ReportProgress(50)

        Dim i As Short = 0
        Dim length As Short = checkboxList.Count
        For Each progStr In checkboxList
            If DistributionPrograms.First Is Nothing Then
                DistributionPrograms.AddFirst(DetermineProgramType(progStr))
            Else
                DistributionPrograms.AddLast(DetermineProgramType(progStr))
            End If
            i += 1
            MineLocationDataBGWorker.ReportProgress(50 + (49 * (i / length)))
        Next
        MineLocationDataBGWorker.ReportProgress(100)
    End Sub


    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        ShowProgramSelectorPanel(False)
    End Sub


    Private Function DetermineProgramType(programStr As String()) As ProgramFile
        Dim filetype = System.IO.Path.GetExtension(programStr(0))
        Dim filename = System.IO.Path.GetFileNameWithoutExtension(programStr(0))
        Dim typeStr = filetype.ToUpper
        Dim filePath = programStr(1)
        Dim program
        'MsgBox(typeStr)
        Select Case typeStr
            Case ".CCF"
                If File.Exists(filePath & "\" & filename & ".H30") Then
                    program = New NonVitalHLC(filename, filePath, "ACE")
                ElseIf File.Exists(filePath & "\" & filename & ".H14") Then
                    program = New VitalHLC(filename, filePath, "ACE")
                Else
                    program = New ElectroLogIXS(filename, filePath)
                End If
            Case ".LOC"
                If File.Exists(filePath & "\" & filename & ".H30") Then
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
        If DistroPathText.Text = "Type" Then
            StartDir = "P:\"
        ElseIf Not Me.DistroPathText.Text = "" Then
            StartDir = Me.DistroPathText.Text
        Else
            StartDir = "P:\"
        End If
        Dim FilesDirectory = GetDistributionFolder(StartDir, "Browse for folder containing distribution files")
        StatusLabel.Text = tempInfoString
        If Not FilesDirectory = "" Then
            DistroPathText.Text = FilesDirectory
            FindFilesAndCreateProgramSelectPanel()
        End If
    End Sub


    Private Function GetInfoFile() As String
        Dim searchStr = "XRL"
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim f = fso.GetFolder(DistroPathText.Text)
        For Each subF In f.SubFolders
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


    Private Function GetNextReaderNum() As String
        Return "1"
    End Function


    Private Sub BackgroundWorker_CreateLetter(sender As Object, e As DoWorkEventArgs)
        Dim GuiData As MainWindowData = e.Argument
        'Creating Letter to place in box with deliverables
        Dim ReaderFilesDir = "C:\MT\"
        Dim readerFileName = GuiData.GetCustomer & GetNextReaderNum()
        Dim savePath = ReaderFilesDir & "\" & readerFileName & ".doc"
        createLetterBGWorker.ReportProgress(20)
        Dim objApp As Word.Application = New Word.Application
        objApp.Visible = False
        Dim distributionLetter As Word.Document = New Word.Document
        Dim filestr = IO.Path.GetFullPath("resources\BlankLetter.doc")
        createLetterBGWorker.ReportProgress(50)
        distributionLetter = objApp.Documents.Add(filestr)
        distributionLetter.Activate()
        Dim objSelection = objApp.Selection
        objApp.Selection.Font.Bold = True
        objApp.Selection.TypeText(vbCrLf & GuiData.GetDistributionDate & vbTab & "File: " & readerFileName & vbCrLf)
        objApp.Selection.Font.Bold = False

        objApp.Selection.TypeParagraph()
        objApp.Selection.TypeText(GuiData.GetRecipientName & vbCrLf &
                                  GuiData.GetCustomer & vbCrLf &
                                  GuiData.GetStreet & vbCrLf &
                                  GuiData.GetCity & ", " & GuiData.GetState & " " & GuiData.GetZipCode & vbCrLf)

        objApp.Selection.TypeParagraph()
        objApp.Selection.Font.Bold = True
        objApp.Selection.TypeText(UCase(Me.locationInfo.GetDivision & " DIVISION / " &
                                  Me.locationInfo.GetSubdivision & " SUBDIVISION / " & Me.locationInfo.GetMilePost & vbCrLf))
        objApp.Selection.Font.Bold = False

        objApp.Selection.TypeParagraph()
        objApp.Selection.TypeText("A package for " & locationInfo.GetLocationName & " has been sent to you via FedEx " &
                                  GuiData.GetShippingMethod & "." & vbCrLf & vbCrLf)
        objApp.Selection.TypeText("Tracking Number: " & GuiData.GetTrackingNumber & vbCrLf & vbCrLf)
        objApp.Selection.TypeText("The Package contains: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf &
                                  "Thank you," & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf &
                                  user.FullName & vbCrLf &
                                  "Xorail" & vbCrLf &
                                  "Email: " & user.Email & vbCrLf &
                                  "Phone: (904) 443-0083" & vbCrLf)
        createLetterBGWorker.ReportProgress(80)
        'Dim logoStr = IO.Path.GetFullPath("wabtecLogoForFiles.png")
        'distributionLetter.Sections(1).Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.InlineShapes.AddPicture(logoStr)
        'Dim myFooterStr As String = "5011 GATE PARKWAY, BLDG. 100 SUITE 400, JACKSONVILLE, FLORIDA 32256   " &
        '                             "PHONE:  904-443-0083    FAX: 904-443-0089"
        'With distributionLetter.Sections(1).Footers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range
        '    .Text = myFooterStr
        '    .Font.Name = "corbel"
        '    .Font.Size = 10
        '    .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        'End With
        distributionLetter.SaveAs(savePath)
        'Dispose the Word objects
        distributionLetter.Close()
        objApp.Quit()
        distributionLetter = Nothing
        objApp = Nothing
        createLetterBGWorker.ReportProgress(100)
    End Sub


    Private Sub CreateLetter()
        ProgressBar.Value = 0
        ProgressBar.Visibility = Visibility.Visible
        TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Normal

        Dim currentGuiData As New MainWindowData(Me.LocationNameText.Text, Me.CustomerComboBox.Text, Me.CustomerJobNumComboBox.Text,
                                                 Me.InternalJobNumComboBox.Text, Me.DistributionDatePicker.SelectedDate,
                                                 Me.TrackingNumberText.Text, Me.InvoiceNumText.Text, Me.ShippingMethodBox.Text,
                                                 Me.RecipientNameText.Text, Me.AddressStreetText.Text, Me.AddressCityText.Text,
                                                 Me.AddressStateBox.Text, Me.AddressZipCodeText.Text)
        createLetterBGWorker.RunWorkerAsync(currentGuiData)
    End Sub


    Private Sub BackgroundWorker_LetterCreationProgressChanged(sender As Object, e As ProgressChangedEventArgs)
        ProgressBar.Value = e.ProgressPercentage
    End Sub


    Private Sub BackgroundWorker_LetterCreationWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        ProgressBar.Visibility = Visibility.Hidden
        TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.None
    End Sub


    Private Sub EnableDataViewFunctions()
        DistributionDataLoaded = True
        'Tabs.Visibility = Visibility.Visible
        DistributionTab.Visibility = Visibility.Visible
        DistributionTabGrid.Visibility = Visibility.Visible

        If DistributionPrograms.Count > 1 Then
            LinkCompareTab.Visibility = Visibility.Visible
            LinksTabGrid.Visibility = Visibility.Visible
            RemoteLinkGrid.Visibility = Visibility.Visible
            AddChassisToLinksTab()
        Else
            LinkCompareTab.Visibility = Visibility.Collapsed
            LinksTabGrid.Visibility = Visibility.Hidden
            RemoteLinkGrid.Visibility = Visibility.Hidden
        End If

        PrintMenu.IsEnabled = True
        PrintToolBttn.IsEnabled = True
        'PrintLinkCompMenuItem.IsEnabled = True
        PrintLocInfoMenuItem.IsEnabled = True

        PrintPreviewTab.Visibility = Visibility.Visible
        ProgramRevisionsTab.Visibility = Visibility.Visible
        PrintingTab.Visibility = Visibility.Visible
        LocationDataFlowDoc = CreateDistInfoDocument()
        LocationInfoViewer.Document = LocationDataFlowDoc
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
        PrintLinkCompMenuItem.IsEnabled = False
        PrintLocInfoMenuItem.IsEnabled = False

        PrintPreviewTab.Visibility = Visibility.Collapsed
        ProgramRevisionsTab.Visibility = Visibility.Collapsed
        PrintingTab.Visibility = Visibility.Collapsed

        LinkComparePreviewSource.IsEnabled = False
        LinkCompareTab.Visibility = Visibility.Collapsed
        LinksTabGrid.Visibility = Visibility.Collapsed
        RemoteLinkGrid.Visibility = Visibility.Collapsed

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
        Dim distributionFldrPath As String = Nothing
        Dim labelsFldrPath As String = Nothing
        Dim searchStr As String

        searchStr = "Distribution"
        distributionFldrPath = FindSubfolderContaining(DistroPathText.Text, searchStr, 0)
        If distributionFldrPath Is Nothing Then
            distributionFldrPath = DistroPathText.Text & "\" & searchStr
            System.IO.Directory.CreateDirectory(distributionFldrPath)
        End If

        searchStr = "Labels"
        labelsFldrPath = FindSubfolderContaining(distributionFldrPath, searchStr, 0)
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

        CreateLetterToolBttn.IsEnabled = True

    End Sub


    Private Sub DisableCreationControls()
        FilterDatabaseByLocation.IsEnabled = False
        InsertToDBToolBttn.IsEnabled = False
        InsertToDBMenuItem.IsEnabled = False
        RefreshDBToolBttn.IsEnabled = False
        DatabaseMenu.IsEnabled = False
        DatabaseTab.Visibility = Visibility.Collapsed

        CreateEmailToolBttn.IsEnabled = False

        CreateLabelsToolBttn.IsEnabled = False

        CreateLetterToolBttn.IsEnabled = False

    End Sub


    Private Sub Exit_Application(sender As Object, e As EventArgs)
        Close()
    End Sub


    Private Sub MyHost_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles MyHost.PreviewKeyDown
        If e.Key = Key.Enter Then
            If Directory.Exists(Me.DistroPathText.Text) Then
                'Enter and path exists
                FindFilesAndCreateProgramSelectPanel()
                tempInfoString = ""
                StatusLabel.Text = tempInfoString
            Else
                tempInfoString = "Invalid path"
            End If
            StatusLabel.Text = tempInfoString
        End If
        Console.Write("pressed " & Int(e.Key) & vbCrLf)
    End Sub


    Private Sub DistroPathText_Click(sender As Object, e As EventArgs) Handles DistroPathText.Click
        If DistroPathText.Text = "Type or paste the directory path here or click the folder icon" Then
            DistroPathText.Text = ""
        End If
    End Sub


    Private Sub DistroPathText_MouseLeave(sender As Object, e As EventArgs) Handles DistroPathText.MouseLeave
        StatusLabel.Text = tempInfoString
        If DistroPathText.Text.Trim = "" Then
            DistroPathText.Text = "Type or paste the directory path here or click the folder icon"
        End If
    End Sub


    Private Sub CreateLetterToolBttn_Click(sender As Object, e As RoutedEventArgs) Handles CreateLetterToolBttn.Click
        CreateLetter()
    End Sub


    Private Sub ValidatePreviewGroup(sender As Object, e As RoutedEventArgs)
        If sender.IsChecked Then
            If sender.Equals(LocationDataPreviewSource) Then
                LinkComparePreviewSource.IsChecked = False
                LocationInfoViewer.Document = LocationDataFlowDoc
                LocationInfoViewer.ViewingMode = FlowDocumentReaderViewingMode.TwoPage
            Else
                LocationDataPreviewSource.IsChecked = False
                LocationInfoViewer.Document = LinkCompareFlowDoc
                LocationInfoViewer.ViewingMode = FlowDocumentReaderViewingMode.Scroll
            End If
        Else
            sender.IsChecked = True
        End If
    End Sub


    Private Sub ItemListPreviewMouseMove(sender As Object, e As MouseEventArgs)
        If (e.LeftButton = MouseButtonState.Pressed And Not e.OriginalSource.GetType.ToString.Equals("System.Windows.Controls.ListViewItem")) Then
            Dim data As New DataObject()
            data.SetData(DataFormats.StringFormat, e.OriginalSource.Text)

            DragDrop.DoDragDrop(sender, data, DragDropEffects.Copy Or DragDropEffects.Move)
        End If
    End Sub


    Private Sub DragAndDropStack_Drop(sender As Object, e As DragEventArgs)
        Dim myStackPanel As StackPanel = sender
        myStackPanel.Effect = Nothing
        If (e.Data.GetDataPresent(DataFormats.StringFormat)) Then
            Dim dataString As String = e.Data.GetData(DataFormats.StringFormat)
            Console.WriteLine(dataString & " dropped on " & myStackPanel.Name)
            If myStackPanel.Equals(MainChassisDropPanel) Then
                RefreshLinksList()
                Dim currentChassis = DistributionPrograms.First
                Do While currentChassis IsNot Nothing
                    If currentChassis.Value.getName.Equals(dataString) Then
                        Dim numOfRemoteChassis = currentChassis.Value.FindRemoteInformation(CombinedRptCheckBox.IsChecked)
                        AddDropPanelsForRemoteChassis(numOfRemoteChassis)
                        LinkCompareInstructions.Text = "Drag all remotes from the list onto thier respective remote number"
                        Exit Do ' found which item was dropped on main chassis, stop searching
                    End If
                    currentChassis = currentChassis.Next
                Loop
                MainHouseDropLabel.Content.Text = dataString
                MainHouseDropLabel.Foreground = System.Windows.Media.Brushes.White
                MainHouseDropLabel.FontWeight = FontWeights.Bold
                MainHouseDropLabel.FontSize = 14
            Else
                Dim currentChassis = DistributionPrograms.First
                Do While currentChassis IsNot Nothing
                    If currentChassis.Value.getName.Equals(dataString) Then
                        'myStackPanel.Tag.ToString.Substring(myStackPanel.Tag.ToString.Length - 1)
                        currentChassis.Value.FindRemoteInformation(CombinedRptCheckBox.IsChecked, myStackPanel.Tag.ToString.Substring(myStackPanel.Tag.ToString.Length - 1))
                        Exit Do ' found the label matching the drop panel, stop searching
                    End If
                    currentChassis = currentChassis.Next
                Loop
                For Each remoteLabel In RemoteLinkGrid.FindChildren(Of Label)
                    If myStackPanel.Tag = remoteLabel.Tag Then
                        remoteLabel.Content.Text = dataString
                        remoteLabel.Foreground = System.Windows.Media.Brushes.White
                        remoteLabel.FontWeight = FontWeights.Bold
                        remoteLabel.FontSize = 14
                        Exit For ' found the label matching the drop panel, stop searching
                    End If
                Next
            End If
            For Each item In ChassisListView.Items
                If item.Content.ToString.Equals(dataString.ToString) Then
                    ChassisListView.Items.Remove(item)
                    Exit For ' found the item to remove from chassis list view, stop searching
                End If
            Next
            If ChassisListView.Items.IsEmpty Then
                CompareLinksButton.IsEnabled = True
                LinkCompareInstructions.Text = "Press the compare button to view the link comparison"
            End If
            RefreshLinksButton.IsEnabled = True
        End If
        e.Handled = True
    End Sub


    Private Sub RefreshLinksList()
        ChassisListView.Items.Clear()
        RemoteLinkGrid.Children.RemoveRange(0, RemoteLinkGrid.Children.Count)
        AddChassisToLinksTab()
        MainHouseDropLabel.Content.Text = "Main Chassis"
        MainHouseDropLabel.Foreground = System.Windows.Media.Brushes.DarkOliveGreen
        MainHouseDropLabel.FontSize = 13
        MainHouseDropLabel.FontWeight = FontWeights.Normal
        CompareLinksButton.IsEnabled = False
        RefreshLinksButton.IsEnabled = False
        LinkCompareInstructions.Text = "Drag main chassis from the list onto the ""Main Chassis"""
    End Sub


    Private Sub DropPanel_DragOver(sender As Object, e As DragEventArgs)
        AddDropSpotBlurEffect(sender)
    End Sub


    Private Sub DropPanel_DragLeave(sender As Object, e As DragEventArgs)
        Dim myStackPanel As StackPanel = sender
        myStackPanel.Effect = Nothing
    End Sub


    Private Sub AddDropSpotBlurEffect(sender As Object)
        Dim myStackPanel As StackPanel = sender

        ' BLUR
        Dim myEffect As New Effects.BlurEffect
        myEffect.Radius = "2"
        myEffect.KernelType = Effects.KernelType.Box

        ' SHADOW
        'Dim myEffect As New Effects.DropShadowEffect
        'myEffect.BlurRadius = "3"
        'myEffect.Color = Colors.Black

        myStackPanel.Effect = myEffect
    End Sub


    Private Function GetTimeStamp()
        Dim monthStr, dayStr, militaryHour, militaryMin As String
        Dim militaryTime
        If Len(Month(Now)) = 1 Then
            monthStr = "0" & Month(Now)
        Else
            monthStr = Month(Now)
        End If
        If Len(Day(Now)) = 1 Then
            dayStr = "0" & Day(Now)
        Else
            dayStr = Day(Now)
        End If
        militaryTime = Split(FormatDateTime(Now, vbShortTime), ":")
        militaryHour = militaryTime(0)
        If Len(militaryTime(0)) = 1 Then
            militaryHour = "0" & militaryTime(0)
        Else
            militaryHour = militaryTime(0)
        End If
        If Len(militaryTime(1)) = 1 Then
            militaryMin = "0" & militaryTime(1)
        Else
            militaryMin = militaryTime(1)
        End If
        Return Year(Now) & monthStr & dayStr & militaryHour & militaryMin & " "
    End Function


    Private Sub CreateLinkCompareFile()
        LinkCompareInstructions.Text = "Use the ""Preview"" tab to view the link comparison"
        CompareLinksButton.IsEnabled = False
        Dim mainChassis As ElectroLogIXS = Nothing
        For Each chassis In DistributionPrograms
            If chassis.GetRemoteNum = 0 Then
                mainChassis = chassis
                Exit For 'found main chassis
            End If
        Next
        If mainChassis IsNot Nothing Then
            Dim newfolderpath As String = ""
            If Directory.Exists("C:\") Then
                newfolderpath = "C:\LinkCompare"
            Else
                newfolderpath = "D:\LinkCompare"
            End If
            If Not Directory.Exists(newfolderpath) Then
                Directory.CreateDirectory(newfolderpath)
            End If
            Dim myFilePath As String = newfolderpath & "\" & GetTimeStamp() & locationInfo.GetLocationName & " Remote Link Comparison.txt"
            Using outputFile As StreamWriter = New StreamWriter(myFilePath, True)
                For i = 0 To mainChassis.GetLinkUpStatus.Count - 1
                    Dim mainLinkInfoArray = Split(mainChassis.GetLinkSetup(i), "       ")
                    If UBound(mainLinkInfoArray) < 0 Then
                        Exit For
                    End If
                    outputFile.WriteLine("Main Link " & i + 1 & " Information:")
                    'Console.WriteLine("Main Link " & i + 1 & " Information:")
                    For j = 0 To UBound(mainLinkInfoArray)
                        outputFile.WriteLine(Trim(mainLinkInfoArray(j)))
                        'Console.WriteLine(Trim(mainLinkInfoArray(j)))
                    Next
                    outputFile.WriteLine("")
                    'Console.WriteLine("")

                    Dim remoteChassis As ElectroLogIXS = Nothing
                    For Each chassis In DistributionPrograms
                        If chassis.GetRemoteNum = i + 1 Then
                            remoteChassis = chassis
                            Exit For 'found remote chassis
                        End If
                    Next
                    Dim remoteLinkInfoArray = Split(remoteChassis.GetLinkSetup(0), "       ")
                    outputFile.WriteLine("Remote " & i + 1 & " Link Information:")
                    'Console.WriteLine("Remote " & i + 1 & " Link Information:")
                    For j = 0 To UBound(remoteLinkInfoArray)
                        outputFile.WriteLine(Trim(remoteLinkInfoArray(j)))
                        'Console.WriteLine(Trim(remoteLinkInfoArray(j)))
                    Next

                    outputFile.WriteLine(vbCrLf & "Linkup Status: (Main) " & mainChassis.GetLinkUpStatus(i) & "   (Remote) " & remoteChassis.GetLinkUpStatus(0) & vbCrLf)
                    'Console.WriteLine(vbCrLf & "Linkup Status: (Main) " & mainChassis.GetLinkUpStatus(i) & "   (Remote) " & remoteChassis.GetLinkUpStatus(0) & vbCrLf)
                    Dim NumInputBits = mainChassis.GetNumInputWords(i) * 8
                    outputFile.WriteLine(vbTab & "M Inputs" & vbTab & "R Outputs")
                    outputFile.WriteLine("____________________________")
                    'Console.WriteLine(vbTab & "M Inputs" & vbTab & "R Outputs")
                    'Console.WriteLine("____________________________")
                    For j = 1 To NumInputBits
                        If Len(mainChassis.GetInputs(i)(j)) > 7 Then
                            outputFile.WriteLine(j & ". " & vbTab & mainChassis.GetInputs(i)(j) & "  " & vbTab & remoteChassis.GetOutputs(0)(j))
                            'Console.WriteLine(j & ". " & vbTab & mainChassis.GetInputs(i)(j) & "  " & vbTab & remoteChassis.GetOutputs(0)(j))
                        Else
                            outputFile.WriteLine(j & ". " & vbTab & mainChassis.GetInputs(i)(j) & vbTab & vbTab & remoteChassis.GetOutputs(0)(j))
                            'Console.WriteLine(j & ". " & vbTab & mainChassis.GetInputs(i)(j) & vbTab & vbTab & remoteChassis.GetOutputs(0)(j))
                        End If
                    Next
                    Dim NumOutputBits = mainChassis.GetNumOutputWords(i) * 8
                    outputFile.WriteLine(vbCrLf & vbTab & "M Outputs" & vbTab & "R Inputs")
                    outputFile.WriteLine("____________________________")
                    'Console.WriteLine(vbCrLf & vbTab & "M Outputs" & vbTab & "R Inputs")
                    'Console.WriteLine("____________________________")
                    For j = 1 To NumOutputBits
                        If Len(mainChassis.GetOutputs(i)(j)) > 7 Then
                            outputFile.WriteLine(j & ". " & vbTab & mainChassis.GetOutputs(i)(j) & "  " & vbTab & remoteChassis.GetInputs(0)(j))
                            'Console.WriteLine(j & ". " & vbTab & mainChassis.GetOutputs(i)(j) & "  " & vbTab & remoteChassis.GetInputs(0)(j))
                        Else
                            outputFile.WriteLine(j & ". " & vbTab & mainChassis.GetOutputs(i)(j) & vbTab & vbTab & remoteChassis.GetInputs(0)(j))
                            'Console.WriteLine(j & ". " & vbTab & mainChassis.GetOutputs(i)(j) & vbTab & vbTab & remoteChassis.GetInputs(0)(j))
                        End If
                    Next
                    outputFile.WriteLine(vbCrLf & vbCrLf)
                    'Console.WriteLine(vbCrLf & vbCrLf)
                Next

                outputFile.Close()
                'Dim DocFont As New Font("Arial", 12)
                Dim ProgInfoStr = My.Computer.FileSystem.ReadAllText(myFilePath)
                Dim paragraph As New Paragraph
                paragraph.Inlines.Add(ProgInfoStr)
                LinkCompareFlowDoc = New FlowDocument(paragraph)
                LinkCompareFlowDoc.FontFamily = New System.Windows.Media.FontFamily("Consolas")
                LinkCompareFlowDoc.FontSize = 14
                LinkCompareFlowDoc.FontStyle = FontStyles.Normal
                LinkComparePreviewSource.IsEnabled = True
                LinkComparePreviewSource.IsChecked = True
                LocationDataPreviewSource.IsChecked = False
                LocationInfoViewer.Document = LinkCompareFlowDoc
                LocationInfoViewer.ViewingMode = FlowDocumentReaderViewingMode.Scroll
                PrintLinkCompMenuItem.IsEnabled = True
            End Using
        End If
    End Sub


    Private Sub PrintListView_PreviewMouseMove(sender As Object, e As MouseEventArgs) Handles PrintListView.PreviewMouseMove
        'Console.WriteLine(e.OriginalSource.GetType.ToString)
        If (e.LeftButton = MouseButtonState.Pressed And e.OriginalSource.GetType.ToString.Equals("System.Windows.Controls.ListView")) Then
            Dim data As New DataObject()
            data.SetData(e.OriginalSource)

            DragDrop.DoDragDrop(sender, data, DragDropEffects.Copy Or DragDropEffects.Move)
        End If
    End Sub


    Private Sub PrintListView_Drop(sender As Object, e As DragEventArgs) Handles PrintListView.Drop
        Dim myStackPanel As StackPanel = sender
        myStackPanel.Effect = Nothing
        If (e.Data.GetDataPresent("System.Windows.Controls.ListView")) Then
            Dim dataObj As ListView = e.Data.GetData("System.Windows.Controls.ListView")

            Dim myStrArrayList As New ArrayList
            Dim myArrayList As New ArrayList
            For Each item As ListViewItem In dataObj.SelectedItems
                If item.IsEnabled Then
                    myArrayList.Add(item)
                    myStrArrayList.Add(item.Tag)
                End If
            Next

            Dim myItemArray = myArrayList.ToArray()
            For Each item In myItemArray
                Console.WriteLine(item.Content.ToString & " dropped on " & myStackPanel.Name)
                item.IsEnabled = False 'disable the printed item but leave it visible in the listview
                item.IsSelected = False 'unselect the disabled item that was printed
                'dataObj.Items.Remove(item) 'remove the printed item from the list completely
            Next

            Dim myStrArray = myStrArrayList.ToArray()
            Dim printCanceled = PrintMySelectedFiles(myStrArray)

            If printCanceled Then ' print was canceled
                For Each item In myItemArray
                    item.IsEnabled = True
                    item.IsSelected = True
                Next
            ElseIf PrintListView.Items.IsEmpty Then
                Console.WriteLine("No more printable items!")
            End If

        End If
        e.Handled = True
    End Sub


    Private Function PrintMySelectedFiles(printFilesArray As Array) As Boolean
        Dim printer As New PrintDialog
        Dim result = printer.ShowDialog()
        If result Then
            ProgressBar.Visibility = Visibility.Visible
            TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Normal
            printFilesBGWorker.RunWorkerAsync(printFilesArray)
        End If

        Return Not result
    End Function


    Private Sub PrintPDF(printFileStr As String)
        Dim prtDoc As New Printing.PrintDocument
        Dim OldPrinter = prtDoc.PrinterSettings.PrinterName
        Dim WshNetwork = CreateObject("WScript.Network")
        WshNetwork.SetDefaultPrinter(RtvPrinter) 'set printer for RTVP to color printer (east 4th floor)

        '---------------------FIND APPLICATION PATH TO ADOBE ACROBAT OR READER-----------------------
        Dim AcrobatPath As String = ""
        Try
            AcrobatPath = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software").
                OpenSubKey("Microsoft").OpenSubKey("Windows").OpenSubKey("CurrentVersion").
                OpenSubKey("App Paths").OpenSubKey("Acrobat.exe").GetValue("")
        Catch e1 As Exception
            Try
                AcrobatPath = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software").
                OpenSubKey("Microsoft").OpenSubKey("Windows").OpenSubKey("CurrentVersion").
                OpenSubKey("App Paths").OpenSubKey("AcroRd32.exe").GetValue("")
            Catch e2 As Exception
                MsgBox("Neither Adobe Acrobat or Adobe Reader are insalled. Please install 
                    one or the other and try to print a pdf again." & vbCrLf &
                    "Path to Adobe Reader is: """ & AcrobatPath & """" & vbCrLf &
                    vbCrLf & e2.ToString)
            End Try
            MsgBox("Path to Adobe Acrobat is: """ & AcrobatPath & """" & vbCrLf &
                   vbCrLf & e1.ToString)
        End Try

        '---------------------Run Adobe Process to print file-----------------------
        Dim processArgs = "/t " & printFileStr & " " & RtvPrinter
        Process.Start(AcrobatPath, processArgs)
        Thread.Sleep(15000)

        Dim procsToKill As Process() = Process.GetProcessesByName("Acrobat")
        For Each proc In procsToKill
            proc.Kill()
            Exit For
        Next

        Console.WriteLine("Printing Printer: " & RtvPrinter & vbCrLf & "Default Printer: " & OldPrinter)
        WshNetwork.SetDefaultPrinter(OldPrinter) 'return to original printer
        Runtime.InteropServices.Marshal.ReleaseComObject(WshNetwork)
        WshNetwork = Nothing
    End Sub


    Private Sub PrintAndDelete(myDocsAray As String())
        Dim myXpsDoc As String = myDocsAray(0)
        Dim tempFile As String = myDocsAray(1)

        Dim defaultPrintQueue As PrintQueue = LocalPrintServer.GetDefaultPrintQueue
        Dim xpsPrintJob = defaultPrintQueue.AddJob(myXpsDoc, myXpsDoc, False)

        File.Delete(tempFile)
        File.Delete(myXpsDoc)
    End Sub


    Private Sub BackgroundWorker_FilePrintWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        ProgressBar.Visibility = Visibility.Hidden
        TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.None
    End Sub


    Private Sub BackgroundWorker_FilePrintingProgressChanged(sender As Object, e As ProgressChangedEventArgs)
        ProgressBar.Value = e.ProgressPercentage
        TaskbarItemInfo.ProgressValue = e.ProgressPercentage / 100
    End Sub


    Private Sub BackgroundWorker_PrintFiles(sender As Object, e As DoWorkEventArgs)
        Dim printFilesArray As Array = e.Argument

        Dim tempFolderProgrammaticallyCreated As Boolean
        Dim tempDirectory As String = DistroPathText.Text & "\Temp"
        If Not Directory.Exists(tempDirectory) Then
            Directory.CreateDirectory(tempDirectory)
            tempFolderProgrammaticallyCreated = True
        End If

        Dim i = 0
        For Each fileToPrint As String In printFilesArray
            Dim j = 0.0
            printFilesBGWorker.ReportProgress((i * 100 / printFilesArray.Length) + (j * 100 / printFilesArray.Length))

            Dim searchStr = "PDF"
            If fileToPrint.Substring(fileToPrint.Length - searchStr.Length).ToUpper.Equals(searchStr) Then
                Console.WriteLine("print PDF: " & fileToPrint)

                j = 0.3
                printFilesBGWorker.ReportProgress((i * 100 / printFilesArray.Length) + (j * 100 / printFilesArray.Length))

                Try
                    PrintPDF(fileToPrint)
                Catch ex As Exception
                    Console.WriteLine(ex)
                End Try

                j = 0.7
                printFilesBGWorker.ReportProgress((i * 100 / printFilesArray.Length) + (j * 100 / printFilesArray.Length))
            Else
                Console.WriteLine("print DOC: " & fileToPrint)

                Dim tempFileName = fileToPrint.Substring(fileToPrint.LastIndexOf("\") + 1)
                Dim tempFile As String = tempDirectory & "\" & tempFileName
                File.Copy(fileToPrint, tempFile)

                j = 0.2
                printFilesBGWorker.ReportProgress((i * 100 / printFilesArray.Length) + (j * 100 / printFilesArray.Length))

                Dim wordObj As New Word.Application
                wordObj.Visible = False

                j = 0.3
                printFilesBGWorker.ReportProgress((i * 100 / printFilesArray.Length) + (j * 100 / printFilesArray.Length))

                Dim doc As Word.Document = wordObj.Documents.Open(tempFile)

                j = 0.4
                printFilesBGWorker.ReportProgress((i * 100 / printFilesArray.Length) + (j * 100 / printFilesArray.Length))

                'set margins
                doc.PageSetup.TopMargin = wordObj.InchesToPoints(0.5)
                doc.PageSetup.BottomMargin = wordObj.InchesToPoints(0.5)
                doc.PageSetup.LeftMargin = wordObj.InchesToPoints(0.5)
                doc.PageSetup.RightMargin = wordObj.InchesToPoints(0.5)

                j = 0.6
                printFilesBGWorker.ReportProgress((i * 100 / printFilesArray.Length) + (j * 100 / printFilesArray.Length))

                Dim myXpsDoc = tempFile.Substring(0, tempFile.Length - 3) & "xps"
                doc.SaveAs(myXpsDoc, Word.WdSaveFormat.wdFormatXPS)
                doc.Close()
                wordObj.Application.Quit()
                wordObj = Nothing

                j = 0.8
                printFilesBGWorker.ReportProgress((i * 100 / printFilesArray.Length) + (j * 100 / printFilesArray.Length))

                Dim PrintsArray As String() = {myXpsDoc, tempFile}
                Dim thread = New Thread(AddressOf PrintAndDelete)
                thread.SetApartmentState(ApartmentState.STA)
                thread.Start(PrintsArray)
                thread.Join()

                j = 0.9
                printFilesBGWorker.ReportProgress((i * 100 / printFilesArray.Length) + (j * 100 / printFilesArray.Length))
            End If
            i += 1
        Next

        If tempFolderProgrammaticallyCreated Then
            Directory.Delete(tempDirectory)
        End If

        printFilesBGWorker.ReportProgress(100)
    End Sub
End Class
