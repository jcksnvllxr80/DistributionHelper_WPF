﻿Imports MahApps.Metro.Controls
Imports System.Drawing
Imports System.IO
Imports System.ComponentModel
Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Printing
Imports System.Drawing.Printing
Imports WPFFolderBrowser
Imports Microsoft.Office.Interop.Outlook
Imports OutlookApp = Microsoft.Office.Interop.Outlook.Application
'Imports Xceed.Wpf.Toolkit

Class MainWindow
    Inherits MetroWindow

    Private InsertToDatabseBGWorker As New BackgroundWorker()
    Private MineLocationDataBGWorker As New BackgroundWorker()
    Private createLetterBGWorker As New BackgroundWorker()
    Private printFilesBGWorker As New BackgroundWorker()
    Private tempInfoString = ""
    Private InitialDistroPathText As String = "Type or paste the directory path here or click the folder icon"
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
            If sender.Name.Equals("PrintLinkCompMenuItem") Then
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
        DistroPathFormHost.Child = DistroPathText

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

        AddHandler PrintListFilterComboBox.SelectionChanged, AddressOf PopulatePrintList
        AddHandler DocCreatorEquipmentTypeText.SelectionChanged, AddressOf DocCreatorEquipmentTypeText_SelectionChanged
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
        'Console.WriteLine(e.Source.Tag)
        StatusLabel.Text = sender.Tag
    End Sub

    Private Sub Text_MouseLeave(sender As Object, e As MouseEventArgs)
        StatusLabel.Text = tempInfoString
    End Sub


    Private Sub CreateDistributionEmail(sender As Object, e As RoutedEventArgs) Handles CreateEmailToolBttn.Click
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
        Dim CC_Recipients As String = "Holmes, Daryl (D.Holmes@xorail.com)"
        Dim subjectStr As String = locationInfo.GetCustomerNumber & "; " & locationInfo.GetLocationName & " – Program Books & Chips"
        CreateEmail(TO_Recipients, CC_Recipients, subjectStr, bodyStr)
    End Sub


    Private Function EmailHeader() As String
        Return vbCr & locationInfo.GetLocationName & vbCr & locationInfo.GetCity & ", " &
            locationInfo.GetState & " / MP. " & locationInfo.GetMilePost & vbCr &
            locationInfo.GetDivision & " Division / " & locationInfo.GetSubdivision & " Subdivision" &
            vbCr & locationInfo.GetCustomerNumber & vbCr & locationInfo.GetInternalNumber & vbCr & vbCr
    End Function


    Private Sub CreateTemplateEmail(sender As Object, e As RoutedEventArgs) Handles MailTemplateMenuItem.Click
        Dim bodyStr As String = EmailHeader()
        Dim TO_Recipients As String = ""
        Dim CC_Recipients As String = ""
        Dim subjectStr As String = locationInfo.GetCustomer & "#: " & locationInfo.GetCustomerNumber &
            "; XRL#: " & locationInfo.GetInternalNumber & "; " & locationInfo.GetLocationName
        CreateEmail(TO_Recipients, CC_Recipients, subjectStr, bodyStr)
    End Sub


    Private Sub CreateSoftwareVerificationEmail(sender As Object, e As RoutedEventArgs) Handles MailSoftwareVerificationMenuItem.Click
        Dim bodyStr As String = EmailHeader() & "An RTVP is required for " & locationInfo.GetLocationName &
            ". In order to begin an RTVP, we need field verification of the in-service software" &
            ". Will you please ask the field to provide this information?"
        Dim TO_Recipients As String = ""
        Dim CC_Recipients As String = "Holmes, Daryl (D.Holmes@xorail.com)"
        Dim subjectStr As String = locationInfo.GetCustomer & "#: " & locationInfo.GetCustomerNumber &
            "; XRL#: " & locationInfo.GetInternalNumber & "; " & "Field Verification of In-Service Software for " &
            locationInfo.GetLocationName
        CreateEmail(TO_Recipients, CC_Recipients, subjectStr, bodyStr)
    End Sub


    Private Sub CreateP_HoursEmail(sender As Object, e As RoutedEventArgs) Handles MailP_HoursMenuItem.Click
        Dim bodyStr As String = EmailHeader() & "I need _________ programming (P) hours for " & locationInfo.GetLocationName &
            ". The scope is _________ ."
        Dim TO_Recipients As String = "Holmes, Daryl (D.Holmes@xorail.com)"
        Dim CC_Recipients As String = ""
        Dim subjectStr As String = locationInfo.GetCustomer & "#: " & locationInfo.GetCustomerNumber &
            "; XRL#: " & locationInfo.GetInternalNumber & "; " & "Need Programming Hours for " & locationInfo.GetLocationName
        CreateEmail(TO_Recipients, CC_Recipients, subjectStr, bodyStr)
    End Sub


    Private Sub CreatePK_HoursEmail(sender As Object, e As RoutedEventArgs) Handles MailPK_HoursMenuItem.Click
        Dim bodyStr As String = EmailHeader() & "I need _________ programming check (PK) hours for " &
            locationInfo.GetLocationName & ". The scope is programming check (paper\simulation\final)."
        Dim TO_Recipients As String = "Holmes, Daryl (D.Holmes@xorail.com)"
        Dim CC_Recipients As String = ""
        Dim subjectStr As String = locationInfo.GetCustomer & "#: " & locationInfo.GetCustomerNumber &
            "; XRL#: " & locationInfo.GetInternalNumber & "; " & "Need Programming Check Hours for " & locationInfo.GetLocationName
        CreateEmail(TO_Recipients, CC_Recipients, subjectStr, bodyStr)
    End Sub


    Private Sub CreateRTVPDistributionEmail(sender As Object, e As RoutedEventArgs) Handles MailRTVP_DistMenuItem.Click
        Dim bodyStr As String = EmailHeader() & "The final software for " & locationInfo.GetLocationName &
            " has been uploaded to RailDOCS under project " & locationInfo.GetCustomerNumber &
            ". The RTVP document reflecting the changes made to the in-service software can be" &
            " found on the testing tab of this job on RailDOCS. Please let me know if there are any questions or concerns."
        Dim TO_Recipients As String = "RTV_ServiceTest@csx.com; Russell Turner <precisionsignal@gmail.com>"
        Dim CC_Recipients As String = "Holmes, Daryl (D.Holmes@xorail.com); Jubin, Tom <T.Jubin@xorail.com>; Grote, Lucas <l.grote@xorail.com>"
        Dim subjectStr As String = locationInfo.GetCustomerNumber & ", " & locationInfo.GetLocationName & " – Review RTV"
        CreateEmail(TO_Recipients, CC_Recipients, subjectStr, bodyStr)
    End Sub


    Private Sub CreatePrelimUploadEmail(sender As Object, e As RoutedEventArgs) Handles MailPrelimUploadMenuItem.Click
        Dim bodyStr As String = EmailHeader() & "The preliminary software for " & locationInfo.GetLocationName &
            " has been uploaded to RailDOCS under project " & locationInfo.GetCustomerNumber &
            ". Please find attached the compare reports reflecting the changes made to the software." &
            " Please let me know if there are any questions or concerns."
        Dim TO_Recipients As String = ""
        Dim CC_Recipients As String = "Holmes, Daryl (D.Holmes@xorail.com)"
        Dim subjectStr As String = locationInfo.GetCustomer & "#: " & locationInfo.GetCustomerNumber &
            "; XRL#: " & locationInfo.GetInternalNumber & "; " & locationInfo.GetLocationName &
            " Preliminary Software Uploaded to RailDOCS"
        CreateEmail(TO_Recipients, CC_Recipients, subjectStr, bodyStr)
    End Sub


    Private Sub CreateFinalUploadEmail(sender As Object, e As RoutedEventArgs) Handles MailFinalUploadMenuItem.Click
        Dim bodyStr As String = EmailHeader() & "The final software for " & locationInfo.GetLocationName &
            " has been uploaded to RailDOCS under project " & locationInfo.GetCustomerNumber &
            ". Please find attached the compare reports reflecting the changes made to the software." &
            "Please let me know if there are any questions Or concerns."
        Dim TO_Recipients As String = ""
        Dim CC_Recipients As String = "Holmes, Daryl (D.Holmes@xorail.com)"
        Dim subjectStr As String = locationInfo.GetCustomer & "#: " & locationInfo.GetCustomerNumber &
            "; XRL#: " & locationInfo.GetInternalNumber & "; " & locationInfo.GetLocationName &
            " Final Software Uploaded to RailDOCS"
        CreateEmail(TO_Recipients, CC_Recipients, subjectStr, bodyStr)
    End Sub


    Private Sub CreateEmail(TO_Recipients As String, CC_Recipients As String, subjectStr As String, bodyStr As String)
        Dim otlApp As New OutlookApp()
        Dim otlNewMail = otlApp.CreateItem(OlItemType.olMailItem)
        Dim WshShell = Type.GetTypeFromProgID("WScript.Shell")
        With otlNewMail
            .Display(subjectStr & " - Message (HTML)")
            .Subject = subjectStr
            .To = TO_Recipients
            .CC = CC_Recipients
        End With
        Dim objDoc = otlApp.ActiveInspector().WordEditor
        Dim objSel = objDoc.Windows(1).Selection
        objSel.InsertBefore(bodyStr)

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
            MsgBox("what to do when an attempt to add to the database Is made but the fields were Not properly filled at some point. this Is know because locationInfo Is = Nothing")
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


    Private Sub RestoreDefaults()
        DistributionTab.IsSelected = True
        DisableCreationControls()
        DisableDataViewFunctions()

        LocationNameText.Text = ""
        CustomerJobNumComboBox.Text = ""
        InternalJobNumComboBox.Text = ""
        CustomerComboBox.Text = ""
        TrackingNumberText.Text = ""
        InvoiceNumText.Text = ""
        ShippingMethodBox.SelectedItem = "Standard (3-5 Days)"
        DistributionDatePicker.SelectedDate = DateTime.Now
        RecipientNameText.Text = ""
        AddressStreetText.Text = ""
        AddressCityText.Text = ""
        AddressPhoneNumberText.Text = ""
        AddressStateBox.SelectedValue = "AL"
        AddressZipCodeText.Text = ""
        AddressPhoneNumberText.Text = ""

        HideDistributionTabFields()

        DistributionTab.Visibility = Visibility.Visible
        DistributionTabGrid.Visibility = Visibility.Visible
        DistributionPrograms.Clear()
        ProgramWrapPanel.Children.RemoveRange(0, ProgramWrapPanel.Children.Count)
        ChassisListView.Items.Clear()
        RemoteLinkGrid.Children.RemoveRange(0, RemoteLinkGrid.Children.Count)
        PrintListView.Items.Clear()
    End Sub


    Private Sub HideDistributionTabFields()
        LocationNameText.Visibility = Visibility.Hidden
        CustomerJobNumComboBox.Visibility = Visibility.Hidden
        InternalJobNumComboBox.Visibility = Visibility.Hidden
        CustomerComboBox.Visibility = Visibility.Hidden
        TrackingNumberText.Visibility = Visibility.Hidden
        InvoiceNumText.Visibility = Visibility.Hidden
        ShippingMethodBox.Visibility = Visibility.Hidden
        DistributionDatePicker.Visibility = Visibility.Hidden
        RecipientNameText.Visibility = Visibility.Hidden
        AddressStreetText.Visibility = Visibility.Hidden
        AddressCityText.Visibility = Visibility.Hidden
        AddressPhoneNumberText.Visibility = Visibility.Hidden
        AddressStateBox.Visibility = Visibility.Hidden
        AddressZipCodeText.Visibility = Visibility.Hidden
        AddressPhoneNumberText.Visibility = Visibility.Hidden

        LocationNameLabel.Visibility = Visibility.Hidden
        CustomerJobNumLabel.Visibility = Visibility.Hidden
        InternalJobNumLabel.Visibility = Visibility.Hidden
        CustomerLabel.Visibility = Visibility.Hidden
        TrackingNumberLabel.Visibility = Visibility.Hidden
        InvoiceNumberLabel.Visibility = Visibility.Hidden
        ShippingMethodLabel.Visibility = Visibility.Hidden
        DistributionDateLabel.Visibility = Visibility.Hidden
        DistributionRecipientLabel.Visibility = Visibility.Hidden
        StreetLabel.Visibility = Visibility.Hidden
        CityLabel.Visibility = Visibility.Hidden
        PhoneNumberLabel.Visibility = Visibility.Hidden
        StateLabel.Visibility = Visibility.Hidden
        ZipCodeLabel.Visibility = Visibility.Hidden
        PhoneNumberLabel.Visibility = Visibility.Hidden

        LocationInputsLabel.Visibility = Visibility.Hidden
        DistributionInputsLabel.Visibility = Visibility.Hidden
        AddressInputsLabel.Visibility = Visibility.Hidden

        DistributionAddressSeparator.Visibility = Visibility.Hidden
        LocationDistributionSeparator.Visibility = Visibility.Hidden
    End Sub


    Private Sub ShowDistributionTabFields()
        LocationNameText.Visibility = Visibility.Visible
        CustomerJobNumComboBox.Visibility = Visibility.Visible
        InternalJobNumComboBox.Visibility = Visibility.Visible
        CustomerComboBox.Visibility = Visibility.Visible
        TrackingNumberText.Visibility = Visibility.Visible
        InvoiceNumText.Visibility = Visibility.Visible
        ShippingMethodBox.Visibility = Visibility.Visible
        DistributionDatePicker.Visibility = Visibility.Visible
        RecipientNameText.Visibility = Visibility.Visible
        AddressStreetText.Visibility = Visibility.Visible
        AddressCityText.Visibility = Visibility.Visible
        AddressPhoneNumberText.Visibility = Visibility.Visible
        AddressStateBox.Visibility = Visibility.Visible
        AddressZipCodeText.Visibility = Visibility.Visible
        AddressPhoneNumberText.Visibility = Visibility.Visible

        LocationNameLabel.Visibility = Visibility.Visible
        CustomerJobNumLabel.Visibility = Visibility.Visible
        InternalJobNumLabel.Visibility = Visibility.Visible
        CustomerLabel.Visibility = Visibility.Visible
        TrackingNumberLabel.Visibility = Visibility.Visible
        InvoiceNumberLabel.Visibility = Visibility.Visible
        ShippingMethodLabel.Visibility = Visibility.Visible
        DistributionDateLabel.Visibility = Visibility.Visible
        DistributionRecipientLabel.Visibility = Visibility.Visible
        StreetLabel.Visibility = Visibility.Visible
        CityLabel.Visibility = Visibility.Visible
        PhoneNumberLabel.Visibility = Visibility.Visible
        StateLabel.Visibility = Visibility.Visible
        ZipCodeLabel.Visibility = Visibility.Visible
        PhoneNumberLabel.Visibility = Visibility.Visible

        LocationInputsLabel.Visibility = Visibility.Visible
        DistributionInputsLabel.Visibility = Visibility.Visible
        AddressInputsLabel.Visibility = Visibility.Visible

        DistributionAddressSeparator.Visibility = Visibility.Visible
        LocationDistributionSeparator.Visibility = Visibility.Visible
    End Sub


    Private Sub FindFilesAndCreateProgramSelectPanel()
        If Directory.Exists(Me.DistroPathText.Text) Then
            StatusLabel.Text = "Looking for software to distribute..."
            Dim j = 0

            RestoreDefaults()

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
                        textBlock.Tag = "Select " & textBlock.Text & " for Distribution"
                        AddHandler textBlock.MouseEnter, AddressOf Text_MouseEnter
                        AddHandler textBlock.MouseLeave, AddressOf Text_MouseLeave

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
                OkButton.Tag = "Accept and mine the checked software"
                AddHandler OkButton.MouseEnter, AddressOf Text_MouseEnter
                AddHandler OkButton.MouseLeave, AddressOf Text_MouseLeave
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
                CancelButton.Tag = "Cancel selecting/mining software"
                AddHandler CancelButton.MouseEnter, AddressOf Text_MouseEnter
                AddHandler CancelButton.MouseLeave, AddressOf Text_MouseLeave
                AddHandler CancelButton.Click, AddressOf CancelButton_Click

                ShowProgramSelectorPanel(True)
                OkButton.Focusable = True
                OkButton.Focus()
            Else
                MsgBox("There are no files to distribute in the selected folder.")
            End If
        Else
            Me.DistroPathText.Text = "Invalid path."
        End If
    End Sub


    Private Sub PopulatePrintList()
        PrintListView.Items.Clear()

        Dim RtvpBelongsInPrintList As Boolean = True
        Dim FileFilterStringArray As String() = {"ALL", "RPT", "LER", "ML2", "DOC", "DOCX"}
        Select Case PrintListFilterComboBox.SelectedValue.Name
            Case "ALL"
                FileFilterStringArray = {"ALL", "RPT", "LER", "ML2", "DOC", "DOCX"}
            Case "XRL"
                FileFilterStringArray = {"PDF", "DOC"}
                'look for {"PDF", "XLSX", "DOC", "DOCX", "XLSM", "URL"} in XRL folder
            Case "Software"
                FileFilterStringArray = {"ALL", "RPT", "LER", "ML2", "DOC", "DOCX"}
                RtvpBelongsInPrintList = False
            Case "ELGX"
                FileFilterStringArray = {"RPT", "LER"}
            Case "ML2"
                FileFilterStringArray = {"ML2", "GN2", "DOC", "DOCX"}
                RtvpBelongsInPrintList = False
            Case "VHLC"
                FileFilterStringArray = {"ALL"}
            Case "EC4"
                FileFilterStringArray = {"DOC", "DOCX"}
                RtvpBelongsInPrintList = False
            Case "Final"
                FileFilterStringArray = {"LOG", "VDR", "VAL"}
                RtvpBelongsInPrintList = False
            Case "RTVP"
                FileFilterStringArray = {}
        End Select

        If Not PrintListFilterComboBox.SelectedValue.Name.Equals("RTVP") Then
            Dim PrintFilesDirectory As String = Me.DistroPathText.Text
            If PrintListFilterComboBox.SelectedValue.Name.Equals("XRL") Then
                Dim ExpectedXRLPath As String = PrintFilesDirectory & "\XRL"
                For Each subf In Directory.GetDirectories(PrintFilesDirectory)
                    If subf.Substring(0, ExpectedXRLPath.Length).Equals(ExpectedXRLPath) Then
                        PrintFilesDirectory = subf
                        Exit For
                    End If
                Next
            End If

            For Each file In Directory.GetFiles(PrintFilesDirectory)
                Dim filesys = CreateObject("Scripting.FileSystemObject")
                Dim filetype = filesys.GetExtensionName(file)
                Dim filename = filesys.GetFileName(file)

                If InStr(filename, "~") = 0 And Not filetype Is Nothing Then 'dont use system files
                    If FileFilterStringArray.Contains(filetype.ToUpper) Then
                        If filetype.ToUpper.Equals("LER") And System.IO.File.Exists(file.Substring(0, file.Length - 3) & "LOC") Or filename.ToString.Contains("_print_list_") Then
                        Else
                            ' Create a ListViewItem
                            Dim printListItem As New ListViewItem()
                            ' Add ListViewItem to form
                            printListItem.Content = filename
                            printListItem.Tag = file
                            printListItem.Background = Media.Brushes.Transparent
                            Me.PrintListView.Items.Add(printListItem)
                        End If
                    End If
                End If
            Next
        End If

        If locationInfo IsNot Nothing And RtvpBelongsInPrintList Then
            AddRtvpToPrintList()
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

        ShowDistributionTabFields()
        PopulatePrintList()

        If locationInfo IsNot Nothing Then
            'set distribution tab fields
            Me.CustomerJobNumComboBox.Text = locationInfo.GetCustomerNumber()
            Me.InternalJobNumComboBox.Text = locationInfo.GetInternalNumber()
            Me.LocationNameText.Text = locationInfo.GetLocationName()
            Me.AddressStateBox.Text = locationInfo.GetState()
            Me.AddressCityText.Text = locationInfo.GetCity()
            locationInfo.SetCustomer(Me.CustomerComboBox.Text)

            'set Doc Creator tab fields
            Me.ProgrammerText.Text = locationInfo.ProgrammerName
            Me.ProgrammerInitialsText.Text = locationInfo.ProgrammerInitials
            Me.ProgramStartDatePicker.Text = locationInfo.StartDate
            Me.DocCreatorDivisionText.Text = locationInfo.GetDivision
            Me.DocCreatorSubdivisionText.Text = locationInfo.GetSubdivision
            Me.DocCreatorSubdivAbbreviationText.Text = locationInfo.GetSubdivAbrev
            Me.DocCreatorSignalRulesText.Text = locationInfo.SignalRules
            Me.DocCreatorDesignerInitialsText.Text = locationInfo.DesignerInitals
            Me.DocCreatorRailroadProjectManagerText.Text = locationInfo.ProjectManager
            Me.DocCreatorRailroadEngineerInitialsText.Text = locationInfo.RailroadEngineer
            For Each item As ComboBoxItem In Me.DocCreatorEquipmentTypeText.Items
                If item.Content.ToString.ToUpper.Equals(locationInfo.EquipmentType.ToUpper) Then
                    Me.DocCreatorEquipmentTypeText.SelectedItem = item
                    Exit For ' found item
                End If
            Next

            If locationInfo.IsRTVP Then
                IsRTVP.IsChecked = True
            Else
                IsRTVP.IsChecked = False
            End If

            AddRtvpToPrintList()
        Else
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


    Sub CheckDivisionAndSub()
        Dim DivBeginStr As String = InStr(1, Me.DocCreatorDivisionText.Text.ToUpper, "DIVISION")
        Dim SubdivBeginStr As String = InStr(1, Me.DocCreatorSubdivisionText.Text.ToUpper, "SUBDIVISION")
        If DivBeginStr > 0 Then
            Me.DocCreatorDivisionText.Text = Trim(Me.DocCreatorDivisionText.Text.Substring(0, DivBeginStr - 1))
        ElseIf SubdivBeginStr > 0 Then
            Me.DocCreatorSubdivisionText.Text = Trim(Me.DocCreatorSubdivisionText.Text.Substring(0, SubdivBeginStr - 1))
        End If
    End Sub


    Private Sub DocCreatorEquipmentTypeText_SelectionChanged()
        If DocCreatorEquipmentTypeText.SelectedItem.Content.Equals("ElectroLogIXS") Then
            If locationInfo IsNot Nothing Then
                'Console.WriteLine("Logicstation Version: " & locationInfo.LogicstationVersion)
                For Each item As ComboBoxItem In Me.DocCreatorAceVersionText.Items
                    If item.Tag.Equals(locationInfo.LogicstationVersion) Then
                        Me.DocCreatorAceVersionText.SelectedItem = item
                        Exit For ' found item
                    End If
                Next
            End If
            Me.DocCreatorAceVersionLabel.Visibility = Visibility.Visible
            Me.DocCreatorAceVersionText.Visibility = Visibility.Visible
        Else
            Me.DocCreatorAceVersionLabel.Visibility = Visibility.Hidden
            Me.DocCreatorAceVersionText.Visibility = Visibility.Hidden
        End If
    End Sub


    Sub AddRtvpToPrintList()
        Dim reducedTestDir As String = Nothing
        If Not locationInfo.GetRTVPfolderNum.Equals(vbNullChar) Then
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
    End Sub


    Private Function FindSubfolderContaining(folderToIterate As String, searchStr As String, startPos As Short) As String
        Dim match As String = Nothing
        If Directory.Exists(folderToIterate) Then
            Dim fso = CreateObject("Scripting.FileSystemObject")
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
        Else
            MsgBox("The folder """ & folderToIterate & """ which was being search for """ & searchStr & """ does not exist.")
        End If
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
        Dim folderBrowser As New WPFFolderBrowserDialog("Browse for folder containing distribution files")
        If DistroPathText.Text = InitialDistroPathText Then
            folderBrowser.InitialDirectory = "P:\"
        ElseIf Not Me.DistroPathText.Text = "" Then
            folderBrowser.InitialDirectory = Me.DistroPathText.Text
        Else
            folderBrowser.InitialDirectory = "P:\"
        End If
        StatusLabel.Text = tempInfoString

        Dim result = folderBrowser.ShowDialog()
        If result.Value Then
            DistroPathText.Text = folderBrowser.FileName
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
        ExportMenuItem.IsEnabled = True
        PrintLocInfoMenuItem.IsEnabled = True
        RailDOCSMenu.IsEnabled = True
        EmailMenu.IsEnabled = True

        PrintPreviewTab.Visibility = Visibility.Visible
        'ReducedTestTab.Visibility = Visibility.Visible
        'ProgramRevisionsTab.Visibility = Visibility.Visible
        DocCreatorTab.Visibility = Visibility.Visible
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
        ExportMenuItem.IsEnabled = False
        ExportRemoteComparisonMenuItem.IsEnabled = False
        PrintLocInfoMenuItem.IsEnabled = False
        RailDOCSMenu.IsEnabled = False
        EmailMenu.IsEnabled = False

        PrintPreviewTab.Visibility = Visibility.Collapsed
        ReducedTestTab.Visibility = Visibility.Collapsed
        ProgramRevisionsTab.Visibility = Visibility.Collapsed
        DocCreatorTab.Visibility = Visibility.Collapsed
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
        Dim labelsStr = Nothing
        Dim doc As XDocument = XDocument.Load("resources\Blank.label")
        Dim labelnode = doc.Descendants("String")
        For Each prog In DistributionPrograms
            If Not prog Is Nothing Then
                If prog.GetEquipType() = "EC4" Then
                    labelnode(0).Value = prog.MAPLabelStr
                    labelnode(1).Value = prog.MAPLabelStr
                    doc.Save(labelPath & "\" & prog.GetName & ".label")
                    labelsStr = labelsStr & vbTab & prog.GetEquipType() & ": " & prog.GetName & ".label" & vbCrLf
                ElseIf {"VHLC", "NVHLC"}.Contains(prog.GetEquipType()) Then
                    labelnode(0).Value = prog.evenLabelStr
                    labelnode(1).Value = prog.oddLabelStr
                    doc.Save(labelPath & "\" & prog.GetName & ".label")
                    labelsStr = labelsStr & vbTab & prog.GetEquipType() & ": " & prog.GetName & ".label" & vbCrLf
                End If
            End If
        Next
        If labelsStr Is Nothing Then
            MessageBox.Show("No labels were created." & vbCrLf & "There were no VHLC or EC4 programs selected.", ":-(")
        Else
            MessageBox.Show("Labels located in:" & vbCrLf & labelPath & "." & vbCrLf &
               vbCrLf & "The following labels were created:" & vbCrLf &
               vbCrLf & labelsStr, "Noice!")
        End If

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


    Private Sub MyHost_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles DistroPathFormHost.PreviewKeyDown
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
        If DistroPathText.Text = InitialDistroPathText Then
            DistroPathText.Text = ""
        End If
    End Sub


    Private Sub DistroPathText_MouseLeave(sender As Object, e As EventArgs) Handles DistroPathText.MouseLeave
        StatusLabel.Text = tempInfoString
        If DistroPathText.Text.Trim = "" Then
            DistroPathText.Text = InitialDistroPathText
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
            data.SetData(DataFormats.StringFormat, e.Source.Content)

            DragDrop.DoDragDrop(sender, data, DragDropEffects.Copy Or DragDropEffects.Move)
        End If
    End Sub


    Private Sub DragAndDropStack_Drop(sender As Object, e As DragEventArgs)
        Dim myStackPanel As StackPanel = sender
        myStackPanel.Effect = Nothing
        If (e.Data.GetDataPresent(DataFormats.StringFormat)) Then
            Dim dataString As String = e.Data.GetData(DataFormats.StringFormat)
            Console.WriteLine(dataString & " dropped on " & myStackPanel.Name)
            Dim removeChassis As Boolean
            If myStackPanel.Equals(MainChassisDropPanel) Then
                RefreshLinksList()
                Dim currentChassis = DistributionPrograms.First
                Do While currentChassis IsNot Nothing
                    If currentChassis.Value.getName.Equals(dataString) Then
                        Dim numOfRemoteChassis = currentChassis.Value.FindRemoteInformation(CombinedRptCheckBox.IsChecked)
                        If numOfRemoteChassis > 0 Then
                            AddDropPanelsForRemoteChassis(numOfRemoteChassis)
                            MainHouseDropLabel.Content.Text = dataString
                            MainHouseDropLabel.Foreground = System.Windows.Media.Brushes.White
                            MainHouseDropLabel.FontWeight = FontWeights.Bold
                            MainHouseDropLabel.FontSize = 14
                            RefreshLinksButton.IsEnabled = True
                            removeChassis = True
                        End If
                        Exit Do ' found which item was dropped on main chassis, stop searching
                    End If
                    currentChassis = currentChassis.Next
                Loop
            Else
                Dim currentChassis = DistributionPrograms.First
                Do While currentChassis IsNot Nothing
                    If currentChassis.Value.getName.Equals(dataString) Then
                        'myStackPanel.Tag.ToString.Substring(myStackPanel.Tag.ToString.Length - 1)
                        currentChassis.Value.FindRemoteInformation(CombinedRptCheckBox.IsChecked, myStackPanel.Tag.ToString.Substring(myStackPanel.Tag.ToString.Length - 1))
                        If currentChassis.Value.GetLinkUpStatus.Count > 0 Then
                            removeChassis = True
                        End If
                        Exit Do ' found the label matching the drop panel, stop searching
                    End If
                    currentChassis = currentChassis.Next
                Loop
                For Each remoteLabel In RemoteLinkGrid.FindChildren(Of Label)
                    If myStackPanel.Tag = remoteLabel.Tag Then
                        If removeChassis Then
                            remoteLabel.Content.Text = dataString
                            remoteLabel.Foreground = System.Windows.Media.Brushes.White
                            remoteLabel.FontWeight = FontWeights.Bold
                            remoteLabel.FontSize = 14
                        End If
                        Exit For ' found the label matching the drop panel, stop searching
                    End If
                Next
            End If
            For Each item In ChassisListView.Items
                If item.Content.ToString.Equals(dataString.ToString) Then
                    If removeChassis Then
                        ChassisListView.Items.Remove(item)
                        LinkCompareInstructions.Text = "Drag all remotes from the list onto thier respective remote number"
                    Else
                        LinkCompareInstructions.Text = "Cannot find "".rpt"" with remote info for the chosen chassis"
                    End If
                    Exit For ' found the item to remove from chassis list view, stop searching
                End If
            Next
            If ChassisListView.Items.IsEmpty Then
                CompareLinksButton.IsEnabled = True
                LinkCompareInstructions.Text = "Press the compare button to view the link comparison"
            End If
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

            Dim myFilePath As String = ""
            If locationInfo IsNot Nothing Then
                myFilePath = newfolderpath & "\" & GetTimeStamp() & locationInfo.GetLocationName & " Remote Link Comparison.txt"
            Else
                myFilePath = newfolderpath & "\" & GetTimeStamp() & "myLocationName Remote Link Comparison.txt"
            End If

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
                ExportRemoteComparisonMenuItem.IsEnabled = True
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

            Dim myStrArray As Array = myStrArrayList.ToArray()
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
            Dim printJobInfo As New PrintJobInfo(printFilesArray, TwoSidedPrintsCheckbox.IsChecked)
            ProgressBar.Visibility = Visibility.Visible
            TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Normal
            printFilesBGWorker.RunWorkerAsync(printJobInfo)
        End If

        Return Not result
    End Function


    Private Sub PrintPDF(printFileStr As String)
        RawPrinterHelper.SendFileToPrinter(RtvPrinter, printFileStr)
    End Sub


    Private Sub PrintAndDelete(myDocsAray As String())
        Dim myXpsDoc As String = myDocsAray(0)
        Dim tempFile As String = myDocsAray(1)
        Dim twoSidedPrints As Boolean = myDocsAray(2)

        Dim thisPrintQueue As PrintQueue = LocalPrintServer.GetDefaultPrintQueue
        If twoSidedPrints Then
            thisPrintQueue.CurrentJobSettings.CurrentPrintTicket.Duplexing = Duplexing.TwoSidedLongEdge
        Else
            thisPrintQueue.CurrentJobSettings.CurrentPrintTicket.Duplexing = Duplexing.OneSided
        End If
        thisPrintQueue.CurrentJobSettings.CurrentPrintTicket.Duplexing = Duplexing.TwoSidedLongEdge
        Dim xpsPrintJob = thisPrintQueue.AddJob(myXpsDoc, myXpsDoc, False)
        Console.WriteLine(xpsPrintJob.PropertiesCollection.ToString)

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
        Dim printJobInfo As PrintJobInfo = e.Argument
        Dim printFilesArray As Array = printJobInfo.FilesToPrint
        Dim twoSided As Boolean = printJobInfo.TwoSidedPrinting

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
                Catch ex As System.Exception
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

                Dim PrintsArray As String() = {myXpsDoc, tempFile, twoSided}
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


    Private Sub ExportToFileMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dim fileContent As New FlowDocument

        Dim mySaveAsDialog As New Microsoft.Win32.SaveFileDialog()
        mySaveAsDialog.InitialDirectory = Me.DistroPathText.Text
        mySaveAsDialog.DefaultExt = ".txt" ' Default file extension
        mySaveAsDialog.Filter = "Text documents (.txt)|*.txt" ' Filter files by extension
        If e.Source.Name.Equals("ExportLocationInfoMenuItem") Then
            If Me.LocationNameText.Text.Equals("") Then
                mySaveAsDialog.FileName = GetTimeStamp() & "myLocationName Location Information File" ' Default file name
            Else
                mySaveAsDialog.FileName = GetTimeStamp() & Me.LocationNameText.Text & " Location Information File" ' Default file name
            End If
            fileContent = LocationDataFlowDoc
        Else
            If Me.LocationNameText.Text.Equals("") Then
                mySaveAsDialog.FileName = GetTimeStamp() & "myLocationName Remote Link Comparison" ' Default file name
            Else
                mySaveAsDialog.FileName = GetTimeStamp() & Me.LocationNameText.Text & " Remote Link Comparison" ' Default file name
            End If
            fileContent = LinkCompareFlowDoc
        End If
        If mySaveAsDialog.ShowDialog Then
            Console.WriteLine(mySaveAsDialog.FileName)
            File.WriteAllText(mySaveAsDialog.FileName, New TextRange(fileContent.ContentStart, fileContent.ContentEnd).Text)
            'Dim filename As String = mySaveAsDialog.FileName
        End If
    End Sub


    Private Sub AboutMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles AboutMenuItem.Click
        'create about window
    End Sub


    Private Sub HelpMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles HelpMenuItem.Click
        'Forms.Help.ShowHelp(Nothing, "https://intranet.xorail.com") 'need to create a help website for the distribution helper
    End Sub


    Private Sub RaildocsPSRMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles RaildocsPSRMenuItem.Click
        OpenBrowserWith("https://www.raildocs.net/CSXTrans/DgnProjTracking/ProjectSummaryReport/index.cgi/view?CSXProjNum=" &
                        locationInfo.GetCustomerNumber)
    End Sub


    Private Sub RaildocsMainJobMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles RaildocsMainJobMenuItem.Click
        OpenBrowserWith("https://www.raildocs.net/CSXTrans/DgnProjTracking/EditProj.cgi?CSXProjNum=" &
                        locationInfo.GetCustomerNumber & "#q=1")
    End Sub


    Private Sub OpenBrowserWith(URL As String)
        Try
            Process.Start(URL)
        Catch ex As System.Exception
            Console.WriteLine(ex)
        End Try
    End Sub


    Private Sub OracleMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles OracleMenuItem.Click
        OpenBrowserWith("http://wcssrv0340.wabtec.com:8009/OA_HTML/RF.jsp?function_id=1028108&resp_id=-1&resp_appl_id=-1" &
                        "&security_group_id=0&lang_code=US&params=n9i84W2HD0fd7VSPSKkmMDdHNtlcBpJgG2aUUG8CvAOeD3BeN5k3UF" &
                        "lWhVh0vFlHYZxiE06CmUbK7s7WQHPPcvtWclFN.R4RJFBQuDgpeqqtb5GRqFmrymW06tTmMosJvP9b5jCJyhYsDu7fcBedQ" &
                        "uZkTTDyKpfxkJgeo2VSpr8ZikbWmBAET2x3oZVWMmmobDiNEOSBXe0Aq7YHL.ssY2Z2Mfs3Xt3t0fN0I1tqHFJw0PFWiWDt" &
                        "Z4F6hEryhQEPn1MBbKRUwtmeLHUL4XmCUP98YnBrASo-6dSbXCRAgkQaqUUB6Ug6iABfQEq0wxF35A49HRDJbCLsEVX5U-R" &
                        "JUXoYPJYAEojt5ucSQGVEPyLATSOl-UMOaLXFedNul8mwKP9mWiWp51Oxg.Bnwh3V0AyUb.eWof7OUGa54-eL5DjaKuYHyZ" &
                        "SbhHMkfInoApVlpNZhaTxr3qMKv-u0y1VSlQ&oas=ROzQIW2KQ8s_Jm-qmAAa2w..#")
    End Sub


    Private Sub SharepointMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles SharepointMenuItem.Click
        OpenBrowserWith("https://wabtec.sharepoint.com/sites/xorail/accounts/CSX/VLC/Lists/Project%20Tracking/AllItems.aspx")
    End Sub

    Private Sub CreateDocsButton_Click(sender As Object, e As RoutedEventArgs) Handles CreateDocsButton.Click
        'stub
    End Sub
End Class




