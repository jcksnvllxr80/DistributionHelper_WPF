Public Class EC4
    Inherits ProgramFile
    Private MAP_crc As String
    Private MAP_sum As String
    Private MNE_crc As String
    Private MNE_sum As String
    Public MapLabelStr As String

    Public Sub New(filename As String, path As String)
        MyBase.New(filename, path, "EC4")
        ReadChecksumLog()
        Me.MapLabelStr = Me.GetName & vbCrLf & "MAP" & vbCrLf & "CS " & Me.GetMAPChecksum & "  CRC " & Me.GetMAPCRC
    End Sub


    Public Overrides Function ToString() As String
        Return Me.GetName & " (" & Me.GetEquipType & ")" & vbCrLf & "MAP: CRC = " &
            Me.MAP_crc & "; Checksum = " & Me.MAP_sum & vbCrLf & "MNE: CRC = " &
            Me.MNE_crc & "; Checksum = " & Me.MNE_sum
    End Function


    Public Sub SetMAPChecksum(chksum As String)
        Me.MAP_sum = chksum
    End Sub


    Public Function GetMAPChecksum() As String
        Return Me.MAP_sum
    End Function


    Public Sub SetMAPCRC(crc As String)
        Me.MAP_crc = crc
    End Sub


    Public Function GetMAPCRC() As String
        Return Me.MAP_crc
    End Function


    Public Sub SetMNEChecksum(chksum As String)
        Me.MNE_sum = chksum
    End Sub

    Public Function GetMNEChecksum() As String
        Return Me.MNE_sum
    End Function

    Public Sub SetMNECRC(crc As String)
        Me.MNE_crc = crc
    End Sub

    Public Function GetMNECRC() As String
        Return Me.MNE_crc
    End Function


    Private Sub ReadChecksumLog()
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim logPathAndName = Me.GetPath & "\CHECKSUM.LOG"
        If System.IO.File.Exists(logPathAndName) Then
            Dim logFile = fso.OpenTextFile(logPathAndName)
            Do Until logFile.AtEndOfStream
                Dim nextLine = logFile.ReadLine
                If InStr(nextLine, ".MAP ") <> 0 Then
                    Dim CrcSum1 = Split(Trim(Mid(nextLine, InStr(nextLine, ".MAP ") + 4)), "  ")
                    Me.SetMAPChecksum(Trim(Right(CrcSum1(0), 4)))
                    Me.SetMAPCRC(Trim(Right(CrcSum1(1), 4)))
                ElseIf InStr(nextLine, ".MNE ") <> 0 Then
                    Dim CrcSum2 = Split(Trim(Mid(nextLine, InStr(nextLine, ".MNE ") + 4)), "  ")
                    Me.SetMNEChecksum(Trim(Right(CrcSum2(0), 4)))
                    Me.SetMNECRC(Trim(Right(CrcSum2(1), 4)))
                End If
            Loop
            logFile.Close
        Else
            MsgBox("checksum.log file does not exist")
        End If
    End Sub


    Public Overrides Sub InsertDistributionToDB(con As SqlClient.SqlConnection, primaryKey As Integer, revNum As Integer)
        Dim cmd As New SqlClient.SqlCommand

        cmd.CommandText = "INSERT INTO Distributions(ID, locationName, programName, date, CRC_h14, checksum_h14, CRC_h15,
            checksum_h15, revision, customer, customerJobNum, internalJobNum, equipmentType) VALUES(" &
            primaryKey & ", '" & My.Windows.MainWindow.LocationNameText.Text & "', '" & Me.GetName & "', 
            '" & My.Windows.MainWindow.DistributionDatePicker.DisplayDate & "', '" & Me.GetMAPCRC & "', 
            '" & Me.GetMAPChecksum & "', '" & Me.GetMNECRC & "', '" & Me.GetMNEChecksum & "',
            '" & revNum & "','" & My.Windows.MainWindow.CustomerComboBox.Text & "', 
            '" & My.Windows.MainWindow.CustomerJobNumComboBox.Text & "', 
            '" & My.Windows.MainWindow.InternalJobNumComboBox.Text & "', '" & Me.GetEquipType & "')"
        cmd.CommandType = CommandType.Text
        cmd.Connection = con
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error:" & vbCrLf & ex.Message)
        End Try

    End Sub

End Class
