Public Class VitalHLC
    Inherits HLC
    Private ValCRC As String

    Public Sub New(filename As String, path As String, compiler As String)
        MyBase.New(filename, path, "VHLC", compiler)
        'MsgBox("Creating V Obj")
        Me.ReadVitalLog()

        Me.evenLabelStr = Me.GetName & vbCrLf & "H14" & vbCrLf & "CS " & Me.GetEvenChecksum & "  CRC " & Me.GetEvenCRC
        Me.oddLabelStr = Me.GetName & vbCrLf & "H15" & vbCrLf & "CS " & Me.GetOddChecksum & "  CRC " & Me.GetOddCRC
    End Sub


    Public Overrides Function ToString() As String
        Return Me.GetName & " (" & Me.GetEquipType & ")" & vbCrLf & "H14: CRC = " & Me.GetEvenCRC &
            "; Checksum = " & Me.GetEvenChecksum & vbCrLf & "H15: CRC = " & Me.GetOddCRC &
            "; Checksum = " & Me.GetOddChecksum & vbCrLf & "Validation CRC = " & Me.ValCRC
    End Function


    Public Sub SetValidationCRC(valcrc As String)
        Me.ValCRC = valcrc
    End Sub


    Public Function GetValidationCRC() As String
        Return Me.ValCRC
    End Function


    Private Sub ReadVitalLog()
        Dim OddChipFlag As Boolean
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim logPathAndName = Me.GetPath & "\" & Me.GetName & ".LOG"
        If System.IO.File.Exists(logPathAndName) Then
            Dim logFile = fso.OpenTextFile(logPathAndName)
            Do Until logFile.AtEndOfStream
                Dim nextLine = logFile.ReadLine
                If InStr(nextLine, "IC14") <> 0 And InStr(nextLine, "EPT-1") <> 0 Then
                    Dim CrcSum1 = Mid(nextLine, InStr(nextLine, "EPT-1") + 10)
                    Me.SetEvenCRC(Left(CrcSum1, 4))
                    Me.SetEvenChecksum(Right(CrcSum1, 4))
                    OddChipFlag = False
                ElseIf InStr(nextLine, "IC15") <> 0 Then
                    Dim CrcSum2 = Mid(nextLine, InStr(nextLine, "EPT-1") + 10)
                    Me.SetOddCRC(Left(CrcSum2, 4))
                    Me.SetOddChecksum(Right(CrcSum2, 4))
                    OddChipFlag = True
                ElseIf InStr(nextLine, "Validation CRC:") <> 0 Then
                    Dim v = Split(Mid(nextLine, InStr(nextLine, "Validation CRC:") + 15), " ")
                    For Each x In v
                        If x = "V2-CRC:" Then
                            Exit For
                        ElseIf x = "" Or x = " " Or x = "  " Then
                        Else
                            If Len(x) >= 8 Then
                                Me.SetValidationCRC(Left(x, 8))
                            End If
                        End If
                    Next
                End If
                If Not OddChipFlag Then
                    Me.SetOddCRC("N/A ")
                    Me.SetOddChecksum("N/A ")
                End If
            Loop
        Else
            MsgBox("Vital log file does not exist.")
        End If
    End Sub


    Public Overrides Sub InsertDistributionToDB(con As SqlClient.SqlConnection, primaryKey As Integer, revNum As Integer)
        Dim cmd As New SqlClient.SqlCommand

        cmd.CommandText = "INSERT INTO Distributions(ID, locationName, programName, date, CRC_h14, checksum_h14, CRC_h15,
            checksum_h15, v_valcrc, revision, customer, customerJobNum, internalJobNum, equipmentType) VALUES(" &
            primaryKey & ", '" & My.Windows.MainWindow.LocationNameText.Text & "', '" & Me.GetName & "', 
            '" & My.Windows.MainWindow.DistributionDatePicker.DisplayDate & "', '" & Me.GetEvenCRC & "', 
            '" & Me.GetEvenChecksum & "', '" & Me.GetOddCRC & "', '" & Me.GetOddChecksum & "',
            '" & Me.GetValidationCRC & "', '" & revNum & "','" & My.Windows.MainWindow.CustomerComboBox.Text & "', 
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
