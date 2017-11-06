Public Class NonVitalHLC
    Inherits HLC

    Public Sub New(filename As String, path As String, compiler As String)
        MyBase.New(filename, path, "NVHLC", compiler)
        'MsgBox("Creating NV Obj")
        If Me.isCompiledInALC Then
            Me.ReadNonVitalLog()
        Else
            Me.ReadNonVitalLog()
        End If

        Me.evenLabelStr = Me.GetName & vbCrLf & "H30" & vbCrLf & "CS " & Me.GetEvenChecksum & "  CRC " & Me.GetEvenCRC
        Me.oddLabelStr = Me.GetName & vbCrLf & "H31" & vbCrLf & "CS " & Me.GetOddChecksum & "  CRC " & Me.GetOddCRC
    End Sub


    Private Sub ReadNonVitalLog()
        Dim chipNameInLog(2) As String
        If Me.isCompiledInALC Then
            chipNameInLog(0) = "IC14"
            chipNameInLog(1) = "IC15"
        Else
            chipNameInLog(0) = "IC30"
            chipNameInLog(1) = "IC31"
        End If
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim logPathAndName = Me.GetPath & "\" & Me.GetName & ".LOG"
        If System.IO.File.Exists(logPathAndName) Then
            Dim logFile = fso.OpenTextFile(logPathAndName)
            Do Until logFile.AtEndOfStream
                Dim nextLine = logFile.ReadLine
                If InStr(nextLine, chipNameInLog(0)) <> 0 Then
                    Dim CrcSum1 = Mid(nextLine, InStr(nextLine, "EPT-1") + 10)
                    Me.SetEvenCRC(Left(CrcSum1, 4))
                    Me.SetEvenChecksum(Right(CrcSum1, 4))
                ElseIf InStr(nextLine, chipNameInLog(1)) <> 0 Then
                    Dim CrcSum2 = Mid(nextLine, InStr(nextLine, "EPT-1") + 10)
                    Me.SetOddCRC(Left(CrcSum2, 4))
                    Me.SetOddChecksum(Right(CrcSum2, 4))
                End If
            Loop
            logFile.Close
        Else
            MsgBox("Non-vital log file does not exist.")
        End If
    End Sub


    Public Overrides Sub InsertDistributionToDB(con As SqlClient.SqlConnection, primaryKey As Integer, revNum As Integer, mainWin As MainWindowData)
        Dim cmd As New SqlClient.SqlCommand

        cmd.CommandText = "INSERT INTO Distributions(ID, locationName, programName, date, CRC_h30, checksum_h30, CRC_h31,
            checksum_h31, revision, customer, customerJobNum, internalJobNum, equipmentType) VALUES(" &
            primaryKey & ", '" & mainWin.GetLocationName & "', '" & Me.GetName & "', 
            '" & mainWin.GetDistributionDate & "', '" & Me.GetEvenCRC & "', 
            '" & Me.GetEvenChecksum & "', '" & Me.GetOddCRC & "', '" & Me.GetOddChecksum & "',
            '" & revNum & "','" & mainWin.GetCustomer & "', 
            '" & mainWin.GetCustomerNumber & "', 
            '" & mainWin.GetInternalNumber & "', '" & Me.GetEquipType & "')"
        cmd.CommandType = CommandType.Text
        cmd.Connection = con
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error:" & vbCrLf & ex.Message)
        End Try
    End Sub

End Class
