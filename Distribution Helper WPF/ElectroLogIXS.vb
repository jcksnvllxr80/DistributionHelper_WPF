Public Class ElectroLogIXS
    Inherits ProgramFile
    Private V_Checksum As String
    Private V_CRC As String
    Private NV_Checksum As String
    Private NV_CRC As String
    Private V_ValCrc As String
    Private NV_ValCrc As String
    Private ConsSum As String
    Private ConsCRC As String
    Private rptPathAndName As String
    Private combinedRPT As Boolean
    Private remoteNum As Short
    Private LinkUpStatus, Inputs, Outputs, linkSetup,
        InternalNetID, ExternalNetID, NumInputWords,
        NumOutputWords As New ArrayList


    Public Sub New(filename As String, path As String)
        MyBase.New(filename, path, "ElectroLogIXS")
        'MsgBox("Creating ElectroLogIXS Obj")
        Me.ReadConsFile()
        Me.GetValidationCRCs()
    End Sub

    Public Function GetLinkUpStatus() As ArrayList
        Return Me.LinkUpStatus
    End Function

    Public Function GetInputs() As ArrayList
        Return Me.Inputs
    End Function

    Public Function GetOutputs() As ArrayList
        Return Me.Outputs
    End Function

    Public Function GetLinkSetup() As ArrayList
        Return Me.linkSetup
    End Function

    Public Function GetInternalNetID() As ArrayList
        Return Me.InternalNetID
    End Function

    Public Function GetExternalNetID() As ArrayList
        Return Me.ExternalNetID
    End Function

    Public Function GetNumInputWords() As ArrayList
        Return Me.NumInputWords
    End Function

    Public Function GetNumOutputWords() As ArrayList
        Return Me.NumOutputWords
    End Function

    Public Function IsCombinedRPT() As String
        Return Me.combinedRPT
    End Function

    Public Function GetRemoteNum() As Short
        Return Me.remoteNum
    End Function

    Public Sub SetNV_Checksum(chksum As String)
        Me.NV_Checksum = chksum
    End Sub

    Public Function GetNV_Checksum() As String
        Return Me.NV_Checksum
    End Function

    Public Sub SetNV_CRC(crc As String)
        Me.NV_CRC = crc
    End Sub

    Public Function GetNV_CRC() As String
        Return Me.NV_CRC
    End Function

    Public Sub SetV_Checksum(chksum As String)
        Me.V_Checksum = chksum
    End Sub

    Public Function GetV_Checksum() As String
        Return Me.V_Checksum
    End Function

    Public Sub SetV_CRC(crc As String)
        Me.V_CRC = crc
    End Sub

    Public Function GetV_CRC() As String
        Return Me.V_CRC
    End Function

    Public Sub SetV_ValCRC(valCrc As String)
        Me.V_ValCrc = valCrc
    End Sub

    Public Function GetV_ValCRC() As String
        Return Me.V_ValCrc
    End Function

    Public Sub SetNV_ValCRC(valCrc As String)
        Me.NV_ValCrc = valCrc
    End Sub

    Public Function GetNV_ValCRC() As String
        Return Me.NV_ValCrc
    End Function

    Public Sub SetConsSum(chksum As String)
        Me.ConsSum = chksum
    End Sub

    Public Function GetConsSum() As String
        Return Me.ConsSum
    End Function

    Public Sub SetConsCRC(crc As String)
        Me.ConsCRC = crc
    End Sub

    Public Function GetConsCRC() As String
        Return Me.ConsCRC
    End Function


    Public Overrides Function ToString() As String
        Return Me.GetName & " (" & Me.GetEquipType & ")" & vbCrLf & "Consolidated: CRC = " &
            Me.ConsCRC & "; Checksum = " & Me.ConsSum & vbCrLf & "Vital: CRC = " & Me.V_CRC &
            "; Checksum = " & Me.V_Checksum & vbCrLf & "Non-Vital: CRC = " & Me.NV_CRC &
            "; Checksum = " & Me.NV_Checksum & vbCrLf & "Vital Validation CRC = " & Me.V_ValCrc &
            vbCrLf & "Non-Vital Validation CRC = " & Me.NV_ValCrc
    End Function


    Sub ReadConsFile()
        Dim g
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim consFile = Me.GetPath & "\" & Me.GetName & "_cons.txt"
        If System.IO.File.Exists(consFile) Then
            Dim f = fso.OpenTextFile(consFile)
            Do Until f.AtEndOfStream
                Dim nextLine = f.ReadLine
                If InStr(nextLine, "Checksum:") <> 0 Then
                    g = Split(Mid(nextLine, 10), " ")
                    For Each h In g
                        If Len(h) = 4 Then
                            Me.SetConsSum(h)
                            Exit For
                        End If
                    Next
                ElseIf InStr(nextLine, "EPT CRC:") <> 0 Then
                    g = Split(Mid(nextLine, 9), " ")
                    For Each h In g
                        If Len(h) = 4 Then
                            Me.SetConsCRC(h)
                            Exit For
                        End If
                    Next
                ElseIf InStr(nextLine, (Me.GetName & "v")) <> 0 Then
                    g = Split(Mid(nextLine, Len(Me.GetName) + 7), " ")
                    Dim sumFlag = 0
                    For Each h In g
                        If Len(h) = 4 And sumFlag = 0 Then
                            Me.SetV_Checksum(h)
                            sumFlag = 1
                        ElseIf Len(h) = 4 And sumFlag = 1 Then
                            Me.SetV_CRC(h)
                            Exit For
                        End If
                    Next
                ElseIf InStr(nextLine, (Me.GetName & "nv")) <> 0 Then
                    g = Split(Mid(nextLine, Len(Me.GetName) + 8), " ")
                    Dim sumFlag = 0
                    For Each h In g
                        If Len(h) = 4 And sumFlag = 0 Then
                            Me.SetNV_Checksum(h)
                            sumFlag = 1
                        ElseIf Len(h) = 4 And sumFlag = 1 Then
                            Me.SetNV_CRC(h)
                            Exit For
                        End If
                    Next
                End If
            Loop
            f.close
            f = Nothing
        Else
            MsgBox("Consolidated text file """ & consFile & """ does not exist.")
        End If
    End Sub


    Public Sub GetValidationCRCs()
        Dim vLog = Me.GetPath & "\" & Me.GetName & "v.log"
        If System.IO.File.Exists(vLog) Then
            Me.V_ValCrc = Me.FindValidationCRC(vLog)
        Else
            MsgBox("Vital log file """ & vLog & """ does not exist.")
        End If

        Dim nvLog = Me.GetPath & "\" & Me.GetName & "nv.log"
        If System.IO.File.Exists(nvLog) Then
            Me.NV_ValCrc = Me.FindValidationCRC(nvLog)
        Else
            MsgBox("Non-vital log file """ & nvLog & """ does not exist.")
        End If
    End Sub


    Public Overrides Sub InsertDistributionToDB(con As SqlClient.SqlConnection, primaryKey As Integer, revNum As Integer, mainWin As MainWindowData)
        Dim cmd As New SqlClient.SqlCommand

        cmd.CommandText = "INSERT INTO Distributions(ID, locationName, programName, date, consCRC, consSum, v_crc,
            v_sum, nv_crc, nv_sum, v_valcrc, nv_valcrc, revision, customer, customerJobNum, internalJobNum, equipmentType)
            VALUES(" & primaryKey & ", '" & mainWin.GetLocationName & "', '" & Me.GetName & "', 
            '" & mainWin.GetDistributionDate & "', '" & Me.GetConsCRC & "', '" & Me.GetConsSum & "', 
            '" & Me.GetV_CRC & "', '" & Me.GetV_Checksum & "', '" & Me.GetNV_CRC & "', '" & Me.GetNV_Checksum & "', 
            '" & Me.GetV_ValCRC & "', '" & Me.GetNV_ValCRC & "', '" & revNum & "', 
            '" & mainWin.GetCustomer & "', 
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


    Public Overloads Function FindRemoteInformation(combinedReport As Boolean) As Short
        ClearRemoteProperties()
        Dim fileEnding = ""
        If Not combinedReport Then
            fileEnding = "v"
        End If
        Me.rptPathAndName = Me.GetPath & "\" & Me.GetName & fileEnding & ".rpt"
        Me.combinedRPT = combinedReport
        Me.remoteNum = 0
        FindVitalLinkInfo()
        FindVitalLinkStatuses()
        Return LinkUpStatus.Count
    End Function


    Public Overloads Sub FindRemoteInformation(combinedReport As Boolean, remoteNumber As Short)
        ClearRemoteProperties()
        Dim fileEnding = ""
        If Not combinedReport Then
            fileEnding = "v"
        End If
        Me.rptPathAndName = Me.GetPath & "\" & Me.GetName & fileEnding & ".rpt"
        Me.combinedRPT = combinedReport
        Me.remoteNum = remoteNumber
        FindVitalLinkInfo()
        FindVitalLinkStatuses()
    End Sub


    Private Sub ClearRemoteProperties()
        remoteNum = Nothing
        LinkUpStatus.Clear()
        Inputs.Clear()
        Outputs.Clear()
        linkSetup.Clear()
        InternalNetID.Clear()
        ExternalNetID.Clear()
        NumInputWords.Clear()
        NumOutputWords.Clear()
    End Sub


    Private Sub FindVitalLinkInfo()
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim rptFile = fso.OpenTextFile(rptPathAndName)
        Do Until rptFile.AtEndOfStream
            Dim nextLine = rptFile.ReadLine
            If InStr(nextLine, "Port 3  Network ID: ") <> 0 Then
                Do
                    nextLine = rptFile.ReadLine
                Loop While InStr(nextLine, "Mode: ") = 0
                For i = 0 To 10
                    Dim linkSetupStr = ""
                    If InStr(nextLine, "Mode: ") <> 0 Then
                        linkSetupStr = linkSetupStr & "  " & nextLine
                    ElseIf InStr(nextLine, "Port ") <> 0 And InStr(nextLine, "Remote ") <> 0 Then
                        Exit For
                    ElseIf InStr(nextLine, "ACE Version:") <> 0 Then
                        If combinedRPT Then
                            Do
                                nextLine = rptFile.ReadLine
                            Loop While InStr(nextLine, "NV-EPT CHECKSUM ") = 0
                        Else
                            Do
                                nextLine = rptFile.ReadLine
                            Loop While InStr(nextLine, "V-EPT CHECKSUM ") = 0
                        End If
                    End If
                    nextLine = rptFile.ReadLine
                    Do
                        If InStr(nextLine, "ACE Version:") <> 0 Then
                            If combinedRPT Then
                                Do
                                    nextLine = rptFile.ReadLine
                                Loop While InStr(nextLine, "NV-EPT CHECKSUM ") = 0
                            Else
                                Do
                                    nextLine = rptFile.ReadLine
                                Loop While InStr(nextLine, "V-EPT CHECKSUM ") = 0
                            End If
                        ElseIf InStr(nextLine, "Network ID Min: ") <> 0 Then
                            linkSetupStr = linkSetupStr & nextLine
                            Dim NetIDs = Split(Trim(nextLine), "Local", 2)
                            ExternalNetID.Add(Trim(Mid(NetIDs(0), InStr(NetIDs(0), "Min: ") + 4, InStr(NetIDs(0), "Network ID Max: ") - InStr(NetIDs(0), "Min: ") - 4)))
                            InternalNetID.Add(Trim(Mid(NetIDs(1), InStr(NetIDs(1), "Min: ") + 4, InStr(NetIDs(1), "Host ") - InStr(NetIDs(1), "Min: ") - 4)))
                        ElseIf InStr(nextLine, "Number Input Bytes: ") <> 0 Then
                            linkSetupStr = linkSetupStr & nextLine
                            Dim NumOfLinkWords = Split(Trim(nextLine), " ")
                            NumInputWords.Add(CInt(NumOfLinkWords(3)))
                            NumOutputWords.Add(CInt(NumOfLinkWords(8)))
                        Else
                            linkSetupStr = linkSetupStr & nextLine
                        End If

                        nextLine = rptFile.ReadLine
                    Loop While InStr(nextLine, "Mode: ") = 0 And InStr(nextLine, "Vital Timer Summary") = 0 And InStr(nextLine, "Rate Table Report") = 0 And
                        Not (InStr(nextLine, "Port ") <> 0 And InStr(nextLine, "Remote ") <> 0)
                    linkSetup.Add(linkSetupStr)
                Next
            End If
        Loop
        rptFile.Close
    End Sub


    Private Sub FindVitalLinkStatuses()
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim rptFile = fso.OpenTextFile(rptPathAndName)
        Do Until rptFile.AtEndOfStream
            Dim nextLine = rptFile.ReadLine
            If InStr(nextLine, "Port 3") <> 0 And InStr(nextLine, "Remote ") <> 0 Then
                For i = 0 To 10
                    Dim InBitList(250) As String
                    Dim OutBitList(250) As String

                    If InStr(nextLine, "Port 3") <> 0 And InStr(nextLine, "Remote ") <> 0 Then

                    ElseIf InStr(nextLine, "Vital Timer Summary") <> 0 Or InStr(nextLine, "Rate Table Report") <> 0 Then
                        Exit For
                    End If

                    nextLine = rptFile.ReadLine
                    Do
                        If InStr(nextLine, "ACE Version:") <> 0 Then
                            If combinedRPT Then
                                Do
                                    nextLine = rptFile.ReadLine
                                Loop While InStr(nextLine, "NV-EPT CHECKSUM ") = 0
                            Else
                                Do
                                    nextLine = rptFile.ReadLine
                                Loop While InStr(nextLine, "V-EPT CHECKSUM ") = 0
                            End If
                        ElseIf InStr(nextLine, "Link Up ") <> 0 Then
                            Dim LinkUpLine = Split(Trim(nextLine), "                   ")
                            Dim LinkUp = Split(Trim(LinkUpLine(1)), "    ")
                            Me.LinkUpStatus.Add(Trim(LinkUp(0)))
                        ElseIf InStr(nextLine, "Input ") <> 0 Then
                            Dim InputLine = Split(Trim(nextLine), "                  ", 2)
                            Dim InputNumberStr = Split(Trim(InputLine(0)), " ")
                            Dim InputNum = CInt(Trim(InputNumberStr(1)))
                            Dim InputBitStr = Split(Trim(InputLine(1)), "     ")
                            InBitList(InputNum) = (Trim(InputBitStr(0)))
                        ElseIf InStr(nextLine, "Output ") <> 0 Then
                            Dim OutputLine = Split(Trim(nextLine), "                  ", 2)
                            Dim OutputNumberStr = Split(Trim(OutputLine(0)), " ")
                            Dim OutputNum = CInt(Trim(OutputNumberStr(1)))
                            Dim OutputBitStr = Split(Trim(OutputLine(1)), "     ")
                            OutBitList(OutputNum) = (Trim(OutputBitStr(0)))
                        End If
                        nextLine = rptFile.ReadLine
                    Loop While InStr(nextLine, "Vital Timer Summary") = 0 And InStr(nextLine, "Rate Table Report") = 0 And
                        Not (InStr(nextLine, "Port 3") <> 0 And InStr(nextLine, "Remote ") <> 0)
                    Me.Inputs.Add(InBitList)
                    Me.Outputs.Add(OutBitList)
                Next
            End If
        Loop
        rptFile.Close
    End Sub

End Class