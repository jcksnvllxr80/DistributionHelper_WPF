Public Class ProgramFile
    Private Name As String
    Private EquipType As String
    Private FilePath As String

    Public Sub New(filename As String, path As String, equip As String)
        'MsgBox("Creating ProgramFile Obj")
        Me.Name = filename
        Me.EquipType = equip
        Me.FilePath = path
    End Sub


    Public Overrides Function ToString() As String
        Return Me.GetName & " (" & Me.GetEquipType & ")"
    End Function


    Public Function GetPath() As String
        Return Me.FilePath
    End Function


    Public Function GetName() As String
        Return Me.Name
    End Function


    Public Function GetEquipType() As String
        Return Me.EquipType
    End Function


    Public Function FindValidationCRC(logfileName) As String
        Dim valCrc As String = ""
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim logFile = fso.OpenTextFile(logfileName)
        Do Until logFile.AtEndOfStream
            Dim nextLine = logFile.ReadLine
            If InStr(nextLine, "Validation CRC:") <> 0 Then
                Dim v = Split(Mid(nextLine, InStr(nextLine, "Validation CRC:") + 15), " ")
                For Each x In v
                    If x = "V2-CRC:" Then
                        Exit For
                    ElseIf x = "" Or x = " " Or x = "  " Then
                    Else
                        If Len(x) >= 8 Then
                            valCrc = Left(x, 8)
                        End If
                    End If
                Next
            End If
        Loop
        Return valCrc
    End Function


    Public Overridable Sub InsertDistributionToDB(con As SqlClient.SqlConnection, primaryKey As Integer, revNum As Integer)

    End Sub
End Class
