Public Class HLC
    Inherits ProgramFile
    Private crcEven As String
    Private sumEven As String
    Private crcOdd As String
    Private sumOdd As String
    Public isVital As Boolean
    Public isCompiledInALC As Boolean
    Private validationCRC As String
    Public evenLabelStr As String
    Public oddLabelStr As String

    Public Sub New(filename As String, path As String, equip As String, compiler As String)
        MyBase.New(filename, path, equip)
        'MsgBox("Creating HLC Obj")
        If equip = "VHLC" Then
            Me.isVital = True
        Else
            Me.isVital = False
        End If

        If compiler = "ALC" Then
            Me.isCompiledInALC = True
        Else
            Me.isCompiledInALC = False
        End If

    End Sub


    Public Overrides Function ToString() As String
        Return Me.GetName & " (" & Me.GetEquipType & ")" & vbCrLf & "H30: CRC = " & Me.crcEven &
            "; Checksum = " & Me.sumEven & vbCrLf & "H31: CRC = " & Me.crcOdd &
            "; Checksum = " & Me.sumOdd
    End Function


    Public Sub SetEvenChecksum(chksum As String)
        Me.sumEven = chksum
    End Sub

    Public Function GetEvenChecksum() As String
        Return Me.sumEven
    End Function

    Public Sub SetEvenCRC(crc As String)
        Me.crcEven = crc
    End Sub

    Public Function GetEvenCRC() As String
        Return Me.crcEven
    End Function

    Public Sub SetOddChecksum(chksum As String)
        Me.sumOdd = chksum
    End Sub

    Public Function GetOddChecksum() As String
        Return Me.sumOdd
    End Function

    Public Sub SetOddCRC(crc As String)
        Me.crcOdd = crc
    End Sub

    Public Function GetOddCRC() As String
        Return Me.crcOdd
    End Function

End Class
