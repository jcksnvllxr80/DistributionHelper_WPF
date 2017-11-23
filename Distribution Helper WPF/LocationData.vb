Public Class LocationData
    Private Customer As String
    Private LocationName As String
    Private MilePost As String
    Private InternalNumber As String
    Private CustomerNumber As String
    Private Division As String
    Private Subdivision As String
    Private SubdivAbrev As String
    Private State As String
    Private City
    Private RTVP As Boolean
    Private RTVPyear As String
    Private RTVPfolderNum As String = Nothing
    Private Filename As String

    Public Sub New(filename As String)
        Me.Filename = filename
        ReadInfoWorksheet()
    End Sub

    Public Overrides Function ToString() As String
        Return "Customer: " & Me.Customer & vbCrLf & Me.CustomerNumber & vbCrLf & Me.InternalNumber &
            vbCrLf & Me.LocationName & If(TypeOf (Me.City) Is String, " / " & Me.City & ", ", ", ") &
            Me.State & " / " & Me.MilePost & vbCrLf & Me.Division & " DIVISION / " & Me.Subdivision &
            " SUBDIVISION" & " (" & Me.SubdivAbrev & ")" & vbCrLf & vbCrLf
    End Function


    Public Function GetLocationName() As String
        Return Me.LocationName
    End Function


    Public Sub SetCustomer(customer As String)
        Me.Customer = customer
    End Sub


    Public Function GetCustomer() As String
        Return Me.Customer
    End Function


    Public Function GetMilePost() As String
        Return Me.MilePost
    End Function


    Public Function GetInternalNumber() As String
        Return Me.InternalNumber
    End Function


    Public Function GetCustomerNumber() As String
        Return Me.CustomerNumber
    End Function


    Public Function GetDivision() As String
        Return Me.Division
    End Function


    Public Function GetSubdivision() As String
        Return Me.Subdivision
    End Function


    Public Function GetSubdivAbrev() As String
        Return Me.SubdivAbrev
    End Function


    Public Function GetState() As String
        Return Me.State
    End Function


    Public Function GetCity() As String
        Return Me.City
    End Function


    Public Function IsRTVP() As Boolean
        Return Me.RTVP
    End Function


    Public Function GetRTVPyear() As String
        Return Me.RTVPyear
    End Function


    Public Function GetRTVPfolderNum() As String
        Return Me.RTVPfolderNum
    End Function


    Private Sub ReadInfoWorksheet()
        Dim objExcel = CreateObject("Excel.Application")
        objExcel.Visible = False
        Dim excelInput = objExcel.Workbooks.Open(Me.Filename)
        With Me
            .LocationName = UCase(excelInput.Sheets("Sheet1").Range("B2").Value)
            .MilePost = UCase(excelInput.Sheets("Sheet1").Range("B3").Value)
            .InternalNumber = UCase(excelInput.Sheets("Sheet1").Range("B4").Value)
            .CustomerNumber = UCase(excelInput.Sheets("Sheet1").Range("B5").Value)
            .Division = UCase(excelInput.Sheets("Sheet1").Range("B6").Value)
            .Subdivision = UCase(excelInput.Sheets("Sheet1").Range("B7").Value)
            .SubdivAbrev = UCase(excelInput.Sheets("Sheet1").Range("B9").Value)
            .State = UCase(excelInput.Sheets("Sheet1").Range("B15").Value)
            .City = UCase(excelInput.Sheets("Sheet1").Range("D1").Value)
            'MsgBox(TypeName(.City) & ": """ & .City & """")
            If UCase(excelInput.Sheets("Sheet1").Range("B18").Value) = 1 Then
                .RTVP = True
            End If
            Dim numStr = UCase(excelInput.Sheets("Sheet1").Range("D4").Value)
            If Not numStr.Equals(vbNullChar) Then
                If Len(.RTVPfolderNum) = 1 Then
                    .RTVPfolderNum = "0" & .RTVPfolderNum
                End If
                .RTVPyear = UCase(excelInput.Sheets("Sheet1").Range("D3").Value)
            End If
        End With
        objExcel.Application.Quit
        objExcel = Nothing
    End Sub

End Class