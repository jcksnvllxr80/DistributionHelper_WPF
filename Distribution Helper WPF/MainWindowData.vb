Public Class MainWindowData
    Private LocationName As String
    Private Customer As String
    Private CustomerJobNum As String
    Private InternalJobNum As String
    Private DistributionDate As Date

    Public Sub New(locName As String, cust As String, customerNumber As String, internalNumber As String, distDate As Date)
        Me.LocationName = locName
        Me.Customer = cust
        Me.CustomerJobNum = customerNumber
        Me.InternalJobNum = internalNumber
        Me.DistributionDate = distDate
    End Sub


    Public Function GetLocationName() As String
        Return Me.LocationName
    End Function


    Public Function GetCustomer() As String
        Return Me.Customer
    End Function


    Public Function GetDistributionDate() As String
        Return Me.DistributionDate
    End Function


    Public Function GetInternalNumber() As String
        Return Me.InternalJobNum
    End Function


    Public Function GetCustomerNumber() As String
        Return Me.CustomerJobNum
    End Function
End Class
