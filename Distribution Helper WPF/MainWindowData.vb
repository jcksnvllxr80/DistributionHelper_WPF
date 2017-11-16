Public Class MainWindowData
    Private LocationName As String
    Private Customer As String
    Private CustomerJobNum As String
    Private InternalJobNum As String
    Private DistributionDate As Date
    Private TrackingNum As String
    Private InvoiceNum As String
    Private ShipMethod As String
    Private Recipient As String
    Private Street As String
    Private City As String
    Private State As String
    Private ZipCode As String


    Public Sub New(locName As String, cust As String, customerNumber As String, internalNumber As String, distDate As Date)
        Me.LocationName = locName
        Me.Customer = cust
        Me.CustomerJobNum = customerNumber
        Me.InternalJobNum = internalNumber
        Me.DistributionDate = distDate
        Me.TrackingNum = ""
        Me.InvoiceNum = ""
        Me.ShipMethod = ""
        Me.Recipient = ""
        Me.Street = ""
        Me.City = ""
        Me.State = ""
        Me.ZipCode = ""
    End Sub


    Public Sub New(locName As String, cust As String, customerNumber As String, internalNumber As String, distDate As Date,
                   trackingNumber As String, invoiceNumber As String, shippingMethod As String, recipientName As String, streetAddress As String,
                   cityAddress As String, stateAddress As String, zipCodeAddress As String)
        Me.LocationName = locName
        Me.Customer = cust
        Me.CustomerJobNum = customerNumber
        Me.InternalJobNum = internalNumber
        Me.DistributionDate = distDate
        Me.TrackingNum = trackingNumber
        Me.InvoiceNum = invoiceNumber
        Me.ShipMethod = shippingMethod
        Me.Recipient = recipientName
        Me.Street = streetAddress
        Me.City = cityAddress
        Me.State = stateAddress
        Me.ZipCode = zipCodeAddress
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

    Public Function GetTrackingNumber() As String
        Return Me.TrackingNum
    End Function


    Public Function GetInvoiceNumber() As String
        Return Me.InvoiceNum
    End Function


    Public Function GetShippingMethod() As String
        Return Me.ShipMethod
    End Function


    Public Function GetRecipientName() As String
        Return Me.Recipient
    End Function

    Public Function GetStreet() As String
        Return Me.Street
    End Function


    Public Function GetCity() As String
        Return Me.City
    End Function


    Public Function GetState() As String
        Return Me.State
    End Function


    Public Function GetZipCode() As String
        Return Me.ZipCode
    End Function
End Class
