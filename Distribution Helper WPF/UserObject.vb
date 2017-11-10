Public Class UserObject
    Public FirstName As String
    Public LastName As String
    Public FullName As String
    Public Email As String
    Public Domain As String
    Public UserName As String
    Public Is64Bit As Boolean

    Public Sub New()
        Dim strNameLastThenFirst = Split(GetObject("LDAP://" & CreateObject("ADSystemInfo").UserName).Get("displayName"), ", ")
        Me.FirstName = strNameLastThenFirst(1)
        Me.LastName = strNameLastThenFirst(0)
        Me.FullName = Me.FirstName & " " & Me.LastName
        Me.UserName = Environment.UserName
        Me.Domain = Environment.UserDomainName
        Me.Email = Me.UserName.ToLower & "@" & Me.Domain.ToLower & ".com"
        Me.Is64Bit = Environment.Is64BitOperatingSystem
    End Sub


    Public Overrides Function ToString() As String
        Return "Name:" & vbTab & Me.FullName & vbCrLf &
               "Email:" & vbTab & Me.Email
    End Function

End Class
