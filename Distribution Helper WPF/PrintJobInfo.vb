Public Class PrintJobInfo
    Private files As Array
    Private twoSided As Boolean

    Public Sub New(filesArray As Array, twoSidedPrinting As Boolean)
        'Console.WriteLine("Creating ProgramFile Obj")
        Me.files = filesArray
        Me.twoSided = twoSidedPrinting
    End Sub


    Property FilesToPrint() As Array
        Get
            Return Me.files
        End Get

        Set(filesArray As Array)
            Me.files = filesArray
        End Set
    End Property


    Property TwoSidedPrinting() As Boolean
        Get
            Return Me.twoSided
        End Get

        Set(twoSidedPrinting As Boolean)
            Me.twoSided = twoSidedPrinting
        End Set
    End Property
End Class
