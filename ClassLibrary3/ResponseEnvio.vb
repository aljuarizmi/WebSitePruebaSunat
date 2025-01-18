Public Class ResponseEnvio
    Public Property codRespuesta As String
    Public Property arcCdr As String
    Public Property indCdrGenerado As String

    Public Property status As String
    Public Property message As String

    'Private _error As String
    'Public Property error() As String
    '    Get
    '        Return _error
    '    End Get
    '    Set(ByVal value As String)
    '        _error = value
    '    End Set
    'End Property

    'Public Property [error] As String
    Public Property [error] As List(Of [error])
End Class
