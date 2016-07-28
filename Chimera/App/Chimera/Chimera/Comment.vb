'' a structy class :)
Public Class Comment

    Public comment As String
    Public name As String

    Public Sub New(c As String, n As String)
        comment = unescape_comment(c)
        name = n
    End Sub

    Private Function unescape_comment(com As String)
        Return com.Replace("\n", vbNewLine).Replace("\t", vbTab)
    End Function

    Public Overrides Function ToString() As String
        Return (name)
    End Function
End Class