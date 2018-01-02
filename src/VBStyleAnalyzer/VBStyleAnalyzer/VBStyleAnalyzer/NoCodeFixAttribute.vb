<AttributeUsage(AttributeTargets.Class, Inherited := False, AllowMultiple := False)>
Public Class NoCodeFixAttribute
    Inherits Attribute

    ' The reason why the DiagnosticAnlyzer does not have a code fix.
    Public ReadOnly Reason As String

    Public Sub New(reason As String)
        Me.Reason = reason
    End Sub

End Class
