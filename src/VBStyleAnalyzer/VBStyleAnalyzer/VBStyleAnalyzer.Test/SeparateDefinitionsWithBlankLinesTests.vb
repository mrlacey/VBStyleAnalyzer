Imports VBStyleAnalyzer
Imports VBStyleAnalyzer.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace VBStyleAnalyzer.Test
    <TestClass>
    Public Class SeparateDefinitionsWithBlankLinesTests
        Inherits CodeFixVerifier

        <TestMethod>
        Public Sub CorrectSpacingTriggersNoDiagnostics()
            Dim test = "
Imports Some.Namespace
Imports Some.Other.Namespace

Imports Totally.Different.Namespace

Interface IAsset
    Event ComittedChange(ByVal Success As Boolean)

    Property Division() As String

    Function GetID() As Integer
End Interface

Class Class1
    Property SomeProperty As Integer

    Property OtherProperty As Integer

    Sub Main()
    End Sub

    Sub Other()
    End Sub
End Class

Class CommentClass1
    Property SomeProperty As Integer

    ' some comment
    Property OtherProperty As Integer

    Sub Main()
    End Sub

    ' a comment
    Sub Other()
    End Sub
End Class

Module Module1
    Property SomeProperty As Integer

    Property OtherProperty As Integer

    Private _anotherProperty As Float
    Public Property AnotherProperty As Float
        Get
            Return _anotherProperty
        End Get

        Set (value as Float)
            _anotherProperty = value
        End Get
     End Property

    Private _readonlyProperty As Float

    Public ReadOnly Property ReadOnlyProperty As Float
        Get
            Return _readonlyProperty
        End Get
     End Property

    Sub Main()
    End Sub

    Sub Other()
    End Sub
End Module

Module Module2
    Property SomeProperty As Integer
End Module

Class Class2
    Inherits BaseClass

    Property SomeProperty As Integer

    Sub Main()
    End Sub
End Class

Class Class3
    Inherits BaseClass
    Implements SomeInterface

    Property SomeProperty As Integer

    Sub Main()
    End Sub
End Class

Class Class4
    Inherits BaseClass
    Implements SomeInterface
    Implements AnotherInterface

    Property SomeProperty As Integer

    Sub Main()
    End Sub
End Class

Class Class5
    Implements SomeInterface

    Property SomeProperty As Integer

    Sub Main()
    End Sub
End Class
"

            VerifyBasicDiagnostic(test, ExpectedDiagnostic(5, 1))
        End Sub

        Private Function ExpectedDiagnostic(line As Integer, column As Integer) As DiagnosticResult
            Return New DiagnosticResult With {.Id = SeparateDefinitionsWithBlankLinesAnalyzer.DiagnosticId,
                .Message = SeparateDefinitionsWithBlankLinesAnalyzer.MessageFormat,
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {
                                                                New DiagnosticResultLocation("Test0.vb", line, column)
                                                            }
                }
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New SeparateDefinitionsWithBlankLinesAnalyzer()
        End Function

    End Class
End Namespace
