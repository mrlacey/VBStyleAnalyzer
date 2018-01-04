Imports VBStyleAnalyzer
Imports VBStyleAnalyzer.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace VBStyleAnalyzer.Test
    <TestClass>
    Public Class CommentWithoutAsterisksTests
        Inherits CodeFixVerifier

        <TestMethod>
        Public Sub NoDiagnosticsShowfNoComments()
            Dim test = ""
            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub EmptyCommentTriggersNothing()

            Dim test = "
Module Module1
    '
    Sub Main()
    End Sub
End Module"

            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub TextCommentTriggersNothing()

            Dim test = "
Module Module1
    ' Something important to know
    Sub Main()
    End Sub
End Module"

            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub AsteriskAtStartTriggersNothing()

            Dim test = "
Module Module1
    '   * Something important to know
    Sub Main()
    End Sub
End Module"

            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub AsteriskAtEndTriggersNothing()

            Dim test = "
Module Module1
    ' Something important to know *
    Sub Main()
    End Sub
End Module"

            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub InvalidCommentFoundAmongstValidOnes()
            Dim test = "
Module Module1
    ' ***********************
    ' * An invalid comment. *
    ' ***********************
    Sub Main()
        ' Valid comment
    End Sub

End Module"

            VerifyBasicDiagnostic(test, ExpectedDiagnostic(3, 5), ExpectedDiagnostic(4, 5), ExpectedDiagnostic(5, 5))
        End Sub

        Private Function ExpectedDiagnostic(line As Integer, column As Integer) As DiagnosticResult
            Return New DiagnosticResult With {.Id = CommentWithoutAsterisksAnalyzer.DiagnosticId,
                .Message = CommentWithoutAsterisksAnalyzer.MessageFormat,
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {
                                                                New DiagnosticResultLocation("Test0.vb", line, column)
                                                            }
                }
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New CommentWithoutAsterisksAnalyzer()
        End Function

    End Class
End Namespace
