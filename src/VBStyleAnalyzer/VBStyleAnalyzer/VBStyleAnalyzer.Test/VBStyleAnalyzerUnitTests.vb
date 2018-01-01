Imports VBStyleAnalyzer
Imports VBStyleAnalyzer.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace VBStyleAnalyzer.Test
    <TestClass>
    Public Class CommentStartsWithSpaceTests
        Inherits CodeFixVerifier

        <TestMethod>
        Public Sub NoDiagnosticsShowfNoComments()
            Dim test = ""
            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub ValidCommentTriggersNothing()

            Dim test = "
Module Module1
    ' A valid comment.
    Sub Main()

    End Sub

End Module"

            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub SingleInvalidCommentDetectedAndCanBeFixed()
            Dim test = "
Module Module1
    'Invalid comment.
    Sub Main()

    End Sub

End Module"
            Dim expected = New DiagnosticResult With {.Id = CommentStartsWithSpaceAnalyzer.DiagnosticId,
                    .Message = CommentStartsWithSpaceAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 3, 5)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected)

            Dim fixtest = "
Module Module1
    ' Invalid comment.
    Sub Main()

    End Sub

End Module"
            VerifyBasicFix(test, fixtest)
        End Sub

        <TestMethod>
        Public Sub MultipleInvalidCommentsFound()
            Dim test = "
Module Module1
    'Invalid comment.
    Sub Main()
        'Another invalid comment.
    End Sub

End Module"
            Dim expected1 = New DiagnosticResult With {.Id = CommentStartsWithSpaceAnalyzer.DiagnosticId,
                    .Message = CommentStartsWithSpaceAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 3, 5)
                                                                }
                    }
            Dim expected2 = New DiagnosticResult With {.Id = CommentStartsWithSpaceAnalyzer.DiagnosticId,
                    .Message = CommentStartsWithSpaceAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 5, 9)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected1, expected2)

            Dim fixtest = "
Module Module1
    ' Invalid comment.
    Sub Main()
        ' Another invalid comment.
    End Sub

End Module"
            VerifyBasicFix(test, fixtest)
        End Sub

        <TestMethod>
        Public Sub InvalidCommentFoundAmongstValidOnes()
            Dim test = "
Module Module1
    ' A valid comment.
    Sub Main()
        'Invalid comment.
    End Sub

End Module"
            Dim expected = New DiagnosticResult With {.Id = CommentStartsWithSpaceAnalyzer.DiagnosticId,
                    .Message = CommentStartsWithSpaceAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 5, 9)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected)

            Dim fixtest = "
Module Module1
    ' A valid comment.
    Sub Main()
        ' Invalid comment.
    End Sub

End Module"
            VerifyBasicFix(test, fixtest)
        End Sub

        Protected Overrides Function GetBasicCodeFixProvider() As CodeFixProvider
            Return New CommentStartsWithSpaceFixProvider()
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New CommentStartsWithSpaceAnalyzer()
        End Function

    End Class
End Namespace
