Imports VBStyleAnalyzer
Imports VBStyleAnalyzer.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace VBStyleAnalyzer.Test
    <TestClass>
    Public Class CommentFormattingTests
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
        Public Sub MultiSpaceCommentTriggersNothing()

            Dim test = "
Module Module1
    '   
    Sub Main()
    End Sub
End Module"

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
        Public Sub CommentWithBulletPointTriggersNothing()

            Dim test = "
Module Module1
    ' A valid comment.
    '   * some point
    Sub Main()

    End Sub

End Module"

            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub MultipleCommentsOnConsecutiveLinesTriggersNothing()

            Dim test = "
Module Module1
    ' A valid comment.
    ' Another valid comment.
    Sub Main()

    End Sub

End Module"

            VerifyBasicDiagnostic(test)
        End Sub
        
        ' ToDo: Handle multi-line comments
        ''<TestMethod>
        Public Sub ValidMultiLineCommentTriggersNothing()

            Dim test = "
Module Module1
    ' A valid comment
    ' over multiple lines.
    Sub Main()

    End Sub

End Module"

            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub ValidCommentOnLineBeforeInvalidDetected()

            Dim test = "
Module Module1
    ' A valid comment.
    ' an invalid comment
    Sub Main()

    End Sub

End Module"

            Dim expected = New DiagnosticResult With {.Id = CommentFormattingAnalyzer.DiagnosticId,
                    .Message = CommentFormattingAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 4, 5)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected)
        End Sub

        <TestMethod>
        Public Sub ValidCommentOnLineAfterInvalidDetected()

            Dim test = "
Module Module1
    ' an invalid comment
    ' A valid comment.
    Sub Main()

    End Sub

End Module"

            Dim expected = New DiagnosticResult With {.Id = CommentFormattingAnalyzer.DiagnosticId,
                    .Message = CommentFormattingAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 3, 5)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected)
        End Sub

        <TestMethod>
        Public Sub SingleInvalidCommentBothIssuesDetected()
            Dim test = "
Module Module1
    ' invalid comment
    Sub Main()

    End Sub

End Module"
            Dim expected = New DiagnosticResult With {.Id = CommentFormattingAnalyzer.DiagnosticId,
                    .Message = CommentFormattingAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 3, 5)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected)
        End Sub

        <TestMethod>
        Public Sub SingleInvalidCommentFirstIssueDetected()
            Dim test = "
Module Module1
    ' invalid comment.
    Sub Main()

    End Sub

End Module"
            Dim expected = New DiagnosticResult With {.Id = CommentFormattingAnalyzer.DiagnosticId,
                    .Message = CommentFormattingAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 3, 5)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected)
        End Sub

        <TestMethod>
        Public Sub SingleInvalidCommentLastIssueDetected()
            Dim test = "
Module Module1
    ' Invalid comment
    Sub Main()

    End Sub

End Module"
            Dim expected = New DiagnosticResult With {.Id = CommentFormattingAnalyzer.DiagnosticId,
                    .Message = CommentFormattingAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 3, 5)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected)
        End Sub

        <TestMethod>
        Public Sub MultipleInvalidCommentsFound()
            Dim test = "
Module Module1
    'invalid comment.
    Sub Main()
        'Another invalid comment
    End Sub

End Module"
            Dim expected1 = New DiagnosticResult With {.Id = CommentFormattingAnalyzer.DiagnosticId,
                    .Message = CommentFormattingAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 3, 5)
                                                                }
                    }
            Dim expected2 = New DiagnosticResult With {.Id = CommentFormattingAnalyzer.DiagnosticId,
                    .Message = CommentFormattingAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 5, 9)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected1, expected2)
        End Sub

        <TestMethod>
        Public Sub InvalidCommentFoundAmongstValidOnes()
            Dim test = "
Module Module1
    ' A valid comment.
    Sub Main()
        'invalid comment
    End Sub

End Module"
            Dim expected = New DiagnosticResult With {.Id = CommentFormattingAnalyzer.DiagnosticId,
                    .Message = CommentFormattingAnalyzer.MessageFormat,
                    .Severity = DiagnosticSeverity.Warning,
                    .Locations = New DiagnosticResultLocation() {
                                                                    New DiagnosticResultLocation("Test0.vb", 5, 9)
                                                                }
                    }

            VerifyBasicDiagnostic(test, expected)
        End Sub

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New CommentFormattingAnalyzer()
        End Function

    End Class
End Namespace
