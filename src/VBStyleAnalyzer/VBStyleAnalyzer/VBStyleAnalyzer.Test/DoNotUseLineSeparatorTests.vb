Imports VBStyleAnalyzer
Imports VBStyleAnalyzer.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace VBStyleAnalyzer.Test
    <TestClass>
    Public Class DoNotUseLineSeparatorTests
        Inherits CodeFixVerifier

        <TestMethod>
        Public Sub NoDiagnosticsShowfNoComments()

            Dim test = "
Module Module1
    Dim a = 1 : Dim b = 2

    Sub DoSomething()
        ' blah blah blah
ALabel:   Debug.WriteLine(""end"")
    End Sub
End Module"
            VerifyBasicDiagnostic(test, ExpectedDiagnostic(3, 15))
        End Sub

        Private Function ExpectedDiagnostic(line As Integer, column As Integer) As DiagnosticResult
            Return New DiagnosticResult With {.Id = DoNotUseLineSeparatorAnalyzer.DiagnosticId,
                .Message = DoNotUseLineSeparatorAnalyzer.MessageFormat,
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {
                                                                New DiagnosticResultLocation("Test0.vb", line, column)
                                                            }
                }
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New DoNotUseLineSeparatorAnalyzer()
        End Function

    End Class
End Namespace
