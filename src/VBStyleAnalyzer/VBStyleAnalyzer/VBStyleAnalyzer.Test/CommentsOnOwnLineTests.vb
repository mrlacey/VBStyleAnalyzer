Imports VBStyleAnalyzer
Imports VBStyleAnalyzer.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace VBStyleAnalyzer.Test
    <TestClass>
    Public Class CommentsOnOwnLineTests
        Inherits CodeFixVerifier

        <TestMethod>
        Public Sub PropertiesVariablesAndClassesNamedIncorrectlyAreIdentified()
            Dim test = "
' Comments here are fine
Dim whitespaceTrivia = allTrivia..ToList() ' This is wrong

Dim something =2
' This is ok too
"

            Dim expected As DiagnosticResult() = {
                                                   ExpectedDiagnostic(3, 44)
            }

            VerifyBasicDiagnostic(test, expected)
        End Sub

        Private Function ExpectedDiagnostic(line As Integer, column As Integer) As DiagnosticResult
            Return New DiagnosticResult With {.Id = CommentsOnOwnLineAnalyzer.DiagnosticId,
                .Message = CommentsOnOwnLineAnalyzer.MessageFormat,
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {
                                                              New DiagnosticResultLocation("Test0.vb", line, column)
                                                            }
                }
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New CommentsOnOwnLineAnalyzer()
        End Function
    End Class
End Namespace
