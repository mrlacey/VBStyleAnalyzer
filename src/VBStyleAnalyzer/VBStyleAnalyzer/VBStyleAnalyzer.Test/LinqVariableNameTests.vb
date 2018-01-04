Imports VBStyleAnalyzer
Imports VBStyleAnalyzer.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace VBStyleAnalyzer.Test
    <TestClass>
    Public Class LinqVariableNameTests
        Inherits CodeFixVerifier

        <TestMethod>
        Public Sub PropertiesVariablesAndClassesNamedIncorrectlyAreIdentified()
            Dim test = "
Dim commentNodes = From node In root.DescendantTrivia() Where node.IsKind(SyntaxKind.CommentTrivia) Select node
Dim commentNodes = From n In root.DescendantTrivia() Where n.IsKind(SyntaxKind.CommentTrivia) Select n

Dim whitespaceTrivia = allTrivia.Where(Function(trivia) Not trivia.IsKind(SyntaxKind.WhitespaceTrivia))
Dim whitespaceTrivia = allTrivia.Where(Function(x) Not x.IsKind(SyntaxKind.WhitespaceTrivia))
"

            Dim expected As DiagnosticResult() = {
                                                     ExpectedDiagnostic(3, 60),
                                                     ExpectedDiagnostic(6, 49)
            }

            VerifyBasicDiagnostic(test, expected)
        End Sub

        Private Function ExpectedDiagnostic(line As Integer, column As Integer) As DiagnosticResult
            Return New DiagnosticResult With {.Id = LinqVariableNameAnalyzer.DiagnosticId,
                .Message = LinqVariableNameAnalyzer.MessageFormat,
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {
                                                                New DiagnosticResultLocation("Test0.vb", line, column)
                                                            }
                }
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New LinqVariableNameAnalyzer()
        End Function
    End Class
End Namespace
