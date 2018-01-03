Imports VBStyleAnalyzer
Imports VBStyleAnalyzer.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace VBStyleAnalyzer.Test
    <TestClass>
    Public Class NamesShouldNotStartWithMyTests
        Inherits CodeFixVerifier

        <TestMethod>
        Public Sub PropertiesVariablesAndClassesNamedIncorrectlyAreIdentified()
            Dim test = "
Class Class1
    Property MyProperty As Integer

    Private Property myOtherProperty As Integer

    Sub Other(myParam As Integer, myOtherParam As String)
    End Sub
End Class

Class MyClass
    Property aProp As Integer

    Property otherProperty As Integer

    Sub Other(aParam As Boolean, myParam As Integer)
    End Sub

    Sub Another(aParam As Boolean, thisParam As Integer)
    End Sub
End Class

Class MyModule
    Property myProperty As Integer

    Property otherProperty As Integer

    Sub Other(aParam As Boolean, myParam As Integer)
    End Sub

    Sub Other2(aParam As Boolean, thisParam As Integer)
    End Sub
End Class
"

            Dim expected As DiagnosticResult() = {
                                                     ExpectedDiagnostic(3, 14),
                                                     ExpectedDiagnostic(5, 22),
                                                     ExpectedDiagnostic(7, 15),
                                                     ExpectedDiagnostic(7, 35),
                                                     ExpectedDiagnostic(11, 7),
                                                     ExpectedDiagnostic(16, 34),
                                                     ExpectedDiagnostic(23, 7),
                                                     ExpectedDiagnostic(24, 14),
                                                     ExpectedDiagnostic(28, 34)
            }

            VerifyBasicDiagnostic(test, expected)
        End Sub

        Private Function ExpectedDiagnostic(line As Integer, column As Integer) As DiagnosticResult
            Return New DiagnosticResult With {.Id = NamesShouldNotStartWithMyAnalyzer.DiagnosticId,
                .Message = NamesShouldNotStartWithMyAnalyzer.MessageFormat,
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {
                                                                New DiagnosticResultLocation("Test0.vb", line, column)
                                                            }
                }
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New NamesShouldNotStartWithMyAnalyzer()
        End Function
    End Class
End Namespace
