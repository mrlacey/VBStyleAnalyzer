Imports VBStyleAnalyzer
Imports VBStyleAnalyzer.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace VBStyleAnalyzer.Test
    <TestClass>
    Public Class PropertyDeclarationBracesTests
        Inherits CodeFixVerifier

        <TestMethod>
        Public Sub NoDiagnosticsShowIfNoProperty()

            Dim test = "
Module Module1

End Module"

            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub NoDiagnosticsShowfNoBraces()

            Dim test = "
Module Module1
    Property SomeProperty As Integer
End Module"

            VerifyBasicDiagnostic(test)
        End Sub

        <TestMethod>
        Public Sub DiagnosticsShownIfBracesIncluded()
            Dim test = "
Module Module1
    Property SomeProperty() As Integer
End Module"

            VerifyBasicDiagnostic(test, ExpectedDiagnostic(3, 14))
        End Sub

        <TestMethod>
        Public Sub DiagnosticsShownIfBracesIncluded_FullProperty()
            Dim test = "
Module Module1
    Private thePropertyValue As String
    Public Property TheProperty() As String
        Get
            Return thePropertyValue
        End Get
        Set(ByVal value As String)
            thePropertyValue = value
        End Set
    End Property
End Module"

            VerifyBasicDiagnostic(test, ExpectedDiagnostic(4, 21))
        End Sub

        <TestMethod>
        Public Sub DiagnosticsShownIfBracesIncluded_WithAccessibilityIndicator()
            Dim test = "
Module Module1
    Public Property SomeProperty() As Integer
End Module"

            VerifyBasicDiagnostic(test, ExpectedDiagnostic(3, 21))
        End Sub

        <TestMethod>
        Public Sub DiagnosticsShownIfBracesIncluded_Readonly()
            Dim test = "
Module Module1
    Property ReadOnly SomeOtherProperty() As Integer
End Module"

            VerifyBasicDiagnostic(test, ExpectedDiagnostic(3, 14))
        End Sub

        <TestMethod>
        Public Sub DiagnosticsShownIfBracesIncluded_WithoutType()
            Dim test = "
Module Module1
    Property AnotherProperty() = ""something""
End Module"

            VerifyBasicDiagnostic(test, ExpectedDiagnostic(3, 14))
        End Sub

        <TestMethod>
        Public Sub DiagnosticsShownIfBracesIncluded_Multiples()
            Dim test = "
Module Module1
    Property SomeProperty() As Integer
    Property ReadOnly SomeOtherProperty() As Integer
    Property AnotherProperty() = ""something""
    Sub DoSomething(id as Integer)
        Dim x As String = ""test String""
        System.Diagnostics.Debug(x)
    End Sub
End Module"

            VerifyBasicDiagnostic(test, ExpectedDiagnostic(3, 14), ExpectedDiagnostic(4, 14), ExpectedDiagnostic(5, 14))
        End Sub

        Private Function ExpectedDiagnostic(line As Integer, column As Integer) As DiagnosticResult
            Return New DiagnosticResult With {.Id = PropertyDeclarationBracesAnalyzer.DiagnosticId,
                .Message = PropertyDeclarationBracesAnalyzer.MessageFormat,
                .Severity = DiagnosticSeverity.Warning,
                .Locations = New DiagnosticResultLocation() {
                                                                New DiagnosticResultLocation("Test0.vb", line, column)
                                                            }
                }
        End Function

        Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
            Return New PropertyDeclarationBracesAnalyzer()
        End Function

    End Class
End Namespace
