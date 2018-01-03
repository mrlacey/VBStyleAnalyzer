Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
<NoCodeFix("Creating a fix for this is low priority")>
Public Class VariablesShouldNotStartWithMyAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "VBSAN001"

    Public Shared ReadOnly Title As String = "Names should not start with ""My"""
    Public Shared ReadOnly MessageFormat As String = "Do not use ""My"" or ""my"" at the start of a name."
    Public Shared ReadOnly Description As String = "Using ""My"" or ""my"" at the start of a name can lead to confusion with `My` objects."

    Private Shared Rule As New DiagnosticDescriptor(DiagnosticId, Title, MessageFormat, Categories.Commenting, defaultSeverity:=DiagnosticSeverity.Warning, isEnabledByDefault:=True, description:=Description)

    Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
        Get
            Return ImmutableArray.Create(Rule)
        End Get
    End Property

    Public Overrides Sub Initialize(context As AnalysisContext)
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None)
        context.RegisterSyntaxTreeAction(AddressOf AnalyzeSyntaxTree)
    End Sub

    Private Sub AnalyzeSyntaxTree(context As SyntaxTreeAnalysisContext)
        Dim root = context.Tree.GetCompilationUnitRoot()
        Dim identiferNodes = From token In root.DescendantTokens() Where token.RawKind.Equals(700) Select token
        Dim identifiers = identiferNodes.ToList()

        For Each identifier As SyntaxToken In identifiers
            If identifier.ValueText.ToUpperInvariant.StartsWith("MY") Then
                Dim diag = Diagnostic.Create(Rule, identifier.GetLocation())
                context.ReportDiagnostic(diag)
            End If
        Next
    End Sub
End Class
