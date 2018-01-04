Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
<NoCodeFix("Creating a fix for this is low priority")>
Public Class LinqVariableNameAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "VBSAM001"

    Public Shared ReadOnly Title As String = "Short LINQ query variable"
    Public Shared ReadOnly MessageFormat As String = "Use meaningful names for LINQ query variables."
    Public Shared ReadOnly Description As String = "Very short variable names are not meaningful and a name that confers meaning more clearly should be used."

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
        Dim tokens = root.DescendantTokens().ToList()

        For i = 0 To tokens.Count - 1
            ' Long form LINQ syntax
            If tokens(i).IsKind(SyntaxKind.WhereKeyword) AndAlso i + 2 <= tokens.Count Then
                If tokens(i + 1).IsKind(SyntaxKind.IdentifierToken) Then
                    If tokens(i + 2).IsKind(SyntaxKind.DotToken) Then
                        Dim variableName = tokens(i + 1).ValueText

                        If variableName.Length < 2 Then
                            Dim diag = Diagnostic.Create(Rule, tokens(i + 1).GetLocation())
                            context.ReportDiagnostic(diag)
                        End If
                    End If
                End If
                ' Inline LINQ syntax
            ElseIf tokens(i).IsKind(SyntaxKind.IdentifierToken) AndAlso tokens(i).ValueText.Equals("Where") AndAlso i + 4 <= tokens.Count Then
                If tokens(i + 1).IsKind(SyntaxKind.OpenParenToken) Then
                    If tokens(i + 2).IsKind(SyntaxKind.FunctionKeyword) Then
                        If tokens(i + 3).IsKind(SyntaxKind.OpenParenToken) Then
                            If tokens(i + 4).IsKind(SyntaxKind.IdentifierToken) Then
                                Dim variableName = tokens(i + 4).ValueText

                                If variableName.Length < 2 Then
                                    Dim diag = Diagnostic.Create(Rule, tokens(i + 4).GetLocation())
                                    context.ReportDiagnostic(diag)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub
End Class
