Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
<NoCodeFix("Creating a fix for this is low priority")>
Public Class CommentWithoutAsterisksAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "VBSAC004"

    Public Shared ReadOnly Title As String = "Don't surround comments with asterisks"
    Public Shared ReadOnly MessageFormat As String = "Do not surround comments with formatted blocks of asterisks."
    Public Shared ReadOnly Description As String = "Remove formatted blocks of asterisks from around comments."

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
        Dim commentNodes = From node In root.DescendantTrivia() Where node.IsKind(SyntaxKind.CommentTrivia) Select node
        Dim comments = commentNodes.ToList()

        For Each comment As SyntaxTrivia In comments
            Dim commentText = comment.ToString().Substring(1).Trim()

            If commentText.StartsWith("*") AndAlso commentText.EndsWith("*") Then
                Dim diag = Diagnostic.Create(Rule, comment.GetLocation())
                context.ReportDiagnostic(diag)
            End If
        Next
    End Sub
End Class
