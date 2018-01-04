Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
<NoCodeFix("Creating a fix for this is low priority")>
Public Class CommentsOnOwnLineAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "VBSAC001"

    Public Shared ReadOnly Title As String = "Put comment on own line"
    Public Shared ReadOnly MessageFormat As String = "Put comments on a separate line instead of at the end of a line of code."
    Public Shared ReadOnly Description As String = "Put comments on a separate line instead of at the end of a line of code."

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
        Dim comments = root.DescendantTrivia().Where(Function(trivia) trivia.IsKind(SyntaxKind.CommentTrivia)).ToList()

        If comments.Any() Then
            Dim fileLines = context.Tree.GetText().Lines

            For Each comment As SyntaxTrivia In comments
                Dim commentPosition = comment.SyntaxTree.GetLineSpan(comment.Span).StartLinePosition
                Dim lineOfInterest = fileLines(commentPosition.Line)

                If Not String.IsNullOrWhiteSpace(lineOfInterest.Text.ToString().Substring(0, commentPosition.Character)) Then
                    Dim diag = Diagnostic.Create(Rule, comment.GetLocation())
                    context.ReportDiagnostic(diag)
                End If
            Next
        End If
    End Sub
End Class
