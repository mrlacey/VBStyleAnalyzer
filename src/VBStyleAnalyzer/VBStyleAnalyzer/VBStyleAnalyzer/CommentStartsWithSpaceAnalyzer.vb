Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
Public Class CommentStartsWithSpaceAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "VBSAC003"

    Public Shared ReadOnly Title As String = "Comments must start with a space"
    Public Shared ReadOnly MessageFormat As String = "The comment does not start with a space."
    Public Shared ReadOnly Description As String = "Insert one space between the comment delimiter (') and the comment text."

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

        For Each node In commentNodes
            Dim commentText = node.ToString()

            If commentText.Length > 2 Then
                If Not commentText.StartsWith("' ") And commentText.Substring(2, 1) IsNot " " Then
                    Dim diag = Diagnostic.Create(Rule, node.GetLocation())
                    context.ReportDiagnostic(diag)
                End If
            End If
        Next
    End Sub
End Class
