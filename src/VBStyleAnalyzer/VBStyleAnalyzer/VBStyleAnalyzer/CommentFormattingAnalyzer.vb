Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
<NoCodeFix("Creating a fix for this is low priority")>
Public Class CommentFormattingAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "VBSAC002"

    Public Shared ReadOnly Title As String = "Comments must be formatted correctly"
    Public Shared ReadOnly MessageFormat As String = "The comment does not start with an uppercase letter and end with a period."
    Public Shared ReadOnly Description As String = "Start comment text with an uppercase letter and end comment text with a period."

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
        Dim lastCommentIndex = comments.Count() - 1

        Dim i = 0

        While i <= lastCommentIndex
            Dim node = comments.Item(i)

            Dim addDiagnostic = Sub()
                                    Dim diag = Diagnostic.Create(Rule, node.GetLocation())
                                    context.ReportDiagnostic(diag)
                                End Sub

            Dim commentText = node.ToString().Substring(1).TrimStart()

            If Not String.IsNullOrWhiteSpace(commentText) Then
                Dim firstChar = commentText.First()

                ' Ignore empty comments, all whitespace, or just symbols
                If Char.IsLetter(firstChar) Then
                    Dim firstCharIsUppercase = firstChar.ToString().ToUpper().Equals(firstChar.ToString())
                    Dim lastCharIsPeriod = commentText.Last().ToString().Equals(".")

                    If Not firstCharIsUppercase OrElse Not lastCharIsPeriod Then
                        addDiagnostic()
                    End If
                End If
            End If

            i = i + 1
        End While
    End Sub
End Class
