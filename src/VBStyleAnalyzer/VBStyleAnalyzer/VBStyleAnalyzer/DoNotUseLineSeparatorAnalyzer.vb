Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
<NoCodeFix("Creating a fix for this is low priority")>
Public Class DoNotUseLineSeparatorAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "VBSAL001"

    Public Shared ReadOnly Title As String = "Multiple statements on one line"
    Public Shared ReadOnly MessageFormat As String = "Use only one statement per line. Don't use the Visual Basic line separator character (:)."
    Public Shared ReadOnly Description As String = "Use only one statement per line. Don't use the Visual Basic line separator character (:)."

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
        Dim colonNodes = From node In root.DescendantTrivia() Where node.IsKind(SyntaxKind.ColonTrivia) Select node
        Dim colons = colonNodes.ToList()

        ' Assuming all colonTrivia are line separators and should be replaced with a NewLine
        For Each colon As SyntaxTrivia In colons
            Dim diag = Diagnostic.Create(Rule, colon.GetLocation())
            context.ReportDiagnostic(diag)
        Next
    End Sub
End Class
