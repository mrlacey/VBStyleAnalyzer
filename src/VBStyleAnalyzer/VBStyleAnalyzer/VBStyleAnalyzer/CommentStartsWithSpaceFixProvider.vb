Imports System.Collections.Immutable
Imports System.Composition
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.VisualBasic

<ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(CommentStartsWithSpaceFixProvider)), [Shared]>
Public Class CommentStartsWithSpaceFixProvider
    Inherits CodeFixProvider

    Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String)
        Get
            Return ImmutableArray.Create(CommentStartsWithSpaceAnalyzer.DiagnosticId)
        End Get
    End Property

    Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
        ' See https://github.com/dotnet/roslyn/blob/master/docs/analyzers/FixAllProvider.md for more information on Fix All Providers
        Return WellKnownFixAllProviders.BatchFixer
    End Function

    Public NotOverridable Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task

        ' TODO: Replace the following code with your own analysis, generating a CodeAction for each fix to suggest

        Dim dia = context.Diagnostics.First()

        ' Register a code action that will invoke the fix.
        context.RegisterCodeFix(
            CodeAction.Create(
                title:=CommentStartsWithSpaceAnalyzer.Title,
                createChangedDocument := Function (c) AddSpaceAtStartOfComment(context.Document, dia, context.CancellationToken),
                equivalenceKey:=CommentStartsWithSpaceAnalyzer.Title),
            dia)

        Return Task.FromResult(False)
    End Function

    Private Async Function AddSpaceAtStartOfComment(document As Document, diag As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
        Dim root = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)

        Dim diagnosticSpan = diag.Location.SourceSpan

        Dim trivia = root.FindToken(diagnosticSpan.Start).Parent.DescendantTrivia().First(Function (d) d.IsKind(SyntaxKind.CommentTrivia))

        Dim currentComment = trivia.ToString()
        Dim newTrivia = SyntaxFactory.CommentTrivia(currentComment.Replace("'", "' "))

        Dim newRoot = root.ReplaceTrivia(trivia, newTrivia)

        Dim updatedDocument = document.WithSyntaxRoot(newRoot)
        Return updatedDocument
    End Function
End Class
