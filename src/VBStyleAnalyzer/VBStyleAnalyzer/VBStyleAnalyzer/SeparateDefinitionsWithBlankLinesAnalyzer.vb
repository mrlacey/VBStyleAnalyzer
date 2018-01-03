Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
<NoCodeFix("Creating a fix for this is low priority")>
Public Class SeparateDefinitionsWithBlankLinesAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "VBSAL002"

    Public Shared ReadOnly Title As String = "Separate definitions with blank lines"
    Public Shared ReadOnly MessageFormat As String = "Method and property definitions should be separated with a single blank line."
    Public Shared ReadOnly Description As String = "Separate method and property definitions with a single blank line."

    Private Shared ReadOnly Rule As New DiagnosticDescriptor(DiagnosticId, Title, MessageFormat, Categories.Layout, defaultSeverity:=DiagnosticSeverity.Warning, isEnabledByDefault:=True, description:=Description)

    Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
        Get
            Return ImmutableArray.Create(Rule)
        End Get
    End Property

    Public Overrides Sub Initialize(context As AnalysisContext)
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None)
        context.EnableConcurrentExecution()
        context.RegisterSyntaxNodeAction(AddressOf HandleTypeDeclaration, ImmutableArray.Create(SyntaxKind.ClassBlock, SyntaxKind.ModuleBlock, SyntaxKind.StructureBlock, SyntaxKind.InterfaceBlock, SyntaxKind.InheritsStatement))
        context.RegisterSyntaxNodeAction(AddressOf HandleCompilationUnit, SyntaxKind.CompilationUnit)
        context.RegisterSyntaxNodeAction(AddressOf HandleInheritenceDeclaration, ImmutableArray.Create(SyntaxKind.InheritsStatement, SyntaxKind.ImplementsStatement))
        context.RegisterSyntaxNodeAction(AddressOf HandleGetterAndSetter, SyntaxKind.EndGetStatement)
    End Sub

    Private Shared Sub HandleGetterAndSetter(context As SyntaxNodeAnalysisContext)
        Dim declaration = TryCast(context.Node, EndBlockStatementSyntax)
        If declaration IsNot Nothing Then
            Dim subsequentFileLines = declaration.SyntaxTree.GetText.ToString().Substring(declaration.Span.End + 2).Split({Environment.NewLine}, StringSplitOptions.None)

            If subsequentFileLines(0).Contains("Set (") Then
                context.ReportDiagnostic(Diagnostic.Create(Rule, declaration.SyntaxTree.GetLocation(declaration.Span)))
            End If
        End If
    End Sub

    Private Shared Sub HandleInheritenceDeclaration(context As SyntaxNodeAnalysisContext)
        Dim declaration = TryCast(context.Node, InheritsOrImplementsStatementSyntax)
        If declaration IsNot Nothing Then
            Dim subsequentText = declaration.SyntaxTree.GetText.ToString().Substring(declaration.Span.End)

            If Not subsequentText.TrimStart().StartsWith("Implements") Then
                If Not subsequentText.StartsWith(Environment.NewLine & Environment.NewLine) Then
                    context.ReportDiagnostic(Diagnostic.Create(Rule, declaration.SyntaxTree.GetLocation(declaration.Span)))
                End If
            End If
        End If
    End Sub

    Private Shared Sub HandleTypeDeclaration(context As SyntaxNodeAnalysisContext)
        Dim typeDeclaration = TryCast(context.Node, TypeBlockSyntax)
        If typeDeclaration IsNot Nothing Then
            Dim members = typeDeclaration.Members
            HandleMemberList(context, members)
        End If
    End Sub

    Private Shared Sub HandleCompilationUnit(context As SyntaxNodeAnalysisContext)
        Dim compilationUnit = TryCast(context.Node, CompilationUnitSyntax)
        If compilationUnit IsNot Nothing Then
            Dim importUnits = compilationUnit.Imports
            Dim members = compilationUnit.Members
            HandleImports(context, importUnits)
            HandleMemberList(context, members)

            If members.Count > 0 AndAlso compilationUnit.Imports.Count > 0 Then
                ReportIfThereIsNoBlankLine(context, importUnits(importUnits.Count - 1), members(0))
            End If
        End If
    End Sub

    Private Shared Sub HandleImports(context As SyntaxNodeAnalysisContext, importStatements As SyntaxList(Of ImportsStatementSyntax))
        If importStatements.Count < 2 Then
            Return
        End If

        Dim previousLine = importStatements(0).SyntaxTree.ToString().Substring(importStatements(0).Span.Start, importStatements(0).Span.Length)
        Dim previousTopNamespace = previousLine.Substring(0, previousLine.IndexOf("."c))
        Dim previousLineSpan = importStatements(0).SyntaxTree.GetLineSpan(importStatements(0).Span)

        For i = 1 To importStatements.Count - 1
            Dim currentLine = importStatements(i).SyntaxTree.ToString().Substring(importStatements(i).Span.Start, importStatements(i).Span.Length)
            Dim currentTopNamespace = currentLine.Substring(0, currentLine.IndexOf("."c))
            Dim currentLineSpan = importStatements(i).SyntaxTree.GetLineSpan(importStatements(i).Span)

            Dim lineDistance = currentLineSpan.StartLinePosition.Line - previousLineSpan.EndLinePosition.Line

            If previousTopNamespace.Equals(currentTopNamespace) Then
                If lineDistance > 1 Then
                    ' Blank line exists where it shouldn't
                    context.ReportDiagnostic(Diagnostic.Create(Rule, importStatements(i).SyntaxTree.GetLocation(importStatements(i).Span)))
                End If
            Else
                If lineDistance <> 2 Then
                    ' Single blank line is missing
                    context.ReportDiagnostic(Diagnostic.Create(Rule, importStatements(i).SyntaxTree.GetLocation(importStatements(i).Span)))
                End If
            End If

            previousLineSpan = currentLineSpan
            previousTopNamespace = currentTopNamespace
        Next
    End Sub

    Private Shared Sub HandleMemberList(context As SyntaxNodeAnalysisContext, ByVal members As SyntaxList(Of StatementSyntax))
        For i = 1 To members.Count - 1
            If Not members(i - 1).ContainsDiagnostics AndAlso Not members(i).ContainsDiagnostics Then
                If Not members(i).IsKind(SyntaxKind.FieldDeclaration) OrElse Not members(i - 1).IsKind(members(i).Kind()) OrElse IsMultiline(CType(members(i - 1), FieldDeclarationSyntax)) Then
                    ReportIfThereIsNoBlankLine(context, members(i - 1), members(i))
                End If
            End If
        Next
    End Sub

    Private Shared Function IsMultiline(fieldDeclaration As FieldDeclarationSyntax) As Boolean
        Dim lineSpan = fieldDeclaration.SyntaxTree.GetLineSpan(fieldDeclaration.Span)
        Dim attributeLists = fieldDeclaration.AttributeLists
        Dim startLine As Integer
        If attributeLists.Count > 0 Then
            Dim lastAttributeSpan = fieldDeclaration.SyntaxTree.GetLineSpan(attributeLists.Last().FullSpan)
            startLine = lastAttributeSpan.EndLinePosition.Line
        Else
            startLine = lineSpan.StartLinePosition.Line
        End If

        Return startLine <> lineSpan.EndLinePosition.Line
    End Function

    Private Shared Sub ReportIfThereIsNoBlankLine(context As SyntaxNodeAnalysisContext, firstNode As SyntaxNode, secondNode As SyntaxNode)
        Dim parent = firstNode.Parent
        Dim allTrivia = parent.DescendantTrivia(TextSpan.FromBounds(firstNode.Span.[End], secondNode.Span.Start), descendIntoTrivia:=True)
        If Not HasEmptyLine(allTrivia) Then
            context.ReportDiagnostic(Diagnostic.Create(Rule, GetDiagnosticLocation(secondNode)))
        End If
    End Sub

    Private Shared Function GetDiagnosticLocation(node As SyntaxNode) As Location
        If node.HasLeadingTrivia Then
            Return node.GetLeadingTrivia()(0).GetLocation()
        End If

        Dim firstToken = node.ChildTokens().FirstOrDefault()
        If firstToken <> Nothing Then
            Return node.ChildTokens().First().GetLocation()
        End If

        Dim nodeLocation = node.GetLocation()
        If nodeLocation <> Nothing Then
            Return nodeLocation
        End If

        Return Location.None
    End Function

    Private Shared Function HasEmptyLine(ByVal allTrivia As IEnumerable(Of SyntaxTrivia)) As Boolean
        allTrivia = allTrivia.Where(Function(x) Not x.IsKind(SyntaxKind.WhitespaceTrivia))
        Dim previousTrivia As SyntaxTrivia = Nothing

        For Each trivia In allTrivia
            If trivia.IsKind(SyntaxKind.EndOfLineTrivia) Then
                If previousTrivia.IsKind(SyntaxKind.EndOfLineTrivia) Then
                    Return True
                End If
            End If

            previousTrivia = trivia
        Next

        Return False
    End Function
End Class
