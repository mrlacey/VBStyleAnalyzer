Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
<NoCodeFix("Creating a fix for this is low priority")>
Public Class PropertyDeclarationBracesAnalyzer
    Inherits DiagnosticAnalyzer

    Public Const DiagnosticId = "VBSAM002"

    Public Shared ReadOnly Title As String = "Do not include braces in property declaration"
    Public Shared ReadOnly MessageFormat As String = "The property declaration includes braces but it should not."
    Public Shared ReadOnly Description As String = "Do not include braces in property  the declaration."

    Private Shared Rule As New DiagnosticDescriptor(DiagnosticId, Title, MessageFormat, Categories.Misc, defaultSeverity:=DiagnosticSeverity.Warning, isEnabledByDefault:=True, description:=Description)

    Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
        Get
            Return ImmutableArray.Create(Rule)
        End Get
    End Property

    Public Overrides Sub Initialize(context As AnalysisContext)
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None)
        context.RegisterSymbolAction(AddressOf AnalyzePropertySymbol, SymbolKind.Property)
    End Sub

    Private Sub AnalyzePropertySymbol(context As SymbolAnalysisContext)
        For Each loc In context.Symbol.Locations
            Dim charsAfterPropertyName As String

            If context.Symbol.Name.Equals("ReadOnly") Then
                Dim lineAfterReadOnlyKeyword = loc.SourceTree.ToString().Substring(loc.SourceSpan.End).TrimStart()
                Dim indexOfNextSpace = lineAfterReadOnlyKeyword.IndexOf(" "C)
                charsAfterPropertyName = lineAfterReadOnlyKeyword.Substring(indexOfNextSpace - 2, 2)
            Else
                charsAfterPropertyName = loc.SourceTree.ToString().Substring(loc.SourceSpan.End, 2)
            End If
            
            If charsAfterPropertyName.Equals("()") Then
                Dim diag = Diagnostic.Create(Rule, loc)
                context.ReportDiagnostic(diag)
            End If
        Next
    End Sub
End Class
