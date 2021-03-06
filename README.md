# VBStyleAnalyzer

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A set of Roslyn code analyzers, specifically aimed at enforcing [Visual Basic Coding Conventions](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/program-structure/coding-conventions) in UWP apps.  
Created for use to confirm consistency in the projects generated by [Windows Template Studio](https://github.com/Microsoft/WindowsTemplateStudio). (WTS also uses other syntax tests to verify code structure, layout, and formatting that are not available as Analyzers.)  
It is intended to be used in companion with [SonarAnalyzer.VisualBasic](https://www.nuget.org/packages/SonarAnalyzer.VisualBasic/).

## Rules

The following rules have been / will be created.

### Naming

- [x] **VBSAN001** Do not use "my" at the start of a name.

### Layout

- [x] **VBSAL001** Use only one statement per line. Don't use the Visual Basic line separator character (:).
- [x] **VBSAL002** Add **exactly** one blank line between method and property definitions.

### Commenting

- [x] **VBSAC001** Put comments on a separate line instead of at the end of a line of code.
- [x] **VBSAC002** Start comment text with an uppercase letter, and end comment text with a period.
- [x] **VBSAC003** Insert one space between the comment delimiter (') and the comment text.
- [x] **VBSAC004** Do not surround comments with formatted blocks of asterisks.

### Misc

- [x] **VBSAM001** Use meaningful names for LINQ query variables.
- [x] **VBSAM002** Do not include braces in property declaration.
- [ ] (?) Use the IsNot keyword instead of Not...Is Nothing.
- [ ] (?) New Keyword: Use short instantiation.
- [ ] (?) Use Handles rather than AddHandler.

(?) = Possible additions
