Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Json
Imports System.Text

Namespace CreoRegenAppVB
    Friend NotInheritable Class BomRuleLimits
        Friend Const MaxRuleLength As Integer = 1000
        Friend Const MaxNestingDepth As Integer = 64
        Friend Const MaxRulesFileBytes As Long = 5L * 1024L * 1024L
        Friend Const MaxFeaturePathLength As Integer = 400
        Friend Const MaxRuleEntries As Integer = 20000

        Private Sub New()
        End Sub
    End Class

    Friend Class BomRuleException
        Inherits Exception

        Friend Sub New(message As String)
            MyBase.New(message)
        End Sub
    End Class

#Region "Tokenizer"

    Friend Enum RuleTokenKind
        Identifier
        Number
        Text
        Op
        EndOfInput
    End Enum

    Friend Class RuleToken
        Friend Property Kind As RuleTokenKind
        Friend Property Value As String
        Friend Property Position As Integer
    End Class

    Friend NotInheritable Class RuleTokenizer
        Private Shared ReadOnly TwoCharOps As String() = {"==", "!=", "<=", ">=", "&&", "||"}
        Private Const SingleCharOps As String = "<>+-*/()=;"

        Private Sub New()
        End Sub

        Friend Shared Function Tokenize(source As String) As List(Of RuleToken)
            If source Is Nothing Then
                source = String.Empty
            End If

            If source.Length > BomRuleLimits.MaxRuleLength Then
                Throw New BomRuleException("Rule exceeds " & BomRuleLimits.MaxRuleLength.ToString(CultureInfo.InvariantCulture) & " characters.")
            End If

            Dim tokens As New List(Of RuleToken)()
            Dim i = 0
            While i < source.Length
                Dim ch = source(i)

                If Char.IsWhiteSpace(ch) Then
                    i += 1
                    Continue While
                End If

                If Char.IsLetter(ch) OrElse ch = "_"c Then
                    Dim start = i
                    While i < source.Length AndAlso (Char.IsLetterOrDigit(source(i)) OrElse source(i) = "_"c)
                        i += 1
                    End While
                    tokens.Add(New RuleToken With {.Kind = RuleTokenKind.Identifier, .Value = source.Substring(start, i - start), .Position = start})
                    Continue While
                End If

                If Char.IsDigit(ch) OrElse (ch = "."c AndAlso i + 1 < source.Length AndAlso Char.IsDigit(source(i + 1))) Then
                    Dim start = i
                    While i < source.Length AndAlso Char.IsDigit(source(i))
                        i += 1
                    End While
                    If i < source.Length AndAlso source(i) = "."c Then
                        i += 1
                        While i < source.Length AndAlso Char.IsDigit(source(i))
                            i += 1
                        End While
                    End If
                    tokens.Add(New RuleToken With {.Kind = RuleTokenKind.Number, .Value = source.Substring(start, i - start), .Position = start})
                    Continue While
                End If

                If ch = "'"c OrElse ch = """"c Then
                    Dim start = i
                    Dim closing = source.IndexOf(ch, i + 1)
                    If closing < 0 Then
                        Throw New BomRuleException("Unterminated text literal at position " & (start + 1).ToString(CultureInfo.InvariantCulture) & ".")
                    End If
                    tokens.Add(New RuleToken With {.Kind = RuleTokenKind.Text, .Value = source.Substring(i + 1, closing - i - 1), .Position = start})
                    i = closing + 1
                    Continue While
                End If

                If i + 1 < source.Length Then
                    Dim pair = source.Substring(i, 2)
                    If Array.IndexOf(TwoCharOps, pair) >= 0 Then
                        tokens.Add(New RuleToken With {.Kind = RuleTokenKind.Op, .Value = pair, .Position = i})
                        i += 2
                        Continue While
                    End If
                End If

                If SingleCharOps.IndexOf(ch) >= 0 Then
                    tokens.Add(New RuleToken With {.Kind = RuleTokenKind.Op, .Value = ch.ToString(), .Position = i})
                    i += 1
                    Continue While
                End If

                Throw New BomRuleException("Unexpected character '" & ch & "' at position " & (i + 1).ToString(CultureInfo.InvariantCulture) & ".")
            End While

            tokens.Add(New RuleToken With {.Kind = RuleTokenKind.EndOfInput, .Value = String.Empty, .Position = source.Length})
            Return tokens
        End Function
    End Class

#End Region

#Region "Values and AST"

    Friend Enum RuleValueKind
        Number
        Text
        Bool
    End Enum

    Friend Class RuleValue
        Friend Property Kind As RuleValueKind
        Friend Property Num As Double
        Friend Property TextValue As String
        Friend Property Bool As Boolean

        Friend Shared Function FromNumber(value As Double) As RuleValue
            Return New RuleValue With {.Kind = RuleValueKind.Number, .Num = value}
        End Function

        Friend Shared Function FromText(value As String) As RuleValue
            Return New RuleValue With {.Kind = RuleValueKind.Text, .TextValue = If(value, String.Empty)}
        End Function

        Friend Shared Function FromBool(value As Boolean) As RuleValue
            Return New RuleValue With {.Kind = RuleValueKind.Bool, .Bool = value}
        End Function

        ' Parameter values arrive as text; numeric-looking text is treated as a number.
        Friend Shared Function FromParameterText(value As String) As RuleValue
            Dim trimmed = If(value, String.Empty).Trim()
            Dim parsed As Double
            If trimmed.Length > 0 AndAlso Double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, parsed) Then
                Return FromNumber(parsed)
            End If

            Return FromText(If(value, String.Empty))
        End Function

        Friend Function ToParameterText() As String
            Select Case Kind
                Case RuleValueKind.Number
                    Return Num.ToString("R", CultureInfo.InvariantCulture)
                Case RuleValueKind.Bool
                    Return If(Bool, "True", "False")
                Case Else
                    Return TextValue
            End Select
        End Function

        Friend Function Describe() As String
            Select Case Kind
                Case RuleValueKind.Number
                    Return "number"
                Case RuleValueKind.Bool
                    Return "true/false"
                Case Else
                    Return "text"
            End Select
        End Function
    End Class

    Friend MustInherit Class RuleNode
        Friend MustOverride Function Evaluate(scope As IDictionary(Of String, String)) As RuleValue
    End Class

    Friend Class LiteralNode
        Inherits RuleNode

        Private ReadOnly _value As RuleValue

        Friend Sub New(value As RuleValue)
            _value = value
        End Sub

        Friend Overrides Function Evaluate(scope As IDictionary(Of String, String)) As RuleValue
            Return _value
        End Function
    End Class

    Friend Class IdentifierNode
        Inherits RuleNode

        Private ReadOnly _name As String

        Friend Sub New(name As String)
            _name = name
        End Sub

        Friend Overrides Function Evaluate(scope As IDictionary(Of String, String)) As RuleValue
            Dim raw As String = Nothing
            If scope Is Nothing OrElse Not scope.TryGetValue(_name, raw) Then
                Throw New BomRuleException("Unknown parameter '" & _name & "' (not available from parent).")
            End If

            Return RuleValue.FromParameterText(raw)
        End Function
    End Class

    Friend Class NegateNode
        Inherits RuleNode

        Private ReadOnly _operand As RuleNode

        Friend Sub New(operand As RuleNode)
            _operand = operand
        End Sub

        Friend Overrides Function Evaluate(scope As IDictionary(Of String, String)) As RuleValue
            Dim value = _operand.Evaluate(scope)
            If value.Kind <> RuleValueKind.Number Then
                Throw New BomRuleException("Unary '-' requires a number, got " & value.Describe() & ".")
            End If

            Return RuleValue.FromNumber(-value.Num)
        End Function
    End Class

    Friend Class BinaryNode
        Inherits RuleNode

        Private Const Epsilon As Double = 0.000000001

        Private ReadOnly _op As String
        Private ReadOnly _left As RuleNode
        Private ReadOnly _right As RuleNode

        Friend Sub New(op As String, left As RuleNode, right As RuleNode)
            _op = op
            _left = left
            _right = right
        End Sub

        Friend Overrides Function Evaluate(scope As IDictionary(Of String, String)) As RuleValue
            Select Case _op
                Case "&&", "||"
                    Dim leftBool = RequireBool(_left.Evaluate(scope))
                    If _op = "&&" AndAlso Not leftBool Then Return RuleValue.FromBool(False)
                    If _op = "||" AndAlso leftBool Then Return RuleValue.FromBool(True)
                    Return RuleValue.FromBool(RequireBool(_right.Evaluate(scope)))
                Case "+", "-", "*", "/"
                    Return EvaluateArithmetic(_left.Evaluate(scope), _right.Evaluate(scope))
                Case Else
                    Return EvaluateComparison(_left.Evaluate(scope), _right.Evaluate(scope))
            End Select
        End Function

        Private Function RequireBool(value As RuleValue) As Boolean
            If value.Kind <> RuleValueKind.Bool Then
                Throw New BomRuleException("'" & _op & "' requires comparisons on both sides, got " & value.Describe() & ".")
            End If

            Return value.Bool
        End Function

        Private Function EvaluateArithmetic(left As RuleValue, right As RuleValue) As RuleValue
            If left.Kind <> RuleValueKind.Number OrElse right.Kind <> RuleValueKind.Number Then
                Throw New BomRuleException("'" & _op & "' requires numbers, got " & left.Describe() & " and " & right.Describe() & ".")
            End If

            Dim result As Double
            Select Case _op
                Case "+"
                    result = left.Num + right.Num
                Case "-"
                    result = left.Num - right.Num
                Case "*"
                    result = left.Num * right.Num
                Case Else
                    If right.Num = 0 Then
                        Throw New BomRuleException("Division by zero.")
                    End If
                    result = left.Num / right.Num
            End Select

            If Double.IsNaN(result) OrElse Double.IsInfinity(result) Then
                Throw New BomRuleException("Arithmetic result is out of range.")
            End If

            Return RuleValue.FromNumber(result)
        End Function

        Private Function EvaluateComparison(left As RuleValue, right As RuleValue) As RuleValue
            If left.Kind = RuleValueKind.Number AndAlso right.Kind = RuleValueKind.Number Then
                Dim tolerance = Epsilon * Math.Max(1.0, Math.Max(Math.Abs(left.Num), Math.Abs(right.Num)))
                Dim equal = Math.Abs(left.Num - right.Num) <= tolerance
                Select Case _op
                    Case "=="
                        Return RuleValue.FromBool(equal)
                    Case "!="
                        Return RuleValue.FromBool(Not equal)
                    Case "<"
                        Return RuleValue.FromBool(Not equal AndAlso left.Num < right.Num)
                    Case ">"
                        Return RuleValue.FromBool(Not equal AndAlso left.Num > right.Num)
                    Case "<="
                        Return RuleValue.FromBool(equal OrElse left.Num < right.Num)
                    Case Else
                        Return RuleValue.FromBool(equal OrElse left.Num > right.Num)
                End Select
            End If

            If _op <> "==" AndAlso _op <> "!=" Then
                Throw New BomRuleException("'" & _op & "' requires numbers, got " & left.Describe() & " and " & right.Describe() & ".")
            End If

            Dim same = String.Equals(left.ToParameterText().Trim(), right.ToParameterText().Trim(), StringComparison.OrdinalIgnoreCase)
            Return RuleValue.FromBool(If(_op = "==", same, Not same))
        End Function
    End Class

#End Region

#Region "Parser"

    Friend Class InheritanceAssignment
        Friend Property Target As String
        Friend Property Expression As RuleNode
    End Class

    Friend Class RuleParser
        Private ReadOnly _tokens As List(Of RuleToken)
        Private _index As Integer
        Private _depth As Integer

        Private Sub New(source As String)
            _tokens = RuleTokenizer.Tokenize(source)
        End Sub

        Friend Shared Function ParseCondition(source As String) As RuleNode
            Dim parser As New RuleParser(source)
            Dim node = parser.ParseOr()
            parser.ExpectEnd()
            Return node
        End Function

        ' Syntax: TARGET = <expr> [; TARGET = <expr>]...
        Friend Shared Function ParseInheritance(source As String) As List(Of InheritanceAssignment)
            Dim parser As New RuleParser(source)
            Dim assignments As New List(Of InheritanceAssignment)()
            Dim targets As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            While parser.Current.Kind <> RuleTokenKind.EndOfInput
                If parser.IsOp(";") Then
                    parser.Advance()
                    Continue While
                End If

                Dim targetToken = parser.Current
                If targetToken.Kind <> RuleTokenKind.Identifier Then
                    Throw parser.ErrorAt(targetToken, "Expected child parameter name")
                End If
                parser.Advance()

                If Not parser.IsOp("=") Then
                    Throw parser.ErrorAt(parser.Current, "Expected '=' after '" & targetToken.Value & "'")
                End If
                parser.Advance()

                If Not targets.Add(targetToken.Value) Then
                    Throw parser.ErrorAt(targetToken, "Parameter '" & targetToken.Value & "' is assigned more than once")
                End If

                assignments.Add(New InheritanceAssignment With {.Target = targetToken.Value, .Expression = parser.ParseOr()})

                If parser.IsOp(";") Then
                    parser.Advance()
                ElseIf parser.Current.Kind <> RuleTokenKind.EndOfInput Then
                    Throw parser.ErrorAt(parser.Current, "Expected ';' between assignments")
                End If
            End While

            Return assignments
        End Function

        Private ReadOnly Property Current As RuleToken
            Get
                Return _tokens(_index)
            End Get
        End Property

        Private Sub Advance()
            If _index < _tokens.Count - 1 Then
                _index += 1
            End If
        End Sub

        Private Function IsOp(value As String) As Boolean
            Return Current.Kind = RuleTokenKind.Op AndAlso Current.Value = value
        End Function

        Private Function ErrorAt(token As RuleToken, message As String) As BomRuleException
            Dim found = If(token.Kind = RuleTokenKind.EndOfInput, "end of rule", "'" & token.Value & "'")
            Return New BomRuleException(message & " at position " & (token.Position + 1).ToString(CultureInfo.InvariantCulture) & " (found " & found & ").")
        End Function

        Private Sub ExpectEnd()
            If Current.Kind <> RuleTokenKind.EndOfInput Then
                If IsOp("=") Then
                    Throw ErrorAt(Current, "Use '==' for comparison")
                End If
                Throw ErrorAt(Current, "Unexpected token")
            End If
        End Sub

        Private Function ParseOr() As RuleNode
            Dim left = ParseAnd()
            While IsOp("||")
                Advance()
                left = New BinaryNode("||", left, ParseAnd())
            End While
            Return left
        End Function

        Private Function ParseAnd() As RuleNode
            Dim left = ParseComparison()
            While IsOp("&&")
                Advance()
                left = New BinaryNode("&&", left, ParseComparison())
            End While
            Return left
        End Function

        Private Function ParseComparison() As RuleNode
            Dim left = ParseAdditive()
            If IsOp("==") OrElse IsOp("!=") OrElse IsOp("<") OrElse IsOp(">") OrElse IsOp("<=") OrElse IsOp(">=") Then
                Dim op = Current.Value
                Advance()
                Return New BinaryNode(op, left, ParseAdditive())
            End If
            Return left
        End Function

        Private Function ParseAdditive() As RuleNode
            Dim left = ParseMultiplicative()
            While IsOp("+") OrElse IsOp("-")
                Dim op = Current.Value
                Advance()
                left = New BinaryNode(op, left, ParseMultiplicative())
            End While
            Return left
        End Function

        Private Function ParseMultiplicative() As RuleNode
            Dim left = ParseUnary()
            While IsOp("*") OrElse IsOp("/")
                Dim op = Current.Value
                Advance()
                left = New BinaryNode(op, left, ParseUnary())
            End While
            Return left
        End Function

        Private Function ParseUnary() As RuleNode
            If IsOp("-") Then
                Dim token = Current
                EnterNesting(token)
                Advance()
                Dim operand = ParseUnary()
                _depth -= 1
                Return New NegateNode(operand)
            End If
            Return ParsePrimary()
        End Function

        Private Function ParsePrimary() As RuleNode
            Dim token = Current
            Select Case token.Kind
                Case RuleTokenKind.Identifier
                    Advance()
                    Return New IdentifierNode(token.Value)
                Case RuleTokenKind.Number
                    Advance()
                    Dim number As Double
                    If Not Double.TryParse(token.Value, NumberStyles.Float, CultureInfo.InvariantCulture, number) Then
                        Throw ErrorAt(token, "Invalid number")
                    End If
                    Return New LiteralNode(RuleValue.FromNumber(number))
                Case RuleTokenKind.Text
                    Advance()
                    Return New LiteralNode(RuleValue.FromText(token.Value))
            End Select

            If IsOp("(") Then
                EnterNesting(token)
                Advance()
                Dim inner = ParseOr()
                If Not IsOp(")") Then
                    Throw ErrorAt(Current, "Expected ')'")
                End If
                Advance()
                _depth -= 1
                Return inner
            End If

            Throw ErrorAt(token, "Expected parameter, number, text or '('")
        End Function

        Private Sub EnterNesting(token As RuleToken)
            _depth += 1
            If _depth > BomRuleLimits.MaxNestingDepth Then
                Throw ErrorAt(token, "Rule is nested too deeply")
            End If
        End Sub
    End Class

#End Region

#Region "Evaluator"

    Friend Class BomConfigSummary
        Friend Property IncludedCount As Integer
        Friend Property ExcludedCount As Integer
        Friend Property ErrorCount As Integer
    End Class

    Friend NotInheritable Class BomConfigEvaluator
        Private Sub New()
        End Sub

        ' Rows must be in pre-order (parent before children), as produced by FlattenBomNode.
        Friend Shared Function Evaluate(rows As IList(Of BomGridRowItem), rootParams As IDictionary(Of String, String)) As BomConfigSummary
            Dim summary As New BomConfigSummary()
            Dim scopes As New Dictionary(Of Integer, Dictionary(Of String, String))()
            Dim excludedNodeIds As New HashSet(Of Integer)()
            Dim rootScope As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If rootParams IsNot Nothing Then
                For Each pair In rootParams
                    rootScope(pair.Key) = pair.Value
                Next
            End If

            For Each row In rows
                row.IsExcluded = False
                row.ConditionError = Nothing
                row.InheritanceError = Nothing
                row.ResolvedInheritance = Nothing
                row.ResolvedAssignments = Nothing

                If row.ParentNodeId = 0 Then
                    scopes(row.NodeId) = rootScope
                    summary.IncludedCount += 1
                    Continue For
                End If

                Dim parentScope As Dictionary(Of String, String) = Nothing
                If Not scopes.TryGetValue(row.ParentNodeId, parentScope) Then
                    parentScope = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                End If

                Dim conditionMet = True
                If Not String.IsNullOrWhiteSpace(row.Conditions) Then
                    Try
                        Dim result = RuleParser.ParseCondition(row.Conditions).Evaluate(parentScope)
                        If result.Kind <> RuleValueKind.Bool Then
                            Throw New BomRuleException("Condition must be a comparison (e.g. WIDTH > 500), got " & result.Describe() & ".")
                        End If
                        conditionMet = result.Bool
                    Catch ex As BomRuleException
                        row.ConditionError = ex.Message
                    End Try
                End If

                row.IsExcluded = excludedNodeIds.Contains(row.ParentNodeId) OrElse Not conditionMet
                If row.IsExcluded Then
                    excludedNodeIds.Add(row.NodeId)
                End If

                Dim childScope As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                If Not String.IsNullOrWhiteSpace(row.Inheritance) Then
                    Try
                        Dim resolvedParts As New List(Of String)()
                        Dim resolvedPairs As New List(Of KeyValuePair(Of String, String))()
                        For Each assignment In RuleParser.ParseInheritance(row.Inheritance)
                            Dim resolvedValue = assignment.Expression.Evaluate(parentScope).ToParameterText()
                            childScope(assignment.Target) = resolvedValue
                            resolvedParts.Add(assignment.Target & "=" & resolvedValue)
                            resolvedPairs.Add(New KeyValuePair(Of String, String)(assignment.Target, resolvedValue))
                        Next
                        row.ResolvedInheritance = String.Join("; ", resolvedParts)
                        row.ResolvedAssignments = resolvedPairs
                    Catch ex As BomRuleException
                        row.InheritanceError = ex.Message
                        childScope.Clear()
                    End Try
                End If
                scopes(row.NodeId) = childScope

                If row.IsExcluded Then
                    summary.ExcludedCount += 1
                Else
                    summary.IncludedCount += 1
                End If

                If row.ConditionError IsNot Nothing OrElse row.InheritanceError IsNot Nothing Then
                    summary.ErrorCount += 1
                End If
            Next

            Return summary
        End Function
    End Class

#End Region

#Region "Persistence"

    <DataContract(Name:="BomRuleFile", [Namespace]:="")>
    Public Class BomRuleFile
        <DataMember(Order:=1)>
        Public Property Version As Integer

        <DataMember(Order:=2)>
        Public Property TopModel As String

        <DataMember(Order:=3)>
        Public Property Rules As List(Of BomRuleEntry)
    End Class

    <DataContract(Name:="BomRuleEntry", [Namespace]:="")>
    Public Class BomRuleEntry
        <DataMember(Order:=1)>
        Public Property FeaturePath As String

        <DataMember(Order:=2)>
        Public Property FileName As String

        <DataMember(Order:=3, EmitDefaultValue:=False)>
        Public Property Condition As String

        <DataMember(Order:=4, EmitDefaultValue:=False)>
        Public Property Inheritance As String
    End Class

    Friend NotInheritable Class BomRulesStore
        Private Const CurrentVersion As Integer = 1
        Private Const FileSuffix As String = ".bomrules.json"

        Private Sub New()
        End Sub

        Friend Shared Function BuildRulesPath(workingDirectory As String, topFileName As String) As Tuple(Of Boolean, String)
            If String.IsNullOrWhiteSpace(workingDirectory) Then
                Return Tuple.Create(False, "Creo working directory is not available.")
            End If

            Dim directoryPath As String
            Try
                directoryPath = Path.GetFullPath(workingDirectory.Trim()).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Catch
                Return Tuple.Create(False, "Creo working directory is invalid.")
            End Try

            If directoryPath.StartsWith("\\", StringComparison.Ordinal) OrElse Not Path.IsPathRooted(directoryPath) Then
                Return Tuple.Create(False, "Rules file must be in a local Creo working directory.")
            End If

            If Not Directory.Exists(directoryPath) Then
                Return Tuple.Create(False, "Creo working directory does not exist.")
            End If

            Dim baseName = Path.GetFileNameWithoutExtension(If(topFileName, String.Empty).Trim())
            If String.IsNullOrWhiteSpace(baseName) OrElse baseName.Contains("..") OrElse baseName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 Then
                Return Tuple.Create(False, "Top model name cannot be used as a rules file name.")
            End If

            Dim rulesPath = Path.GetFullPath(Path.Combine(directoryPath, baseName.ToLowerInvariant() & FileSuffix))
            If Not String.Equals(Path.GetDirectoryName(rulesPath).TrimEnd(Path.DirectorySeparatorChar), directoryPath, StringComparison.OrdinalIgnoreCase) Then
                Return Tuple.Create(False, "Rules file path is invalid.")
            End If

            Return Tuple.Create(True, rulesPath)
        End Function

        Friend Shared Function Load(rulesPath As String) As BomRuleFile
            Dim info As New FileInfo(rulesPath)
            If info.Length > BomRuleLimits.MaxRulesFileBytes Then
                Throw New BomRuleException("Rules file is larger than the 5 MB limit.")
            End If

            Dim ruleFile As BomRuleFile
            Using stream As New FileStream(rulesPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                ruleFile = TryCast(CreateSerializer().ReadObject(stream), BomRuleFile)
            End Using

            If ruleFile Is Nothing Then
                Throw New BomRuleException("Rules file is empty or invalid.")
            End If

            If ruleFile.Rules Is Nothing Then
                ruleFile.Rules = New List(Of BomRuleEntry)()
            End If

            If ruleFile.Rules.Count > BomRuleLimits.MaxRuleEntries Then
                Throw New BomRuleException("Rules file contains too many entries.")
            End If

            Return ruleFile
        End Function

        Friend Shared Sub Save(rulesPath As String, ruleFile As BomRuleFile)
            Using stream As New FileStream(rulesPath, FileMode.Create, FileAccess.Write, FileShare.None)
                Using writer = JsonReaderWriterFactory.CreateJsonWriter(stream, New UTF8Encoding(False), False, True)
                    CreateSerializer().WriteObject(writer, ruleFile)
                    writer.Flush()
                End Using
            End Using
        End Sub

        Friend Shared Function CreateFromRows(topModel As String, rows As IEnumerable(Of BomGridRowItem)) As BomRuleFile
            Dim ruleFile As New BomRuleFile With {
                .Version = CurrentVersion,
                .TopModel = topModel,
                .Rules = New List(Of BomRuleEntry)()
            }

            For Each row In rows
                If row.ParentNodeId = 0 Then
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(row.Conditions) AndAlso String.IsNullOrWhiteSpace(row.Inheritance) Then
                    Continue For
                End If

                ruleFile.Rules.Add(New BomRuleEntry With {
                    .FeaturePath = row.FeaturePath,
                    .FileName = row.FileName,
                    .Condition = NullIfBlank(row.Conditions),
                    .Inheritance = NullIfBlank(row.Inheritance)
                })
            Next

            Return ruleFile
        End Function

        ' Returns the number of applied rules; mismatches are reported through warnings.
        Friend Shared Function ApplyToRows(ruleFile As BomRuleFile, rows As IEnumerable(Of BomGridRowItem), warnings As List(Of String)) As Integer
            Dim rowsByPath As New Dictionary(Of String, BomGridRowItem)(StringComparer.Ordinal)
            For Each row In rows
                If row.ParentNodeId <> 0 AndAlso Not String.IsNullOrEmpty(row.FeaturePath) AndAlso Not rowsByPath.ContainsKey(row.FeaturePath) Then
                    rowsByPath(row.FeaturePath) = row
                End If
            Next

            Dim applied = 0
            For Each entry In ruleFile.Rules
                If entry Is Nothing Then
                    Continue For
                End If

                Dim featurePath = If(entry.FeaturePath, String.Empty).Trim()
                If Not IsValidFeaturePath(featurePath) Then
                    warnings.Add("Skipped rule with invalid feature path '" & Truncate(featurePath, 60) & "'.")
                    Continue For
                End If

                If Len(entry.Condition) > BomRuleLimits.MaxRuleLength OrElse Len(entry.Inheritance) > BomRuleLimits.MaxRuleLength Then
                    warnings.Add("Skipped rule at " & featurePath & ": rule text is too long.")
                    Continue For
                End If

                Dim row As BomGridRowItem = Nothing
                If Not rowsByPath.TryGetValue(featurePath, row) Then
                    warnings.Add("Skipped rule at " & featurePath & " (" & Truncate(entry.FileName, 60) & "): component not found in BOM.")
                    Continue For
                End If

                If Not String.Equals(If(entry.FileName, String.Empty).Trim(), If(row.FileName, String.Empty).Trim(), StringComparison.OrdinalIgnoreCase) Then
                    warnings.Add("Skipped rule at " & featurePath & ": expected " & Truncate(entry.FileName, 60) & " but BOM has " & row.FileName & ".")
                    Continue For
                End If

                row.Conditions = If(entry.Condition, String.Empty)
                row.Inheritance = If(entry.Inheritance, String.Empty)
                applied += 1
            Next

            Return applied
        End Function

        Private Shared Function CreateSerializer() As DataContractJsonSerializer
            Return New DataContractJsonSerializer(GetType(BomRuleFile), New DataContractJsonSerializerSettings With {
                .MaxItemsInObjectGraph = BomRuleLimits.MaxRuleEntries * 6 + 16
            })
        End Function

        Private Shared Function IsValidFeaturePath(featurePath As String) As Boolean
            If featurePath.Length = 0 OrElse featurePath.Length > BomRuleLimits.MaxFeaturePathLength Then
                Return False
            End If

            For Each ch In featurePath
                If Not (Char.IsDigit(ch) OrElse ch = "/"c OrElse ch = "-"c) Then
                    Return False
                End If
            Next

            Return True
        End Function

        Private Shared Function NullIfBlank(value As String) As String
            Return If(String.IsNullOrWhiteSpace(value), Nothing, value.Trim())
        End Function

        Private Shared Function Truncate(value As String, maxLength As Integer) As String
            Dim text = If(value, String.Empty)
            Return If(text.Length <= maxLength, text, text.Substring(0, maxLength) & "...")
        End Function
    End Class

#End Region
End Namespace
