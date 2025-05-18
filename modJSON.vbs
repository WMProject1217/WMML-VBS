Option Explicit

Const JSONErrCode = 2000

Private Function IsSpace(Char)
    IsSpace = (Char = vbCr) Or (Char = vbLf) Or (Char = vbTab) Or (Char = " ")
End Function

Private Function GetPositionString(Ctx)
    GetPositionString = "line " & Ctx("LineNo") & " column " & Ctx("Column")
End Function

Private Function IsEndOfString(Ctx)
    IsEndOfString = Ctx("I") > Ctx("Length")
End Function

Private Function PeekChar(Ctx)
    If Ctx("I") <= Ctx("Length") Then
        PeekChar = Mid(Ctx("JSONString"), Ctx("I"), 1)
    Else
        PeekChar = ""
    End If
End Function

Private Sub SkipChar(Ctx, PeekedChar)
    Ctx("I") = Ctx("I") + 1
    If PeekedChar = vbLf Then
        Ctx("LineNo") = Ctx("LineNo") + 1
        Ctx("Column") = 1
    Else
        Ctx("Column") = Ctx("Column") + 1
    End If
End Sub

Private Function GetChar(Ctx)
    GetChar = PeekChar(Ctx)
    SkipChar Ctx, GetChar
End Function

Private Sub SkipSpaces(Ctx)
    Dim CurChar
    Do
        CurChar = PeekChar(Ctx)
        If Not IsSpace(CurChar) Then Exit Do
        SkipChar Ctx, CurChar
    Loop
End Sub

Private Function HexCharToVal(HexCharAsc)
    If HexCharAsc >= &H30 And HexCharAsc <= &H39 Then
        HexCharToVal = HexCharAsc - &H30
    ElseIf HexCharAsc >= &H41 And HexCharAsc <= &H46 Then
        HexCharToVal = HexCharAsc - &H41 + 10
    ElseIf HexCharAsc >= &H61 And HexCharAsc <= &H66 Then
        HexCharToVal = HexCharAsc - &H61 + 10
    Else
        Err.Raise JSONErrCode, "JSON Parser"
    End If
End Function

Private Function ParseString(Ctx, IsObjectKey)
    Dim CurChar, Escape, EscapeHex, HexNumDigits, HexVal
    Dim StartLineNo, StartColumn
    
    StartLineNo = Ctx("LineNo")
    StartColumn = Ctx("Column") - 1
    
    Do
        CurChar = GetChar(Ctx)
        If Len(CurChar) = 0 Then Err.Raise JSONErrCode, "JSON Parser", "Unterminated string starting at line " & StartLineNo & " column " & StartColumn
        If Escape Then
            If EscapeHex Then
                HexVal = HexVal * &H10 + HexCharToVal(Asc(CurChar))
                HexNumDigits = HexNumDigits + 1
                If HexNumDigits = 4 Then
                    If IsObjectKey And HexVal < &H20 Then Err.Raise JSONErrCode, "JSON Parser", "Invalid control character at " & GetPositionString(Ctx)
                    ParseString = ParseString & Chr(HexVal)
                    EscapeHex = False
                    Escape = False
                End If
            Else
                Escape = False
                Select Case CurChar
                    Case """": ParseString = ParseString & CurChar
                    Case "\": ParseString = ParseString & CurChar
                    Case "/": ParseString = ParseString & CurChar
                    Case "b": ParseString = ParseString & vbBack
                    Case "f": ParseString = ParseString & vbFormFeed
                    Case "n": ParseString = ParseString & vbLf
                    Case "r": ParseString = ParseString & vbCr
                    Case "t": ParseString = ParseString & vbTab
                    Case "u"
                        Escape = True
                        EscapeHex = True
                        HexNumDigits = 0
                        HexVal = 0
                        Err.Description = "Invalid \uXXXX escape at " & GetPositionString(Ctx)
                    Case Else
                        Err.Raise JSONErrCode, "JSON Parser", "Invalid \escape at " & GetPositionString(Ctx)
                End Select
            End If
        Else
            If IsObjectKey And Asc(CurChar) < &H20 Then Err.Raise JSONErrCode, "JSON Parser", "Invalid control character at " & GetPositionString(Ctx)
            If CurChar = "\" Then
                Escape = True
            ElseIf CurChar = """" Then
                Exit Do
            Else
                ParseString = ParseString & CurChar
            End If
        End If
    Loop
End Function

Private Function GetNumeric(Ctx)
    Dim CurChar
    Do
        CurChar = PeekChar(Ctx)
        If IsNumeric(CurChar) Then
            GetNumeric = GetNumeric & CurChar
            SkipChar Ctx, CurChar
        Else
            Exit Do
        End If
    Loop
    If Len(GetNumeric) = 0 Then Err.Raise JSONErrCode, "JSON Parser", "Expecting value at " & GetPositionString(Ctx)
End Function

Private Function NumericToInteger(Numeric)
    On Error Resume Next
    NumericToInteger = CLng(Numeric)
    If Err.Number <> 0 Then
        Err.Clear
        NumericToInteger = CCur(Numeric)
    End If
End Function

Private Function ParseNumber(Ctx, FirstChar)
    Dim IsSigned, NumberString, CurChar, IsSignedExp
    
    If FirstChar = "-" Then
        IsSigned = True
        SkipChar Ctx, FirstChar
    End If
    
    NumberString = GetNumeric(Ctx)
    ParseNumber = NumericToInteger(NumberString)
    
    CurChar = PeekChar(Ctx)
    If CurChar = "." Then
        SkipChar Ctx, CurChar
        NumberString = GetNumeric(Ctx)
        ParseNumber = CDbl(ParseNumber) + CDbl(NumberString) / (10 ^ Len(NumberString))
    End If
    
    CurChar = PeekChar(Ctx)
    If LCase(CurChar) = "e" Then
        SkipChar Ctx, CurChar
        CurChar = PeekChar(Ctx)
        If CurChar = "-" Then
            SkipChar Ctx, CurChar
            IsSignedExp = True
        End If
        NumberString = GetNumeric(Ctx)
        If IsSignedExp Then NumberString = "-" & NumberString
        ParseNumber = CDbl(ParseNumber) * (10 ^ CDbl(NumberString))
    End If
    
    If IsSigned Then ParseNumber = -ParseNumber
End Function

Function IsEmptyArray(TestArray)
    On Error Resume Next
    Dim U
    U = UBound(TestArray)
    If Err.Number = 0 Then
        IsEmptyArray = False
    Else
        IsEmptyArray = True
    End If
End Function

Function ParseList(Ctx)
    Dim CurChar
    Dim RetList()
    Dim ItemCount
    
    SkipSpaces Ctx
    CurChar = PeekChar(Ctx)
    If CurChar = "]" Then
        SkipChar Ctx, CurChar
        ParseList = RetList
        Exit Function
    End If
    
    ReDim RetList(8)
    Do
        ParseSubString Ctx, RetList(ItemCount)
        ItemCount = ItemCount + 1
        If ItemCount >= UBound(RetList) + 1 Then 
            ReDim Preserve RetList(ItemCount * 3 / 2 + 1)
        End If
        
        SkipSpaces Ctx
        CurChar = PeekChar(Ctx)
        If CurChar = "]" Then
            SkipChar Ctx, CurChar
            If ItemCount Then
                ReDim Preserve RetList(ItemCount - 1)
            Else
                Erase RetList
            End If
            ParseList = RetList
            Exit Function
        ElseIf CurChar = "," Then
            SkipChar Ctx, CurChar
        Else
            Err.Raise JSONErrCode, "JSON Parser", "Unexpected `" & CurChar & "` at " & GetPositionString(Ctx)
        End If
    Loop
End Function

Private Function ParseObject(Ctx)
    Dim JObject, SubItem, CurChar, KeyName, IsNotFirst
    
    Set JObject = CreateObject("Scripting.Dictionary")
    
    SkipSpaces Ctx
    CurChar = PeekChar(Ctx)
    If CurChar = "}" Then
        SkipChar Ctx, CurChar
        Set ParseObject = JObject
        Exit Function
    End If
    
    IsNotFirst = False
    
    Do
        CurChar = PeekChar(Ctx)
        If CurChar = """" Then
            SkipChar Ctx, CurChar
            KeyName = ParseString(Ctx, True)
        ElseIf CurChar = "'" Then
            Err.Raise JSONErrCode, "JSON Parser", "Expecting property name enclosed in double quotes at " & GetPositionString(Ctx)
        Else
            Err.Raise JSONErrCode, "JSON Parser", "Key name must be string at " & GetPositionString(Ctx)
        End If
        
        SkipSpaces Ctx
        CurChar = PeekChar(Ctx)
        If CurChar <> ":" Then Err.Raise JSONErrCode, "JSON Parser", "Expecting ':' delimiter at " & GetPositionString(Ctx)
        SkipChar Ctx, CurChar
        SkipSpaces Ctx
        ParseSubString Ctx, SubItem
        JObject.Add KeyName, SubItem
        
        SkipSpaces Ctx
        CurChar = PeekChar(Ctx)
        If CurChar = "}" Then
            SkipChar Ctx, CurChar
            Exit Do
        ElseIf CurChar = "," Then
            SkipChar Ctx, CurChar
            SkipSpaces Ctx
        Else
            Err.Raise JSONErrCode, "JSON Parser", "Expecting ',' delimiter at " & GetPositionString(Ctx)
        End If
    Loop
    
    Set ParseObject = JObject
End Function

Private Function ParseBoolean(Ctx, ExpectedValue)
    Dim CurChar, Word, ExpectedWord, I
    
    If ExpectedValue = False Then
        ExpectedWord = "false"
    Else
        ExpectedWord = "true"
    End If
    
    For I = 1 To Len(ExpectedWord)
        CurChar = GetChar(Ctx)
        If Len(CurChar) Then 
            Word = Word & CurChar 
        Else 
            Err.Raise JSONErrCode, "JSON Parser", "Expecting value at " & GetPositionString(Ctx)
        End If
    Next
    
    If Word = ExpectedWord Then
        ParseBoolean = ExpectedValue
    Else
        Err.Raise JSONErrCode, "JSON Parser", "Unknown identifier `" & Word & "` at " & GetPositionString(Ctx)
    End If
End Function

Private Sub ParseSubString(Ctx, outParsed)
    SkipSpaces Ctx
    If IsEndOfString(Ctx) Then Err.Raise JSONErrCode, "JSON Parser", "Expecting value at " & GetPositionString(Ctx)
    
    Dim CurChar
    CurChar = PeekChar(Ctx)
    
    If CurChar = """" Then
        SkipChar Ctx, CurChar
        outParsed = ParseString(Ctx, True)
    ElseIf IsNumeric(CurChar) Or CurChar = "-" Then
        outParsed = ParseNumber(Ctx, CurChar)
    ElseIf CurChar = "[" Then
        SkipChar Ctx, CurChar
        outParsed = ParseList(Ctx)
    ElseIf CurChar = "{" Then
        SkipChar Ctx, CurChar
        Set outParsed = ParseObject(Ctx)
    ElseIf CurChar = "t" Then
        outParsed = ParseBoolean(Ctx, True)
    ElseIf CurChar = "f" Then
        outParsed = ParseBoolean(Ctx, False)
    Else
        Err.Raise JSONErrCode, "JSON Parser", "Unexpected `" & CurChar & "` at " & GetPositionString(Ctx)
    End If
End Sub

Private Function NewParserContext(JSONString)
    Set NewParserContext = CreateObject("Scripting.Dictionary")
    NewParserContext.Add "JSONString", JSONString
    NewParserContext.Add "I", 1
    NewParserContext.Add "Length", Len(JSONString)
    NewParserContext.Add "LineNo", 1
    NewParserContext.Add "Column", 1
End Function

Function ParseJSONString(JSONString)
    Dim Ctx
    Set Ctx = NewParserContext(JSONString)
    ParseSubString Ctx, ParseJSONString
    SkipSpaces Ctx
    If Not IsEndOfString(Ctx) Then Err.Raise JSONErrCode, "JSON Parser", "Extra data at " & GetPositionString(Ctx)
End Function

Sub ParseJSONString2(JSONString, ReturnParsed)
    Dim Ctx
    Set Ctx = NewParserContext(JSONString)
    ParseSubString Ctx, ReturnParsed
    SkipSpaces Ctx
    If Not IsEndOfString(Ctx) Then Err.Raise JSONErrCode, "JSON Parser", "Extra data at " & GetPositionString(Ctx)
End Sub

Private Function Hex4(Value)
    Hex4 = Right("000" & Hex(Value), 4)
End Function

Private Function EscapeString(SourceStr)
    Dim I, EI, CurChar, CharCode, ToAppend
    
    EI = Len(SourceStr)
    For I = 1 To EI
        CurChar = Mid(SourceStr, I, 1)
        CharCode = Asc(CurChar)
        Select Case CharCode
            Case 0
                ToAppend = "\0"
            Case 1, 2, 3, 4, 5, 6, 7
                ToAppend = "\u" & Hex4(CharCode)
            Case 11
                ToAppend = "\u" & Hex4(CharCode)
            Case 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31
                ToAppend = "\u" & Hex4(CharCode)
            Case 8
                ToAppend = "\b"
            Case 9
                ToAppend = "\t"
            Case 10
                ToAppend = "\n"
            Case 12
                ToAppend = "\f"
            Case 13
                ToAppend = "\r"
            Case 92
                ToAppend = "\\"
            Case Else
                ToAppend = CurChar
        End Select
        EscapeString = EscapeString & ToAppend
    Next
End Function

Function JSONToString(JSONData, Indent, IndentChar, CurIndentLevel)
    If IsArray(JSONData) Then
        If IsEmptyArray(JSONData) Then
            JSONToString = "[]"
            Exit Function
        End If
        
        Dim I, U
        U = UBound(JSONData)
        JSONToString = "["
        
        If IsEmpty(CurIndentLevel) Then CurIndentLevel = 0
        If IsEmpty(Indent) Then Indent = 0
        If IsEmpty(IndentChar) Then IndentChar = " "
        
        CurIndentLevel = CurIndentLevel + 1
        
        If Indent > 0 Then 
            JSONToString = JSONToString & vbCrLf & String(Indent * CurIndentLevel, IndentChar)
        End If
        
        For I = 0 To U
            JSONToString = JSONToString & JSONToString(JSONData(I), Indent, IndentChar, CurIndentLevel)
            If I <> U Then
                JSONToString = JSONToString & ","
                If Indent > 0 Then 
                    JSONToString = JSONToString & vbCrLf & String(Indent * CurIndentLevel, IndentChar)
                End If
            End If
        Next
        
        CurIndentLevel = CurIndentLevel - 1
        If Indent > 0 Then 
            JSONToString = JSONToString & vbCrLf & String(Indent * CurIndentLevel, IndentChar)
        End If
        
        JSONToString = JSONToString & "]"
    ElseIf TypeName(JSONData) = "Dictionary" Then
        Dim Key, IsNotFirst
        If JSONData.Count = 0 Then
            JSONToString = "{}"
            Exit Function
        End If
        
        JSONToString = "{"
        
        If IsEmpty(CurIndentLevel) Then CurIndentLevel = 0
        If IsEmpty(Indent) Then Indent = 0
        If IsEmpty(IndentChar) Then IndentChar = " "
        
        If Indent > 0 Then 
            JSONToString = JSONToString & vbCrLf & String(Indent * (CurIndentLevel + 1), IndentChar)
        End If
        
        IsNotFirst = False
        For Each Key In JSONData.Keys
            If IsNotFirst Then
                JSONToString = JSONToString & ","
                If Indent > 0 Then 
                    JSONToString = JSONToString & vbCrLf & String(Indent * (CurIndentLevel + 1), IndentChar)
                End If
            End If
            JSONToString = JSONToString & """" & Key & """: " & JSONToString(JSONData(Key), Indent, IndentChar, CurIndentLevel + 1)
            IsNotFirst = True
        Next
        
        If Indent > 0 Then 
            JSONToString = JSONToString & vbCrLf & String(Indent * CurIndentLevel, IndentChar)
        End If
        
        JSONToString = JSONToString & "}"
    Else
        Select Case VarType(JSONData)
            Case vbString
                JSONToString = """" & EscapeString(JSONData) & """"
            Case vbBoolean
                If JSONData Then
                    JSONToString = "true"
                Else
                    JSONToString = "false"
                End If
            Case vbNull
                JSONToString = "null"
            Case Else
                If IsNumeric(JSONData) Then
                    JSONToString = CStr(JSONData)
                    If Left(JSONToString, 1) = "." Then
                        JSONToString = "0" & JSONToString
                    Else
                        JSONToString = Replace(JSONToString, "-.", "-0.")
                    End If
                    JSONToString = Replace(LCase(JSONToString), "e+", "e")
                Else
                    Err.Raise JSONErrCode, "JSON Parser", "Unknown variant type `" & VarType(JSONData) & "`"
                End If
        End Select
    End If
End Function