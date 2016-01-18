Attribute VB_Name = "ExcelCore"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : GetExcelApplication
' Author    : bradley_handziuk
' Date      : 3/6/2014
' Purpose   : Don't use this function if you are going to be creating then destroying an instance quickly.
'---------------------------------------------------------------------------------------
Public Function GetWordApplication(Optional ByRef WasANewInstanceReturned As Boolean) As Word.Application
 
    If WordInstanceCount > 0 Then
        Set GetWordApplication = GetObject(, "Word.Application")
        WasANewInstanceReturned = False
    Else
        Set GetWordApplication = New Word.Application
        WasANewInstanceReturned = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : InstanceCount
' Author    : bradley_handziuk
' Date      : 3/6/2014
' Purpose   :   Gets the instance count of excel applications
' http://www.mrexcel.com/forum/excel-questions/400446-visual-basic-applications-check-if-excel-already-open.html
'---------------------------------------------------------------------------------------
Function WordInstanceCount() As Integer
    Dim objList As Object, objType As Object, strObj As String
    strObj = "WINWORD.exe"
    Set objType = GetObject("winmgmts:").ExecQuery("select * from win32_process where name='" & strObj & "'")
    WordInstanceCount = objType.Count
End Function

Public Function GetReplicateNumber(rng As Range)
    Dim lastCharacter As String
    lastCharacter = Right(Split(rng, "-")(1), 1)
    Select Case lastCharacter
        Case "C":
            GetReplicateNumber = "Replicate C"
            Exit Function
        Case "B":
            GetReplicateNumber = "Replicate B"
            Exit Function
        Case Else:
            GetReplicateNumber = "Replicate A"
            Exit Function
    End Select
    
    
    
    
End Function

Public Function ConvertFractionFeetToDecimalFeet(inputFeet)
    Dim ft, inch
    Dim ftHolder, inHolder
    If inputFeet = "" Then Exit Function
    ftHolder = Split(inputFeet, "'")
    If UBound(ftHolder) > 0 Then
        If InStr(1, ftHolder(1), """") > 0 Then
            inHolder = Split(ftHolder(1), """")
            inHolder(0) = Replace(inHolder(0), "-", "+")
            inHolder(0) = Replace(inHolder(0), " ", "+")
            inHolder(0) = Evaluate(inHolder(0))
            inch = inHolder(0) / 12
        End If
    End If
    ft = ftHolder(0)
    
    
    ConvertFractionFeetToDecimalFeet = IIf(IsNumeric(ft + inch), CDbl(ft + inch), ft + inch)
End Function


Public Function GetSheetName(cell As Range)
    GetSheetName = cell.Parent.Name
End Function

'---------------------------------------------------------------------------------------
' Procedure : WordCount
' Author    : bradley_handziuk
' Date      : 1/9/2015
' Purpose   :
'---------------------------------------------------------------------------------------
Function WordCount(c)

    Dim words() As String
    words = Split(c, " ")
    WordCount = UBound(words) + 1

End Function

Public Function GetStuffInParens(rng)
    
    Dim startOfParens As Integer, endOfParens As Integer
    
    startOfParens = InStr(1, rng, "(") + 1
    endOfParens = InStr(1, rng, ")") - 1
    
    GetStuffInParens = Mid(rng, startOfParens, endOfParens - startOfParens + 1)
    

End Function

Public Function GetStuffNotInParens(rng)
    
    Dim startOfParens As Integer, endOfParens As Integer
    
    startOfParens = InStr(1, rng, "(")
    endOfParens = InStr(1, rng, ")")
    
    GetStuffNotInParens = Replace(rng, Mid(rng, startOfParens, endOfParens - startOfParens + 1), "")

End Function

'---------------------------------------------------------------------------------------
' Procedure : PunctuationCount
' Author    : bradley_handziuk
' Date      : 1/9/2015
' Purpose   :
'---------------------------------------------------------------------------------------
Function PunctuationCount(c)
    Dim commas As Integer, periods As Integer
    Dim originalLength As Integer
    originalLength = Len(c)
    commas = originalLength - Len(Replace(c, ",", ""))
    periods = originalLength - Len(Replace(c, ".", ""))

    PunctuationCount = commas + periods
End Function

'---------------------------------------------------------------------------------------
' Procedure : AnyCompare
' Author    : bradley_handziuk
' Date      : 11/21/2013
' Purpose   : Compares CompareThisValue to all the values in AgainstThisList.
'               The first value which satisfies CompareThisValue Using AgainstThisList[x]
'               Returns True. Otherwise False is returned
'               If Using is not a logical operator from this list an error is thrown (<, <=, >=, >, =)
'---------------------------------------------------------------------------------------
Function AnyCompare(CompareThisValue, AgainstThisList, Optional Using = "=")
    Dim c As Range
    If Using = "=" Or Using = ">" Or Using = ">=" Or Using = "<=" Or Using = "<" Then
        For Each c In AgainstThisList
            If c.value <> "" Then
            If Evaluate(CompareThisValue & Using & c.value) Then
                AnyCompare = True
                Exit Function
            End If
            End If
        Next c
    Else
        Err.Raise 10    ''bad input
        Exit Function
    End If
    AnyCompare = False
End Function
'---------------------------------------------------------------------------------------
' Procedure : SplitWs
' Author    : bradley_handziuk
' Date      : 1/9/2015
' Purpose   :
'---------------------------------------------------------------------------------------
Function SplitWs(rng, delimiter As String, index As Integer) As Variant
    
    Dim tempArr
    tempArr = Split(rng, delimiter)
    If UBound(tempArr) >= index Then
        SplitWs = tempArr(index)
    End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : GetRGB
' Author    : bradley_handziuk
' Date      : 4/8/2014
' Purpose   : Converts long numbers into RGB values. Printed as a string.
'---------------------------------------------------------------------------------------
Function GetRGB(color As Long) As String
    Dim R As Long, G As Long, B As Long
    R = color Mod 256
    G = (color \ 256) Mod 256
    B = (color \ 256 \ 256) Mod 256
    GetRGB = "{" & R & "," & G & "," & B & "}"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Reverse
' Author    : bradley_handziuk
' Date      : 1/21/2014
' Purpose   : Reverses the string
'---------------------------------------------------------------------------------------
Public Function Reverse(str As String) As String
    Reverse = StrReverse(str)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetSigFigs
' Author    : bradley_handziuk
' Date      : 1/21/2014
' Purpose   : Returns value with the number of significant digits specified
'---------------------------------------------------------------------------------------
Function GetSigFigs(value, sigfigs)
    If sigfigs <= 0 Then
        GetSigFigs = value
        Exit Function
    End If
    If Not IsNumeric(value) Then
        GetSigFigs = value
        Exit Function
    End If
    Dim isNegative As Boolean
    isNegative = value < 0
    value = Abs(value)
    
    Dim exp As Integer
    Dim tmp As Double

    Dim ret
   value = CDec(value)
    If Abs(value) < 1 Then
        exp = Get_First_Non_Zero_Char(value, 1) - 2
    End If
   
    exp = exp + sigfigs - Len(CStr(Fix(value)))
    tmp = value * 10 ^ (exp)
   
    tmp = Round(tmp, 0)
    ret = tmp * 10 ^ (-exp) 'calcualtes the final _number_
    If Abs(value) < 1 Then
        ret = "0." & String(exp - Len(CStr(tmp)), "0") & tmp
        
    ElseIf ret <> Fix(ret) And (Len(CStr(ret)) - Len(".")) <> sigfigs Then 'right of the AND means it is already formatted as it should be
         ret = ret & String(sigfigs - Len(CStr(ret)) + Len("."), "0")
         
    ElseIf Abs(ret) > 1 And Len(ret) < sigfigs Then    'right of the AND asks if any additional formatting is even necessary
         ret = ret & "." & String(sigfigs - Len(CStr(ret)), "0")
         
    ElseIf ret = 1 And sigfigs - 1 > 0 Then
        ret = ret & "." & String(sigfigs - 1, "0")
    End If
    
    If isNegative Then
        ret = ret * -1
    End If
    GetSigFigs = CStr(ret)
End Function



'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : bradley_handziuk
' Date      : 4/15/2014
' Purpose   :
'---------------------------------------------------------------------------------------
Function PadRight(text As Variant, totalLength As Integer, padCharacter As String) As String
    If totalLength - Len(CStr(text)) > 0 Then
        PadRight = CStr(text) & String(totalLength - Len(CStr(text)), padCharacter)
    Else
        PadRight = text
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : PadLeft
' Author    : bradley_handziuk
' Date      : 4/15/2014
' Purpose   :
'---------------------------------------------------------------------------------------
Function PadLeft(text As Variant, totalLength As Integer, padCharacter As String) As String
    If totalLength - Len(CStr(text)) > 0 Then
        PadLeft = String(totalLength - Len(CStr(text)), padCharacter) & CStr(text)
    Else
        PadLeft = text
    End If
End Function

Function CollectionToArray(c As Collection) As String()
    If c.Count = 0 Then Exit Function
    Dim a() As String: ReDim a(0 To c.Count - 1)
    Dim i As Integer
    For i = 1 To c.Count
        a(i - 1) = c.Item(i)
    Next
    CollectionToArray = a
End Function

'---------------------------------------------------------------------------------------
' Procedure : Join
' Author    : bradley_handziuk
' Date      : 1/21/2014
' Purpose   : This method will concatenate all the items in rng() together using seperator
'---------------------------------------------------------------------------------------
Public Function Join(seperator As String, ExcludeEmpty As Boolean, ParamArray params()) As String

    Dim cell As Variant, param As Variant
    Dim joinedString As String
   Dim c As Range
   
    For Each param In params
        If IsArray(param) Then
            For Each cell In param
            Set c = cell
            
                If Not (cell.EntireRow.Hidden Or cell.EntireColumn.Hidden) Then
                    If (Not ExcludeEmpty And Len(cell) = 0) Or Len(cell) > 0 Then
                        joinedString = joinedString & cell & seperator
                    End If
                End If
            Next cell
        ElseIf TypeName(param) = "Collection" Then
            Dim aNewCollection As Collection
            Set aNewCollection = param
            Join = Join(seperator, ExcludeEmpty, CollectionToArray(aNewCollection))
        Else
            If TypeName(param) = "String" Then
                joinedString = joinedString & param & seperator
            End If
        End If
    Next param

        joinedString = Left(joinedString, Len(joinedString) - Len(seperator))
        Join = joinedString
End Function

Public Function JoinStrings(seperator As String, rng As Variant, Optional ExcludeEmpty As Boolean = False) As String
   ' If rng Is Nothing Then Exit Function
    
    Dim cell As Variant
    Dim joinedString As String
    
    For Each cell In rng
        If (Not ExcludeEmpty And Len(cell) = 0) Or Len(cell) > 0 Then
            joinedString = joinedString & cell & seperator
        End If
    Next cell
    If Len(joinedString) > 0 Then
        joinedString = Left(joinedString, Len(joinedString) - Len(seperator))
    End If
    JoinStrings = joinedString
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : IsFormula
' Author    : bradley_handziuk
' Date      : 1/21/2014
' Purpose   : Returns true if the range reference's content is a formula
'---------------------------------------------------------------------------------------
Function IsFormula(rng As Range) As Boolean
    IsFormula = rng.HasFormula
End Function

'---------------------------------------------------------------------------------------
' Procedure : Get_First_Non_Zero_Char
' Author    : bradley_handziuk
' Date      : 1/21/2014
' Purpose   : Does exactly what the name says (gets first non zero/non '.' character
'---------------------------------------------------------------------------------------
Function Get_First_Non_Zero_Char(value, depth As Integer) As Integer
    'Debug.Print Mid(value, depth, 1)
    If Mid(value, depth, 1) = "0" Or Mid(value, depth, 1) = "." Then
        
        depth = Get_First_Non_Zero_Char(value, depth + 1)
    End If
    Get_First_Non_Zero_Char = depth
End Function



Private Sub ExportAllCharts()
Dim c As Chart
    For Each c In Charts
        c.Export ActiveWorkbook.path & "\" & Replace(c.Name, "_chart", "") & ".png"
        If Err.number = 76 Then
           ' MkDir ThisWorkbook.Path & "\"
            c.Export ActiveWorkbook.path & "\" & Replace(c.Name, "_chart", "") & ".png"
        End If
        On Error GoTo 0
    Next c
End Sub








'---------------------------------------------------------------------------------------
' Procedure : Convert_Decimal
' Author    : bradley_handziuk
' Date      : 4/7/2014
' Purpose   : from here http://support.microsoft.com/kb/213449 but there was a mistake in the original
'---------------------------------------------------------------------------------------
Function Convert_Decimal(Degree_Deg As String) As Double
   ' Declare the variables to be double precision floating-point.
   Dim degrees As Double
   Dim minutes As Double
   Dim seconds As Double
   ' Set degree to value before "°" of Argument Passed.
   degrees = val(Left(Degree_Deg, InStr(1, Degree_Deg, "°") - 1))
   ' Set minutes to the value between the "°" and the "'"
   ' of the text string for the variable Degree_Deg divided by
   ' 60. The Val function converts the text string to a number.
   minutes = val(Mid(Degree_Deg, InStr(1, Degree_Deg, "°") + 1, _
             InStr(1, Degree_Deg, "'") - InStr(1, Degree_Deg, _
             "°") - 1)) / 60
    ' Set seconds to the number to the right of "'" that is
    ' converted to a value and then divided by 3600.
    seconds = val(Mid(Degree_Deg, InStr(1, Degree_Deg, "'") + _
            1, Len(Degree_Deg) - InStr(1, Degree_Deg, "'") - 1)) _
            / 3600
   Convert_Decimal = degrees + minutes + seconds
End Function

'---------------------------------------------------------------------------------------
' Procedure : UnpivotData
' Author    : bradley_handziuk
' Date      : 4/4/2014
' Purpose   : This will do the same thing as making Multiple Consolidation Ranges then showing detail (alt + D, P)
'---------------------------------------------------------------------------------------
Sub UnpivotData()
    Dim OBJECT_REQUIRED_ERROR As Integer, UNABLE_TO_SHOW_DETAIL_ERROR As Integer
    OBJECT_REQUIRED_ERROR = 424
    UNABLE_TO_SHOW_DETAIL_ERROR = 1004

    Dim sourceRange As Range, sourceSheet As Worksheet
    On Error Resume Next
    Set sourceRange = Application.InputBox("Select cell(s)", Type:=8, Title:="Select range to unpivot...")
    If Err.number = OBJECT_REQUIRED_ERROR Then
        Exit Sub
    End If
    On Error GoTo 0
    Application.ScreenUpdating = False
    Dim fullRangeAddress As String
    fullRangeAddress = "'" & sourceRange.Worksheet.Name & "'!" & sourceRange.Address(ReferenceStyle:=xlR1C1, external:=False)

    Dim pvtchache  As PivotCache, newPivotTable As PivotTable
    Set pvtchache = sourceRange.Worksheet.Parent.PivotCaches.Create(SourceType:=xlConsolidation, SourceData:=Array(fullRangeAddress), Version:=xlPivotTableVersion14)
    Set newPivotTable = pvtchache.CreatePivotTable(TableDestination:="", DefaultVersion:=xlPivotTableVersion14)            'TableName:="PivotTable2",

    Dim pivotWorksheet As Worksheet, numberOfNewColumns As Integer, detailWorksheet As Worksheet

    Set pivotWorksheet = ActiveSheet  'the last function will make this the active worksheet as default behavior
    On Error Resume Next
    pivotWorksheet.Cells(newPivotTable.RowRange.Rows.Count + 1, newPivotTable.ColumnRange.Columns.Count + 1).ShowDetail = True    'go to the bottom right corner (grand grand total)
    If Err.number = UNABLE_TO_SHOW_DETAIL_ERROR Then
        MsgBox "Could not find the Grand Total cell. You'll have to double click it yourself to finish."
    End If

    Set detailWorksheet = ActiveSheet
    numberOfNewColumns = Len(detailWorksheet.Range("A2")) - Len(Replace(detailWorksheet.Range("A2"), "|", ""))

    Dim i As Integer
    Do While i < numberOfNewColumns
        detailWorksheet.Columns(2).Insert
        i = i + 1
    Loop

    Dim txtToColsRange As Range
    With detailWorksheet
        Application.DisplayAlerts = False
        .Range("A1") = sourceRange.Cells(1, 1)
        Set txtToColsRange = .Range(.Range("A1"), .Range("A1").End(xlDown))
    End With
    txtToColsRange.TextToColumns DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", _
        FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

    detailWorksheet.Cells.EntireColumn.AutoFit

    pivotWorksheet.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub MakeTextNumbers()
    '###0.00#######_);###0.00########
    Dim R As Long, c As Long
    Dim MaxRow As Long, MaxCol As Long
    Dim ws As Worksheet, rng As Range
    Set ws = ActiveSheet
    Set rng = ws.Cells(1, 1)
    MaxCol = rng.End(xlToRight).Column
    MaxRow = rng.End(xlDown).Row
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For R = 2 To MaxRow
        For c = 2 To MaxCol
            ws.Cells(R, c).Formula = ws.Cells(R, c).Formula
        Next c
    Next R
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub


    





