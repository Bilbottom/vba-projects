Attribute VB_Name = "modFunctions"
Option Explicit
Option Private Module


Public Function GetUserPath() As String
    Let GetUserPath = "C:\Users\" & Environ("username") & "\"
End Function


Public Function IsLoaded(FormName As String) As Boolean
    '''
    ' Check is a Form is loaded or not.
    '
    ' https://stackoverflow.com/a/50563889/8213085
    '''
    Dim Frm As Object

    For Each Frm In VBA.UserForms
        If Frm.Name = FormName Then
            Let IsLoaded = True
            Exit Function
        End If
    Next Frm

    Let IsLoaded = False
End Function


Public Function IsInArray(StringToFind As String, vArray As Variant) As Boolean
    '''
    ' Check whether `StringToFind` is within `vArray`. VBA equivalent of the SQL
    ' `IN` operator.
    '
    ' https://stackoverflow.com/a/38268261/8213085
    '''
    Dim i As Long

    For i = LBound(vArray) To UBound(vArray)
        If vArray(i) = StringToFind Then
            Let IsInArray = True
            Exit Function
        End If
    Next i

    Let IsInArray = False
End Function


Private Sub PrintArrayEntries(vArray As Variant)
    '''
    ' Print the contents of an array.
    '''
    Dim i As Long

    For i = LBound(vArray) To UBound(vArray)
        Debug.Print vArray(i)
    Next i
End Sub


Public Function UniqueStr(sStr As String, sDelim As String) As String
    '''
    ' Return only the unique elements in a string.
    '
    ' This expects the string `sStr` to be a list of values delimited by `sDelim`.
    '
    ' https://superuser.com/a/1505673
    '''
    Dim dic      As Object
    Dim strArr() As String
    Dim strPart  As Variant
    Dim temp     As String
    Dim key      As Variant

    Set dic = CreateObject("Scripting.Dictionary")
    Let strArr = Split(sStr, sDelim)
    Let temp = ""

    For Each strPart In strArr
        On Error Resume Next
            dic.Add Trim(strPart), Trim(strPart)
        On Error GoTo 0
    Next strPart

    For Each key In dic
        Let temp = temp & key & sDelim
    Next key

    Let UniqueStr = Left(temp, Len(temp) - Len(sDelim))
End Function


Public Function SortStr(ByVal sStr As String, ByVal sDelim As String) As String
    '''
    ' Returns a sorted list of elements from a string as a string.
    '
    ' This expects the string `sStr` to be a list of values delimited by `sDelim`.
    '
    ' https://www.get-digital-help.com/sort-values-in-a-cell-using-a-custom-delimiter-vba/
    '''
    Dim vItems    As Variant
    Dim MaxVal    As Variant
    Dim MaxIndex  As Integer
    Dim i         As Integer
    Dim j         As Integer

    Let vItems = Split(sStr, sDelim)

    For i = UBound(vItems) To 0 Step -1
        Let MaxVal = vItems(i)
        Let MaxIndex = i

        For j = 0 To i
            If vItems(j) > MaxVal Then
                Let MaxVal = vItems(j)
                Let MaxIndex = j
            End If
        Next j

        If MaxIndex < i Then
            vItems(MaxIndex) = vItems(i)
            vItems(i) = MaxVal
        End If
    Next i

    Let SortStr = Join(vItems, sDelim)
End Function


Public Function RegexpReplace( _
    ByVal sInput As String, _
    ByVal sPattern As String, _
    Optional ByVal sReplace As String = "", _
    Optional ByVal bGlobal As Boolean = True, _
    Optional ByVal bIgnoreCase As Boolean = False _
) As String
    '''
    ' VBA implementation of the common Regex 'replace' function.
    '
    ' Requires Microsoft VBScript Regular Expressions 5.5
    '
    ' bGlobal: True returns all sPattern matches, False returns the first
    '''

    Dim RegEx As New RegExp

    With RegEx
        .Global = bGlobal
        .IgnoreCase = bIgnoreCase
        .Pattern = sPattern
    End With

    If RegEx.Test(sInput) Then
        Let RegexpReplace = RegEx.Replace(sInput, sReplace)
    Else
        Let RegexpReplace = sInput
    End If

    Set RegEx = Nothing
End Function


Public Function RegexpMatch( _
    ByVal sInput As String, _
    ByVal sPattern As String, _
    Optional ByVal bIgnoreCase As Boolean = False _
) As Boolean
    '''
    ' VBA implementation of the common Regex 'match' function.
    '
    ' Requires Microsoft VBScript Regular Expressions 5.5
    '''

    Dim RegEx As New RegExp

    With RegEx
        .IgnoreCase = bIgnoreCase
        .Pattern = sPattern
    End With

    Let RegexpMatch = RegEx.Test(sInput)
    Set RegEx = Nothing
End Function


Public Function EncodeBase64(ByVal sText As String) As String
    '''
    ' https://stackoverflow.com/a/169945
    '''
    Dim arrData() As Byte
    Dim objXML    As MSXML2.DOMDocument
    Dim objNode   As MSXML2.IXMLDOMElement

    Let arrData = StrConv(sText, vbFromUnicode)
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData

    Let EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing
End Function
