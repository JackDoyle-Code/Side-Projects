Attribute VB_Name = "Module2"
Function split_string_regex(str As String, chr As String)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    Dim pattern As String
    pattern = chr & "{2,}"
    
    With regex
        .Global = True
        .IgnoreCase = True
        .pattern = pattern
    End With
    
    str = regex.Replace(str, chr)
    split_string_regex = Split(str, chr)

End Function

Function split_string_regex_num(str As String, chr As String, num As Integer)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    Dim pattern As String
    pattern = chr & "{2,}"
    
    With regex
        .Global = True
        .IgnoreCase = True
        .pattern = pattern
    End With
    
    str = regex.Replace(str, chr)
    split_string = Split(str, chr)
    split_string_regex_num = split_string(num)

End Function
