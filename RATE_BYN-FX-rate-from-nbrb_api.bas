Attribute VB_Name = "Module1"
Function RATE(ByVal cur_code As String, Optional ByVal xdate As String)
Attribute RATE.VB_Description = "Retuns BYN the exchange rate of the BYN per unit of the specified currency from the nbrb.by"
Attribute RATE.VB_ProcData.VB_Invoke_Func = " \n9"
If xdate = Null Or xdate = "" Then xdate = Date

On Error GoTo Error1

Dim http As Object
Set http = CreateObject("MSXML2.XMLHTTP")

xxxdate = Format(xdate, "yyyy-m-d")
xdate_url = "https://www.nbrb.by/API/ExRates/Rates?onDate=" & xxxdate & "&Periodicity=0"
http.Open "GET", xdate_url, False
http.Send
response = Mid(http.responseText, 3, (Len(http.responseText) - 4))
txt = Replace(response, Chr(34), "")
stringi = Split(txt, "},{")

For i = 0 To UBound(stringi)
    element = Split(stringi(i), ",")
        decomposition_cname = Split(element(2), ":")
        c_name = decomposition_cname(1)
        
        If c_name = cur_code Then
            decomposition_cscale = Split(element(3), ":")
            c_scale = decomposition_cscale(1)
            decomposition_cval = Split(element(5), ":")
            c_val = decomposition_cval(1)
            norm_cval = Replace(c_val, ".", ",")
            cur = CDec(norm_cval) / CDec(c_scale)
        End If
                    
        If cur_code = "BYN" Then
            cur = 1
        End If
        
Next i
   
    If xdate < CDate("01.07.2016") Then
        RATE = cur / 10000
    Else
        RATE = cur
    End If
    
        If RATE <= 0 Or cur = "" Then RATE = "#Error!"
            If norm_cval = "" Then RATE = "Wrong ISO4217 currency code!"
                If response = "" Then RATE = "No data as of " & Format(xdate, "yyyy-m-d") & "!"
    
GoTo Ends:
Error1:
MsgBox ("Îøèáêà! Âàëþòà: " & cur_code & "| Äàòà: " & Format(xdate, "yyyy-m-d") & " | Êóðñ: " & RATE)
Ends:

End Function

Sub RegisterDescriptionRATE()

    Dim D0 As String, D1 As String, D2 As String
                    D0 = "Retuns BYN the exchange rate of the BYN per unit of the specified currency from the website of the National Bank of the Republic of Belarus"
                    D1 = "ISO4217 currency code as string , (Example: 'USD')"
                    D2 = "Date. By default - today"

    Application.MacroOptions _
        Macro:="RATE", _
            Description:=D0, _
            ArgumentDescriptions:=Array(D1, D2), _
            HasMenu:=True, _
            MenuText:="RATE"
            
End Sub

Sub UnregisterRATE()

    Application.MacroOptions _
        Macro:="RATE", _
            Description:=Empty, _
            ArgumentDescriptions:=Empty, _
            Category:=Empty, _
            HasMenu:=Empty, _
            MenuText:=Empty
    
End Sub


