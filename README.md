# ExcelQuickstats
Excel with VB code to import Quickstats data using the NASS Quickstats API
Blog post contains more explaination https://www.jeffkoenig.com/quickstats-api-using-excel/
To use you will need to request an api key from NASS https://quickstats.nass.usda.gov/api

Sub NASS_QS_API()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

    Dim objRequest As Object
    Dim strUrl As String
    Dim api_key As String: api_key = Sheets("config").Range("B1")
    Dim params As String: params = Sheets("config").Range("B2")
    Dim strResponse() As String
    Dim strResponse2 As String
    Dim blnAsync As Boolean
    Dim Lines As Long
    Dim Line As Long
    
    strUrl = "https://quickstats.nass.usda.gov/api/api_GET/?"
    strUrl = strUrl & "key=" & api_key
    strUrl = strUrl & params
    strUrl = strUrl & "&format=csv"
    blnAsync = True

    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    With objRequest
        .Open "GET", strUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        
            While objRequest.readyState <> 4  'wait for response
                DoEvents
            Wend
        
        strResponse() = Split(.ResponseText, Chr(10)) 'split response into lines vbCrLf
    End With
   
    Lines = UBound(strResponse)  'get number of lines
    For Line = 0 To Lines - 1
        strResponse2 = Mid(strResponse(Line), 2, Len(strResponse(Line)) - 2)
        Sheets("Sheet1").Range(Cells(Line + 1, 1), Cells(Line + 1, 39)) = _
            Split(strResponse2, Chr(34) & "," & Chr(34))
    Next
   
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

