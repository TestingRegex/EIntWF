Option Explicit

' Constants for the API endpoint and key
Const ApiEndpoint As String = "https://api.example.com/"
Const ApiKey As String = "IFfREUlhk1GazoYUgvea"

' Function to make a GET request to the API
Function GetApiResponse(endpoint As String) As String
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

    ' Construct the API request URL
    Dim url As String
    url = ApiEndpoint & endpoint

    ' Open the connection and send the request
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "Authorization", "Bearer " & ApiKey
    xmlhttp.send
    'asdfdsf
    ' Return the API response
    GetApiResponse = xmlhttp.responseText
End Function

' Example usage of the API function
Sub TestApi()
    Dim response As String
    response = GetApiResponse("example-endpoint")

    ' Handle the API response as needed
    MsgBox response
End Sub