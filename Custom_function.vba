Function OpenAI(prompt As String) As String

Dim url As String, apiKey As String
Dim response As Object, json As String

apiKey = "your API KEY"
url = "https://api.openai.com/v1/engines/text-davinci-003/completions"

Set response = CreateObject("MSXML2.XMLHTTP")
response.Open "POST", url, False
response.setRequestHeader "Content-Type", "application/json"
response.setRequestHeader "Authorization", "Bearer " + apiKey
response.Send "{""prompt"":""" & prompt & """,""max_tokens"":1024}"

json = response.responseText
OpenAI = Split(Mid(json, InStr(json, """text"":""") + 8), """")(0)
OpenAI = Replace(OpenAI, "\n", "")


End Function
