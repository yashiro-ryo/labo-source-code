Sub Main()
    ' シートからデータを取得
    Dim ApiKey As String
    ApiKey = Range("B1").Value
    Dim MessageBody As String
    MessageBody = Range("B2").Value
    
    ' パラメータのJSON変換
    Dim Json As Object
    Set Json = JsonConverter.ParseJson("{""model"": ""gpt-3.5-turbo"",""messages"":[{""role"": ""user"", ""content"": """"}]}")
    ' Json内にメッセージ本文を格納
    Json("messages")(1)("content") = MessageBody
    Debug.Print JsonConverter.ConvertToJson(Json)
    
    Dim httpReq As New XMLHTTP60   '「Microsoft XML, v6.0」を参照設定
    Dim params As New Dictionary   '「Microsoft Scripting Runtime」を参照設定
    
    ' POSTサンプル
    ' headerにapi keyを添付する
    ' content-typeにjsonを指定する
    With httpReq
      .Open "POST", "https://api.openai.com/v1/chat/completions"
      .setRequestHeader "Authorization", "Bearer " + ApiKey
      .setRequestHeader "Content-Type", "application/json"
      .send JsonConverter.ConvertToJson(Json)
    End With

    Do While httpReq.readyState < 4
        DoEvents
    Loop

    Debug.Print httpReq.responseText
    
    Dim ResponseJson As Object
    ' responseをJSONに変換
    Set ResponseJson = JsonConverter.ParseJson(httpReq.responseText)
    
    ' セルへresponseを格納
    Range("B3").Value = ResponseJson("choices")(1)("message")("content")
End Sub