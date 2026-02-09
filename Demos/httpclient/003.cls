Module demoHttpClitent

    Public Sub TestGet()
        MsgBox TBMAN.HttpClient.Send(ReqGet, "http://a-vi.com/tbman/hello/woeoio").ReturnText()
    End Sub
        
    Public Sub TestPost()
        On Error GoTo EH
        Dim Body As String
        With New cJson
            .Item("username") = "dengwei"
            .Item("password") = "123456"
            Body = .Encode()
        End With
        With New cHttpClient
            .SetRequestContentType(JsonString)
            .SendPost("https://api.vb6.pro/tbman/login?ret=0", Body)
            With .ReturnJson()
                Debug.Print .Encode(, 2, True)
                If .Root("code") = 0 Then
                    MsgBox("wellcome to www.vb6.pro")
                Else
                    MsgBox(.Root("msg"))
                End If
            End With
        End With
        Exit Sub
        EH:
        MsgBox(Err.Description)
    End Sub

End Module