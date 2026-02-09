Module Tests

	Sub	TestHttpClient()
        On Error GoTo EH
        With New cHttpClient
            'Make json body
            With .RequestDataJson
                With .NewItem("consignee")
                    With .NewItem("address")
                        .Item("city") = "Paderborn"
                        .Item("countryCode") = "DE"
                        .Item("postCode") = "33100"
                    End With
                End With
                With .NewItems("lines")
                    With .NewItem()
                        .Item("content") = "furniture"
                        .Item("unitHeight") = 120
                        .Item("unitLength") = 120
                        .Item("unitPackageType") = "EP"
                        .Item("unitQuantity") = 1
                        .Item("unitWeight") = 200
                        .Item("unitWidth") = 80
                    End With
                End With
                .Item("product") = "STANDARD"
                With .NewItem("consignee")
                    With .NewItem("address")
                        .Item("city") = "Berlin"
                        .Item("countryCode") = "DE"
                        .Item("postCode") = "10115"
                    End With
                End With
                .Item("pickupOn") = Format$(Now, "yyyy-MM-dd")
                'Check json data
                Debug.Print "Request Data: ",.Encode(, 2, True)
            End With
            'You can open debug info object form build in.
            .DebugStart = True
            'Add any header
            .RequestHeaders.Item("X-API-KEY") = "YOUR API KEY HERE"
            'You can set Content-Type like this:
            .SetRequestContentType(EnumRequestContentType.JsonString)
            'Post 
            .SendPost("https://vb6.pro/echo/?hello=woeoio")
            'Server retun json data
            With .ReturnJson()
                MsgBox(.Root("server")("time"))
                'Show server reopnse, you can use .ReturnText()
                Debug.Print "Server Response: ", .Encode(, 2, True)
            End With
            QH:
            'Show deubg info
            Debug.Print "Debug Info: ", .DebugInfo.Encode(, 2, True)
            Exit Sub
            EH:
            'Do some thing...
            MsgBox(Err.Description)
            Resume QH
        End With
    End Sub

End Module