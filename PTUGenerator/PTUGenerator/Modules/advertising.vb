Module advertising
    Dim script As String = ""
    Dim keyWords As String = ""
    Dim fileAds As String = "ads.html"

    Friend Sub InitializedAds()
        keyWords = "computer, IT, department, software, pgc, perfecto, dalton, perfecom, pcom, system, program, develop, hardware, technical, MIS"
        script = "<script async src=""http://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js""></script>"
        script &= vbCrLf & "<!-- Windows Apps -->"
        script &= vbCrLf & "<ins class=""adsbygoogle"""
        script &= vbCrLf & "style=""display:block"""
        script &= vbCrLf & "data-ad-client=""ca-pub-3858294701571956"""
        script &= vbCrLf & "data-ad-slot=""3360323427"""
        script &= vbCrLf & "data-ad-format=""auto""></ins>"
        script &= vbCrLf & "<script>"
        script &= vbCrLf & "(adsbygoogle = window.adsbygoogle || []).push({});"
        script &= vbCrLf & "</script>"

        script = "<div style=""visibility:hidden;"">" & keyWords & "</div>" & vbCrLf & script
    End Sub

    Friend Function DisplayAds() As String
        Dim exePath As String = Application.StartupPath() & "\"
        createAds()

        Return exePath & fileAds
    End Function

    Private Sub createAds()
        If Not System.IO.File.Exists(fileAds) Then
            System.IO.File.CreateText(fileAds).Dispose()
        End If

        Dim objWriter As New System.IO.StreamWriter(fileAds)
        objWriter.WriteLine(script)
        objWriter.Close()
        Console.WriteLine("Ads written")
    End Sub
End Module
