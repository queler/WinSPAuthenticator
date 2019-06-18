Attribute VB_Name = "spo"
Option Explicit
' Some Helper functions for a SPO request

Public Enum SP_RQ_TYPE
    Add
    Delete
    Update
    Read
End Enum

'    usage:Dim digest As String
'    digest = getRequestDigest(spClient)
'    Set rq = createBasicSPRequest(Delete, digest)
'    rq.Resource = "Web/Lists/getbytitle('qpop')/Items({id})"
Public Function createBasicSPRequest(op As SP_RQ_TYPE, Optional rqDigest As String = "") As WebRequest
    Const SPO_UA As String = "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"
    Set createBasicSPRequest = New WebRequest
    With createBasicSPRequest
        .Format = json
        .UserAgent = SPO_UA
        .ContentType = "application/json;odata=verbose"
        .Accept = "application/json;odata=verbose"
        Select Case op
            Case SP_RQ_TYPE.Add
                .Method = Httppost
                If rqDigest <> "" Then .SetHeader "X-RequestDigest", rqDigest
            Case SP_RQ_TYPE.Delete
                .Method = Httppost
                .AddHeader "X-HTTP-Method", "DELETE"
                .AddHeader "IF-MATCH", "*"
                If rqDigest <> "" Then .SetHeader "X-RequestDigest", rqDigest
            Case SP_RQ_TYPE.Read
                .Method = HttpGet
            Case SP_RQ_TYPE.Update
                .Method = Httppost
                .AddHeader "X-HTTP-Method", "MERGE"
                .AddHeader "IF-MATCH", "*"
                If rqDigest <> "" Then .SetHeader "X-RequestDigest", rqDigest
            Case Else
                Err.Raise vbError + 20
        End Select
    End With
End Function

Public Function getRequestDigest(client As WebClient) As String
    Dim dr As WebRequest
    Set dr = createBasicSPRequest(Read) 'vanilla rq is a read
    dr.Method = Httppost                'but we need a post for the digest
    dr.Body = ""
    dr.Resource = "contextinfo"

    Dim resp As WebResponse:
    Dim i As Integer: For i = 1 To 2
        Set resp = client.Execute(dr)
        If resp.StatusCode = Ok Then
            Exit For
        Else
            'there was a bug where i wasn't able to authenticate on my first request but I
            'think i squashed it
            Debug.Print i
            If i >= 2 Then Err.Raise vbError + 400, "Function AddDigestToRequest", "two retries, still failed"
        End If
    Next i
    Dim digest As String: digest = resp.Data("d")("GetContextWebInformation")("FormDigestValue")
    getRequestDigest = digest
End Function
