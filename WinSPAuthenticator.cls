VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinSPAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Sharepoint with windows auth
' (c) Adam Queler
'
' Description
' With a big assist from https://blogs.msdn.microsoft.com/omarv/2012/11/15/developing-windows-8-store-apps-for-sharepoint-online-with-sso-single-sign-on/
' Requires WindowsAuthenticator
' Windows only
' @class
' @implements IWebAuthenticator v4.*
' @author Adam Queler
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup authenticator
''
'IOMode Constants
Private originalRq As WebRequest
' sharepoint likes this UA
Const SPO_UA As String = "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"
'IOMode Constants
Const ForAppending As Integer = 8
Const ForReading As Integer = 1
Const ForWriting As Integer = 2
Dim spSite As String
Const rtfa As String = "rtFa"
Const fedAuth As String = "FedAuth"
Private cookieRtfa As String
Private cookieFedAuth As String
Public GotCookies As Boolean
Public Sub Setup(Optional altSiteURL As String = "")
    'only need to run this is you are sending something other than a pathless host of the base url for authentication
    ' i.e. if base url is host.sharepoint.com/somesite/path, default will use host.sharepoint.com for authentication purposes
     spSite = altSiteURL

End Sub
Private Sub Class_Initialize()
GotCookies = False
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal client As WebClient, ByRef request As WebRequest)
    ' e.g Add headers, cookies, etc.
    loadCookies client
    If cookieFedAuth = "" Or cookieRtfa = "" Then 'Or headerDigest = "" Then
        WebHelpers.LogDebug "cookie blank getting new ones"
        getCookiesFromSPO client
    Else
        LogDebug "probably loaded cookies", "WinSPAuthenticator"
    End If
    Set originalRq = request
    Set request = originalRq.Clone
    request.AddCookie rtfa, cookieRtfa
    request.AddCookie fedAuth, cookieFedAuth
    
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param in|out {WebResponse} Response to request
''
Private Sub IWebAuthenticator_AfterExecute(ByVal client As WebClient, ByVal request As WebRequest, ByRef Response As WebResponse)
    ' e.g. Handle 401 Unauthorized or other issues
    If Response.StatusCode = FORBIDDEN Then 'refresh
        LogDebug "forbidden"
        If Not GotCookies Then
            LogDebug "but we haven't gotten new cookies yet in this rq"
            getCookiesFromSPO client
            Set request = originalRq
            request.AddCookie rtfa, cookieRtfa
            request.AddCookie fedAuth, cookieFedAuth
            Set Response = client.Execute(request)
            Set request = originalRq
        End If
        
    End If
End Sub
''
' Hook for updating http before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {WinHttpRequest} Http
''
Private Sub IWebAuthenticator_PrepareHttp(ByVal client As WebClient, ByVal request As WebRequest, ByRef Http As Object)
    ' e.g. Update option, headers, etc.
End Sub

''
' Hook for updating cURL before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {String} Curl
''
Private Sub IWebAuthenticator_PrepareCurl(ByVal client As WebClient, ByVal request As WebRequest, ByRef Curl As String)
    ' e.g. Add flags to cURL
End Sub
Private Function hash(s As String)
    hash = Left(MD5(s, "Base64"), 22)
End Function

Public Sub loadCookies(client As WebClient)
    Dim fso As Object ' As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
    Dim fp As String
    fp = fso.BuildPath(Environ$("APPDATA"), hash("WinSPAuthenticator" & client.BaseUrl))
    If fso.FolderExists(fp) Then
        Dim pRtfa As String: pRtfa = fso.BuildPath(fp, hash(rtfa))
        Dim pFedAuth As String: pFedAuth = fso.BuildPath(fp, hash(fedAuth))
        'Dim pDigest As String: pDigest = fso.BuildPath(fp, hash("digest"))
        If fso.fileexists(pRtfa) And fso.fileexists(pFedAuth) Then 'And fso.fileexists(pDigest) Then
            Dim ts ' As Scripting.TextStream
            Set ts = fso.OpenTextFile(pRtfa)
            cookieRtfa = (ts.ReadAll)
            ts.Close
            Set ts = fso.OpenTextFile(pFedAuth)
            cookieFedAuth = (ts.ReadAll)
            ts.Close
        End If
    End If
End Sub

Public Sub saveCookies(client As WebClient)
    Dim fso ' As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
    Dim fp As String
    fp = fso.BuildPath(Environ$("APPDATA"), hash("WinSPAuthenticator" & client.BaseUrl))
    If Not fso.FolderExists(fp) Then
        fso.CreateFolder fp
    End If
    Dim ts As Object 'Scripting.TextStream
    Dim pRtfa As String: pRtfa = fso.BuildPath(fp, hash(rtfa))
    Dim pFedAuth As String: pFedAuth = fso.BuildPath(fp, hash(fedAuth))
    Set ts = fso.OpenTextFile(pRtfa, ForWriting, True)
    ts.write cookieRtfa
    ts.Close
    Set ts = Nothing
    Set ts = fso.OpenTextFile(pFedAuth, ForWriting, True)
    ts.write cookieFedAuth
    ts.Close
End Sub
Public Sub getCookiesFromSPO(client As WebClient)
    Call WebHelpers.LogDebug("getting cookies", "WinSPAuthenticator")
    Const spowssigninUri   As String = "_forms/default.aspx?wa=wsignin1.0"
    Const msoStsUrl As String = "https://login.microsoftonline.com/extSTS.srf"
    Dim realmClient As New WebClient
    Dim wa As New WindowsAuthenticator
    Set realmClient.Authenticator = wa
    realmClient.BaseUrl = "https://login.microsoftonline.com/"
    If spSite = "" Then
        Dim parts As Dictionary
        Set parts = GetUrlParts(client.BaseUrl)
        spSite = Replace(client.BaseUrl, parts("Path"), "")
        spSite = Replace(spSite, parts("Querystring"), "")
        spSite = Replace(spSite, parts("Hash"), "")
    End If
    Dim rx As New WebRequest
    rx.UserAgent = SPO_UA
    rx.Resource = "GetUserRealm.srf"
    rx.ResponseFormat = json
    rx.RequestFormat = FormUrlEncoded
    rx.AddBodyParameter "login", Environ("username") & "@" & Environ("userdnsdomain") ' might be different for other people
    rx.Method = Httppost
    Dim resp As WebResponse
    Set resp = realmClient.Execute(rx)
    Dim authURL As String
    authURL = resp.Data("AuthURL"): realmClient.BaseUrl = authURL
    Dim authRx As New WebRequest: Set authRx = New WebRequest
    With authRx
        .Format = PlainText
        .UserAgent = SPO_UA
        .Accept = "text/html; charset=utf-8"
        .Method = HttpGet
        .ContentType = "text/html; charset=utf-8"
        .Resource = ""
        .Body = ""
    End With
    Set resp = realmClient.Execute(authRx)
    Dim dom As Object: Set dom = CreateObject("MSXML2.DOMDocument.6.0") 'New MSXML2.DOMDocument60
    dom.Async = False
    dom.LoadXML resp.Content
    dom.LoadXML dom.SelectSingleNode("//form/input[@name='wresult']/@value").NodeValue
    Dim assertion As String
    dom.setProperty "SelectionNamespaces", "xmlns:saml=""urn:oasis:names:tc:SAML:1.0:assertion"""
    assertion = dom.SelectSingleNode("//saml:Assertion").Xml
    Dim env As String
    env = ParameterizeSoapRequestTokenMsgWithAssertion(spSite, assertion, msoStsUrl)
    realmClient.BaseUrl = msoStsUrl
    Set authRx = New WebRequest
    With authRx
        .Format = PlainText
        .UserAgent = SPO_UA
        .Accept = "application/soap+xml; charset=utf-8"
        .Method = HttpGet
        .ContentType = .Accept
        .Resource = ""
        .Body = env
    End With
    Set resp = realmClient.Execute(authRx)
    Dim secToken As String, secTokenExp
    dom.LoadXML resp.Content
    dom.setProperty "SelectionNamespaces", "xmlns:wsa='http://www.w3.org/2005/08/addressing' xmlns:wsse='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd' xmlns:wsu='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd' xmlns:wsp='http://schemas.xmlsoap.org/ws/2004/09/policy' xmlns:wst='http://schemas.xmlsoap.org/ws/2005/02/trust' xmlns:S='http://www.w3.org/2003/05/soap-envelope'"
    secToken = dom.SelectSingleNode("//wsse:BinarySecurityToken").Text
    secTokenExp = ParseIso(dom.SelectSingleNode("//wsu:Expires").Text)
    realmClient.BaseUrl = spSite
    Set rx = New WebRequest
    rx.Resource = spowssigninUri
    rx.Method = Httppost
    rx.Body = secToken
    rx.Format = FormUrlEncoded
    Set resp = realmClient.Execute(rx)
    cookieRtfa = FindInKeyValues(resp.Cookies, rtfa)
    cookieFedAuth = FindInKeyValues(resp.Cookies, fedAuth)
    saveCookies client
    GotCookies = True
End Sub
Function ParameterizeSoapRequestTokenMsgWithAssertion(spSite As String, samlAssertion As String, stsUrl As String) As String
    
    Dim samlRTString  As String
    samlRTString = "<s:Envelope xmlns:s=""http://www.w3.org/2003/05/soap-envelope"" xmlns:a=""http://www.w3.org/2005/08/addressing"" xmlns:u=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"">" & _
                    "  <s:Header>" & _
                    "    <a:Action s:mustUnderstand=""1"">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>" & _
                    "    <a:ReplyTo>" & _
                    "      <a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>" & _
                    "    </a:ReplyTo>" & _
                    "    <a:To s:mustUnderstand=""1"">[toUrl]</a:To>" & _
                    "    <o:Security s:mustUnderstand=""1"" xmlns:o=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"">[assertion]" & _
                    "    </o:Security>" & _
                    "  </s:Header>" & _
                    "  <s:Body>" & _
                    "<t:RequestSecurityToken xmlns:t=""http://schemas.xmlsoap.org/ws/2005/02/trust"">" & _
                    "      <wsp:AppliesTo xmlns:wsp=""http://schemas.xmlsoap.org/ws/2004/09/policy"">" & _
                    "        <a:EndpointReference>" & _
                    "          <a:Address>[url]</a:Address>" & _
                    "        </a:EndpointReference>" & _
                    "      </wsp:AppliesTo>" & _
                    "      <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>" & _
                    "      <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>" & _
                    "      <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>" & _
                    "    </t:RequestSecurityToken>" & _
                    "  </s:Body>" & _
"</s:Envelope>"
    samlRTString = Replace(samlRTString, "[assertion]", samlAssertion)
    samlRTString = Replace(samlRTString, "[url]", spSite)
    samlRTString = Replace(samlRTString, "[toUrl]", stsUrl)

    ParameterizeSoapRequestTokenMsgWithAssertion = samlRTString
        
End Function
