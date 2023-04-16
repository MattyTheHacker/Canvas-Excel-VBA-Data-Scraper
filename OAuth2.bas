Attribute VB_Name = "OAuth2"
Option Explicit

Sub GetClientIDAndSecret()

    ' redirect users to request canvas access
    Dim url As String
    url = "https://canvas.bham.ac.uk/login/oauth2/auth?client_id=XXX&response_type=code&state=YYY&redirect_uri=https://example.com/oauth2response"

End Sub
