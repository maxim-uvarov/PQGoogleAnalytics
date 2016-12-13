// the purpose of this query is to retrieve "Refresh token"

let
    app_credentials= "client_id=155297956885-088obc06926s8kolll6kdqkd9u842n56.apps.googleusercontent.com&client_secret=bDlhVaxg2CGccGBT4xyYakSA",
    code = "code=" & authToken,

    url = code & "&" & app_credentials & "&redirect_uri=urn:ietf:wg:oauth:2.0:oob&grant_type=authorization_code",
    GetJson = Web.Contents("https://accounts.google.com/o/oauth2/token",
      [
        Headers = [#"Content-Type"="application/x-www-form-urlencoded"],
        Content = Text.ToBinary(url)
      ]
    ),
    refreshTokenOutput = try Json.Document(GetJson)[refresh_token] otherwise "Bad authToken. Go to google analytics, get a new one, paste it into authToken parameter, get back"
in
    refreshTokenOutput
