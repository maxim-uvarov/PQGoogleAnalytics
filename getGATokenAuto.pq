let

// the purpose for this query is to get from excel "app_credentials" table client_id and client_secret for app on google

app_credentials =
"client_id=155297956885-088obc06926s8kolll6kdqkd9u842n56.apps.googleusercontent.com&client_secret=bDlhVaxg2CGccGBT4xyYakSA",

url = app_credentials & "&refresh_token=" & refreshToken & "&grant_type=refresh_token",
  GetJson = Web.Contents("https://accounts.google.com/o/oauth2/token",
    [
      Headers = [#"Content-Type"="application/x-www-form-urlencoded"],
      Content = Text.ToBinary(url)
    ]
  ),
    FormatAsJson = Json.Document(GetJson)[access_token]
in
    FormatAsJson
