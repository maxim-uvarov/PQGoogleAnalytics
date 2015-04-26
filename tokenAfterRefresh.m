// Function, which is used every time, when main function in operation. It retrieves "access token" using "refresh token"

let
  Source = Excel.CurrentWorkbook(){[Name="authorisation_token"]}[Content],
  refresh_token = Source{0}[refresh_token],
  url = app_credentials & "&refresh_token=" & refresh_token & "&grant_type=refresh_token",
  GetJson = Web.Contents("https://accounts.google.com/o/oauth2/token",
    [
      Headers = [#"Content-Type"="application/x-www-form-urlencoded"],
      Content = Text.ToBinary(url)
    ]
  ),
    FormatAsJson = Json.Document(GetJson),
    #"Converted to Table" = Record.ToTable(FormatAsJson),
    #"Transposed Table" = Table.Transpose(#"Converted to Table"),
    #"First Row as Header" = Table.PromoteHeaders(#"Transposed Table")
in
    #"First Row as Header"