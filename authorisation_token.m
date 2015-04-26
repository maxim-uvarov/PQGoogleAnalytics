// the purpose of this query is to retrieve "Refresh token"

let
  Source = Excel.CurrentWorkbook(){[Name="auth_token_table"]}[Content],
  code = "code=" & Source{0}[authentication_code],
  url = code & "&" & app_credentials & "&redirect_uri=urn:ietf:wg:oauth:2.0:oob&grant_type=authorization_code",
  GetJson = Web.Contents("https://accounts.google.com/o/oauth2/token",
    [
      Headers = [#"Content-Type"="application/x-www-form-urlencoded"],
      Content = Text.ToBinary(url)
    ]
  ),
    FormatAsJson = Json.Document(GetJson),
    #"Converted to Table" = Record.ToTable(FormatAsJson),
    #"Transposed Table" = Table.Transpose(#"Converted to Table"),
    #"First Row as Header" = Table.PromoteHeaders(#"Transposed Table"),
    #"Removed Columns" = Table.RemoveColumns(#"First Row as Header",{"access_token"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"refresh_token", "token_type", "expires_in"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Reordered Columns",{"refresh_token"})
in
  #"Removed Other Columns"