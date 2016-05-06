let

// the purpose for this query is to get from excel "app_credentials" table client_id and client_secret for app on google

app_credentials =
let
    Source = Excel.CurrentWorkbook(){[Name="app_credentials"]}[Content],
    #"Merged Columns" = Table.CombineColumns(Source,{"parameter_name", "token"},Combiner.CombineTextByDelimiter("=", QuoteStyle.None),"Merged"),
    #"Transposed Table" = Table.Transpose(#"Merged Columns"),
    #"Merged Columns1" = Table.CombineColumns(#"Transposed Table",{"Column1", "Column2"},Combiner.CombineTextByDelimiter("&", QuoteStyle.None),"Merged"),
    Merged = #"Merged Columns1"{0}[Merged]
in
    Merged,

// main app


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
    #"First Row as Header" = Table.PromoteHeaders(#"Transposed Table"),
    access_token = #"First Row as Header"{0}[access_token]
in
    access_token
