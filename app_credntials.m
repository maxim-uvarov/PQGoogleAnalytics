// the purpose for this query is to get from excel "app_credentials" table client_id and client_secret for app on google

let
    Source = Excel.CurrentWorkbook(){[Name="app_credentials"]}[Content],
    #"Merged Columns" = Table.CombineColumns(Source,{"parameter_name", "token"},Combiner.CombineTextByDelimiter("=", QuoteStyle.None),"Merged"),
    #"Transposed Table" = Table.Transpose(#"Merged Columns"),
    #"Merged Columns1" = Table.CombineColumns(#"Transposed Table",{"Column1", "Column2"},Combiner.CombineTextByDelimiter("&", QuoteStyle.None),"Merged"),
    Merged = #"Merged Columns1"{0}[Merged]
in
    Merged