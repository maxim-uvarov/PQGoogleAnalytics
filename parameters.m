// Query parameters function - retrieves report config data from Excel table

let
    Source = Excel.CurrentWorkbook(){[Name="query_params"]}[Content],
    #"Filtered Rows" = Table.SelectRows(Source, each ([Parameter_value] <> null)),
    #"Merged Columns" = Table.CombineColumns(#"Filtered Rows",{"Parameter_name", "Parameter_value"},Combiner.CombineTextByDelimiter("=", QuoteStyle.None),"Merged"),
    #"Transposed Table" = Table.Transpose(#"Merged Columns"),
    columnnames= Table.ColumnNames(#"Transposed Table"),
    #"Merged Columns1" = Table.CombineColumns(#"Transposed Table",columnnames,Combiner.CombineTextByDelimiter("&", QuoteStyle.None),"Merged"),
    Merged = "?" & #"Merged Columns1"{0}[Merged]
in
    Merged