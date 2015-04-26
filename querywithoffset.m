// function to get table from google analytics. Used in main Query

let
    querywithoffset = (offset_1) => 

let
    analytics = "https://www.googleapis.com/analytics/v3/data/ga",
    access_token = "&access_token=" & tokenAfterRefresh{0}[access_token],
    startindex = "&start-index=" &Number.ToText (offset_1),
    url = analytics & parameters & access_token & startindex,
    GaJson = Json.Document(Web.Contents(url)),

    columnHeaders11 = GaJson[columnHeaders],
    #"Table from List11" = Table.FromList(columnHeaders11, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expand Column1" = Table.ExpandRecordColumn(#"Table from List11", "Column1", {"name"}, {"Column1.name"}),
    columnsnames = #"Expand Column1"[Column1.name],
    
    rows = GaJson[rows],
    #"Table from List" = Table.FromList(rows, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Added Custom" = Table.AddColumn(#"Table from List", "table", each Table.FromList([Column1],Splitter.SplitByNothing(), null, null, ExtraValues.Error)),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Custom", each Table.Transpose([table])),
    #"Expand Custom" = Table.ExpandTableColumn(#"Added Custom1", "Custom", Table.ColumnNames(#"Added Custom1"{0}[Custom]), columnsnames),

    tabletoappend  = Table.RemoveColumns(#"Expand Custom",{"Column1", "table"})
in
    tabletoappend
in
    querywithoffset