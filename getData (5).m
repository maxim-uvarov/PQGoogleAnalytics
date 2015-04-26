let
    analytics = "https://www.googleapis.com/analytics/v3/data/ga",
    access_token = "&access_token=" & tokenAfterRefresh{0}[access_token],
    url = analytics & parameters & access_token,

// Getting first page Json
    GaJson = Json.Document(Web.Contents(url)),

// Calculating pages of report
    Custom1 = Number.RoundDown( GaJson [totalResults] / GaJson[itemsPerPage]),

// Creating table of numbers to use querywithoutoffset function
    Custom2 = List.Numbers(1,Custom1),
    #"Table from List1" = Table.FromList(Custom2, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Multiplied Column" = Table.TransformColumns(#"Table from List1", {{"Column1", each List.Product({_, GaJson [itemsPerPage]})+1}}),

// Setting names for headers (retrieving them from first json)
    columnHeaders11 = GaJson[columnHeaders],
    #"Table from List11" = Table.FromList(columnHeaders11, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expand Column1" = Table.ExpandRecordColumn(#"Table from List11", "Column1", {"name"}, {"Column1.name"}),
    columnsnames = #"Expand Column1"[Column1.name],

// Creating table from first json
    rows = GaJson[rows],
    #"Table from List" = Table.FromList(rows, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Added Custom" = Table.AddColumn(#"Table from List", "table", each Table.FromList([Column1],Splitter.SplitByNothing(), null, null, ExtraValues.Error)),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Custom", each Table.Transpose([table])),
    #"Expand Custom" = Table.ExpandTableColumn(#"Added Custom1", "Custom", Table.ColumnNames(#"Added Custom1"{0}[Custom]), columnsnames),

// Querying api for each row (paging)
    #"Added Custom2" = Table.AddColumn(#"Multiplied Column", "table", each querywithoffset([Column1])),
    #"Removed Other Columns" = Table.SelectColumns(#"Added Custom2",{"table"}),
    #"Expand table" = try Table.ExpandTableColumn(#"Removed Other Columns", "table", columnsnames, columnsnames) otherwise Table.FromList({}),
    
// Appending all created tables
    tabletoappend  = Table.RemoveColumns(#"Expand Custom",{"Column1", "table"}),
    Append = Table.Combine({tabletoappend,#"Expand table"})
in
    Append