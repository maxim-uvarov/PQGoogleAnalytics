/* PQGA - PQGoogleAnalytics - custom, unoffical Google Analytics connector for Power BI

homepage: https://github.com/maxim-uvarov/PQGoogleAnalytics

link to get auth token: https://accounts.google.com/o/oauth2/auth?scope=https://www.googleapis.com/auth/analytics.readonly&response_type=code&access_type=offline&redirect_uri=urn:ietf:wg:oauth:2.0:oob&approval_prompt=force&client_id=155297956885-088obc06926s8kolll6kdqkd9u842n56

Example invoked query: 
PQGA("72013428", "ga:sessions", "ga:source,ga:hostname", "2016-01-01", "2016-11-01", null, null, null)

*/




let
    Source = (ids as text, metrics as text, dimensions as nullable text, startDate as text, endDate as text, optional sort as text, optional filters as text, optional segment as text) => 

    let

    // helper function to get 1 page of report

        queryWithOffsetForOneTable = (ids_with_ga as text, startDate as text, endDate as text, metrics as text, optional dimensions as text, optional sort as text, optional filters as text, optional segment as text, optional startIndex as text) => 

        let
        // Creating url to query data from google analytics

            recordToApply = [],
            quest1 = Record.AddField(recordToApply, "ids", ids_with_ga), 
            quest2 = Record.AddField(quest1, "start-date", startDate), 
            quest3 = Record.AddField(quest2, "end-date", endDate), 
            quest4 = Record.AddField(quest3, "metrics", metrics), 
            quest5 = if dimensions = null then quest4 else Record.AddField(quest4, "dimensions", dimensions), 
            quest6 = if sort = null then quest5 else Record.AddField(quest5, "sort", sort), 
            quest7 = if filters = null then quest6 else Record.AddField(quest6, "filters", filters), 
            quest8 = if segment = null then quest7 else Record.AddField(quest7, "segment", segment), 
            quest9 = Record.AddField(quest8, "start-index", startIndex), 
            quest10 = Record.AddField(quest9, "access_token", getGATokenAuto),
            quest11 = Record.AddField(quest10, "max-results", "10000"),
            quest12 = Uri.BuildQueryString(quest11),
            url = "https://www.googleapis.com/analytics/v3/data/ga?" & quest12,

            // Getting data from Google Analytics
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
            tabletoappend, 

    // start of main program

        ids_with_ga = if Text.Contains(ids, "ga:") 
            then ids 
            else "ga:" & ids, 
        recordToApply = [],
        quest1 = Record.AddField(recordToApply, "ids", ids_with_ga), 
        quest2 = Record.AddField(quest1, "start-date", startDate), 
        quest3 = Record.AddField(quest2, "end-date", endDate), 
        quest4 = Record.AddField(quest3, "metrics", metrics), 
        quest5 = if dimensions = null then quest4 else Record.AddField(quest4, "dimensions", dimensions), 
        quest6 = if sort = null then quest5 else Record.AddField(quest5, "sort", sort), 
        quest7 = if filters = null then quest6 else Record.AddField(quest6, "filters", filters), 
        quest8 = if segment = null then quest7 else Record.AddField(quest7, "segment", segment), 

        quest10 = Record.AddField(quest8, "access_token", getGATokenAuto),
        quest11 = Record.AddField(quest10, "max-results", "10000"),
        quest12 = Uri.BuildQueryString(quest11),
        url = "https://www.googleapis.com/analytics/v3/data/ga?" & quest12,


    // Getting first page Json
        GaJson = Json.Document(Web.Contents(url)),

    // Calculating pages of report
        Custom1 = Number.RoundDown( GaJson [totalResults] / GaJson[itemsPerPage]),

    // Creating table of numbers to use in querywithoutoffset function
        Custom2 = List.Numbers(1,Custom1),
        #"Table from List1" = Table.FromList(Custom2, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        #"Multiplied Column" = Table.TransformColumns(#"Table from List1", {{"Column1", each Number.ToText(List.Product({_, GaJson [itemsPerPage]})+1)}}),

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

    // Querying api for each page of report with 10k results (paging)
        #"Added Custom2" = Table.AddColumn(#"Multiplied Column", "table", each queryWithOffsetForOneTable(ids_with_ga, startDate, endDate, metrics, dimensions, sort, filters, segment, [Column1])),
        #"Removed Other Columns" = Table.SelectColumns(#"Added Custom2",{"table"}),
        #"Expand table" = try Table.ExpandTableColumn(#"Removed Other Columns", "table", columnsnames, columnsnames) otherwise Table.FromList({}),
        
    // Appending all created tables
        tabletoappend  = Table.RemoveColumns(#"Expand Custom",{"Column1", "table"}),
        Append = Table.Combine({tabletoappend,#"Expand table"})
    in
        Append
in
    Source
