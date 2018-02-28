// PQGA
let
    Source = (idsVar as text, metricsVar as text, dimensionsVar as nullable text, startDateVar as date, endDateVar as date, filtersVar as nullable text, segmentVar as nullable text, intervalsVar as number) as table => 

let
 
fn_configRecord2 = (startDate, endDate) =>

let
    tokenBuffer = Text.From(getGATokenAuto), 
    startDateSelect = 
        if Date.AddDays(endDate, -intervalsVar) < startDate 
        then Date.ToText(startDate, "yyyy-MM-dd") 
        else Date.ToText(Date.AddDays(endDate, -intervalsVar), "yyyy-MM-dd") ,

    result = [
        ids = if Text.Contains(idsVar, "ga:") then idsVar else "ga:" & idsVar,
        #"start-date" = startDateSelect,
        #"end-date" = Date.ToText(endDate, "yyyy-MM-dd"),
        metrics = metricsVar,
        access_token = tokenBuffer,
        #"max-results" = "10000",
        #"start-index"="1"
    ],

    dimCheck= if dimensionsVar = null then result else result & [dimensions = dimensionsVar], 
    quest7 = if filtersVar = null then dimCheck else dimCheck & [filters = filtersVar], 
    quest8 = if segmentVar = null then quest7 else quest7 & [segment = segmentVar],


    quest9 = if endDate = null then null else quest8
in
    quest9,





fn_getData2 = (configRecord1, optional lastParameter) =>
let
    fn_GaJsonGetNoDelay= () => 
        if Text.Contains(configRecord1[metrics], "mcf:") 
        then Json.Document(Web.Contents("https://www.googleapis.com/analytics/v3/data/mcf", [Query = configRecord1, ManualStatusHandling={401}])) 
        else Json.Document(Web.Contents("https://www.googleapis.com/analytics/v3/data/ga", [Query = configRecord1, ManualStatusHandling={401}])),
    GaJsonGet = Function.InvokeAfter(fn_GaJsonGetNoDelay, #duration(0,0,0,0.2)),
    HeadersFromJson = try List.Transform(GaJsonGet[columnHeaders], each _[name]) otherwise {},
    RowsFromJson = try GaJsonGet[rows] otherwise {},

    checkDate = if Date.AddDays( Date.From(Uri.Parts(GaJsonGet[selfLink])[Query][#"end-date"]), -intervalsVar -1) >= startDateVar 
                then Date.AddDays(
                    Date.From(
                        Uri.Parts(
                            GaJsonGet[selfLink]
                            )[Query][#"end-date"]
                        )
                    , -intervalsVar -1)
                else null,
    
    res = [
        data=#table(HeadersFromJson, RowsFromJson), 
        samplingTrueOrFalse = try GaJsonGet[containsSampledData] otherwise null, 
        nextPage = try Uri.Parts(GaJsonGet[nextLink])[Query] & [access_token = getGATokenAuto] 
            otherwise fn_configRecord2(startDateVar, checkDate), 
        hadPreviousParameter = lastParameter, 
        SampleSize = try GaJsonGet[sampleSpace] otherwise null
]     
in
    res,


//    ids = "ga:72013428", metrics = "ga:sessions", dimensions = "ga:medium,ga:hostname,ga:date,ga:browser",  startDate = "2012-01-01", endDate = "2016-11-01", sort = null, filters = null, segment = null, 
    tokenBuffer = Text.From(getGATokenAuto), 
    
    generateDataConfig = List.Generate(
        ()=>[res = fn_getData2(fn_configRecord2(startDateVar, endDateVar), [])],
        each [res][hadPreviousParameter] <> null,
        each [res = fn_getData2([res][nextPage], [res][nextPage])]
    ),
    #"Converted to Table" = Table.FromList(generateDataConfig, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"res"}, {"res"}),
    #"Expanded res" = Table.ExpandRecordColumn(#"Expanded Column1", "res", {"data",  "samplingTrueOrFalse"}, {"data", "samplingTrueOrFalse"}),
    #"Expanded data" = Table.ExpandTableColumn(#"Expanded res", "data", Table.ColumnNames(#"Expanded res"[data]{0}) )
in
    #"Expanded data",

  documentation = [
   Documentation.Name = "PQGA",
   Documentation.Description = "Returns data from Google Analytics via API as a table",
   Documentation.LongDescription = "Returns data from Google Analytics via APIs: Core v3 reporting api and MCF reporting api as a table. 
There is intervalsVar. It splits requests to api into defined number of days (It's made to minimise sampling).",
   Documentation.Category = "function",
   Documentation.Examples = {[Description = "Example of Query from MCF Api", Code = "= PQGA2(""ga:72013428"", ""mcf:totalConversions"", ""mcf:sourceMediumPath"", #date(2014, 12, 1), #date(2017, 12, 1), null, null, 20)", Result = "#table({""data""}, {{""data""}})"]}
  ]
 
in
  Value.ReplaceType(Source, Value.ReplaceMetadata(Value.Type(Source), documentation))



//              .         .                                                                      .         .                
//             ,8.       ,8.                   .8.          `8.`8888.      ,8'  8 8888          ,8.       ,8.               
//            ,888.     ,888.                 .888.          `8.`8888.    ,8'   8 8888         ,888.     ,888.              
//           .`8888.   .`8888.               :88888.          `8.`8888.  ,8'    8 8888        .`8888.   .`8888.             
//          ,8.`8888. ,8.`8888.             . `88888.          `8.`8888.,8'     8 8888       ,8.`8888. ,8.`8888.            
//         ,8'8.`8888,8^8.`8888.           .8. `88888.          `8.`88888'      8 8888      ,8'8.`8888,8^8.`8888.           
//        ,8' `8.`8888' `8.`8888.         .8`8. `88888.         .88.`8888.      8 8888     ,8' `8.`8888' `8.`8888.          
//       ,8'   `8.`88'   `8.`8888.       .8' `8. `88888.       .8'`8.`8888.     8 8888    ,8'   `8.`88'   `8.`8888.         
//      ,8'     `8.`'     `8.`8888.     .8'   `8. `88888.     .8'  `8.`8888.    8 8888   ,8'     `8.`'     `8.`8888.        
//     ,8'       `8        `8.`8888.   .888888888. `88888.   .8'    `8.`8888.   8 8888  ,8'       `8        `8.`8888.       
//    ,8'         `         `8.`8888. .8'       `8. `88888. .8'      `8.`8888.  8 8888 ,8'         `         `8.`8888.      
//                                                                                                                          
//        8 8888      88 `8.`888b           ,8' .8.          8 888888888o.      ,o888888o.  `8.`888b           ,8'          
//        8 8888      88  `8.`888b         ,8' .888.         8 8888    `88.  . 8888     `88. `8.`888b         ,8'           
//        8 8888      88   `8.`888b       ,8' :88888.        8 8888     `88 ,8 8888       `8b `8.`888b       ,8'            
//        8 8888      88    `8.`888b     ,8' . `88888.       8 8888     ,88 88 8888        `8b `8.`888b     ,8'             
//        8 8888      88     `8.`888b   ,8' .8. `88888.      8 8888.   ,88' 88 8888         88  `8.`888b   ,8'              
//        8 8888      88      `8.`888b ,8' .8`8. `88888.     8 888888888P'  88 8888         88   `8.`888b ,8'               
//        8 8888      88       `8.`888b8' .8' `8. `88888.    8 8888`8b      88 8888        ,8P    `8.`888b8'                
//        ` 8888     ,8P        `8.`888' .8'   `8. `88888.   8 8888 `8b.    `8 8888       ,8P      `8.`888'                 
//          8888   ,d8P          `8.`8' .888888888. `88888.  8 8888   `8b.   ` 8888     ,88'        `8.`8'                  
//           `Y88888P'            `8.` .8'       `8. `88888. 8 8888     `88.    `8888888P'           `8.`                   
