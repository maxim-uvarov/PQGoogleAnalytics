section GetGoogleAnalyticsBL;



client_id = "299222177909-8mvc09uilhdrick7kqjvf4hsi3tp4duc.apps.googleusercontent.com";
client_secret = "ELtgNWb9teqd20vCqDdARaok";

redirect_uri = "https://preview.powerbi.com/views/oauthredirect.html";
windowWidth = 1200;
windowHeight = 1000;


[DataSource.Kind="GetGoogleAnalyticsBL", Publish="GetGoogleAnalyticsBL.UI"]

shared GetGoogleAnalyticsBL.Data = Value.ReplaceType(GetGoogleAnalyticsBL.Stat, FuncType);

FuncType =type function (
       idsVar  as (type text meta [
        Documentation.FieldCaption = "ID представления Google Analytics:",
        Documentation.SampleValues = {"ga:123456789"}
    ]),
     token  as (type text meta [
        Documentation.FieldCaption = "Ваш refresh token:",
        Documentation.FieldDescription = "Enter your Google Analytics refresh Token",
        Documentation.SampleValues = {"1//gfgFG33qJfk23mbh1fghKMnqhh3p"}
    ]),
     
 metricsVar  as (type text meta [
        Documentation.FieldCaption = "Список показателей через запятую:",
        Documentation.SampleValues = {"ga:sessions,ga:goal2Completions"}
    ]),

     dimensionsVar  as (type text meta [
        Documentation.FieldCaption = "Список параметров через запятую:",
        Documentation.SampleValues = {"ga:date,ga:campaign"}
    ]),

    startDateVar  as (type text meta [
        Documentation.FieldCaption = "Дата начала:",
        Documentation.SampleValues = {"2020-02-28"}
    ]),

     endDateVar  as (type text meta [
        Documentation.FieldCaption = "Дата окончания:",
        Documentation.SampleValues = {"yesterday"}
    ]),


     intervalsVar as (type number meta [
        Documentation.FieldCaption = "Интервал отсутствия семплирования в днях:",
        Documentation.SampleValues = {"3"}
    ]),

 filtersVar as (type  nullable text meta [
        Documentation.FieldCaption = "Фильтры Google Analytics:",
        Documentation.SampleValues = {"ga:source=@yandex"}
    ]),


   segmentVar as (type nullable  text meta [
        Documentation.FieldCaption = "Сегменты Google Analytics:",
        Documentation.SampleValues = {"gaid::-10"}

    ])
    

    )
    as table meta [
        Documentation.Name = "Забираем данные из Google Analytics",
        Documentation.LongDescription = 
        "",
        Documentation.Examples = {
        [
            Description = "Возвращает данные из Google Analytics по сессиям и Цели 2",
            Code = "GetGoogleAnalyticsBL(""ga:17201325"", ""ga:sessions,ga:goal2Completions"", ""ga:date"", ""2020-01-01"", ""yesterday"", null, null, 3)",
            Result = "Данные по сессиям и цели, разбитые по датам."
        ],
        [
            Description = "Возвращает пути конверсии по всем конверсиям",
            Code = "GetGoogleAnalyticsBL(""ga:17201325"", ""mcf:totalConversions"", ""mcf:sourceMediumPath"", ""500daysAgo"", ""yesterday"", null, null, 20)",
            Result ="Отчёт ""пути конверсий""."
        ],
        [
            Description = "Первая ссылка - на список параметров и показателей. Вторая ссылка - на список фильтров.",
            Code = "https://ga-dev-tools.appspot.com/dimensions-metrics-explorer/",
            Result = "https://developers.google.com/analytics/devguides/reporting/core/v3/reference#filters"
                    ],

[
            Description = "Первая ссылка - на MCF API. Вторая ссылка - на сайт автора функции.",
            Code = "https://developers.google.com/analytics/devguides/reporting/mcf/dimsmets?hl=ru",
            Result = "https://learn.needfordata.ru/"
                    ]

        }
    ];

getToken= (tkn as text)=>
let
tok=Text.Replace(tkn,"\",""),
app_credentials ="client_id="&client_id&"&client_secret="&client_secret,
url = app_credentials & "&refresh_token=" & tok & "&grant_type=refresh_token",
GetJson = Web.Contents("https://accounts.google.com/o/oauth2/token",  [ManualStatusHandling={404,400},Headers = [#"Content-Type"="application/x-www-form-urlencoded"], Content = Text.ToBinary(url) ]),
FormatAsJson = try Json.Document(GetJson)[access_token] otherwise "Bad refresh token"
in
FormatAsJson;





GetGoogleAnalyticsBL.Stat = (idsVar as text, token as text, metricsVar as text, dimensionsVar as  text, startDateVar as text, endDateVar as text, intervalsVar as number,  filtersVar as nullable text,  segmentVar as nullable text )  =>
     
 let
   final_token=  getToken(token),

 fn_configRecord2 = (startDate, endDate) =>

let
    tokenBuffer = Text.From(final_token), 
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

    checkDate = if Date.AddDays( Date.From(Uri.Parts(GaJsonGet[selfLink])[Query][#"end-date"]), -intervalsVar -1) >= 
Date.From(dateFunc(startDateVar))
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
        chunkStart = try GaJsonGet[query][#"start-date"] otherwise null,
        chunkEnd = try GaJsonGet[query][#"end-date"] otherwise null,
        nextPage = try Uri.Parts(GaJsonGet[nextLink])[Query] & [access_token = final_token] 
            otherwise fn_configRecord2(Date.From(dateFunc(startDateVar)), checkDate), 
        hadPreviousParameter = lastParameter, 
        SampleSize = try GaJsonGet[sampleSpace] otherwise null
]     
in
    res,


//    ids = "ga:72013428", metrics = "ga:sessions", dimensions = "ga:medium,ga:hostname,ga:date,ga:browser",  startDate = "2012-01-01", endDate = "2016-11-01", sort = null, filters = null, segment = null, 
    tokenBuffer = Text.From(final_token), 
    
    generateDataConfig = List.Generate(
        ()=>[res = fn_getData2(fn_configRecord2(Date.From(dateFunc(startDateVar)), Date.From(dateFunc(endDateVar))), [])],
        each [res][hadPreviousParameter] <> null,
        each [res = fn_getData2([res][nextPage], [res][nextPage])]
    ),
    #"Converted to Table" = Table.FromList(generateDataConfig, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"res"}, {"res"}),
    #"Expanded res" = Table.ExpandRecordColumn(#"Expanded Column1", "res", {"data",  "samplingTrueOrFalse", "chunkStart", "chunkEnd"}, {"data", "samplingTrueOrFalse", "chunkStart", "chunkEnd"}),
    #"Expanded data" = Table.ExpandTableColumn(#"Expanded res", "data", Table.ColumnNames(#"Expanded res"[data]{0}) )
in
    #"Expanded data";


dateFunc = (dat)=>
       if dat = "yesterday" then Date.ToText(Date.AddDays(Date.From(DateTime.LocalNow()) , -1),"yyyy-MM-dd") 
       else if dat = "today" then Date.ToText(Date.From(DateTime.LocalNow()),"yyyy-MM-dd")
       else if (try Number.From(Text.BeforeDelimiter(dat,"daysAgo")) otherwise 0) >0 then Date.ToText(Date.AddDays(Date.From(DateTime.LocalNow()),-1*Number.From(Text.BeforeDelimiter(dat,"daysAgo"))),"yyyy-MM-dd") 
       else dat;




GetGoogleAnalyticsBL = [
    TestConnection = (dataSourcePath) => { "GetGoogleAnalyticsBL.Data",
    Json.Document(dataSourcePath)[idsVar],
     Json.Document(dataSourcePath)[token],
    Json.Document(dataSourcePath)[metricsVar],
    Json.Document(dataSourcePath)[dimensionsVar],
    Json.Document(dataSourcePath)[startDateVar],
    Json.Document(dataSourcePath)[endDateVar],
    Json.Document(dataSourcePath)[intervalsVar],
  Json.Document(dataSourcePath)[filtersVar]?,
    Json.Document(dataSourcePath)[segmentVar]?
 
    },
    Authentication = [
        // Key = []
        // UsernamePassword = [],
        // Windows = [],
        Implicit = []
    ],
    Label = "GoogleAnalytics BL Connector"
];




GetGoogleAnalyticsBL.UI = [
    Beta = false,
    Category = "Other",
    ButtonText = { "Google Analytics BL Connector", "Google Analytics Connector Help" },
    LearnMoreUrl = "https://www.youtube.com/channel/UCqSLNYLN_3mf6nA1Xf7fu1g/",
    SourceImage = GetGoogleAnalyticsBL.Icons,
    SourceTypeImage = GetGoogleAnalyticsBL.Icons
];



GetGoogleAnalyticsBL.Icons = [
    Icon16 = { Extension.Contents("GetGoogleAnalyticsBL16.png"), Extension.Contents("GetGoogleAnalyticsBL20.png"), Extension.Contents("GetGoogleAnalyticsBL24.png"), Extension.Contents("GetGoogleAnalyticsBL32.png") },
    Icon32 = { Extension.Contents("GetGoogleAnalyticsBL32.png"), Extension.Contents("GetGoogleAnalyticsBL40.png"), Extension.Contents("GetGoogleAnalyticsBL48.png"), Extension.Contents("GetGoogleAnalyticsBL64.png") }
];