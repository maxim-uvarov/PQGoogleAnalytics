section GetGoogleAnalytics;


client_id="299222177909-e41dco25ctoh065p4gmh0lrt9demb20h.apps.googleusercontent.com";
client_secret="zwUbG2pRDYD1AGI0ljcuDhxJ";

redirect_uri = "https://preview.powerbi.com/views/oauthredirect.html";
windowWidth = 1200;
windowHeight = 1000;


[DataSource.Kind="GetGoogleAnalytics", Publish="GetGoogleAnalytics.UI"]

shared GetGoogleAnalytics.Data = Value.ReplaceType(GetGoogleAnalytics.Stat, FuncType);

FuncType =type function (
       idsVar  as (type text meta [
        Documentation.FieldCaption = "ID представления Google Analytics:",
        Documentation.SampleValues = {"ga:123456789"}
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

optional filtersVar as (type   text meta [
        Documentation.FieldCaption = "Фильтры Google Analytics:",
        Documentation.SampleValues = {"ga:source=@yandex"}
    ]),


 optional  segmentVar as (type   text meta [
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
            Code = "GetGoogleAnalytics(""ga:17201325"", ""ga:sessions,ga:goal2Completions"", ""ga:date"", ""2020-01-01"", ""yesterday"", null, null, 3)",
            Result = "Данные по сессиям и цели, разбитые по датам."
        ],
        [
            Description = "Возвращает пути конверсии по всем конверсиям",
            Code = "GetGoogleAnalytics(""ga:17201325"", ""mcf:totalConversions"", ""mcf:sourceMediumPath"", ""500daysAgo"", ""yesterday"", null, null, 20)",
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






GetGoogleAnalytics.Stat = (idsVar as text, metricsVar as text, dimensionsVar as  text, startDateVar as text, endDateVar as text, intervalsVar as number, optional filtersVar as text, optional segmentVar as  text )  =>
     
 let


 fn_configRecord2 = (startDate, endDate) =>

let
    tokenBuffer = Text.From(Extension.CurrentCredential()[access_token]), 
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
        nextPage = try Uri.Parts(GaJsonGet[nextLink])[Query] & [access_token = Extension.CurrentCredential()[access_token]] 
            otherwise fn_configRecord2(Date.From(dateFunc(startDateVar)), checkDate), 
        hadPreviousParameter = lastParameter, 
        SampleSize = try GaJsonGet[sampleSpace] otherwise null
]     
in
    res,


//    ids = "ga:72013428", metrics = "ga:sessions", dimensions = "ga:medium,ga:hostname,ga:date,ga:browser",  startDate = "2012-01-01", endDate = "2016-11-01", sort = null, filters = null, segment = null, 
    tokenBuffer = Text.From(Extension.CurrentCredential()[access_token]), 
    
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




GetGoogleAnalytics = [
    TestConnection = (dataSourcePath) => { "GetGoogleAnalytics.Data",
    Json.Document(dataSourcePath)[idsVar],
    Json.Document(dataSourcePath)[metricsVar],
    Json.Document(dataSourcePath)[dimensionsVar],
    Json.Document(dataSourcePath)[startDateVar],
    Json.Document(dataSourcePath)[endDateVar],
    Json.Document(dataSourcePath)[intervalsVar],
  Json.Document(dataSourcePath)[filtersVar]?,
    Json.Document(dataSourcePath)[segmentVar]?
 
    },
    Authentication = [
        OAuth = [
            StartLogin = StartLogin,
            FinishLogin = FinishLogin,
            Logout = Logout,
            Refresh = Refresh,
            Label = "Google Analytics Auth"
        ]
    ],
    Label = "Google Analytics Connector"
];


StartLogin = (resourceUrl, state, display) =>
    let
        AuthorizeUrl = "https://accounts.google.com/o/oauth2/v2/auth?" & Uri.BuildQueryString([
            scope = "https://www.googleapis.com/auth/analytics.readonly",
            access_type = "offline",
            include_granted_scopes = "true",
            client_id = client_id,
            state = state,
            redirect_uri = redirect_uri,
            prompt="consent",
            response_type = "code"])
    in
        [
            LoginUri = AuthorizeUrl,
            CallbackUri = redirect_uri,
            WindowHeight = windowHeight,
            WindowWidth = windowWidth,
            Context = null
        ];


FinishLogin = (context, callbackUri, state) =>
    let
        Parts = Uri.Parts(callbackUri)[Query]
    in
        TokenMethod("authorization_code","code" ,Parts[code]);



TokenMethod = (grantType,tokenField,code) =>
    let
        string1=[client_id = client_id,
                client_secret = client_secret,
                grant_type = grantType, 
                redirect_uri = redirect_uri],
        string2=Record.AddField(string1,tokenField,code),
        
        Response = Web.Contents("https://accounts.google.com/o/oauth2/token", [
            Content = Text.ToBinary(Uri.BuildQueryString(string2)),
            Headers=[#"Content-type" = "application/x-www-form-urlencoded", #"Accept" = "application/json"]]),
        Parts = Json.Document(Response)
    in
        Parts;


Refresh = (resourceUrl, refresh_token) => TokenMethod("refresh_token", "refresh_token", refresh_token);

Logout = (token) => "https://accounts.google.com/o/oauth2/revoke?token=" & token;

GetGoogleAnalytics.UI = [
    Beta = false,
    Category = "Other",
    ButtonText = { "Google Analytics Connector", "Google Analytics Connector Help" },
    LearnMoreUrl = "https://www.youtube.com/channel/UCqSLNYLN_3mf6nA1Xf7fu1g/",
    SourceImage = GetGoogleAnalytics.Icons,
    SourceTypeImage = GetGoogleAnalytics.Icons
];



GetGoogleAnalytics.Icons = [
    Icon16 = { Extension.Contents("GoogleAnalytics16.png"), Extension.Contents("GoogleAnalytics20.png"), Extension.Contents("GoogleAnalytics24.png"), Extension.Contents("GoogleAnalytics32.png") },
    Icon32 = { Extension.Contents("GoogleAnalytics32.png"), Extension.Contents("GoogleAnalytics40.png"), Extension.Contents("GoogleAnalytics48.png"), Extension.Contents("GoogleAnalytics64.png") }
];