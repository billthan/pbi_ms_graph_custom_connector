section MyGraph;

//
// OAuth configuration settings
//
client_id = Text.FromBinary(Extension.Contents("client_id"));
tenant_id = Text.FromBinary(Extension.Contents("tenant_id"));
resource = "https://graph.microsoft.com/"; // Microsoft Graph resource URI

[DataSource.Kind="MyGraph", Publish="MyGraph.UI"]
shared MyGraph.Contents = Value.ReplaceType(MyGraphImpl, MyGraphType);

MyGraphImpl = (graphUrl as text) as table =>

let
    accessToken = Text.From(Extension.CurrentCredential()[access_token]),  
    
    // Recursively get tables
    GetData = (url as text) =>
    let
        response = try Web.Contents(url, [
            Headers = [
                Authorization = "Bearer " & accessToken,
                #"Content-Type" = "application/json"
            ]
        ]) otherwise null,

        result = if response <> null then
                    let
                        json = Json.Document(response),  
                        // Extract the records and the next link if present
                        values = json[value],
                        nextLink = try Record.Field(json, "@odata.nextLink") otherwise null,
                                                valuesTable = Table.FromList(values, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
                        
                        combined = if nextLink <> null then
                                       valuesTable & @GetData(nextLink)
                                    else
                                       valuesTable
                    in
                        combined
                 else
                    error "No response or error"
    in
        result,

    resultData = GetData(graphUrl)

in
    resultData;

MyGraphType = type function (
    graphUrl as (type text meta [
        Documentation.FieldCaption = "Graph API URL",
        Documentation.FieldDescription = "Enter the Microsoft Graph API URL",
        Documentation.SampleValues = {"https://graph.microsoft.com/v1.0/me"}
    ]))
    as table meta [
        Documentation.Name = "Microsoft Graph Connector",
        Documentation.LongDescription = "Connects to Microsoft Graph API using Aad authentication",
        Documentation.Examples = {[ 
            Description = "Fetches user details",
            Code = "MyGraph.Contents(""https://graph.microsoft.com/v1.0/me"")"
        ]}
    ];

//
// Data Source definition
//
MyGraph = [
    TestConnection = (dataSourcePath) => { "MyGraph.Contents", dataSourcePath },
    Authentication = [
        Aad = [
            AuthorizationUri = "https://login.microsoftonline.com/" & tenant_id & "/oauth2/authorize",
            Resource = resource,
            Scope = "User.Read.All ServicePrincipalEndpoint.Read.All Group.Read.All"
        ]
    ],
    Label = "Microsoft Graph Connector"
];

//
// UI Export definition
//
MyGraph.UI = [
    Beta = true,
    ButtonText = { "MyGraph.Contents", "Connect to Microsoft Graph" },
    SourceImage = MyGraph.Icons,
    SourceTypeImage = MyGraph.Icons
];

MyGraph.Icons = [
    Icon16 = { Extension.Contents("MyGraph16.png"), Extension.Contents("MyGraph20.png"), Extension.Contents("MyGraph24.png"), Extension.Contents("MyGraph32.png") },
    Icon32 = { Extension.Contents("MyGraph32.png"), Extension.Contents("MyGraph40.png"), Extension.Contents("MyGraph48.png"), Extension.Contents("MyGraph64.png") }
];
