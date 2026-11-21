// =====================================================================
// BACKLOG REDUCTION SUCCESSES - "Trend Details" version
// Power Query (M) - every Summary number is calculated here.
//
// Create each block below as its own query (Home > New Source > Blank
// Query > Advanced Editor), using the query name shown in the header.
// Only "Summary" and "Settings" are loaded to the model; everything
// else: right-click > uncheck "Enable load".
//
// Column positions follow the formulas in the workbook:
//   Trends Details : B = case number, N = action created date,
//                    S = action type ("Initial C"), V = disposition date
//                    (B Case_Number, N Inv Created Dt, S Inv Action Type,
//                     V Inv Disposition Dt - confirmed from the sheet)
//   Staff          : A = Pay Period, I = Titles (K "Agent" = BA when I is FE/FA)
//   Pay Periods    : A = pay period code, C = pay period end date
//   BTP Case List  : C = Transferred To Region, F = Inv. Determination Date
//                    ("-" in F is text, so it is ignored like COUNTIFS ">0")
// Positions are 0-based in M (A=0, B=1 ... K=10, N=13, S=18, V=21).
// NOTE: Excel.Workbook drops leading EMPTY columns. If column A of a
// sheet is completely empty, subtract 1 from that sheet's positions.
// =====================================================================


// ---------------------------------------------------------------------
// PARAMETERS (Home > Manage Parameters > New)
// ---------------------------------------------------------------------
// pFilePath           Text     full path to the customer workbook
// pBaselineDate       Date     1/7/2026      (cell AK5 in Excel)
// pBTPStartDate       Date     5/1/2026      (N/A cut-off used in C7/D7)
// pFallbackAgents     Decimal  373           (IFERROR fallback in C6/D6)
// pPercentDenominator Decimal  20371         (hard-coded in E5 =D5/20371)


// ---------------------------------------------------------------------
// fnSheet  - returns one sheet with headers promoted
// ---------------------------------------------------------------------
(sheetName as text) as table =>
let
    Source   = Excel.Workbook(File.Contents(pFilePath), null, true),
    Sheet    = Source{[Item = sheetName, Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true])
in
    Promoted


// ---------------------------------------------------------------------
// fnExcelNum - turns a cell into the number Excel would compare with.
// Dates/datetimes become Excel serials (time of day kept), numbers stay
// numbers, anything else (text, blank) returns null.
// Mirrors ISNUMBER() and the >= / <= date tests in the formulas.
// ---------------------------------------------------------------------
(v as any) as nullable number =>
    if v is datetime then Number.From(v)
    else if v is date then Number.From(v)
    else if v is number then v
    else null


// ---------------------------------------------------------------------
// qCohort  (not loaded)
// The cases behind C5 and D5:
//   unique case numbers with an "Initial C" action created on or before
//   the baseline, and no disposition date or a disposition after it.
// One row per case, with the date the case stops being open.
// OpenUntilSerial = null means it is still open (blank disposition).
// ---------------------------------------------------------------------
let
    Src      = fnSheet("Trends Details"),
    Cols     = Table.ColumnNames(Src),
    Kept     = Table.SelectColumns(Src, {Cols{1}, Cols{13}, Cols{18}, Cols{21}}),
    Renamed  = Table.RenameColumns(Kept, {
                   {Cols{1},  "CaseNumber"},
                   {Cols{13}, "ActionCreated"},
                   {Cols{18}, "ActionType"},
                   {Cols{21}, "Disposition"}}),

    BaselineSerial = Number.From(pBaselineDate),

    // S = "Initial C"  (Excel "=" is not case-sensitive)
    InitialC = Table.SelectRows(Renamed, each
                   (try ([ActionType] is text and Text.Lower([ActionType]) = "initial c") otherwise false)),

    // ISNUMBER(N) * (N <= baseline)
    AddCreated = Table.AddColumn(InitialC, "CreatedSerial",
                   each try fnExcelNum([ActionCreated]) otherwise null, type nullable number),
    CreatedOk  = Table.SelectRows(AddCreated, each
                   [CreatedSerial] <> null and [CreatedSerial] <= BaselineSerial),

    // V = "" -> open. A text value in V also counts as open, because
    // Excel treats text as greater than any date in V >= date.
    AddDisp  = Table.AddColumn(CreatedOk, "DispSerial",
                   each try fnExcelNum([Disposition]) otherwise null, type nullable number),

    // (V = "") + (V >= baseline + 1)
    InCohort = Table.SelectRows(AddDisp, each
                   [DispSerial] = null or [DispSerial] >= BaselineSerial + 1),

    // UNIQUE(case number). A case is open on a date if ANY of its
    // qualifying rows is still open, so keep the latest disposition
    // (or null if any row has none).
    CaseText = Table.TransformColumns(InCohort,
                   {{"CaseNumber", each if _ = null then "" else Text.From(_), type text}}),
    PerCase  = Table.Group(CaseText, {"CaseNumber"},
                   {{"OpenUntilSerial",
                     each if List.Contains([DispSerial], null) then null
                          else List.Max([DispSerial]),
                     type nullable number}},
                   GroupKind.Global, Comparer.OrdinalIgnoreCase)
in
    PerCase


// ---------------------------------------------------------------------
// qReportDates  (not loaded - Summary carries these columns)
// The C4 drop-down: the baseline date, then every month end through the
// current month (1/7/2026, 1/31/2026 ... 11/30/2026).
// ---------------------------------------------------------------------
let
    Base          = pBaselineDate,
    FirstMonthEnd = Date.EndOfMonth(Base),
    LastMonthEnd  = Date.EndOfMonth(Date.From(DateTime.LocalNow())),
    MonthEnds     = List.Generate(
                        () => FirstMonthEnd,
                        each _ <= LastMonthEnd,
                        each Date.EndOfMonth(Date.AddMonths(_, 1))),
    AllDates      = List.Distinct({Base} & MonthEnds),
    AsTable       = Table.FromList(AllDates, Splitter.SplitByNothing(),
                        type table [ReportDate = date]),
    Sorted        = Table.Sort(AsTable, {{"ReportDate", Order.Ascending}}),
    WithOrder     = Table.AddIndexColumn(Sorted, "DateOrder", 1, 1, Int64.Type),
    WithLabel     = Table.AddColumn(WithOrder, "ReportDateLabel",
                        each Date.ToText([ReportDate], [Format = "MMMM d, yyyy", Culture = "en-US"]),
                        type text)
in
    WithLabel


// ---------------------------------------------------------------------
// qAgentsByDate  (not loaded)
// The LET() in C6/D6, for every report date:
//   eligible pay periods = end date between the 1st of the month and the
//   report date AND the code appears on the Staff sheet;
//   take the latest end date, look up its code (first match, like
//   XLOOKUP), count Staff rows for that code with K = "BA".
//   No pay period or zero agents -> IFERROR fallback of 373.
// ---------------------------------------------------------------------
let
    // Pay Periods: A = code, C = end date
    PP      = fnSheet("Pay Periods"),
    PPc     = Table.ColumnNames(PP),
    PP1     = Table.SelectColumns(PP, {PPc{0}, PPc{2}}),
    PP2     = Table.RenameColumns(PP1, {{PPc{0}, "PeriodCode"}, {PPc{2}, "EndDate"}}),
    PP3     = Table.TransformColumns(PP2, {
                  {"PeriodCode", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type nullable text},
                  {"EndDate", fnExcelNum, type nullable number}}),
    PP4     = Table.SelectRows(PP3, each [PeriodCode] <> null and [PeriodCode] <> ""),
    PPIdx   = Table.Buffer(Table.AddIndexColumn(PP4, "SheetOrder", 0, 1, Int64.Type)),

    // Staff: A = Pay Period, I = Titles.
    // K (Agent) is itself a formula: =IF(OR(I2="FE",I2="FA"),"BA","")
    // so it is rebuilt here from Titles instead of relying on K.
    // try/otherwise guards against #N/A cells (Region/Titles lookups).
    St      = fnSheet("Staff"),
    Stc     = Table.ColumnNames(St),
    St1     = Table.SelectColumns(St, {Stc{0}, Stc{8}}),
    St2     = Table.RenameColumns(St1, {{Stc{0}, "PeriodCode"}, {Stc{8}, "Titles"}}),
    St3a    = Table.TransformColumns(St2, {
                  {"PeriodCode", each try (if _ = null then null else Text.Upper(Text.Trim(Text.From(_)))) otherwise null, type nullable text},
                  {"Titles",     each try (if _ = null then null else Text.Upper(Text.From(_))) otherwise null, type nullable text}}),
    St3     = Table.AddColumn(St3a, "AgentGroup", each
                  if [Titles] = "FE" or [Titles] = "FA" then "BA" else "", type text),

    // COUNTIF(staffPP, code) > 0  and  COUNTIFS(staffPP, code, staffAgent, "BA")
    StaffCounts = Table.Buffer(Table.Group(St3, {"PeriodCode"},
                  {{"BACount", each List.Count(List.Select([AgentGroup], each _ = "BA")), Int64.Type}})),
    StaffedCodes = List.Buffer(StaffCounts[PeriodCode]),

    // hasStaff filter, and drop rows with no usable end date
    Staffed = Table.Buffer(Table.SelectRows(PPIdx, each
                  [EndDate] <> null and List.Contains(StaffedCodes, [PeriodCode]))),

    Dates   = Table.SelectColumns(qReportDates, {"ReportDate"}),

    AddBA   = Table.AddColumn(Dates, "BAAgents", (r) =>
        let
            d          = Number.From(r[ReportDate]),
            monthStart = Number.From(Date.StartOfMonth(r[ReportDate])),
            eligible   = List.Select(Staffed[EndDate], each _ >= monthStart and _ <= d),
            latestEnd  = if List.IsEmpty(eligible) then null else List.Max(eligible),
            // XLOOKUP against the full Pay Periods list: first code with that end date
            latestPP   = if latestEnd = null then null
                         else Table.First(Table.SelectRows(PPIdx, each [EndDate] = latestEnd))[PeriodCode],
            baCount    = if latestPP = null then null
                         else List.First(Table.SelectRows(StaffCounts, each [PeriodCode] = latestPP)[BACount], 0)
        in
            baCount, type nullable number),

    AddUsed = Table.AddColumn(AddBA, "AgentsUsed", each
                  if [BAAgents] = null or [BAAgents] = 0 then pFallbackAgents else [BAAgents],
                  type number),
    AddFlag = Table.AddColumn(AddUsed, "UsedFallback", each
                  [BAAgents] = null or [BAAgents] = 0, type logical)
in
    AddFlag


// ---------------------------------------------------------------------
// qBTPCaseCount  (not loaded)  ->  C7
// COUNTIF('BTP Case List'!C2:C500000, "<>")
// ---------------------------------------------------------------------
let
    Src = fnSheet("BTP Case List"),
    ColC = Table.Column(Src, Table.ColumnNames(Src){2}),
    Cnt  = List.Count(List.Select(ColC, each (try (_ <> null and _ <> "") otherwise true)))
in
    Cnt


// ---------------------------------------------------------------------
// qBTPDispositions  (not loaded)  ->  feeds D7
// Disposition dates in column F that are numbers > 0
// (COUNTIFS ">0" ignores text and blanks).
// ---------------------------------------------------------------------
let
    Src   = fnSheet("BTP Case List"),
    ColF  = Table.Column(Src, Table.ColumnNames(Src){5}),
    Nums  = List.Transform(ColF, each try fnExcelNum(_) otherwise null),
    Valid = List.Buffer(List.Select(Nums, each _ <> null and _ > 0))
in
    Valid


// ---------------------------------------------------------------------
// Summary  (LOAD)
// One row per report date x metric - the Excel Summary table for every
// value in the C4 drop-down.
// ---------------------------------------------------------------------
let
    Cohort        = List.Buffer(qCohort[OpenUntilSerial]),
    BaselineCount = List.Count(Cohort),          // every cohort case is open at the baseline (20,343)
    BTPTotal      = qBTPCaseCount,
    BTPDisp       = qBTPDispositions,

    Joined     = Table.NestedJoin(qReportDates, {"ReportDate"}, qAgentsByDate, {"ReportDate"}, "A", JoinKind.LeftOuter),
    WithAgents = Table.ExpandTableColumn(Joined, "A", {"BAAgents", "AgentsUsed", "UsedFallback"}),

    AddRows = Table.AddColumn(WithAgents, "Rows", (r) =>
        let
            d        = Number.From(r[ReportDate]),

            // C5: cohort cases still open on the report date (V blank or V >= C4 + 1)
            aged     = List.Count(List.Select(Cohort, each _ = null or _ >= d + 1)),
            // D5: open at baseline minus open now
            running  = BaselineCount - aged,
            agents   = r[AgentsUsed],

            // C7 / D7 / E7
            btpOn    = r[ReportDate] >= pBTPStartDate,
            btpTotal = if btpOn then BTPTotal else null,
            btpDisp  = if btpOn then List.Count(List.Select(BTPDisp, each _ < d + 1)) else null,
            // IFERROR(D7/C7, 0): "N/A"/"N/A" errors, so Excel shows 0.00% before May
            btpPct   = if btpOn and btpTotal <> 0 then btpDisp / btpTotal else 0
        in
            {
                [MetricOrder = 1, Metric = "Aged Pending Cases",
                    AsOfValue = aged,          RunningReductionValue = running,
                    PercentValue = running / pPercentDenominator],              // E5 = D5/20371
                [MetricOrder = 2, Metric = "Per Agent Caseload",
                    AsOfValue = aged / agents, RunningReductionValue = running / agents,
                    PercentValue = null],                                         // merged with E5 in Excel
                [MetricOrder = 3, Metric = "Case Reallocation Initiative",
                    AsOfValue = btpTotal,      RunningReductionValue = btpDisp,
                    PercentValue = btpPct]
            }),

    Slim      = Table.SelectColumns(AddRows, {"ReportDate", "DateOrder", "ReportDateLabel", "BAAgents", "AgentsUsed", "UsedFallback", "Rows"}),
    ExpList   = Table.ExpandListColumn(Slim, "Rows"),
    ExpRec    = Table.ExpandRecordColumn(ExpList, "Rows",
                    {"MetricOrder", "Metric", "AsOfValue", "RunningReductionValue", "PercentValue"}),

    Typed     = Table.TransformColumnTypes(ExpRec, {
                    {"MetricOrder", Int64.Type}, {"Metric", type text},
                    {"AsOfValue", type number}, {"RunningReductionValue", type number},
                    {"PercentValue", type number}, {"BAAgents", type number},
                    {"AgentsUsed", type number}}),

    // Display columns, formatted like the Excel cells (N/A, 9,316, 54.13%)
    AddAsOf   = Table.AddColumn(Typed, "As of", each
                    if [AsOfValue] = null then "N/A" else Number.ToText([AsOfValue], "N0", "en-US"), type text),
    AddRun    = Table.AddColumn(AddAsOf, "Running Reduction", each
                    if [RunningReductionValue] = null then "N/A" else Number.ToText([RunningReductionValue], "N0", "en-US"), type text),
    AddPct    = Table.AddColumn(AddRun, "Total Percent Reduced", each
                    if [PercentValue] = null then "" else Number.ToText([PercentValue], "0.00%", "en-US"), type text),

    Sorted    = Table.Sort(AddPct, {{"DateOrder", Order.Ascending}, {"MetricOrder", Order.Ascending}})
in
    Sorted


// ---------------------------------------------------------------------
// Settings  (LOAD) - one row, used for the header text
// ---------------------------------------------------------------------
let
    T = #table(
        type table [BaselineDate = date, ReductionSinceDate = date,
                    BaselineCount = Int64.Type, PercentDenominator = number],
        {{ pBaselineDate,
           Date.AddDays(pBaselineDate, 1),
           List.Count(qCohort[OpenUntilSerial]),
           pPercentDenominator }})
in
    T
