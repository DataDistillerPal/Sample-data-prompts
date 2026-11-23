// =====================================================================
// BACKLOG REDUCTION SUCCESSES - "Trend Details" version
// Power Query (M) - every Summary number is calculated here.
//
// Create each block below as its own query (Home > New Source > Blank
// Query > Advanced Editor), using the query name shown in the header.
//
// DESIGN: SOURCE QUERIES -> DIMENSIONS / FACT -> SUMMARY
//   Today every source is a sheet in this workbook (self-contained, the
//   same lists the Excel formulas use). Later the pipeline will deliver:
//     - Staff        as a SQL table (raw staff columns only)
//     - Titles       as a SQL dimension
//     - Office       as a SQL dimension
//     - Pay Periods  as a SQL dimension
//   Only the four src* queries read a source. Each one returns FIXED
//   column names. When the SQL tables arrive, replace the body of the
//   src* query (an example is in each one) and keep the same output
//   column names - nothing downstream changes.
//
// Loaded to the model : Summary, Settings, Staff, Titles, Office, Pay Periods
// Not loaded          : everything else (right-click > uncheck "Enable load")
//
// MODEL RELATIONSHIPS (many-to-one, single direction):
//   Staff[TitleKey]    -> Titles[TitleKey]
//   Staff[OfficeKey]   -> Office[OfficeKey]
//   Staff[Pay Period]  -> Pay Periods[Pay Period]
//   Summary and Settings stay disconnected (already calculated).
//
// Sheet columns used (0-based positions in M: A=0, B=1 ... V=21):
//   Trends Details : B Case_Number, N Inv Created Dt, S Inv Action Type,
//                    V Inv Disposition Dt
//   Staff          : A Pay Period, B Org Desc, C..G staff details,
//                    H Position Title Opm  (I..M are formulas - rebuilt)
//                    Q position -> R title code   (list used by Titles I)
//                    T Org Desc -> U region       (list used by Region J)
//   Pay Periods    : A Pay Period, B Start, C End, D Fiscal Year, E Month
//   BTP Case List  : C Transferred To Region, F Inv. Determination Date
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


// =====================================================================
// HELPERS
// =====================================================================

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
// fnToWhole - same helper as the operations build
// Turns 1, "1", "01", 1.0 into 1; anything else into null.
// ---------------------------------------------------------------------
(v as any) as nullable number =>
    let
        t = if v = null then null else try Text.Trim(Text.From(v)) otherwise null,
        n = if t = null or t = "" then null else try Number.From(t) otherwise null
    in
        if n = null then null else Int64.From(n)


// ---------------------------------------------------------------------
// fnKey - one join key for every text lookup
// Upper case (XLOOKUP ignores case), removes non-printing characters,
// trims, collapses double / non-breaking spaces. Blank -> null.
// Used on BOTH sides of every join.
// ---------------------------------------------------------------------
(v as any) as nullable text =>
    let
        t     = if v = null then null else try Text.From(v) otherwise null,
        parts = if t = null then {} else Text.SplitAny(Text.Clean(t), " #(00A0)#(tab)"),
        k     = Text.Upper(Text.Combine(List.Select(parts, each _ <> ""), " "))
    in
        if k = "" then null else k


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


// =====================================================================
// SOURCE QUERIES (not loaded) - the ONLY place a source is read
// =====================================================================

// ---------------------------------------------------------------------
// srcStaff
// Output: Pay Period, Org Desc, <C..G as in the sheet>, Position Title Opm
// (raw staff columns only - the same shape the SQL Staff table will have)
// ---------------------------------------------------------------------
let
    S     = fnSheet("Staff"),
    C     = Table.ColumnNames(S),
    AH    = Table.SelectColumns(S, List.FirstN(C, 8)),                 // A:H only
    Named = Table.RenameColumns(AH, {{C{0}, "Pay Period"}, {C{1}, "Org Desc"}, {C{7}, "Position Title Opm"}}),
    // drop fully blank rows (rows that only held the Q:U lists)
    Rows  = Table.SelectRows(Named, each [Pay Period] <> null or [Position Title Opm] <> null)
in
    Rows
// LATER (SQL):
// let
//     Src  = Sql.Database("server", "database"),
//     T    = Src{[Schema = "dbo", Item = "Staff"]}[Data],
//     Rows = Table.RenameColumns(T, {{"PayPeriod", "Pay Period"}, {"OrgDesc", "Org Desc"},
//                                   {"PositionTitle", "Position Title Opm"}})
// in  Rows


// ---------------------------------------------------------------------
// srcTitles
// Output: Position, Title
// Today: the Q:R list on the Staff sheet (what the Titles formula uses).
// ---------------------------------------------------------------------
let
    S     = fnSheet("Staff"),
    C     = Table.ColumnNames(S),
    Sel   = Table.RenameColumns(Table.SelectColumns(S, {C{16}, C{17}}), {{C{16}, "Position"}, {C{17}, "Title"}}),
    Rows  = Table.SelectRows(Sel, each [Position] <> null)
in
    Rows
// LATER (SQL): Sql.Database(...){[Schema = "dbo", Item = "Titles"]}[Data]
//              renamed to Position, Title (keep Series if you need it)


// ---------------------------------------------------------------------
// srcOffice
// Output: Org Desc, Region
// Today: the T:U list on the Staff sheet (what the Region formula uses).
// ---------------------------------------------------------------------
let
    S     = fnSheet("Staff"),
    C     = Table.ColumnNames(S),
    Sel   = Table.RenameColumns(Table.SelectColumns(S, {C{19}, C{20}}), {{C{19}, "Org Desc"}, {C{20}, "Region"}}),
    Rows  = Table.SelectRows(Sel, each [Org Desc] <> null)
in
    Rows
// LATER (SQL): Sql.Database(...){[Schema = "dbo", Item = "Office"]}[Data]
//              renamed to Org Desc, Region


// ---------------------------------------------------------------------
// srcPayPeriods
// Output: Pay Period, Start Date, End Date, Fiscal Year, Month
// ---------------------------------------------------------------------
let
    S     = fnSheet("Pay Periods"),
    C     = Table.ColumnNames(S),
    Sel   = Table.SelectColumns(S, List.FirstN(C, 5)),
    Named = Table.RenameColumns(Sel, {{C{0}, "Pay Period"}, {C{1}, "Start Date"}, {C{2}, "End Date"},
                                      {C{3}, "Fiscal Year"}, {C{4}, "Month"}})
in
    Named
// LATER (SQL): Sql.Database(...){[Schema = "dbo", Item = "PayPeriods"]}[Data]
//              renamed to the five names above


// =====================================================================
// DIMENSIONS (loaded)
// Each is unique on its key (first row wins, like XLOOKUP), so a join
// can never add Staff rows and the relationship is a clean many-to-one.
// =====================================================================

// ---------------------------------------------------------------------
// Titles   Position, Title, TitleKey
// "(Unmatched)" row with Title = "Check" for staff titles not in the list
// ---------------------------------------------------------------------
let
    S       = srcTitles,
    Typed   = Table.TransformColumns(S, {{"Title", each if _ = null then null else Text.Trim(Text.From(_)), type text},
                                        {"Position", each Text.From(_), type text}}),
    Keyed   = Table.AddColumn(Typed, "TitleKey", each fnKey([Position]), type text),
    Valid   = Table.SelectRows(Keyed, each [TitleKey] <> null),
    Unique  = Table.Distinct(Valid, {"TitleKey"}),
    Unknown = #table(type table [Position = text, Title = text, TitleKey = text],
                     {{"(Unmatched)", "Check", "<UNMATCHED>"}}),
    Final   = Table.Combine({Unique, Unknown})
in
    Final


// ---------------------------------------------------------------------
// Office   Org Desc, Region, OfficeKey
// Region is a whole number (same as the operations build): "OPS" -> blank
// ---------------------------------------------------------------------
let
    S       = srcOffice,
    Typed   = Table.TransformColumns(S, {{"Org Desc", each Text.From(_), type text},
                                        {"Region", fnToWhole, Int64.Type}}),
    Keyed   = Table.AddColumn(Typed, "OfficeKey", each fnKey([Org Desc]), type text),
    Valid   = Table.SelectRows(Keyed, each [OfficeKey] <> null),
    Unique  = Table.Distinct(Valid, {"OfficeKey"}),
    Unknown = #table(type table [#"Org Desc" = text, Region = Int64.Type, OfficeKey = text],
                     {{"(Unmatched)", null, "<UNMATCHED>"}}),
    Final   = Table.Combine({Unique, Unknown})
in
    Final


// ---------------------------------------------------------------------
// Pay Periods   Pay Period, Start Date, End Date, Fiscal Year, Month
// Sheet order is kept (XLOOKUP first-match in the agent calculation).
// ---------------------------------------------------------------------
let
    S      = srcPayPeriods,
    Whole  = Table.TransformColumns(S, {{"Pay Period", fnToWhole, Int64.Type}, {"Fiscal Year", fnToWhole, Int64.Type}}),
    Typed  = Table.TransformColumnTypes(Whole, {{"Start Date", type date}, {"End Date", type date}, {"Month", type text}}),
    Rows   = Table.SelectRows(Typed, each [Pay Period] <> null),
    Unique = Table.Distinct(Rows, {"Pay Period"})
in
    Unique


// =====================================================================
// FACT
// =====================================================================

// ---------------------------------------------------------------------
// Staff  (loaded)
// Raw columns A:H, the rebuilt formula columns, and the relationship keys:
//   Titles      = XLOOKUP(Position Title Opm, Q, R)   -> "Check" if not found
//   Region      = XLOOKUP(Org Desc, T, U)
//   Agent       = IF(OR(Titles="FE", Titles="FA"), "BA", "")
//   Month       = Pay Periods[Month]
//   Fiscal Year = Pay Periods[Fiscal Year]
//   TitleKey / OfficeKey = "<UNMATCHED>" when not in the dimension, so every
//   row relates to a real dimension row.
// Region, Month and Fiscal Year are kept for Excel parity; with the
// relationships in place they can be dropped later and read from the dims.
// ---------------------------------------------------------------------
let
    S        = srcStaff,
    TypedPP  = Table.TransformColumns(S, {{"Pay Period", fnToWhole, Int64.Type}}),
    RowsIn   = Table.RowCount(TypedPP),

    Keys     = Table.AddColumn(
                   Table.AddColumn(TypedPP, "__TKey", each fnKey([Position Title Opm]), type text),
                   "__OKey", each fnKey([Org Desc]), type text),

    TitlesD  = Table.Buffer(Table.SelectColumns(Titles, {"TitleKey", "Title"})),
    OfficeD  = Table.Buffer(Table.SelectColumns(Office, {"OfficeKey", "Region"})),
    PPD      = Table.Buffer(Table.SelectColumns(#"Pay Periods", {"Pay Period", "Month", "Fiscal Year"})),

    JT       = Table.ExpandTableColumn(Table.NestedJoin(Keys, {"__TKey"},     TitlesD, {"TitleKey"},   "__t", JoinKind.LeftOuter), "__t", {"TitleKey", "Title"}, {"TitleKey", "Titles"}),
    JO       = Table.ExpandTableColumn(Table.NestedJoin(JT,   {"__OKey"},     OfficeD, {"OfficeKey"},  "__o", JoinKind.LeftOuter), "__o", {"OfficeKey", "Region"}),
    JP       = Table.ExpandTableColumn(Table.NestedJoin(JO,   {"Pay Period"}, PPD,     {"Pay Period"}, "__p", JoinKind.LeftOuter), "__p", {"Month", "Fiscal Year"}),

    Fixed    = Table.TransformColumns(JP, {
                   {"TitleKey",  each _ ?? "<UNMATCHED>", type text},
                   {"OfficeKey", each _ ?? "<UNMATCHED>", type text},
                   {"Titles",    each _ ?? "Check",       type text}}),
    // Agent: Excel "=" ignores case
    AddAgent = Table.AddColumn(Fixed, "Agent", each
                   if Text.Upper([Titles]) = "FE" or Text.Upper([Titles]) = "FA" then "BA" else "", type text),

    Cleaned  = Table.RemoveColumns(AddAgent, {"__TKey", "__OKey"}),
    Ordered  = Table.ReorderColumns(Cleaned,
                   Table.ColumnNames(S) & {"Titles", "Region", "Agent", "Month", "Fiscal Year", "TitleKey", "OfficeKey"}),
    Typed    = Table.TransformColumnTypes(Ordered, {{"Region", Int64.Type}, {"Month", type text}, {"Fiscal Year", Int64.Type}}),

    // Safety net: a join must never change the number of Staff rows
    Checked  = if Table.RowCount(Typed) = RowsIn then Typed
               else error Error.Record("Staff join check",
                    "Row count changed from " & Text.From(RowsIn) & " to " & Text.From(Table.RowCount(Typed)) &
                    " - a dimension has duplicate keys.")
in
    Checked


// =====================================================================
// SUMMARY CALCULATION
// =====================================================================

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
//   report date AND the pay period appears in Staff;
//   take the latest end date, look up its pay period (first match, like
//   XLOOKUP), count Staff rows for that pay period with Agent = "BA".
//   No pay period or zero agents -> IFERROR fallback of 373.
// ---------------------------------------------------------------------
let
    PP      = Table.SelectColumns(#"Pay Periods", {"Pay Period", "End Date"}),
    PPNum   = Table.AddColumn(PP, "EndSerial", each
                  if [End Date] = null then null else Number.From([End Date]), type nullable number),
    PPIdx   = Table.Buffer(PPNum),

    // COUNTIF(staffPP, code) > 0  and  COUNTIFS(staffPP, code, staffAgent, "BA")
    StaffCounts  = Table.Buffer(Table.Group(Table.SelectColumns(Staff, {"Pay Period", "Agent"}), {"Pay Period"},
                       {{"BACount", each List.Count(List.Select([Agent], each _ = "BA")), Int64.Type}})),
    StaffedCodes = List.Buffer(StaffCounts[Pay Period]),

    Staffed = Table.Buffer(Table.SelectRows(PPIdx, each
                  [EndSerial] <> null and List.Contains(StaffedCodes, [Pay Period]))),

    Dates   = Table.SelectColumns(qReportDates, {"ReportDate"}),

    AddBA   = Table.AddColumn(Dates, "BAAgents", (r) =>
        let
            d          = Number.From(r[ReportDate]),
            monthStart = Number.From(Date.StartOfMonth(r[ReportDate])),
            eligible   = List.Select(Staffed[EndSerial], each _ >= monthStart and _ <= d),
            latestEnd  = if List.IsEmpty(eligible) then null else List.Max(eligible),
            // XLOOKUP against the full Pay Periods list: first pay period with that end date
            latestPP   = if latestEnd = null then null
                         else Table.First(Table.SelectRows(PPIdx, each [EndSerial] = latestEnd))[Pay Period],
            baCount    = if latestPP = null then null
                         else List.First(Table.SelectRows(StaffCounts, each [Pay Period] = latestPP)[BACount], 0)
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
    Src  = fnSheet("BTP Case List"),
    ColC = Table.Column(Src, Table.ColumnNames(Src){2}),
    Cnt  = List.Count(List.Select(ColC, each (try (_ <> null and _ <> "") otherwise true)))
in
    Cnt


// ---------------------------------------------------------------------
// qBTPDispositions  (not loaded)  ->  feeds D7
// Disposition dates in column F that are numbers > 0
// (COUNTIFS ">0" ignores text such as "-" and blanks).
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
