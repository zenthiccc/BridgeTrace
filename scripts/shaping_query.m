// Generic Excel Power Query (M) that was originally run by hand.
// src/shaping.py reproduces these steps in pandas so no Excel refresh is needed.
let
    Source = Excel.Workbook(File.Contents("<PATH TO EXPORT>.xlsx"), null, true),
    RawSheet = Source{[Item = "Sheet1", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(RawSheet, [PromoteAllScalars = true]),
    ChangedType = Table.TransformColumnTypes(PromotedHeaders, {
        {"Department", type text}, {"Employee ID", type text}, {"Employee Name", type text},
        {"Leave Code", type text}, {"Start Date", type date}, {"End Date", type date},
        {"Last Modified", type date}, {"Start Time", type datetime}, {"End Time", type datetime},
        {"Leave Days", type number}, {"Leave Hours", type number}
    }),
    KeptColumns = Table.SelectColumns(ChangedType, {
        "Department", "Employee ID", "Employee Name", "Leave Code", "Start Date", "End Date",
        "Last Modified", "Start Time", "End Time", "Leave Days", "Leave Hours"
    }),
    AddedDocs = Table.AddColumn(KeptColumns, "Supporting Documents", each null),
    AddedNotes = Table.AddColumn(AddedDocs, "Reviewer Notes", each null),
    TargetLeaveOnly = Table.SelectRows(AddedNotes, each List.Contains({"SL", "SLH", "SLX"}, [Leave Code])),
    Sorted = Table.Sort(TargetLeaveOnly, {{"Last Modified", Order.Descending}})
in
    Sorted
