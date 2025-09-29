let
    file                = Excel.CurrentWorkbook(){[Name = "Line_List_File"]}[Content]{0}[Column1],
    Source              = Excel.Workbook(File.Contents(file), null, true),
    First_Sheet         = Source{0}[Data],
    PromotedHeaders     = Table.PromoteHeaders(First_Sheet, [PromoteAllScalars = true]),
    RenamedColumns      = Table.RenameColumns(
        PromotedHeaders,
        {
            {"Size Range sold As", "Size 1"},
            {"Column11", "Size 2"},
            {"Column12", "Size 3"},
            {"Column13", "Size 4"},
            {"Column14", "Size 5"},
            {"Column15", "Size 6"},
            {"Column16", "Size 7"}
        }
    ),
    ChangedType         = Table.TransformColumnTypes(
        RenamedColumns,
        {
            {"Season", type text},
            {"Item Status", type text},
            {"Brand", type text},
            {"Gender", type text},
            {"Style Code", type text},
            {"Theme", type text},
            {"Description", type text},
            {"Description 2", type text},
            {"Colour", type text},
            {"Size 1", type text},
            {"Size 2", type text},
            {"Size 3", type text},
            {"Size 4", type text},
            {"Size 5", type text},
            {"Size 6", type text},
            {"Size 7", type text},
            {"Notes About Product", type text}
        }
    ),
    TrimmedText         = Table.TransformColumns(
        ChangedType,
        List.Transform(
            {"Colour", "Size 1", "Size 2", "Size 3", "Size 4", "Size 5", "Size 6", "Size 7", "Style Code"},
            each {_, Text.Trim, type text}
        )
    ),
    AddedCustom         = Table.AddColumn(
        TrimmedText,
        "Item Key",
        each [Style Code] & "_" & Text.Replace([Colour], "/", "~")
    ),
    RemovedOtherColumns = Table.SelectColumns(
        AddedCustom,
        {
            "Season", "Item Status", "Brand", "Gender", "Style Code", "Theme",
            "Description", "Description 2", "Colour",
            "Size 1", "Size 2", "Size 3", "Size 4", "Size 5", "Size 6", "Size 7",
            "Item Key"
        }
    ),
    ReorderedColumns    = Table.ReorderColumns(
        RemovedOtherColumns,
        {
            "Item Key", "Season", "Item Status", "Brand", "Gender", "Style Code", "Theme",
            "Description", "Description 2", "Colour",
            "Size 1", "Size 2", "Size 3", "Size 4", "Size 5", "Size 6", "Size 7"
        }
    ),
    #"Change Item Key Type" = Table.TransformColumnTypes(
        ReorderedColumns,
        {{"Item Key", type text}}
    ),

    replacements = {
        {"Size 1", {
            {"UK 8-11 / EU 42-45", "8-11"},
            {"O/S (175cm x 61cm - 5mm)", "O/S"},
            {"4-7 (EU 37-40)", "4-7"},
            {"YOUTH O/S", "O/S"},
            {"O/S (60cm x 60cm)", "O/S"},
            {"US 7", "6"},
            {"US 3", "2"}
        }},
        {"Size 2", {
            {" YRS", ""},
            {"US 4", "3"},
            {"US 8", "7"}
        }},
        {"Size 3", {
            {" YRS", ""},
            {"US 9", "8"},
            {"US 5", "4"}
        }},
        {"Size 4", {
            {" YRS", ""},
            {"US 10", "9"},
            {"US 6", "5"}
        }},
        {"Size 5", {
            {" YRS", ""},
            {"US 11", "10"},
            {"US 7", "6"}
        }},
        {"Size 6", {
            {" YRS", ""},
            {"US 12", "11"}
        }},
        {"Size 7", {
            {" YRS", ""},
            {"US 13", "12"}
        }}
    },

    CleanedSizes = List.Accumulate(
        replacements,
        #"Change Item Key Type",
        (state, current) =>
            List.Accumulate(
                current{1},
                state,
                (s, r) => Table.ReplaceValue(
                    s,
                    r{0},
                    r{1},
                    Replacer.ReplaceText,
                    { current{0} }
                )
            )
    )
in
    CleanedSizes
