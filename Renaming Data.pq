let
    JoinType = Excel.CurrentWorkbook(){[Name = "Join_Kind"]}[Content]{0}[Column1],

    JoinVal =
        if JoinType = "Inner" then 0
        else if JoinType = "LeftOuter" then 1
        else if JoinType = "RightOuter" then 2
        else if JoinType = "FullOuter" then 3
        else if JoinType = "LeftAnti" then 4
        else if JoinType = "RightAnti" then 5
        else 1,

    Source              = #"Image Data",
    MergedQueries       = Table.NestedJoin(
                            Source,
                            {"Item Key"},
                            #"Line Data",
                            {"Item Key"},
                            "Line Data",
                            JoinVal
                          ),
    ExpandedLineData    = Table.ExpandTableColumn(
                            MergedQueries,
                            "Line Data",
                            {
                                "Season", "Item Status", "Brand", "Gender", "Style Code", "Theme",
                                "Description", "Description 2", "Colour",
                                "Size 1", "Size 2", "Size 3", "Size 4", "Size 5", "Size 6", "Size 7"
                            },
                            {
                                "Season", "Item Status", "Brand", "Gender", "Style Code", "Theme",
                                "Description", "Description 2", "Colour",
                                "Size 1", "Size 2", "Size 3", "Size 4", "Size 5", "Size 6", "Size 7"
                            }
                          ),
    ReorderedColumns    = Table.ReorderColumns(
                            ExpandedLineData,
                            {
                                "Item Key", "Season", "Item Status", "Brand", "Theme", "Gender", "Style Code",
                                "Description", "Description 2", "Colour",
                                "Size 1", "Size 2", "Size 3", "Size 4", "Size 5", "Size 6", "Size 7",
                                "Size (KB)", "Shot", "Name"
                            }
                          )
in
    ReorderedColumns
