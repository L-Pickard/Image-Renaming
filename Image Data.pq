let
    image_folder                = Excel.CurrentWorkbook(){[Name = "Image_Folder"]}[Content]{0}[Column1],
    Source                      = Folder.Files(image_folder),
    ExpandedAttributes          = Table.ExpandRecordColumn(Source, "Attributes", {"Size"}, {"Size"}),
    WithSizeKB                  = Table.AddColumn(ExpandedAttributes, "Size (KB)", each [Size] / 1024, type number),
    WithItemKey                 = Table.AddColumn(WithSizeKB, "Item Key", each Text.BeforeDelimiter([Name], "_", 1), type text),
    WithShot                    = Table.AddColumn(WithItemKey, "Shot", each Text.BetweenDelimiters([Name], "_", ".", 1, 0), type text),
    RemovedOtherColumns         = Table.SelectColumns(WithShot, {"Shot", "Item Key", "Size (KB)", "Name"}),
    ReorderedColumns            = Table.ReorderColumns(RemovedOtherColumns, {"Item Key", "Name", "Shot", "Size (KB)"})
in
    ReorderedColumns
