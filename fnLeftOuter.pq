let fnLeftOuter=(TableLeft as text, KeysLeft as text, TableRight as text, KeysRight as text, FieldsRight as text, FieldExpand as text) as table =>
    let
        
        pTableLeft= Excel.CurrentWorkbook(){[Name=TableLeft]}[Content],
        pKeyListLeft=Text.Split(KeysLeft,","),
        pTableRight=Excel.CurrentWorkbook(){[Name=TableRight]}[Content],
        pKeyListRight=Text.Split(KeysRight, ","),
        pListFields=Text.Split(FieldsRight, ","),    
        pExpandField=FieldExpand,
        pNewColumn="__NewColumn__",
        pExpandFieldList=Text.Split(pExpandField,""),
        pExpandFieldTable=List.Transform(pListFields, each pExpandField & "." & _),

        Source = Table.NestedJoin(pTableLeft,pKeyListLeft,pTableRight,pKeyListRight,pNewColumn,JoinKind.LeftOuter),
        
        RenameTmpColumns = Table.RenameColumns(Source,{{pNewColumn, pExpandField}}),   

        Expand = Table.ExpandTableColumn(RenameTmpColumns, pExpandField, pListFields, pExpandFieldTable)

    in
        Expand
in
    fnLeftOuter