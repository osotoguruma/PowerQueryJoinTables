let fnLeftOuter=(TableLeft as text, KeysLeft as text, TableRight as text, KeysRight as text, FieldsRight as text, FieldExpand as text, ModeJoinKind as text) as table =>
    let
        
        pTableLeft= Excel.CurrentWorkbook(){[Name=Text.Trim(TableLeft)]}[Content],
        pKeyListLeft=Text.Split(KeysLeft,","),
        pTableRight=Excel.CurrentWorkbook(){[Name=Text.Trim(TableRight)]}[Content],
        pKeyListRight=Text.Split(Text.Trim(KeysRight), ","),
        pListFields=Text.Split(Text.Trim(FieldsRight), ","),    
        pExpandField=Text.Trim(FieldExpand),
        pNewColumn="__NewColumn__",
        pExpandFieldList=Text.Split(pExpandField,""),
        pExpandFieldTable=List.Transform(pListFields, each pExpandField & "." & _),
        ModeJoin=Text.Trim(Text.Upper(ModeJoinKind)),

        Source = if ModeJoin = "LO"
                then
                    Table.NestedJoin(pTableLeft,pKeyListLeft,pTableRight,pKeyListRight,pNewColumn,JoinKind.LeftOuter)
                else 
                    if ModeJoin ="RO"
                    then 
                        Table.NestedJoin(pTableLeft,pKeyListLeft,pTableRight,pKeyListRight,pNewColumn,JoinKind.RightOuter) 
                    else 
                        if ModeJoin = "FO"
                        then 
                            Table.NestedJoin(pTableLeft,pKeyListLeft,pTableRight,pKeyListRight,pNewColumn,JoinKind.FullOuter) 
                            else 
                                if ModeJoin = "IN"
                                then 
                                    Table.NestedJoin(pTableLeft,pKeyListLeft,pTableRight,pKeyListRight,pNewColumn,JoinKind.Inner)
                                else 
                                     if ModeJoin ="LA"
                                     then 
                                        Table.NestedJoin(pTableLeft,pKeyListLeft,pTableRight,pKeyListRight,pNewColumn,JoinKind.LeftAnti)
                                    else
                                        if ModeJoin = "RA"
                                        then 
                                            Table.NestedJoin(pTableLeft,pKeyListLeft,pTableRight,pKeyListRight,pNewColumn,JoinKind.RightAnti)
                                        else
                                            null,        
        
        RenameTmpColumns = Table.RenameColumns(Source,{{pNewColumn, pExpandField}}),  

        Expand = Table.ExpandTableColumn(RenameTmpColumns, pExpandField, pListFields, pExpandFieldTable)

    in
        Expand
in
    fnLeftOuter