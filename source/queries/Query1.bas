Operation =1
Option =0
Begin InputTables
    Name ="Table1"
    Name ="Table2"
End
Begin OutputColumns
    Expression ="Table1.*"
    Expression ="Table2.*"
End
Begin Joins
    LeftTable ="Table1"
    RightTable ="Table2"
    Expression ="Table1.ID = Table2.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbByte "PublishToWeb" ="1"
Begin
    Begin
        dbText "Name" ="Table1.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Table1.Field1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Table1.Field2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Table2.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Table2.Name1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Table2.Name2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1151
    Bottom =539
    Left =-1
    Top =-1
    Right =816
    Bottom =294
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Table1"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Table2"
        Name =""
    End
End
