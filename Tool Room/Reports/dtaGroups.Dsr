VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} dtaGroups 
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3315
   _ExtentX        =   5847
   _ExtentY        =   14182
   FolderFlags     =   5
   TypeLibGuid     =   "{4B152026-5EE0-4A24-9DB2-B54B7134CBDF}"
   TypeInfoGuid    =   "{9535C720-765C-4A9D-9408-5DFAD27C37ED}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "conDb"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   8
   BeginProperty Recordset1 
      CommandName     =   "cmdSummary"
      CommDispId      =   1002
      RsDispId        =   1009
      CommandText     =   "Select Distinctrow format(short_month,'mmmm') as MonthHeader, Year from tblMonthTable"
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "MonthHeader"
         Caption         =   "MonthHeader"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   65656
         Scale           =   0
         Type            =   204
         Name            =   "Year"
         Caption         =   "Year"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Year"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   1
      BeginProperty Aggregate1 
         Name            =   "Item_Total"
         AggOn           =   "cmdDetails"
         AggField        =   "item_id"
         AggType         =   1
         AggFunction     =   3
         Precision       =   8
         Size            =   4
         Scale           =   8
         Type            =   131
         Name            =   "Item_Total"
         Caption         =   "Item_Total"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "cmdDetails"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"dtaGroups.dsx":0000
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdSummary"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "item_id"
         Caption         =   "item_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "qty"
         Caption         =   "qty"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "date_added"
         Caption         =   "date_added"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "DateRelated"
         Caption         =   "DateRelated"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Available"
         Caption         =   "Available"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Unavailable"
         Caption         =   "Unavailable"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "year"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   2
      BeginProperty Relation1 
         ParentField     =   "Year"
         ChildField      =   "year"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "MonthHeader"
         ChildField      =   "DateRelated"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "cmdCategory"
      CommDispId      =   1011
      RsDispId        =   1017
      CommandText     =   "Select description, dYear from tblCategories"
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   16711935
         Scale           =   0
         Type            =   204
         Name            =   "dYear"
         Caption         =   "dYear"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   1
      BeginProperty Aggregate1 
         Name            =   "Item_Count"
         AggOn           =   "cmdCDetails"
         AggField        =   "item_id"
         AggType         =   1
         AggFunction     =   3
         Precision       =   8
         Size            =   4
         Scale           =   8
         Type            =   131
         Name            =   "Item_Count"
         Caption         =   "Item_Count"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "cmdCDetails"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * from tblItemList where format(reg_date, 'yyyy') = dYear"
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdCategory"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "item_id"
         Caption         =   "item_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "category"
         Caption         =   "category"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "location"
         Caption         =   "location"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "remarks"
         Caption         =   "remarks"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "image_name"
         Caption         =   "image_name"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "reg_date"
         Caption         =   "reg_date"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   2
      BeginProperty Relation1 
         ParentField     =   "description"
         ChildField      =   "category"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "dYear"
         ChildField      =   "dYear"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "cmdLocation"
      CommDispId      =   1018
      RsDispId        =   1031
      CommandText     =   "Select description, remarks, dYear from tblLocations"
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      GrandTotal      =   "GrandTotal1"
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "remarks"
         Caption         =   "remarks"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   510
         Scale           =   0
         Type            =   204
         Name            =   "dYear"
         Caption         =   "dYear"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   1
      BeginProperty Aggregate1 
         Name            =   "Item_Total"
         AggOn           =   "cmdLDetails"
         AggField        =   "item_id"
         AggType         =   1
         AggFunction     =   3
         Precision       =   8
         Size            =   4
         Scale           =   8
         Type            =   131
         Name            =   "Item_Total"
         Caption         =   "Item_Total"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "cmdLDetails"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * from tblItemList where format(reg_date, 'yyyy') = dYear"
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdLocation"
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "item_id"
         Caption         =   "item_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "category"
         Caption         =   "category"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "location"
         Caption         =   "location"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "remarks"
         Caption         =   "remarks"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "image_name"
         Caption         =   "image_name"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "reg_date"
         Caption         =   "reg_date"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   2
      BeginProperty Relation1 
         ParentField     =   "dYear"
         ChildField      =   "dYear"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "description"
         ChildField      =   "location"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "cmdStatus"
      CommDispId      =   1032
      RsDispId        =   1036
      CommandText     =   "Select description, remarks, include, system, dYear from tblStatus"
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "remarks"
         Caption         =   "remarks"
      EndProperty
      BeginProperty Field3 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "include"
         Caption         =   "include"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "system"
         Caption         =   "system"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   510
         Scale           =   0
         Type            =   204
         Name            =   "dYear"
         Caption         =   "dYear"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   1
      BeginProperty Aggregate1 
         Name            =   "Item_Total"
         AggOn           =   "cmdSDetails"
         AggField        =   "stat.item_id"
         AggType         =   1
         AggFunction     =   3
         Precision       =   8
         Size            =   4
         Scale           =   8
         Type            =   131
         Name            =   "Item_Total"
         Caption         =   "Item_Total"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "cmdSDetails"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"dtaGroups.dsx":0216
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdStatus"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "record_no"
         Caption         =   "record_no"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "stat.item_id"
         Caption         =   "stat.item_id"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "stat.qty"
         Caption         =   "stat.qty"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "reg.item_id"
         Caption         =   "reg.item_id"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "reg.qty"
         Caption         =   "reg.qty"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "date_added"
         Caption         =   "date_added"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   2
      BeginProperty Relation1 
         ParentField     =   "description"
         ChildField      =   "status"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "dYear"
         ChildField      =   "dYear"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "dtaGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
