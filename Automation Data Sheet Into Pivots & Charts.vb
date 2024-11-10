' VBA code creates three pivot tables & three line charts based on the three button selection on the work sheet. These pivots & charts 
' are based on the data sheet in the workbook. Note the pivot tables & line charts are dynamic based on the three button selection and appear on same sheet.



'Button one ************************

Sub Comp_Month_Year()
'
' Comp_Month_Year Macro
'

'
    Application.ScreenUpdating = False ' so screen will not flicker
    'when macro running
    
    ActiveSheet.PivotTables("PivotTable3").ClearTable
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Company")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Month")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Year")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Revenue"), "Sum of Revenue", xlSum
        
    Application.ScreenUpdating = True
    
End Sub


'Button two ************************
Sub Quarter_Company()
'
' Quarter_Company Macro
'

'
    Application.ScreenUpdating = False ' so screen will not flicker
    'when macro running
    
    ActiveSheet.PivotTables("PivotTable3").ClearTable
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Quarter")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Company")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Revenue"), "Sum of Revenue", xlSum
    
    Application.ScreenUpdating = True
End Sub


'Button three ************************
Sub Year_Month()
'
' Year_Month Macro
'

'
Application.ScreenUpdating = False ' so screen will not flicker
    'when macro running

    ActiveSheet.PivotTables("PivotTable3").ClearTable
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Year")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Month")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Revenue"), "Sum of Revenue", xlSum
        
    Application.ScreenUpdating = True
End Sub