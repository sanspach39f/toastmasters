Attribute VB_Name = "Toastmasters_Create_Report"
' ------------------------------------------------------------------------------------------------
' ------------------ Samuel Anspach, District 4 Statistician ----------------------------
' ------------------------------------------------------------------------------------------------
' Work Log:
' Date ------------- Tasks Accomplished
' -----------------------------------------------
' 09/17/2017 ---- First created macro page, worked on creating the datatable, _
'                            relearning VBA, and creating the different sheets for different _
'                            pivot tables.
'
' 09/24/2017 ---- Creating code sections for other sheets/tables as well as setting _
                            up the Generate_Report Section. Finishing up macro by completing _
                            all reports and including the educational awards part
' ------------------------------------------------------------------------------------------------
' ---The Below Section of Code is used create the data table from the CSV
' ------------------------------------------------------------------------------------------------



' ------------------------------------------------------------------------------------------------
' ---------- The Below Section of Code is used to generate the metrics for
' ---------- everything except the educational awards. That will be called after
' ---------- getting the data online.
' ------------------------------------------------------------------------------------------------

Sub Generate_Reports()

    define_dataset
    lucky_7_sheet
    early_achievers
    smedley_stretch
    september_sanity
    educational_awards_sheet

End Sub


' ------------------------------------------------------------------------------------------------
' -----The Below Section of Code is used to turn our CSV data into a table
' ------------------------------------------------------------------------------------------------

Sub define_dataset()
    Dim sht As Worksheet
    Set sht = Sheets("Club_Performance")
    Sheets("Club_Performance").Activate
    Dim lastrow As Long
    Dim lastcolumn As Long
    Dim startcell As Range
    Set startcell = Range("A1")
    lastrow = Cells(sht.Rows.Count, startcell.Column).End(xlUp).Row
    
    'The last row contains non-data information and needs to be cleared
    Range(Cells(lastrow, 1), Cells(lastrow, 2)).ClearContents
    
    ' We need to now change our bottom row to the next one up
    lastrow = lastrow - 1
    lastcolumn = Cells(startcell.Row, sht.Columns.Count).End(xlToLeft).Column
    
    'Columns(lastcolumn).Select
    'MsgBox lastcolumn
    sht.Range(sht.Cells(1, 1), sht.Cells(lastrow, lastcolumn)).Select
    Worksheets("Club_Performance").ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "Data_Table"
    
End Sub

' ------------------------------------------------------------------------------------------------
' -------The Below Section of Code is used to construct the Lucky 7 page -------
' ------------------------------------------------------------------------------------------------

Sub lucky_7_sheet()

    Dim startcell As Range
    Dim lastcell As Range
    Dim Lucky_7 As Worksheet
    Set Lucky_7 = Worksheets.Add
    ActiveSheet.Name = "Lucky_7"
    
    
    'Below we create our pivot table in the Lucky_7 Sheet
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data_Table" _
    ).CreatePivotTable TableDestination:="Lucky_7!R5C1", _
    TableName:="Lucky_7_Table"

    ActiveSheet.PivotTables("Lucky_7_Table").AddDataField ActiveSheet.PivotTables( _
        "Lucky_7_Table").PivotFields("Off. Trained Round 1"), _
        "Sum of Off. Trained Round 1", xlSum

    With ActiveSheet.PivotTables("Lucky_7_Table").PivotFields("Club Name")
        .Orientation = xlRowField
        .Position = 1
    End With
    
     ActiveSheet.PivotTables("Lucky_7_Table").PivotFields("Club Name").AutoSort _
        xlDescending, "Sum of Off. Trained Round 1", ActiveSheet.PivotTables( _
        "Lucky_7_Table").PivotColumnAxis.PivotLines(1), 1
        
    ActiveSheet.PivotTables("Lucky_7_Table").PivotFields("Club Name").PivotFilters. _
        Add2 Type:=xlValueIsGreaterThanOrEqualTo, DataField:=ActiveSheet. _
        PivotTables("Lucky_7_Table").PivotFields("Sum of Off. Trained Round 1"), _
        Value1:=7
        


    Rows("1:4").Delete
    Set startcell = Range("C2")
    Set lastcell = Cells(ActiveSheet.Rows.Count, 2).End(xlUp).Offset(-1, 1)
    
    lastcell.Activate
     
    Range(startcell, lastcell).FormulaR1C1 = "=""<li>""&RC[-2]&""<li>"""
    

End Sub

' ------------------------------------------------------------------------------------------------
' ------------------------------- End of Lucky 7 page code -------------------------------
' ------------------------------------------------------------------------------------------------

' ------------------------------------------------------------------------------------------------
' ---The Below Section of Code is used to construct the Early_Achievers Sheet
' ------------------------------------------------------------------------------------------------

Sub early_achievers()
    
    Dim startcell As Range
    Dim lastcell As Range
    Dim early_achievers As Worksheet
    Set early_achievers = Worksheets.Add
    ActiveSheet.Name = "Early_Achievers"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data_Table" _
    ).CreatePivotTable TableDestination:="Early_Achievers!R5C1", _
    TableName:="Early_Achievers_Table"
    
    ActiveSheet.PivotTables("Early_Achievers_Table").AddDataField ActiveSheet.PivotTables( _
        "Early_Achievers_Table").PivotFields("Goals Met"), _
        "Total Goals Met", xlSum

    With ActiveSheet.PivotTables("Early_Achievers_Table").PivotFields("Club Name")
        .Orientation = xlRowField
        .Position = 1
    End With

     ActiveSheet.PivotTables("Early_Achievers_Table").PivotFields("Club Name").AutoSort _
        xlDescending, "Total Goals Met", ActiveSheet.PivotTables( _
        "Early_Achievers_Table").PivotColumnAxis.PivotLines(1), 1
        
        
    ActiveSheet.PivotTables("Early_Achievers_Table").PivotFields("Club Name").PivotFilters. _
        Add2 Type:=xlValueIsGreaterThanOrEqualTo, DataField:=ActiveSheet. _
        PivotTables("Early_Achievers_Table").PivotFields("Total Goals Met"), _
        Value1:=5
        
    Rows("1:4").Delete
    Set startcell = Range("C2")
 
    Set lastcell = Cells(ActiveSheet.Rows.Count, 2).End(xlUp).Offset(-1, 1)
    
    
    Range(startcell, lastcell).FormulaR1C1 = "=GETPIVOTDATA(""Goals Met"",R[-1]C[-2],""Club Name"", ""Rhino Business Club"")"
    
    
    startcell = startcell.Offset(, 1)
    lastcell = lastcell.Offset(, 1)
    
    Range(startcell, lastcell).FormulaR1C1 = "=""<tr><td>""&RC[-3]&""</td><td align'=""""center"""">""&RC[-1]&""</td><td>""&R[3]C[-3]&""</td><td align'=""""center"""">""&R[3]C[-1]&""</td></tr>"""
    
    
        
End Sub

' ------------------------------------------------------------------------------------------------
' ------------------------- End of Early Achievers Sheet code --------------------------
' ------------------------------------------------------------------------------------------------

' ------------------------------------------------------------------------------------------------
' -The Below Section of Code is used to construct the Smedley_Stretch Sheet
' ------------------------------------------------------------------------------------------------

Sub smedley_stretch()

    Dim startcell As Range
    Dim lastcell As Range
    Dim smedley_stretch As Worksheet
    Set smedley_stretch = Worksheets.Add
    ActiveSheet.Name = "Smedley_Stretch"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data_Table" _
    ).CreatePivotTable TableDestination:="Smedley_Stretch!R5C1", _
    TableName:="Smedley_Stretch_Table"
    
    With ActiveSheet.PivotTables("Smedley_Stretch_Table").PivotFields("Club Name")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("Smedley_Stretch_Table").AddDataField ActiveSheet. _
        PivotTables("Smedley_Stretch_Table").PivotFields("New Members"), _
        "New Members 1", xlSum
    ActiveSheet.PivotTables("Smedley_Stretch_Table").AddDataField ActiveSheet. _
        PivotTables("Smedley_Stretch_Table").PivotFields("Add. New Members"), _
        "New Members 2", xlSum
        
    ActiveSheet.PivotTables("Smedley_Stretch_Table").CalculatedFields.Add _
        "Total New Members", "='New Members' +'Add. New Members'", True
        
    ActiveSheet.PivotTables("Smedley_Stretch_Table").PivotFields( _
        "Total New Members").Orientation = xlDataField
    
    ActiveSheet.PivotTables("Smedley_Stretch_Table").PivotFields("Club Name"). _
        PivotFilters.Add2 Type:=xlValueIsGreaterThanOrEqualTo, DataField:= _
        ActiveSheet.PivotTables("Smedley_Stretch_Table").PivotFields( _
        "Sum of Total New Members"), Value1:=7
        
         
    Rows("1:4").Delete
    Set startcell = Range("E3")
 
    Set lastcell = Cells(ActiveSheet.Rows.Count, 4).End(xlUp).Offset(-1, 1)
    
    
    Range(startcell, lastcell).Formula = "=""<li>""&A3&"" (""&C3&"")</li>"""
    

End Sub

' ------------------------------------------------------------------------------------------------
' ------------------------ End of Smedley Stretch Sheet code --------------------------
' ------------------------------------------------------------------------------------------------

' ------------------------------------------------------------------------------------------------
' ---The Below Section of Code is used to construct the Sept._Sanity Sheet---
' ------------------------------------------------------------------------------------------------

Sub september_sanity()

    Dim startcell As Range
    Dim lastcell As Range
    Dim september_sanity As Worksheet
    Set september_sanity = Worksheets.Add
    ActiveSheet.Name = "September_Sanity"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data_Table" _
    ).CreatePivotTable TableDestination:="September_Sanity!R5C1", _
    TableName:="September_Sanity_Table"
    
    With ActiveSheet.PivotTables("September_Sanity_Table").PivotFields("Club Name")
        .Orientation = xlRowField
        .Position = 1
    End With

    ActiveSheet.PivotTables("September_Sanity_Table").AddDataField ActiveSheet. _
        PivotTables("September_Sanity_Table").PivotFields("Mem. Base"), _
        "Base Membership", xlSum
    ActiveSheet.PivotTables("September_Sanity_Table").AddDataField ActiveSheet. _
        PivotTables("September_Sanity_Table").PivotFields("Active Members"), _
        "Currently Active Members", xlSum

    ActiveSheet.PivotTables("September_Sanity_Table").CalculatedFields.Add _
        "Club Renewal Percentage", "='Active Members'/'Mem. Base'", True
    ActiveSheet.PivotTables("September_Sanity_Table").PivotFields( _
        "Club Renewal Percentage").Orientation = xlDataField
    ActiveSheet.PivotTables("September_Sanity_Table").PivotFields("Club Name"). _
        AutoSort xlDescending, "Sum of Club Renewal Percentage", ActiveSheet. _
        PivotTables("September_Sanity_Table").PivotColumnAxis.PivotLines(3), 1
    ActiveSheet.PivotTables("September_Sanity_Table").PivotFields("Club Name"). _
        PivotFilters.Add2 Type:=xlValueIsGreaterThanOrEqualTo, DataField:= _
        ActiveSheet.PivotTables("September_Sanity_Table").PivotFields( _
        "Sum of Club Renewal Percentage"), Value1:=0.75
        
    With ActiveSheet.PivotTables("September_Sanity_Table").PivotFields( _
        "Mem. dues on time Oct")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("September_Sanity_Table").PivotFields( _
        "Mem. dues on time Oct").ClearAllFilters
    ActiveSheet.PivotTables("September_Sanity_Table").PivotFields( _
        "Mem. dues on time Oct").CurrentPage = "1"
    Range("E7", Range("D7").End(xlDown).Offset(-1, 1)).Formula = "=TEXT(D7*100,0)&""%"""
    Range("F7", Range("D7").End(xlDown).Offset(-1, 2)).Formula = "=""<tr><td>""&A7&""</td><td>""&E7&""</td><td>""&""  ""&""</td><td>""&A57&""</td><td>""&E19&""</td></tr>"""
    
        
        
        
    
End Sub


' ------------------------------------------------------------------------------------------------
' ----------------------- End of September Sanity Sheet code -------------------------
' ------------------------------------------------------------------------------------------------

' ------------------------------------------------------------------------------------------------
' The Below Section of Code is used to construct the Educational Awards Sheet
' ------------------------------------------------------------------------------------------------

Sub educational_awards_sheet()

    Dim educational_awards_dataset As Worksheet
    Set educational_awards_dataset = Worksheets.Add
    ActiveSheet.Name = "Educational_Awards_Dataset"

    Range("A6").Value = "Go to the link below and copy the educational awards."
    Range("A7").Value = "http://reports.toastmasters.org/reports/dprReports.cfm?r=3&d=4&s=Date&sortOrder=0"

    MsgBox "Get the data for the educational awards online and then run the next macro for the counts"

End Sub

' ------------------------------------------------------------------------------------------------
' -------------- End of the Educational Awards Dataset Sheet code ----------------
' ------------------------------------------------------------------------------------------------

Sub Generate_Reports_Education()

    educational_awards_dataset
    awards_per_division
    awards_by_Type

End Sub



Sub educational_awards_dataset()

    Dim lastrow As Long
    Dim lastcolumn As Long
    Dim startcell As Range
    Dim i As Integer

    Set startcell = Range("A1")

    lastrow = Cells(ActiveSheet.Rows.Count, startcell.Column).End(xlUp).Row

    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Months from today"
    Range(Cells(2, 9), Cells(lastrow, 9)).Formula = "=DATEDIF(E2,TODAY(),""m"")"

    lastcolumn = Cells(startcell.Row, ActiveSheet.Columns.Count).End(xlToLeft).Column
    ActiveSheet.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(lastrow, lastcolumn)).Select
    Worksheets("Educational_Awards_Dataset").ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "Educational_Awards_Data"
    ActiveSheet.ListObjects("Educational_Awards_Data").Range.AutoFilter Field:=9, Criteria1:="0", Operator:=xlOr, Criteria2:="1"

        
    Range("J2", Range("I2").End(xlDown).Offset(, 1)).Formula = "=""<tr><td>""&F3&""</td><td>""&G3&""</tr></tr>"""
    
    
End Sub

Sub awards_per_division()

    Dim Division_Awards As Worksheet
    Set Division_Awards = Worksheets.Add
    ActiveSheet.Name = "Division_Awards"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Educational_Awards_Data" _
    ).CreatePivotTable TableDestination:="Division_Awards!R5C1", _
    TableName:="Division_Awards_Table"

    ActiveSheet.PivotTables("Division_Awards_Table").AddDataField ActiveSheet.PivotTables( _
        "Division_Awards_Table").PivotFields("Award"), _
        "Awards", xlCount

    With ActiveSheet.PivotTables("Division_Awards_Table").PivotFields("Division")
        .Orientation = xlRowField
        .Position = 1
    End With
    Rows("1:4").Delete
    Range("C1").Value = "Division,Awards"
    Range("C2", "C5").Formula = "=""Division ""&A2&"",""&B2"
    

End Sub

Sub awards_by_Type()

    Dim awards_by_Type As Worksheet
    Set awards_by_Type = Worksheets.Add
    ActiveSheet.Name = "Awards_by_Type"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Educational_Awards_Data" _
    ).CreatePivotTable TableDestination:="Awards_by_Type!R5C1", _
    TableName:="Awards_by_Type_Table"

    ActiveSheet.PivotTables("Awards_by_Type_Table").AddDataField ActiveSheet.PivotTables( _
        "Awards_by_Type_Table").PivotFields("Award"), _
        "Awards", xlCount

    With ActiveSheet.PivotTables("Awards_by_Type_Table").PivotFields("Award")
        .Orientation = xlRowField
        .Position = 1
    End With

    Rows("1:4").Delete
    Range("C1").Value = "Web"
    Range("D1").Value = "Award,Achieved"
    Range("C2").Formula = "=GETPIVOTDATA(""Award"",$A$1,""Award"",""CC"")"
    Range("C3").Formula = "=GETPIVOTDATA(""Award"",$A$1,""Award"",""ACB"")+GETPIVOTDATA(""Award"",$A$1,""Award"",""ACS"")+GETPIVOTDATA(""Award"",$A$1,""Award"",""ACS"")"
    Range("C4").Formula = "=GETPIVOTDATA(""Award"",$A$1,""Award"",""CL"")"
    Range("C5").Formula = "=GETPIVOTDATA(""Award"",$A$1,""Award"",""ALB"")+GETPIVOTDATA(""Award"",$A$1,""Award"",""ALS"")"
    Range("C6").Formula = "=GETPIVOTDATA(""Award"",$A$1,""Award"",""LDREXC"")"
    Range("C7").Formula = "=GETPIVOTDATA(""Award"",$A$1,""Award"",""DTM"")"
    Range("D2").Formula = "=""Competent Communicator,""&C2"
    Range("D3").Formula = "=""Advanced Communicator,""&C3"
    Range("D4").Formula = "=""Competent Leader,""&C4"
    Range("D5").Formula = "=""Advanced Leader,""&C5"
    Range("D6").Formula = "=""Leadership Excellence,""&C6"
    Range("D7").Formula = "=""Distinguished Toastmaster,""&C7"

End Sub

