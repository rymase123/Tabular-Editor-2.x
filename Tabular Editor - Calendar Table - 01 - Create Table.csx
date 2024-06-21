// To use this C# Script:
//
// 1. Run the script
// 2. Select the column that has the earliest date
// 3. Select the column that has the latest date

// List of all DateTime columns in the model
var _dateColumns = Model.AllColumns.Where(c => c.DataType == DataType.DateTime ).ToList();

// Select the column with the earliest date in the model
try
{
    string _EarliestDate = 
        SelectColumn(
            _dateColumns, 
            null, 
            "Select the Column with the Earliest Date:"
        ).DaxObjectFullName;
    
    try
    {
        // Select the column with the latest date in the model
        string _LatestDate = 
            SelectColumn(
                _dateColumns, 
                null, 
                "Select the Column with the Latest Date:"
            ).DaxObjectFullName;
        
// R.Mason - check if Measure exists already.
            
//Define the measure name to check

 string measureName = "RefDate";

// Use LINQ to check if the measure exists in the model
bool measureExists = Model.AllMeasures.Any(m => m.Name == measureName);

// Output the result of whether each measure already exists. If it does exist, then do nothing, if it doesnt exist
// then create the measure

    if (measureExists) 
    {
        
   ////if it exists. do nothing!
   
    } 
    else   
    {
        // Create measure for reference date
        var _RefDateMeasure = _dateColumns[0].Table.AddMeasure(
            "RefDate",
            "CALCULATE ( MAX ( " + _LatestDate + " ), REMOVEFILTERS ( ) )"
        );
    }
     
     
            

        
        
        // Formatted date table DAX
        // Based on date table from https://www.sqlbi.com/topics/date-table/
        // To adjust, copy everything after the @" into a DAX query window & replace
        
        var _DateDaxExpression = @"-- Reference date for the latest date in the report
        -- Until when the business wants to see data in reports
        VAR _Refdate_Measure = [RefDate]
        VAR _Today = TODAY ( )
        
        -- Replace with ""Today"" if [RefDate] evaluates blank
        VAR _Refdate = IF ( ISBLANK ( _Refdate_Measure ), _Today, _Refdate_Measure )
            VAR _RefYear        = YEAR ( _Refdate )
            VAR _RefQuarter     = _RefYear * 100 + QUARTER(_Refdate)
            VAR _RefMonth       = _RefYear * 100 + MONTH(_Refdate)
            VAR _RefWeek_EU     = _RefYear * 100 + WEEKNUM(_Refdate, 2)
        
        -- Earliest date in the model scope
        VAR _EarliestDate       = DATE ( YEAR ( MIN ( " + _EarliestDate + @" ) ) - 2, 1, 1 )
        VAR _EarliestDate_Safe  = MIN ( _EarliestDate, DATE ( YEAR ( _Today ) + 1, 1, 1 ) )
        
        -- Latest date in the model scope
        VAR _LatestDate_Safe    = DATE ( YEAR ( _Refdate ) + 2, 12, 1 )
        
        ------------------------------------------
        -- Base calendar table
        VAR _Base_Calendar      = CALENDAR ( _EarliestDate_Safe, _LatestDate_Safe )
        ------------------------------------------
        
        
        
        ------------------------------------------
        VAR _IntermediateResult = 
            ADDCOLUMNS ( _Base_Calendar,
        
                    ------------------------------------------
                    ""Year Number"",           --|
                    YEAR ([Date]),                          --|-- Year
                ""Year"" ,                                       --|
                FORMAT ( [Date], ""YYYY"" ),          --|                                   
                                                            
                                                            
                    ------------------------------------------
        
                    ------------------------------------------
                    ""Quarter"",       --|
                    ""Q"" &                                   --|-- Quarter
                    CONVERT(QUARTER([Date]), STRING) &      --|
                    "" "" &                                   --|
                    CONVERT(YEAR([Date]), STRING),          --|
                                                            --|
                ""Year Quarter"",        --|
                    YEAR([Date]) * 100 + QUARTER([Date]),   --|
                    ------------------------------------------
        
                    ------------------------------------------
                ""Month Year"",          --|
                    FORMAT ( [Date], ""MMM YY"" ),            --|-- Month
                                                            --|
                ""Year Month"",          --|
                    YEAR([Date]) * 100 + MONTH([Date]),     --|
                                                            --|
                ""Month Name"",                  --|
                    FORMAT ( [Date], ""MMM"" ),               --|
                                                            --|
                ""Month Number"",                  --|
                    MONTH ( [Date] ),                       --|
                    
                ""Month"",                  --|
                    FORMAT([Date] , ""YYYY-MM"") ,                 --|                    
                    
                    
                    ------------------------------------------
   
                ""YYYYMMDD"",                                 --| -- Day
                    YEAR ( [Date] ) * 10000                 --|
                    +                                       --|
                    MONTH ( [Date] ) * 100                  --|
                    +                                       --|
                    DAY ( [Date] ),                         --|
                    ------------------------------------------
        
        
                    ------------------------------------------
                ""IsDateInScope"",                            --|
                    [Date] <= _Refdate                      --|-- Boolean
                    &&                                      --|
                    YEAR([Date]) > YEAR(_EarliestDate),     --|
                                                            --|
                ""IsBeforeThisMonth"",                        --|
                    [Date] <= EOMONTH ( _Refdate, -1 ),     --|
                                                            --|
                ""IsLastMonth"",                              --|
                    [Date] <= EOMONTH ( _Refdate, 0 )       --|
                    &&                                      --|
                    [Date] > EOMONTH ( _Refdate, -1 ),      --|
                                                            --|
                ""IsYTD"",                                    --|
                    MONTH([Date])                           --|
                    <=                                      --|
                    MONTH(EOMONTH ( _Refdate, 0 )),         --|
                                                            --|
                ""IsActualToday"",                            --|
                    [Date] = _Today,                        --|
                                                            --|
                ""IsRefDate"",                                --|
                    [Date] = _Refdate,                      --|
                                                            --|
                ""IsHoliday"",                                --|
                    MONTH([Date]) * 100                     --|
                    +                                       --|
                    DAY([Date])                             --|
                        IN {0101, 0501, 1111, 1225},        --|
                                                            --|
                ""IsWeekday"",                                --|
                    WEEKDAY([Date], 2)                      --|
                        IN {1, 2, 3, 4, 5})                 --|
                    ------------------------------------------
        
        VAR _Result = 
            
                    --------------------------------------------
            ADDCOLUMNS (                                      --|
                _IntermediateResult,                          --|-- Boolean #2
                ""IsThisYear"",                                 --|
                [Year Number]          --|
                        = _RefYear,                           --|
                                                            --|
                ""IsThisMonth"",                                --|
                    [Year Month]         --|
                        = _RefMonth,                          --|
                                                            --|
                ""IsThisQuarter"",                              --|
                    [Year Quarter]       --|
                        = _RefQuarter                        --|
             
            )                                                 --|
                    --------------------------------------------
                    
        RETURN 
            _Result";
        
        // Create date table
        var _date = Model.AddCalculatedTable(
        "Calendar",
            _DateDaxExpression
        );
        
      
        
        // Mark as date table
        _date.DataCategory = "Time";
       
     
        //-------------------------------------------------------------------------------------------//
        
        
        Info ( "Created a new, organized 'Date' table based on the template in the C# Script.\nThe Earliest Date is taken from " + _EarliestDate + "\nThe Latest Date is taken from " + _LatestDate );
    
        }
       catch (Exception ex )
{
    Output( ex );        
}
}
catch
{
    Error( "Earliest column not selected! Ending script without making changes." );
}

