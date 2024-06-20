// For all the columns in the date table:
foreach( var t in Selected.Tables )
    
        foreach (var c in t.Columns )
        {
        c.DisplayFolder = "7. Boolean Fields";
        c.IsHidden = true;
        
        // Organize the date table into folders
            if ( ( c.DataType == DataType.DateTime & c.Name.Contains("Date") ) )
                {
                    c.DisplayFolder = "4. Calendar Date";
                c.IsHidden = false;
                c.IsKey = true;
                }
        
            if ( c.Name == "YYMMDDDD" )
                {
                    c.DisplayFolder = "4. Calendar Date";
                c.IsHidden = true;
                }
        
            if ( c.Name.Contains("Year") & c.DataType != DataType.Boolean )
                {
                c.DisplayFolder = "1. Year";
                c.IsHidden = false;
                }
        

        
            if ( c.Name.Contains("Month") & c.DataType != DataType.Boolean )
                {
                c.DisplayFolder = "3. Month";
                c.IsHidden = false;
                }
        
            if ( c.Name.Contains("Quarter") & c.DataType != DataType.Boolean )
                {
                c.DisplayFolder = "2. Quarter";
                c.IsHidden = false;
                }
        
        }
        
    
        