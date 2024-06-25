// Prompt for calendar table selection

var dateTable = SelectTable(label: "Select your date table");
if(dateTable == null) return;

// Set the sort by column


(dateTable.Columns["Quarter"] as CalculatedTableColumn).SortByColumn = (Selected.Table.Columns["Year Quarter"] as CalculatedTableColumn);
(dateTable.Columns["Month Name"] as CalculatedTableColumn).SortByColumn = (Selected.Table.Columns["Month Number"] as CalculatedTableColumn);

// Create Hierarchy
var h = dateTable.AddHierarchy("Calendar Hierarchy");
h.AddLevel("Year");
h.AddLevel("Quarter");
h.AddLevel("Month");
h.AddLevel("Date");

// Assign Columns to Folders

    
        foreach (var c in dateTable.Columns )
        {
        c.DisplayFolder = "5. Boolean Fields";
        c.IsHidden = true;
        
        // Organize the date table into folders
            if ( ( c.DataType == DataType.DateTime & c.Name.Contains("Date") ) )
                {
                    c.DisplayFolder = "4. Date";
                c.IsHidden = false;
                c.IsKey = true;
                }
        
            if ( c.Name == "YYMMDDDD" )
                {
                    c.DisplayFolder = "4. Date";
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
        
    
        