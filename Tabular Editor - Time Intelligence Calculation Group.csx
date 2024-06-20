#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
//
// CHANGELOG:
// '2021-05-01 / B.Agullo / 
// '2021-05-17 / B.Agullo / added affected measure table
// '2021-06-19 / B.Agullo / data label measures
// '2021-07-10 / B.Agullo / added flag expression to avoid breaking already special format strings
// '2021-09-23 / B.Agullo / added code to prompt for parameters (code credit to Daniel Otykier) 
// '2021-09-27 / B.Agullo / added code for general name 
// '2022-10-11 / B.Agullo / added MMT and MWT calc item groups
// '2024-06-12 / R.Mason / removed calc item groups that I do not use and added ones that I do use
// '2024-06-13 / R.Mason / added check for Label measures. Also added code to impact all measures, but you can edit the table after the fact



// by Bernat Agulló
// twitter: @AgulloBernat
// www.esbrina-ba.com/blog
//
// REFERENCE: 
// Check out https://www.esbrina-ba.com/time-intelligence-the-smart-way/ where this script is introduced
// 
// FEATURED: 
// this script featured in GuyInACube https://youtu.be/_j0iTUo2HT0
//
// THANKS:
// shout out to Johnny Winter for the base script and SQLBI for daxpatterns.com

//select the measures that you want to be affected by the calculation group
//before running the script. 
//measure names can also be included in the following array (no need to select them) 
string[] preSelectedMeasures = {}; //include measure names in double quotes, like: {"Profit","Total Cost"};

//AT LEAST ONE MEASURE HAS TO BE AFFECTED!, 
//either by selecting it or typing its name in the preSelectedMeasures Variable

// **** NOTE THAT THE MEASURES FOR LABEL FORMATTING ARE NOT WORKING CORRECTLY FOR ANYTHING OUTSIDE OF YEAR CALCS. WILL HAVE TO RETURN 
// TO THESE AT SOME POINT, BUT LOW PRIORITY FOR NOW - R.MASON ****

//
// ----- do not modify script below this line -----
//


string affectedMeasures = "{";

int i = 0; 

for (i=0;i<preSelectedMeasures.GetLength(0);i++){
  
    if(affectedMeasures == "{") {
    affectedMeasures = affectedMeasures + "\"" + preSelectedMeasures[i] + "\"";
    }else{
        affectedMeasures = affectedMeasures + ",\"" + preSelectedMeasures[i] + "\"" ;
    }; 
    
};

//*** this part of the code lets you select the measures to load to the Time Intellegence Affected Measures table  (the impacted measures)

//if (Selected.Measures.Count != 0) {
// foreach(var m in Selected.Measures) {

// dont forget to close the bracket below the foreach code

//***    

foreach(var m in Model.AllMeasures) { // if you just want load all the measures, use this code. Otherwise replace with the Selected.Measures code above.
        if(affectedMeasures == "{") {
        affectedMeasures = affectedMeasures + "\"" + m.Name + "\"";
        }else{
            affectedMeasures = affectedMeasures + ",\"" + m.Name + "\"" ;
        };
    };  
// if you use the Selected.Measures code, then make sure you close the curly bracket:
//};


//check that by either method at least one measure is affected
if(affectedMeasures == "{") { 
    Error("No measures affected by calc group"); 
    return; 
};

string calcGroupName; 
string columnName; 

if(Model.CalculationGroups.Any(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group")) {
    calcGroupName = Model.CalculationGroups.Where(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group").First().Name;
    
}else {
    calcGroupName = Interaction.InputBox("Provide a name for your Calc Group", "Calc Group Name", "Time Intelligence", 740, 400);
}; 

if(calcGroupName == "") return;


if(Model.CalculationGroups.Any(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group")) {
    columnName = Model.Tables.Where(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group").First().Columns.First().Name;
    
}else {
    columnName = Interaction.InputBox("Provide a name for your Calc Group Column", "Calc Group Column Name", calcGroupName, 740, 400);
}; 

if(columnName == "") return;

string affectedMeasuresTableName; 

if(Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table")) {
    affectedMeasuresTableName = Model.Tables.Where(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table").First().Name;

} else {
    affectedMeasuresTableName = Interaction.InputBox("Provide a name for affected measures table", "Affected Measures Table Name", calcGroupName  + " Affected Measures", 740, 400);

};

if(affectedMeasuresTableName == "") return;

string affectedMeasuresColumnName = Interaction.InputBox("Provide a name for affected measures table column name", "Affected Measures Table Column Name", "Measure", 740, 400);

if(Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table")) {
    affectedMeasuresColumnName = Model.Tables.Where(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table").First().Columns.First().Name;

} else {
    affectedMeasuresColumnName = Interaction.InputBox("Provide a name for affected measures column", "Affected Measures Table Column Name", "Measure", 740, 400);

};


if(affectedMeasuresColumnName == "") return;
//string affectedMeasuresColumnName = "Affected Measure"; 

string labelAsValueMeasureName = "Label as Value Measure"; 
string labelAsFormatStringMeasureName = "Label as format string"; 


 // '2021-09-24 / B.Agullo / model object selection prompts! 
var factTable = SelectTable(label: "Select your fact table");
if(factTable == null) return;

var factTableDateColumn = SelectColumn(factTable.Columns, label: "Select the main date column");
if(factTableDateColumn == null) return;

Table dateTableCandidate = null;

if(Model.Tables.Any
    (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table" 
        || x.Name == "Date" 
        || x.Name == "Calendar")){
            dateTableCandidate = Model.Tables.Where
                (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table" 
                    || x.Name == "Date" 
                    || x.Name == "Calendar").First();

};

var dateTable = 
    SelectTable(
        label: "Select your date table",
        preselect:dateTableCandidate);

if(dateTable == null) {
    Error("You just aborted the script"); 
    return;
} else {
    dateTable.SetAnnotation("@AgulloBernat","Time Intel Date Table");
}; 


Column dateTableDateColumnCandidate = null; 

if(dateTable.Columns.Any
            (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Date Column" || x.Name == "Date")){
    dateTableDateColumnCandidate = dateTable.Columns.Where
        (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Date Column" || x.Name == "Date").First();
};

var dateTableDateColumn = 
    SelectColumn(
        dateTable.Columns, 
        label: "Select the date column",
        preselect: dateTableDateColumnCandidate);

if(dateTableDateColumn == null) {
    Error("You just aborted the script"); 
    return;
} else { 
    dateTableDateColumn.SetAnnotation("@AgulloBernat","Time Intel Date Table Date Column"); 
}; 

Column dateTableYearColumnCandidate = null;
if(dateTable.Columns.Any(x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Year Column" || x.Name == "Year")){
    dateTable.Columns.Where
        (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Year Column" || x.Name == "Year").First();
};

var dateTableYearColumn = 
    SelectColumn(
        dateTable.Columns, 
        label: "Select the year column", 
        preselect:dateTableYearColumnCandidate);

if(dateTableYearColumn == null) {
    Error("You just abourted the script"); 
    return;
} else {
    dateTableYearColumn.SetAnnotation("@AgulloBernat","Time Intel Date Table Year Column"); 
};


//these names are for internal use only, so no need to be super-fancy, better stick to datpatterns.com model
string ShowValueForDatesMeasureName = "ShowValueForDates";
string dateWithSalesColumnName = "DateWith" + factTable.Name;

// '2021-09-24 / B.Agullo / I put the names back to variables so I don't have to tough the script
string factTableName = factTable.Name;
string factTableDateColumnName = factTableDateColumn.Name;
string dateTableName = dateTable.Name;
string dateTableDateColumnName = dateTableDateColumn.Name;
string dateTableYearColumnName = dateTableYearColumn.Name; 

// '2021-09-24 / B.Agullo / this is for internal use only so better leave it as is 
string flagExpression = "UNICHAR( 8204 )"; 

string calcItemProtection = "<CODE>"; //default value if user has selected no measures
string calcItemFormatProtection = "<CODE>"; //default value if user has selected no measures

// check if there's already an affected measure table
if(Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table")) {
    //modifying an existing calculated table is not risk-free
    Info("Make sure to include measure names to the table " + affectedMeasuresTableName);
} else { 
    // create calculated table containing all names of affected measures
    // this is why you need to enable 
    if(affectedMeasures != "{") { 
        
        affectedMeasures = affectedMeasures + "}";
        
        string affectedMeasureTableExpression = 
            "SELECTCOLUMNS(" + affectedMeasures + ",\"" + affectedMeasuresColumnName + "\",[Value])";

        var affectedMeasureTable = 
            Model.AddCalculatedTable(affectedMeasuresTableName,affectedMeasureTableExpression);
        
        affectedMeasureTable.FormatDax(); 
        affectedMeasureTable.Description = 
            "Measures affected by " + calcGroupName + " calculation group." ;
        
        affectedMeasureTable.SetAnnotation("@AgulloBernat","Time Intel Affected Measures Table"); 
       
        // this causes error
        // affectedMeasureTable.Columns[affectedMeasuresColumnName].SetAnnotation("@AgulloBernat","Time Intel Affected Measures Table Column");

        affectedMeasureTable.IsHidden = true;     
        
    };
};

//if there where selected or preselected measures, prepare protection code for expresion and formatstring
string affectedMeasuresValues = "VALUES('" + affectedMeasuresTableName + "'[" + affectedMeasuresColumnName + "])";

calcItemProtection = 
    "SWITCH(" + 
    "   TRUE()," + 
    "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," + 
    "   <CODE> ," + 
    "   ISSELECTEDMEASURE([" + labelAsValueMeasureName + "])," + 
    "   <LABELCODE> ," + 
    "   SELECTEDMEASURE() " + 
    ")";
    
    
calcItemFormatProtection = 
    "SWITCH(" + 
    "   TRUE() ," + 
    "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," + 
    "   <CODE> ," + 
    "   ISSELECTEDMEASURE([" + labelAsFormatStringMeasureName + "])," + 
    "   <LABELCODEFORMATSTRING> ," +
    "   SELECTEDMEASUREFORMATSTRING() " + 
    ")";

   
string dateColumnWithTable = "'" + dateTableName + "'[" + dateTableDateColumnName + "]"; 
string yearColumnWithTable = "'" + dateTableName + "'[" + dateTableYearColumnName + "]"; 
string factDateColumnWithTable = "'" + factTableName + "'[" + factTableDateColumnName + "]";
string dateWithSalesWithTable = "'" + dateTableName + "'[" + dateWithSalesColumnName + "]";
string calcGroupColumnWithTable = "'" + calcGroupName + "'[" + columnName + "]";

//check to see if a table with this name already exists
//if it doesnt exist, create a calculation group with this name
if (!Model.Tables.Contains(calcGroupName)) {
  var cg = Model.AddCalculationGroup(calcGroupName);
  cg.Description = "Calculation group for time intelligence. Availability of data is taken from " + factTableName + ".";
  cg.SetAnnotation("@AgulloBernat","Time Intel Calc Group"); 
};

//set variable for the calc group
Table calcGroup = Model.Tables[calcGroupName];

//if table already exists, make sure it is a Calculation Group type
if (calcGroup.SourceType.ToString() != "CalculationGroup") {
  Error("Table exists in Model but is not a Calculation Group. Rename the existing table or choose an alternative name for your Calculation Group.");
  return;
};

//adds the two measures that will be used for label as value, label as format string 
//built in check to see if the Label measures exists or not. If they exist, then do nothing, otherwise create them.

// Define the measure name to check
string measureName = "Label as format string";
string measureName2 = "Label as Value Measure";
// Use LINQ to check if the measure exists in the model
bool measureExists = Model.AllMeasures.Any(m => m.Name == measureName);
bool measureExists2 = Model.AllMeasures.Any(m => m.Name == measureName2);


// Output the result of whether each measure already exists. If it does exist, then do nothing, if it doesnt exist
// then create the measure

    if (measureExists) 
    {
        
   ////if it exists. do nothing!
   
    } 
    else   
    {
     var labelAsValueMeasure = calcGroup.AddMeasure(labelAsValueMeasureName,"");
    }
 
    if (measureExists2) 
    {
     ////if it exists. do nothing! 
    } 
    else   
    {
     var labelAsFormatStringMeasure = calcGroup.AddMeasure(labelAsFormatStringMeasureName,"0");
    }

//var labelAsFormatStringMeasure = calcGroup.AddMeasure(labelAsFormatStringMeasureName,"0");
//labelAsFormatStringMeasure.Description = "Use this measure to show the year evaluated in charts"; 

//by default the calc group has a column called Name. If this column is still called Name change this in line with specfied variable
if (calcGroup.Columns.Contains("Name")) {
  calcGroup.Columns["Name"].Name = columnName;

};

calcGroup.Columns[columnName].Description = "Select value(s) from this column to apply time intelligence calculations.";
calcGroup.Columns[columnName].SetAnnotation("@AgulloBernat","Time Intel Calc Group Column"); 


//Only create them if not in place yet (reruns)
if(!Model.Tables[dateTableName].Columns.Any(C => C.GetAnnotation("@AgulloBernat") == "Date with Data Column")){
    string DateWithSalesCalculatedColumnExpression = 
        dateColumnWithTable + " <= MAX ( " + factDateColumnWithTable + ")";

    Column dateWithDataColumn = dateTable.AddCalculatedColumn(dateWithSalesColumnName,DateWithSalesCalculatedColumnExpression);
    dateWithDataColumn.SetAnnotation("@AgulloBernat","Date with Data Column");
};

if(!Model.Tables[dateTableName].Measures.Any(M => M.Name == ShowValueForDatesMeasureName)) {
    string ShowValueForDatesMeasureExpression = 
        "VAR LastDateWithData = " + 
        "    CALCULATE ( " + 
        "        MAX (  " + factDateColumnWithTable + " ), " + 
        "        REMOVEFILTERS () " +
        "    )" +
        "VAR FirstDateVisible = " +
        "    MIN ( " + dateColumnWithTable + " ) " + 
        "VAR Result = " +  
        "    FirstDateVisible <= LastDateWithData " +
        "RETURN " + 
        "    Result ";

    var ShowValueForDatesMeasure = dateTable.AddMeasure(ShowValueForDatesMeasureName,ShowValueForDatesMeasureExpression); 

    ShowValueForDatesMeasure.FormatDax();
};



//defining expressions and formatstring for each calc item
string Base = 
    "/*Base*/ " + 
    "SELECTEDMEASURE()";

string Baselabel = 
    "SELECTEDVALUE(" + yearColumnWithTable + ")";

string MTD = 
    "/*MTD*/" + 
    "IF (" +
    "    [" + ShowValueForDatesMeasureName + "]," + 
    "    CALCULATE (" +
    "        " + Base+ "," + 
    "        DATESMTD (" +  dateColumnWithTable + " )" + 
    "   )" + 
    ") ";
    

string MTDlabel = Baselabel + "& \" MTD\""; 
    
string QTD = 
    "/*QTD*/" + 
    "IF (" +
    "    [" + ShowValueForDatesMeasureName + "]," + 
    "    CALCULATE (" +
    "        " + Base+ "," + 
    "        DATESQTD (" +  dateColumnWithTable + " )" + 
    "   )" + 
    ") ";
    

string QTDlabel = Baselabel + "& \" QTD\"";     
    
string YTD = 
    "/*YTD*/" + 
    "IF (" +
    "    [" + ShowValueForDatesMeasureName + "]," + 
    "    CALCULATE (" +
    "        " + Base+ "," + 
    "        DATESYTD (" +  dateColumnWithTable + " )" + 
    "   )" + 
    ") ";
    

string YTDlabel = Baselabel + "& \" YTD\"";   


string T3M = 
    "/*T3M*/" + 
    "IF (" +
    "    [" + ShowValueForDatesMeasureName + "]," + 
    "    CALCULATE (" +
    "        " + Base + "," + 
    "        DATESINPERIOD (" +  dateColumnWithTable + ", LASTDATE( " + dateColumnWithTable + " ) , -3 , MONTH )" + 
    "   )" + 
    ") ";
    

string T3Mlabel = Baselabel + "& \" T3M\"";   


string TTM = 
    "/*TTM*/" + 
    "IF (" +
    "    [" + ShowValueForDatesMeasureName + "]," + 
    "    CALCULATE (" +
    "        " + Base + "," + 
    "        DATESINPERIOD (" +  dateColumnWithTable + ", LASTDATE( " + dateColumnWithTable + " ) , -12 , MONTH )" + 
    "   )" + 
    ") ";
    

    string TTMlabel = Baselabel + "& \" TTM\"";   

    
string ITD = 
    "/*ITD*/" + 
    "VAR MAX_DATE = MAX( " + dateColumnWithTable + ")" + 
    "RETURN " + "/*blankspace*/" +
    "   IF (" +
    "    [" + ShowValueForDatesMeasureName + "]," + 
    "    CALCULATE (" +
    "        " + Base+ "," + 
    "        " + dateColumnWithTable + "<= MAX_DATE , " + 
    "        ALL ('" +  dateTableName + "' )" + 
    "   )" + 
    ") ";
    

string ITDlabel = Baselabel + "& \" ITD\"";   

string PMTD = 
    "/*PMTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ MTD + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, MONTH ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";
    

string PMTDlabel = 
    "/*PMTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ MTDlabel + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, MONTH ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";     
    
string PQTD = 
    "/*PQTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ QTD + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, QUARTER ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";
    

string PQTDlabel = 
    "/*PQTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ QTDlabel + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, QUARTER ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";     
    
string PYTD = 
    "/*PYTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ YTD + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";
    

string PYTDlabel = 
    "/*PYTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ YTDlabel + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") "; 
    
       
    
string PYMTD = 
    "/*PYMTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ MTD + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";
  
  

string PYMTDlabel = 
    "/*PYMTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ MTDlabel + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") "; 
    
    
string PYQTD = 
    "/*PYQTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ QTD + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";
    

string PYQTDlabel = 
    "/*PYQTD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ QTDlabel + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") "; 
       
    
string PYITD = 
       "/*PYITD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ ITD + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") ";
    

string PYITDlabel = 
    "/*PYITD*/ " +
    "IF (" + 
    "    [" + ShowValueForDatesMeasureName + "], " + 
    "    CALCULATE ( " + 
    "        "+ ITDlabel + ", " + 
    "        CALCULATETABLE ( " + 
    "            DATEADD ( " + dateColumnWithTable + " , -1, YEAR ), " + 
    "            " + dateWithSalesWithTable + " = TRUE " +  
    "        ) " + 
    "    ) " + 
    ") "; 
       

string ΔMoM = 
    "/*ΔMoM*/ " + 
    "VAR ValueCurrentPeriod = " + MTD + " " + 
    "VAR ValuePreviousPeriod = " + PMTD + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string ΔMoMlabel = 
    "/*ΔMoM*/ " + 
    "VAR ValueCurrentPeriod = " + MTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PMTDlabel + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string ΔQoQ = 
    "/*ΔQoQ*/ " + 
    "VAR ValueCurrentPeriod = " + QTD + " " + 
    "VAR ValuePreviousPeriod = " + PQTD + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string ΔQoQlabel = 
    "/*ΔQoQ*/ " + 
    "VAR ValueCurrentPeriod = " + QTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PQTDlabel + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";
    
string ΔYoY = 
    "/*ΔYoY*/ " + 
    "VAR ValueCurrentPeriod = " + YTD + " " + 
    "VAR ValuePreviousPeriod = " + PYTD + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string ΔYoYlabel = 
    "/*ΔYoY*/ " + 
    "VAR ValueCurrentPeriod = " + YTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYTDlabel + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";
  
    
string MOMTDpct = 
    "/*ΔMOM%*/" + 
    "VAR ValueCurrentPeriod = " + MTD + " " + 
    "VAR ValuePreviousPeriod = " + PMTD + " " + 
    "VAR CurrentMinusPreviousPeriod = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "VAR Result = " + 
    "DIVIDE ( "  + 
    "    CurrentMinusPreviousPeriod," + 
    "    ValuePreviousPeriod" + 
    ") " + 
    "RETURN " + 
    "  Result";


string MOMTDpctLabel = 
    "/*ΔMOM%*/ " +
    "VAR ValueCurrentPeriod = " + MTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PMTDlabel + " " + 
    "VAR Result = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" + 
    " ) " +  
    "RETURN " + 
    "  Result";

    
string QOQTDpct = 
    "/*ΔQOQ%*/" + 
    "VAR ValueCurrentPeriod = " + QTD + " " + 
    "VAR ValuePreviousPeriod = " + PQTD + " " + 
    "VAR CurrentMinusPreviousPeriod = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "VAR Result = " + 
    "DIVIDE ( "  + 
    "    CurrentMinusPreviousPeriod," + 
    "    ValuePreviousPeriod" + 
    ") " + 
    "RETURN " + 
    "  Result";


string QOQTDpctLabel = 
    "/*ΔQOQ%*/ " +
    "VAR ValueCurrentPeriod = " + QTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PQTDlabel + " " + 
    "VAR Result = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" + 
    " ) " +  
    "RETURN " + 
    "  Result";

    
    
string YOYTDpct = 
    "/*ΔYOY%*/" + 
    "VAR ValueCurrentPeriod = " + YTD + " " + 
    "VAR ValuePreviousPeriod = " + PYTD + " " + 
    "VAR CurrentMinusPreviousPeriod = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "VAR Result = " + 
    "DIVIDE ( "  + 
    "    CurrentMinusPreviousPeriod," + 
    "    ValuePreviousPeriod" + 
    ") " + 
    "RETURN " + 
    "  Result";


string YOYTDpctLabel = 
    "/*ΔYOY%*/ " +
   "VAR ValueCurrentPeriod = " + YTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYTDlabel + " " + 
    "VAR Result = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" + 
    " ) " +  
    "RETURN " + 
    "  Result";

string ΔMoM_PY = 
    "/*ΔMoM_PY*/ " + 
    "VAR ValueCurrentPeriod = " + MTD + " " + 
    "VAR ValuePreviousPeriod = " + PYMTD + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string ΔMoM_PYlabel = 
    "/*ΔMoM_PY*/ " + 
    "VAR ValueCurrentPeriod = " + MTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYMTDlabel + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string ΔQoQ_PY = 
    "/*ΔQoQ_PY*/ " + 
    "VAR ValueCurrentPeriod = " + QTD + " " + 
    "VAR ValuePreviousPeriod = " + PYQTD + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string ΔQoQ_PYlabel = 
    "/*ΔQoQ_PY*/ " + 
    "VAR ValueCurrentPeriod = " + QTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYQTDlabel + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";
    
string ΔYoY_PY = 
    "/*ΔYoY_PY*/ " + 
    "VAR ValueCurrentPeriod = " + YTD + " " + 
    "VAR ValuePreviousPeriod = " + PYTD + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";

string ΔYoY_PYlabel = 
    "/*ΔYoY_PY*/ " + 
    "VAR ValueCurrentPeriod = " + YTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYTDlabel + " " +
    "VAR Result = " + 
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod" + 
    " ) " +  
    "RETURN " + 
    "   Result ";
  
    
string MOMTD_PYpct = 
    "/*ΔMOM%_PY*/" + 
    "VAR ValueCurrentPeriod = " + MTD + " " + 
    "VAR ValuePreviousPeriod = " + PYMTD + " " + 
    "VAR CurrentMinusPreviousPeriod = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "VAR Result = " + 
    "DIVIDE ( "  + 
    "    CurrentMinusPreviousPeriod," + 
    "    ValuePreviousPeriod" + 
    ") " + 
    "RETURN " + 
    "  Result";


string MOMTD_PYpctLabel = 
    "/*ΔMOM%_PY*/ " +
    "VAR ValueCurrentPeriod = " + MTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYMTDlabel + " " + 
    "VAR Result = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" + 
    " ) " +  
    "RETURN " + 
    "  Result";

    
string QOQTD_PYpct = 
    "/*ΔQOQ%_PY*/" + 
    "VAR ValueCurrentPeriod = " + QTD + " " + 
    "VAR ValuePreviousPeriod = " + PYQTD + " " + 
    "VAR CurrentMinusPreviousPeriod = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "VAR Result = " + 
    "DIVIDE ( "  + 
    "    CurrentMinusPreviousPeriod," + 
    "    ValuePreviousPeriod" + 
    ") " + 
    "RETURN " + 
    "  Result";


string QOQTD_PYpctLabel = 
    "/*ΔQOQ%_PY*/ " +
    "VAR ValueCurrentPeriod = " + QTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYQTDlabel + " " + 
    "VAR Result = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" + 
    " ) " +  
    "RETURN " + 
    "  Result";

    
    
 string YOYTD_PYpct = 
    "/*ΔYOY%_PY*/" + 
    "VAR ValueCurrentPeriod = " + YTD + " " + 
    "VAR ValuePreviousPeriod = " + PYTD + " " + 
    "VAR CurrentMinusPreviousPeriod = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod - ValuePreviousPeriod" + 
    " ) " +  
    "VAR Result = " + 
    "DIVIDE ( "  + 
    "    CurrentMinusPreviousPeriod," + 
    "    ValuePreviousPeriod" + 
    ") " + 
    "RETURN " + 
    "  Result";


string YOYTD_PYpctLabel = 
    "/*ΔYOY%_PY */ " +
   "VAR ValueCurrentPeriod = " + YTDlabel + " " + 
    "VAR ValuePreviousPeriod = " + PYTDlabel + " " + 
    "VAR Result = " +
    "IF ( " + 
    "    NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), " + 
    "     ValueCurrentPeriod & \" vs \" & ValuePreviousPeriod & \" (%)\"" + 
    " ) " +  
    "RETURN " + 
    "  Result";
    
    
string defFormatString = "SELECTEDMEASUREFORMATSTRING()";

//if the flag expression is already present in the format string, do not change it, otherwise apply % format. 
string pctFormatString = 
"IF(" + 
"\n FIND( "+ flagExpression + ", SELECTEDMEASUREFORMATSTRING(), 1, - 1 ) <> -1," + 
"\n SELECTEDMEASUREFORMATSTRING()," + 
"\n \"#,##0.# %\"" + 
"\n)";


//the order in the array also determines the ordinal position of the item    
string[ , ] calcItems = 
    {
        {"Base",      Base,         defFormatString,    "Current year",             Baselabel},
        {"MTD",     MTD,        defFormatString,    "Month-to-date",             MTDlabel},
        {"QTD",     QTD,        defFormatString,    "Quarter-to-date",             QTDlabel},
        {"YTD",     YTD,        defFormatString,    "Year-to-date",             YTDlabel},
        {"ITD",     ITD,        defFormatString,    "Inception-to-date",             ITDlabel},
        {"T3M",     T3M,        defFormatString,    "Trailing 3-months",             T3Mlabel},
        {"TTM",     TTM,        defFormatString,    "Trailing 12-months",             TTMlabel},
        {"PMTD",      PMTD,         defFormatString,    "Previous Month-to-date",            PMTDlabel},
        {"PQTD",      PQTD,         defFormatString,    "Previous Quarter-to-date",            PQTDlabel},
        {"PYTD",      PYTD,         defFormatString,    "Previous Year-to-date",            PYTDlabel},
        {"PYMTD",     PYMTD,        defFormatString,    "Previous year Month-to-date",             PYMTDlabel},
        {"PYQTD",     PYQTD,        defFormatString,    "Previous year Quarter-to-date",             PYQTDlabel},
        {"PYITD",     PYITD,        defFormatString,    "Previous year Inception-to-date",             PYITDlabel},
        {"ΔMoM",     ΔMoM,        defFormatString,    "Month-over-month",             ΔMoMlabel},
        {"ΔQoQ",     ΔQoQ,        defFormatString,    "Quarter-over-quarter",             ΔQoQlabel},
        {"ΔYoY",     ΔYoY,        defFormatString,    "Year-over-year",             ΔYoYlabel},
        {"ΔMoM%",     MOMTDpct,        pctFormatString,    "Month-over-month %",             MOMTDpctLabel},
        {"ΔQoQ%",     QOQTDpct,        pctFormatString,    "Quarter-over-quarter %",             QOQTDpctLabel},
        {"ΔYoY%",     YOYTDpct,        pctFormatString,    "Year-over-year %",             YOYTDpctLabel},
        {"ΔMoM (PY)",     ΔMoM_PY,        defFormatString,    "Month-over-previous-year-month",             ΔMoM_PYlabel},
        {"ΔQoQ (PY)",     ΔQoQ_PY,        defFormatString,    "Quarter-over-previous-year-quarter",             ΔQoQ_PYlabel},
        {"ΔYoY (PY)",     ΔYoY_PY,        defFormatString,    "Year-over-previous-year",             ΔYoY_PYlabel},
        {"ΔMoM% (PY)",     MOMTD_PYpct,        pctFormatString,    "Month-over-previous-year-month %",             MOMTD_PYpctLabel},
        {"ΔQoQ% (PY)",     QOQTD_PYpct,        pctFormatString,    "Quarter-over-previous-year-quarter %",             QOQTD_PYpctLabel},
        {"ΔYoY% (PY)",     YOYTD_PYpct,        pctFormatString,    "Year-over-previous-year %",             YOYTD_PYpctLabel},
        

  
    };

    
int j = 0;


//create calculation items for each calculation with formatstring and description
foreach(var cg in Model.CalculationGroups) {
    if (cg.Name == calcGroupName) {
        for (j = 0; j < calcItems.GetLength(0); j++) {
            
            string itemName = calcItems[j,0];
            
            string itemExpression = calcItemProtection.Replace("<CODE>",calcItems[j,1]);
            itemExpression = itemExpression.Replace("<LABELCODE>",calcItems[j,4]); 
            
            string itemFormatExpression = calcItemFormatProtection.Replace("<CODE>",calcItems[j,2]);
            itemFormatExpression = itemFormatExpression.Replace("<LABELCODEFORMATSTRING>","\"\"\"\" & " + calcItems[j,4] + " & \"\"\"\"");
            
            //if(calcItems[j,2] != defFormatString) {
            //    itemFormatExpression = calcItemFormatProtection.Replace("<CODE>",calcItems[j,2]);
            //};

            string itemDescription = calcItems[j,3];
            
            if (!cg.CalculationItems.Contains(itemName)) {
                var nCalcItem = cg.AddCalculationItem(itemName, itemExpression);
                nCalcItem.FormatStringExpression = itemFormatExpression;
                nCalcItem.FormatDax();
                nCalcItem.Ordinal = j; 
                nCalcItem.Description = itemDescription;
                
            };




        };

        
    };
};