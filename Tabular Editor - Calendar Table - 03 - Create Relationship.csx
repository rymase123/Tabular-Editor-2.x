 // Prompt for Table Selection
 
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
        
string factTableName = factTable.Name;
string dateTableName = dateTable.Name;
string dateTableDateColumnName = dateTableDateColumn.Name;
string factTableDateColumnName = factTableDateColumn.Name;




var fromColumn = Model.Tables[factTableName].Columns[factTableDateColumnName] ; // Enter the 'from' part of the relationship
var toColumn = Model.Tables[dateTableName].Columns[dateTableDateColumnName ]; // Enter the 'to' part of the relationship

{
    var r = Model.AddRelationship();
    r.FromColumn = fromColumn;
    r.ToColumn = toColumn;
    r.FromCardinality = RelationshipEndCardinality.Many;
    r.ToCardinality = RelationshipEndCardinality.One;
    r.CrossFilteringBehavior = CrossFilteringBehavior.OneDirection; //CrossFilteringBehavior.BothDirections
}