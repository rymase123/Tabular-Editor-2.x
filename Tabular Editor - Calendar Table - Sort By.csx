
{
    // Set the sort by column

(Selected.Table.Columns["Quarter"] as CalculatedTableColumn).SortByColumn = (Selected.Table.Columns["Year Quarter"] as CalculatedTableColumn);
(Selected.Table.Columns["Month Name"] as CalculatedTableColumn).SortByColumn = (Selected.Table.Columns["Month Number"] as CalculatedTableColumn);


}
