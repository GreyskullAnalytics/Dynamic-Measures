// by Johnny Winter
// www.greyskullanalytics.com

#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;

if (Selected.Measures.Count == 0) {
  Error("Select one or more measures");
  return;
}
//select the measures you want to add to your Dynamic Measure and then run this script
//change the next 4 string variables for different naming conventions



//add the name of your calculation group here
string calcGroupName = Interaction.InputBox("Provide a name for your Calculation Group", "Calculation Group Name", "@Dynamic Measure", 740, 400);
if(calcGroupName == "") return;

//add the name for the column you want to appear in the calculation group
string columnName = Interaction.InputBox("Provide a name for the Calculation Group measure selection column", "Column Name", "Measure Selection", 740, 400);
if(columnName == "") return;

//check to see if a table with this name already exists
//if it doesnt exist, create a calculation group with this name
if (!Model.Tables.Contains(calcGroupName)) {
  var cg = Model.AddCalculationGroup(calcGroupName);
  cg.Description = "Contains dynamic measures and a column called " + columnName + ". The contents of the dynamic measures can be controlled by selecting values from " + columnName + ".";
};
//set variable for the calc group
Table calcGroup = Model.Tables[calcGroupName];

//if table already exists, make sure it is a Calculation Group type
if (calcGroup.SourceType.ToString() != "CalculationGroup") {
  Error("Table exists in Model but is not a Calculation Group. Rename the existing table or choose an alternative name for your Calculation Group.");
  return;
};

//add the name of the dynamic measure here
string measureName = Interaction.InputBox("Provide a name for the Dynamic Measure", "Dynamic Measure Name", "Dynamic Measure", 740, 400);
if(measureName == "") return;

//create a default value for the dynamic measure. 
string measureDefault = "BLANK()";

//by default the calc group has a column called Name. If this column is still called Name change this in line with specfied variable
if (calcGroup.Columns.Contains("Name")) {
  calcGroup.Columns["Name"].Name = columnName;
};

//check to see if dynamic measure has been created, if not create it now
//if a measure with that name alredy exists elsewhere in the model, throw an error
if (!calcGroup.Measures.Contains(measureName)) {
  foreach(var m in Model.AllMeasures) {
    if (m.Name == measureName) {
      Error("This measure name already exists in table " + m.Table.Name + ". Either rename the existing measure or choose a different name for the measure in your Calculation Group.");
      return;
    };
  };
  var newMeasure = calcGroup.AddMeasure(
  measureName, measureDefault);
};

//create calculation items based on selected measures, including check to make sure calculation item doesnt exist
foreach(var cg in Model.CalculationGroups) {
  if (cg.Name == calcGroupName) {
    foreach(var m in Selected.Measures) {
      if (!cg.CalculationItems.Contains(m.Name)) {
        var newCalcItem = cg.AddCalculationItem(
        m.Name, "IF ( " + "ISSELECTEDMEASURE ( [" + measureName + "] ), " + "[" + m.Name + "], " + "SELECTEDMEASURE() )");
        newCalcItem.FormatStringExpression = "IF ( " + "ISSELECTEDMEASURE ( [" + measureName + "] ),\"" + m.FormatString + "\", SELECTEDMEASUREFORMATSTRING() )";
      };
    };
  };
};
