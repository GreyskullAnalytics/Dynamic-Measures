// by Johnny Winter
// www.greyskullanalytics.com

//select the measures you want to add to your Dynamic Measure and then run this script
//change the next 4 string variables for different naming conventions

//add the name of your calculation group here
string calcGroupName = "@Dynamic Measure";

//add the name for the column you want to appear in the calculation group
string columnName = "Measure Selection";

//add the name of the dynamic measure here
string measureName = "Dynamic Measure";

//create a default value for the dynamic measure. 
string measureDefault = "BLANK()";
//
// ----- do not modify script below this line -----
//

if (Selected.Measures.Count == 0) {
  Error("Select one or more measures");
  return;
}

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

