#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;

// by Johnny Winter
// www.greyskullanalytics.com
// '2021-10-15 / B.Agullo / dynamic parameters by B.Agullo / 

// Instructions:
//select the measures you want to add to your Dynamic Measure and then run this script (or store it as macro)

//
// ----- do not modify script below this line -----
//

if (Selected.Measures.Count == 0) {
  Error("Select one or more measures");
  return;
}

string calcGroupName = Interaction.InputBox("Provide a name for your Calc Group", "Calc Group Name", "Dynamic Measure", 740, 400);
if(calcGroupName == "") return;

string columnName = Interaction.InputBox("Calc Group column name", "Column Name", calcGroupName, 740, 400);
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



string measureName = Interaction.InputBox("Dynamic Measure Name (cannot be named \"" + columnName +"\")", "Measure Name", "Placeholder for Dynamic Measure", 740, 400);
if(measureName == "") return;

string switchSuffix = Interaction.InputBox("suffix for the SWITCH dynamic measure", "Suffix for switch", "SWITCH", 740, 400);
if(switchSuffix == "") return;

string formattedSuffix = Interaction.InputBox("suffix for the FORMATTED dynamic measure", "Suffix for formatted", "FORMATTED", 740, 400);
if(formattedSuffix == "") return;

string measureDefault = Interaction.InputBox("Measure default value", "Default Value", "BLANK()", 740, 400);
if(measureDefault == "") return;

//by default the calc group has a column called Name. If this column is still called Name change this in line with specfied variable
if (calcGroup.Columns.Contains("Name")) {
  calcGroup.Columns["Name"].Name = columnName;
};
calcGroup.Columns[columnName].Description = "Select value(s) from this column to control the contents of the dynamic measures.";

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
  newMeasure.Description = "Control the content of this measure by selecting values from " + columnName + ".";
};

//create calculation items based on selected measures, including check to make sure calculation item doesnt exist
foreach(var cg in Model.CalculationGroups) {
  if (cg.Name == calcGroupName) {
    foreach(var m in Selected.Measures) {
      if (!cg.CalculationItems.Contains(m.Name)) {
        var newCalcItem = cg.AddCalculationItem(
        m.Name, "IF ( " + "ISSELECTEDMEASURE ( [" + measureName + "] ), " + "[" + m.Name + "], " + "SELECTEDMEASURE() )");
        // '2021-10-15 / B.Agullo / double quotes in format string need to be doubled to be preserved
        newCalcItem.FormatStringExpression = "IF ( " + "ISSELECTEDMEASURE ( [" + measureName + "] ),\"" + m.FormatString.Replace("\"","\"\"") + "\", SELECTEDMEASUREFORMATSTRING() )";
        newCalcItem.FormatDax();
      };
    };
  };
};

//check to see if SWITCH dynamic measure has been created, if not create it now
//if a measure with that name alredy exists elsewhere in the model, throw an error
string switchMeasureName = measureName + " " + switchSuffix;
if (!calcGroup.Measures.Contains(switchMeasureName)) {
  foreach(var m in Model.AllMeasures) {
      if (m.Name == switchMeasureName) {
      Error("This measure name already exists in table " + m.Table.Name + ". Either rename the existing measure or choose a different name for the measure in your Calculation Group.");
      return;
    };
  };
  var newMeasure = calcGroup.AddMeasure(switchMeasureName);
  newMeasure.Description = "Control the content of this measure by selecting values from " + columnName + ".";
};

//check to see if FORMATTED dynamic measure has been created, if not create it now
//if a measure with that name alredy exists elsewhere in the model, throw an error
string formattedMeasureName = measureName + " " + formattedSuffix;
if (!calcGroup.Measures.Contains(formattedMeasureName)) {
  foreach(var m in Model.AllMeasures) {
      if (m.Name == formattedMeasureName) {
      Error("This measure name already exists in table " + m.Table.Name + ". Either rename the existing measure or choose a different name for the measure in your Calculation Group.");
      return;
    };
  };
var newMeasure = calcGroup.AddMeasure(formattedMeasureName);
  newMeasure.Description = "Control the content of this measure by selecting values from " + columnName + ".";
};

//create DAX for SWITCH and FORMATTED measures
string switchItemList = "";
string formattedItemList = "";
foreach(var cg in Model.CalculationGroups) {
  if (cg.Name == calcGroupName) {
      foreach (var ci in cg.CalculationItems) {
          switchItemList = switchItemList + "\"" + ci.Name + "\", [" + ci.Name + "],";
          string formatString = "";          
          foreach(var m in Model.AllMeasures) {
              if (m.Name == ci.Name) {
                  formatString = m.FormatString;
              };
          }; 
          formattedItemList = formattedItemList + "\"" + ci.Name + "\", FORMAT( [" + ci.Name + "], \"" + formatString.Replace("\"","\"\"") + "\"),";
      };
    };
};

//assign SWITCH measure DAX
Measure switchMeasure = calcGroup.Measures[switchMeasureName];
switchMeasure.Expression = 
    "SWITCH ( SELECTEDVALUE('" + calcGroupName + "'[" + columnName + "])," +
    switchItemList +
    measureDefault + ")";
    switchMeasure.FormatDax();

//assign FORMATTED measure DAX
Measure formattedMeasure = calcGroup.Measures[formattedMeasureName];
formattedMeasure.Expression = 
    "SWITCH ( SELECTEDVALUE('" + calcGroupName + "'[" + columnName + "])," +
    formattedItemList +
    measureDefault + ")";
    formattedMeasure.FormatDax();

    
//check to see if Display measure has been created, if not create it now
//if a measure with that name alredy exists elsewhere in the model, throw an error
string displayMeasureName = "Display selected " + measureName + "(s)";
if (!calcGroup.Measures.Contains(displayMeasureName)) {
  foreach(var m in Model.AllMeasures) {
      if (m.Name == displayMeasureName) {
      Error("This measure name already exists in table " + m.Table.Name + ". Either rename the existing measure or choose a different name for the measure in your Calculation Group.");
      return;
    };
  };
var newMeasure = calcGroup.AddMeasure(displayMeasureName);
  newMeasure.Description = "This measure displays a concatenated list of selections from " + columnName + ".";
};
calcGroup.Measures[displayMeasureName].Expression = 
    "CONCATENATEX('" + calcGroupName + "', '" + calcGroupName + "'[" + columnName + "], \", \")";