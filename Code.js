// Globals
// This is the configuration of the system
var Base_Dir_ID = "1ufmdIwTijqfL79-IzRITwujL4UDf8XwZ";
var Backup_Dir_ID = "18tsT32WmOf-DAavLQZyyMe6ZTpPacaxW"; 
var Output_Dir_ID = "1iT07s-0Ddmm91ib4jVwIjvTSh0Ich7Cv";

//
// function is_not_dept_specs_sheet
//
// Params:
//     sheet: Sheet
//
// Returns:
//     true/false
//
// Requirements:
//     nothing
//
// Description:
//     Rturns true if the give sheet is not a department specifications sheet, false otherwise.
//
// Notes:
//     The array of specification and administration sheets should be part of the confifuration of CTables19.
//
function is_not_dept_specs_sheet(sheet)
{
      var sheet_name = sheet.getName();
      
      var array = ["Welcome", "Depts", "Items", "Groups", "MasterSheet", "ΚΑΕ", "ΚΑΕ-Είδη", "TemplateList", "TemplateSpecs", "SandBox", "ToDo", "Admin", "BudgetCheck", "Trash", "Dept-Budgets", "TestDept-1111"];
      
      if ( array.indexOf(sheet_name) != -1 )
      {
          return true;
      }
      else
      {
          return false;
      }
}

//
// function check_budgets
//
// Parameters:
//     nothing
//
// Returns:
//     nothing
//
// Requirements:
//     Requires the existence of the sheet with the name "BudgetCheck".
//
// Description:
//     This function operates on the curretly active spreadsheet.
//     Iterates over the department specification sheets 
//
function check_budgets()
{
    var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = active_spreadsheet.getSheets();
    
    var d_sheet = active_spreadsheet.getSheetByName("Depts");
    var d_data = d_sheet.getRange(1, 1, 23, 7).getValues();
    
    var b_sheet = active_spreadsheet.getSheetByName("BudgetCheck");
    b_sheet.clear();
    var c_date = Utilities.formatDate(new Date(), "Europe/Athens", "yyyy-MM-dd HH:mm:ss");
    
    for (i=0; i<sheets.length; i++)
    {

      if ( is_not_dept_specs_sheet(sheets[i]) )
          continue;
      
      dept_name = sheets[i].getRange(1,2).getValue();
      dept_kae  = sheets[i].getRange(2,2).getValue();
      dept_cost = sheets[i].getRange(3,4).getValue();
      
      for (k=0; k<d_data.length; k++)
      {
        if (d_data[k][0] === dept_name )
        {
            Logger.log(dept_name + " found");
            Logger.log("Dept KAE: " + dept_kae);
            Logger.log("Dept Cost: " +  dept_cost);
            var kae_col = -1;
            
            // Lookup the KAE column
            for (j=2; j<d_data[0].length; j++)
            {
                if ( dept_kae === Number(d_data[0][j]))
                {
                    kae_col = j;
                    break;
                }
            }            

            if ( dept_cost > d_data[k][kae_col] )
            {
                // do something
                Logger.log("Dept Limit: " + d_data[k][kae_col]);
                Logger.log(dept_name + " εκτός προϋπολογισμού στον ΚΑΕ " + dept_kae);
                //Browser.msgBox(dept_name + " εκτός προϋπολογισμού στον ΚΑΕ " + dept_kae);
                b_sheet.appendRow([dept_name, dept_kae, "Εκτός προϋπολογισμού"]);
            }
            else
            {
              b_sheet.appendRow([dept_name, dept_kae, "OK"]);
            }
            
            b_sheet.autoResizeColumns(1, 3);
            
            Logger.log("===================");
            break;
        } // if given dept found
      }
    }
}

// Use this only once!!
//function createDeptDirectories()
//{
//    var a_ss = SpreadsheetApp.getActiveSpreadsheet();
//    var depts_sheet = a_ss.getSheetByName("Depts");
//    //var pt19_id = "18x81BnJTQZnEhqhC3TSneaKiCdBdj_7J";
//    //var pt20_id = "1a-wMQQO7DX7sYGJKVqkGj_XKsKuWt8T9";
//    
//    var dept_names = depts_sheet.getRange(2, 1, 22, 1).getValues();
//    
//    for (k=0; k<dept_names.length; k++)
//    {
//        Logger.log(dept_names[k]);
//        var f1 = DriveApp.getFolderById(Base_Dir_ID);
//        f1.createFolder(dept_names[k]);
//    }
//}

function spreadsheetBackup()
{
  var a_ss = SpreadsheetApp.getActiveSpreadsheet();
  var current_name = a_ss.getName();
  //a_ss.copy(current_name + "-" + Utilities.formatDate(new Date(), "Europe/Athens", "yyyyMMddHHmmss"));
  //var Backup_Folder_ID = "1mCNe3rNQ6lVq443ZVj7MhHVy3yn6fq6f";

  //var destFolder = DriveApp.getFolderById("1YygvOeZajkK8Ovn1hIsXIME38FJyWjV8");
  var destFolder = DriveApp.getFolderById(Backup_Dir_ID);
  DriveApp.getFileById(a_ss.getId()).makeCopy(current_name + "-" + Utilities.formatDate(new Date(), "Europe/Athens", "yyyyMMddHHmmss"), destFolder);
  
}

function moveSpreadsheetToOutputDir(new_ss)
{
   //var output_dir_id = "13kU8KolbGtq-w-IgTIODAGeY6lLSyZvH";

   var file = DriveApp.getFileById(new_ss.getId());
   file.getParents().next().removeFile(file);
   DriveApp.getFolderById(Output_Dir_ID).addFile(file);
}

//
// function sheetName
//
// Parameters:
// 
// Returns:
//
// Description:
//     Returns the name of the active sheet.
//
function sheetName() 
{
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

//
// function BOM_per_dept
//
// Parameters:
//    nothing
// 
// Returns:
//    nothing
//
// Description:
//     Creates a new spreadsheet with a list of equipment for each department.
//     The name of the new spreadsheet is "BOM_per_dept-<TIMESTAMP>".
//     The fields of information that are listed are: <ΟΜΑΔΑ>, <ΠΕΡΙΓΡΑΦΗ>, <ID>, <ΤΕΜΑΧΙΑ>, <ΚΟΣΤΟΣ/ΤΕΜΑΧΙΟ>, <ΕΝΔΕΙΚΤΙΚΟ ΜΟΝΤΕΛΟ>.
//
function BOM_per_dept()
{
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = active_spreadsheet.getSheets();
  
  var new_ss = SpreadsheetApp.create("BOM_per_dept-" + Utilities.formatDate(new Date(), "Europe/Athens", "yyyyMMddHHmmss"));
  
  // Iterate all sheets of current active spreadsheet.
  for (var i=0; i<sheets.length; i++)
  {
    var sheet_name = sheets[i].getName();
    var dept_kae  = sheets[i].getRange(2,2).getValue();
    
    if ( is_not_dept_specs_sheet(sheets[i]) )
          continue;
          
//    if ( ["Depts", "Items", "Groups", "MasterSheet", "ΚΑΕ", "ΚΑΕ-Είδη", "TemplateList", "TemplateSpecs", "SandBox", "ToDo", "Admin", "BudgetCheck", "Trash"].indexOf(sheet_name) != -1 )
//    { 
//      // Ignore the meta-sheets (i.e. sheets that contain meta-information.
//      // We are interesting only on sheets that contain equipment.
//      continue;
//    }
    
    var lastRow = sheets[i].getLastRow();
    
    // Get as much data as possible. Every getValue() is expensive.
    var col2_data = sheets[i].getRange(1,2,lastRow,1).getValues();
    var col3_data = sheets[i].getRange(1,3,lastRow,1).getValues();
    
    // Iteratet over rows (i.e. run vertically over column 2)
    for (j=0; j<col2_data.length; j++)
    {
      if ( String(col2_data[j]) === "BEGIN" )
      {
        // Here a table of an item is found.
        
        var department = String(col2_data[0]);
        var group = String(col3_data[j-1]);
        var table_title = String(col2_data[j+1]);
        var table_id    = String(col3_data[j+1]);
        var num_of_items = Number(col3_data[j-2]);
        var cost_per_item = Number(col3_data[j-3]);
        var indicating_model = String(col3_data[j-8]);
        
        dsheet = new_ss.getSheetByName(department);
        if (dsheet == null)
        {
          // If there is not yet sheet for this department, create a new sheet
          dsheet = new_ss.insertSheet(department);
          var values =[["ΟΜΑΔΑ", "ΠΕΡΙΓΡΑΦΗ", "ID", "ΤΕΜΑΧΙΑ", "ΚΟΣΤΟΣ/ΤΕΜ.", "ΚΑΕ", "ΕΝΔΕΙΚΤΙΚΟ ΜΟΝΤΕΛΟ", "ΕΓΚΡΙΝΕΤΑΙ"]];
          dsheet.getRange(1, 1, 1, 8).setValues(values);
          dsheet.getRange(1, 1, 1, 8).setFontWeight("bold");
        }

        var accepted = "ΟΧΙ";
        if (is_accepted(sheets[i], j+1) )
            accepted = "ΝΑΙ";
        var dlastRow = dsheet.getLastRow();
        var values = [[group, table_title, table_id, num_of_items, cost_per_item, dept_kae, indicating_model, accepted]];

        dsheet.getRange(dlastRow+1, 1, 1, 8).setValues(values);
        dsheet.autoResizeColumns(1, 8);
      }
    }
    //sheets[i].autoResizeColumns(1, 7);
  } // loop over sheets
  
  moveSpreadsheetToOutputDir(new_ss);

} // BOM_per_dept


//
// function BOM_per_group
// 
// Parameters:
//     BOM: data srtucture
//
// Returns:
//     nothing
//
// Description:
//     To be called by createPTTables.
//     Creates a new spreadsheet with a list of equipment per group.
//     This spreasheet is to be delivered to suppliers in order to know
//     which items goes to which department.
//     The BOM data structure is being created 
// 
// Creates a list of equipment per procurement group and per department.
// The format of the list is:
// <ITEM DESCRIPTION><ID><DEPARTMENT><NUM_OF_ITEMS><COST/ITEM><INDICATING_MODEL>
//
function BOM_per_group(BOM)
{
  var new_ss = SpreadsheetApp.create("BOM_per_group-" + Utilities.formatDate(new Date(), "Europe/Athens", "yyyyMMddHHmmss"));
  
  for (group in BOM)
  {
    // create a sheet
    // insert values
    var group_sheet = new_ss.insertSheet(group)
    var num_of_cols = 7;
    var num_of_rows = BOM[group].length;
    group_sheet.getRange(1, 1, 1, num_of_cols).setValues([["ΠΕΡΙΓΡΑΦΗ", "ID", "ΤΜΗΜΑ", "ΤΕΜΑΧΙΑ", "ΚΟΣΤΟΣ/ΤΕΜΑΧΙΟ", "KAE", "ΕΝΔEIKTIKO ΜΟΝΤΕΛΟ"]]);
    group_sheet.getRange(2, 1, num_of_rows, num_of_cols).setValues(BOM[group]);
    group_sheet.getRange(1, 1, num_of_rows+1, num_of_cols).setHorizontalAlignment("center");
    group_sheet.autoResizeColumns(1, num_of_cols);
  }
  
  moveSpreadsheetToOutputDir(new_ss);
}

//
// function
// To be called by createPTTables.
// Returns true if the conformity table that starts from row <row> in <sheet>
// is accepted by the Coordinator--Auditors--Department commitee.
//
function is_accepted(sheet, row)
{
    var a_data = sheet.getRange(row+1, 4, 1 ,5).getValues(); //+1 because we count array positions
    //Logger.log("a_data:");
    //Logger.log(a_data);
    //Logger.log(String(a_data[0][0]));
    //Logger.log(typeof a_data[0][0]);
    //if ( a_data[0][0] === true ||  // Departmental Veto
    if ( a_data[0][1] === false || // Auditor 1 approvement
         a_data[0][2] === false || // Auditor 2 approvement
         a_data[0][3] === false || // Auditor 3 approvement
         a_data[0][4] === false )  // Coordinator approvement
    {
      //Logger.log("Table not acceted");
      return false;
    }
    else
    {
      //Logger.log("Table acceted");
      return true;
     }
}

//
// function is_template (to be called by createPTTables)
// 
// Parameters:
//     table: the array of the "landmarks" of the conformity table.
//     col2_data: array of data of column 2 in the current sheet
//     col3_data: array of data of column 3 in the current sheet
//     active_spreadsheet: the active spreadsheet
//
// Returns:
//     true if it's a template.
//     false otherwise
//
// Description:
//     Scans the sheet of the templates ("TemplateSpecs") and if
//     conformity table that is indicated by table is found, then
//     true is returnt, otherwise false.
//
function is_template(table, col2_data, col3_data, active_spreadsheet)
{
  var template_specs_sheet = active_spreadsheet.getSheetByName("TemplateSpecs");
  var template_specs_last_row = template_specs_sheet.getDataRange().getLastRow();
  var t_col2_data = template_specs_sheet.getRange(1,2,template_specs_last_row,1).getValues();
  var t_col3_data = template_specs_sheet.getRange(1,3,template_specs_last_row,1).getValues();
  
  var retval = new Object();
  
  table_id = String(col3_data[table["BEGIN"]+1]);
  for (i=0; i<t_col2_data.length; i++)
  {
    if ( String(t_col2_data[i]) === "BEGIN" )
    {
        if ( String(t_col3_data[i+1]) === table_id )
        {
           retval["IS_TEMPLATE"] = true;
           break;
        }
    }
  }
  
  if (i < t_col2_data.length)
  {
    retval["LANDMARKS"] = table_landmarks(i, t_col2_data);
  }
  else
    retval["IS_TEMPLATE"] = false;

  //Logger.log(retval);
  return retval;
  
}

//
// To be called by createPTTables.
// Returns the group_name of the group that its description is <group_descrption>
//
function get_group_name(group_description, g_data1, g_data2)
{
    var group_name = "GROUP_NAME";
    for (k=0; k<g_data1.length; k++)
    {
      if ( String(g_data2[k]) === group_description )
      {
          group_name = String(g_data1[k]);
          break;
      }
    }
    
    return group_name;
}



//
// To be called by createPTTables.
// Creates and returns a new Sheet for a procurement group.
//
function newGroupSheet(ss, group_name, group_description)
{
    var group_sheet = ss.insertSheet(group_name);
    
    group_sheet.getRange(1, 2).setValue("Ομάδα " + group_name + ": " + group_description);
    group_sheet.getRange(1, 2).setFontSize(14);
    group_sheet.getRange(2, 2).setValue(0);
    group_sheet.getRange(2, 2).setNumberFormat("€ 00.00");
    group_sheet.getRange(2, 2).setFontSize(14);
    group_sheet.getRange(2, 2).setHorizontalAlignment("center");
    group_sheet.getRange(1,2,2,4).setBackground("lightgray");
    group_sheet.getRange(1,2,2,4).setFontWeight("bold");
    group_sheet.getRange(1,2,2,4).setBorder(true, true, true, true, true, true);

    return group_sheet;
}

function get_group_sheet(ss, group_name, group_description)
{
  var group_sheet = ss.getSheetByName(group_name);

  if (group_sheet == null)
    group_sheet = newGroupSheet(ss, group_name, group_description);
  
  return group_sheet;
}

//
// function insertPTTable:
//
// Parameters:
//     table: indexes of table limits (DOCS, BEGIN, END)
//     table_counters: tables counters for each sheet in the PTTables
//     col2_data: data of column 2 of current sheet
//     col3_data: data of column 3 of current sheet
//     group_sheet: sheet into which table is going to be inserted
//     num_of_items: total number of items of the specific table
//
// To be called by createPTTables. Inserts another table
//
function insertPTTable(table, table_counters, col2_data, col3_data, group_sheet, num_of_items)
{
    //Logger.log(table);
    var g_lastRow = group_sheet.getLastRow();
    var group_name = group_sheet.getName();
    
    if (table["ALREADY_EXISTS"])
    {
        // Do something: just update num_of_items
        //Logger.log("Exists!!" + table["TABLE_ID"]);
        // table_id should also be inside table
        
        var g_col3_data = group_sheet.getRange(1,3,g_lastRow,1).getValues();
        var row_found = -1;
        for (k=0; k<g_col3_data.length; k++)
        {
            //Logger.log("11111111111111111111: " + g_col3_data[k]);
            //Logger.log("22222222222222222222:" + table["TABLE_ID"]);
           if ( String(g_col3_data[k]) === table["TABLE_ID"] )
           {
               row_found = k+1;
               break;
           }
        }
//        if (row_found === -1)
//        {
//            Logger.log("AAAAAAAAAAAAAAAAAAAAAAAAA");
//            Logger.log("table_id: " + table["TABLE_ID"] + " " + typeof table["TABLE_ID"]);
//            Logger.log(g_col3_data);
//        }
//        var values = [[group_name + table_counters[group_name] + ".1", "Τεμάχια", ">=" + num_of_items, "", "" ]];
//        group_sheet.getRange(row_found+2,1,1,5).setValues(values);
        
        var values = [["Τεμάχια", ">=" + num_of_items, "", "" ]];
        group_sheet.getRange(row_found+2,2,1,4).setValues(values);
        
        return -213;
    }
    
    // Create Table Title

    var values = [
      [group_name + table_counters[group_name], String(col2_data[table["BEGIN"]+1]), String(col3_data[table["BEGIN"]+1]), table["COST"], "" ],
      ["", "Προδιαγραφή", "Απαίτηση", "Απάντηση Προμηθευτή", "Σχόλιο"],
      [group_name + table_counters[group_name] + ".1", "Τεμάχια", ">=" + num_of_items, "", "" ],
    ];
    // Range starting from (go_lastRow+4,1) and expands to 2 rows and three columns
    group_sheet.getRange(g_lastRow+4, 1, 3, 5).setValues(values);
    group_sheet.getRange(g_lastRow+4, 4, 1, 1).setNumberFormat("Μέγιστο κόστος ανά τεμάχιο: € 00.00");
    group_sheet.getRange(g_lastRow+4, 4, 1, 1).setHorizontalAlignment("center");
    var this_range = group_sheet.getRange(g_lastRow+4, 1, 2, 5);
    this_range.setBackground("lightgray");
    this_range.setFontWeight("bold");


    var data_pointer = table["BEGIN"]+1+2; //reading from dataPointer array elements and writing to rowPointer rows
    var rowPointer = g_lastRow+4+3;
    values = [];
    for (var table_row_counter = 2; table_row_counter<=table["SPECLEN"]-2; table_row_counter++)
    {
       values.push([group_name + table_counters[group_name] + "." + table_row_counter, col2_data[data_pointer], col3_data[data_pointer] ]);
       rowPointer++;
       data_pointer++;
    }
    group_sheet.getRange(g_lastRow+4+3,1, table["SPECLEN"]-3, 3).setValues(values);
    
    // formatting
    group_sheet.autoResizeColumns(1, 6);
    group_sheet.setColumnWidth(5, 200);
    group_sheet.setColumnWidth(2, 500);
    group_sheet.setColumnWidth(5, 200);
    //group_sheet.setColumnWidth(6,100);
    group_sheet.getRange(g_lastRow+4,1, table["SPECLEN"], 5).setBorder(true, true, true, true, true, true);
    group_sheet.getRange(g_lastRow+4,1, table["SPECLEN"], 5).setWrap(true);
    group_sheet.getRange(g_lastRow+4,3, table["SPECLEN"], 1).setHorizontalAlignment("center");
}

function table_landmarks(j, col2_data)
{
    var table = new Object();
    table["DOCS"] = j;
    table["END"] = j;
    table["BEGIN"] = j;
    for (k=j; k<col2_data.length; k++)
    {
      if ( String(col2_data[k]) === "BEGIN" )
      {
        table["BEGIN"] = k;
      }
      if ( String(col2_data[k]) === "END" )
      {
        table["END"] = k;
        break;
      }
    }
    table["SPECLEN"] = table["END"] - table["BEGIN"];
    
    return table;
}

function get_template_col_data(k, active_spreadsheet)
{
  var t_sheet = active_spreadsheet.getSheetByName("TemplateSpecs");
  var lastRow = t_sheet.getLastRow();
  
  return t_sheet.getRange(1,k, lastRow, 1).getValues();
}

//
// function createPTTables
//
// Scan all spreadsheet and create conformity tables
// with groups and budget, ready for Procurment Tendering.
//
// Scan all sheets except for Depts, Groups, Items

/*
 * Creates bla bla
 * @function
 * @argument {none}
 * @returns {none}
 */
function createPTTables()
{
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = active_spreadsheet.getSheets();
  
  var new_ss = SpreadsheetApp.create("PTTables-" + Utilities.formatDate(new Date(), "Europe/Athens", "yyyyMMddHHmmss"));

  var sheet_of_groups = active_spreadsheet.getSheetByName("Groups");
  var sheet_of_groups_lastRow = sheet_of_groups.getLastRow();
  var col1_group_data = sheet_of_groups.getRange(1,1,sheet_of_groups_lastRow,1).getValues();
  var col2_group_data = sheet_of_groups.getRange(1,2,sheet_of_groups_lastRow,1).getValues();
  
  var BOM = new Object(); // BOM per group
  
  var Group_Budget = new Object();
  
  // Table counters for each group
  var table_counters = new Object();
  for (i=1; i<col1_group_data.length; i++)
  { 
    table_counters[col1_group_data[i]] = 0;
  }
  
  // counts items with unique ID
  var unique_items_counter = new Object();

  var template_col2_data = get_template_col_data(2, active_spreadsheet);
  var template_col3_data = get_template_col_data(3, active_spreadsheet);
  //Logger.log(template_col2_data);
  //Logger.log(template_col3_data);

  for (var i=0; i<sheets.length; i++)
  { // Scann all sheets
  
    var sheet_name = sheets[i].getName();
    
    if ( is_not_dept_specs_sheet(sheets[i]) )
          continue;
          
    var KAE = sheets[i].getRange(2, 2).getValue();
    
//    if ( ["Depts", "Items", "Groups", "MasterSheet", "ΚΑΕ", "ΚΑΕ-Είδη", "TemplateList", "TemplateSpecs", "SandBox", "ToDo", "Admin", "BudgetCheck", "Trash"].indexOf(sheet_name) != -1 )
//    {
//      continue;
//    }
    
    //var sheet_data = sheets[i].getDataRange();
    var lastRow = sheets[i].getLastRow();
    
    // Get as much data as possible. Every getValue() is expensive.
    var col2_data = sheets[i].getRange(1,2,lastRow,1).getValues();
    var col3_data = sheets[i].getRange(1,3,lastRow,1).getValues();
    
    for (j=0; j<col2_data.length; j++)
    {
      //Logger.log("j: " + j);
      // As soon as i get the word "DOCS" i know that it's a table with its at word "END"
      if (String(col2_data[j]) === "DOCS")
      {
        var table = table_landmarks(j,col2_data);
        
        // If table is not yet accepted, go to the next
        table["IS_ACCEPTED"] = is_accepted(sheets[i], table["BEGIN"]+1);
        //table["IS_ACCEPTED"] = true; // for the moment
        if ( table["IS_ACCEPTED"] === false )
        {
          j = table["END"];
          continue; // For the moment accept everything until development is finished
        }
        
        table["TEMPLATE"] = is_template(table, col2_data, col3_data, active_spreadsheet); 

        var group_description = String(col3_data[table["DOCS"] + 8]);
        var group_name = get_group_name(group_description, col1_group_data, col2_group_data);
        var group_sheet = get_group_sheet(new_ss, group_name, group_description);
        
//        if ( group_name != "ComputerParts" )
//        {
//            continue;
//        }
        
        if ( !(group_name in Group_Budget) )
        {
          Group_Budget[group_name] = 0;
          Logger.log("0000000000000000000");
        }
        Group_Budget[group_name] += ( Number(col3_data[table["BEGIN"]-2])*Number(col3_data[table["BEGIN"]-3]) ); //items*cost_per_item
        Logger.log("================");
        Logger.log("ITEM: " + String(col2_data[table["BEGIN"]+1]) );
        Logger.log("Num_Of_Items: " + String(col3_data[table["BEGIN"]-2]));
        Logger.log("Cost_per_item: " + String(col3_data[table["BEGIN"]-3]));
        Logger.log("Group Name: " + group_name);
        Logger.log(Group_Budget);
        Logger.log("================");
        
        if ( !(group_name in BOM) )
        {
          BOM[group_name] = []; // Creates a new group
        }
        BOM[group_name].push([
          String(col2_data[table["BEGIN"]+1]), // Item description
          String(col3_data[table["BEGIN"]+1]), // Table_it
          String(col2_data[0]),                // Department
          String(col3_data[table["BEGIN"]-2]), // Number of Items
          String(col3_data[table["BEGIN"]-3]), // Cost per item
          String(KAE),
          String(col3_data[table["DOCS"]+1]),  // Model
        ]);

        // Εντώ είναι το προμπλέμα. Μεταφέρθηκε παρακάτω.
        // table_counters[group_name]++;
        
        //Logger.log("col3_data[table[BEGIN]+1] = " + col3_data[table["BEGIN"]+1]);
        var table_id = String(col3_data[table["BEGIN"]+1]);
        table["TABLE_ID"] = table_id;
        //Logger.log("Checking " + table_id);
        //Logger.log("33333333333333333333333333: " + unique_items_counter[table_id] + " " + table_id);
        //var ui_empty = Object.keys(unique_items_counter).length === 0 && unique_items_counter.constructor === Object;
        //Logger.log("4444444444444444444: " + ui_empty);
        //Logger.log("5555555555555555555"  + table_id in unique_items_counter);
        //if ( !(table_id in unique_items_counter) )
        var table_id_exists = unique_items_counter.hasOwnProperty(table_id);
        //for (var lala in unique_items_counter)
        //{
        //    Logger.log(lala + ": " + unique_items_counter[lala]);
        //}
        if ( !table_id_exists )
        {
          Logger.log("AAAAAAAAAAAA (table exists)");
          table_counters[group_name]++;
          unique_items_counter[table_id] = Number(col3_data[table["BEGIN"]-2]);
          table["ALREADY_EXISTS"] = false;
          table["COST"] = Number(col3_data[table["BEGIN"]-3]);
          if ( table["TEMPLATE"]["IS_TEMPLATE"] )
          {
            table["TEMPLATE"]["LANDMARKS"]["ALREADY_EXISTS"] = false;
            table["TEMPLATE"]["LANDMARKS"]["COST"] = Number(col3_data[table["BEGIN"]-3]);
          }
        }
        else
        {
          table["ALREADY_EXISTS"] = true;
          if ( table["TEMPLATE"]["IS_TEMPLATE"] )
          {
            table["TEMPLATE"]["LANDMARKS"]["ALREADY_EXISTS"] = true;
          }
          unique_items_counter[table_id] += Number(col3_data[table["BEGIN"]-2]);
        }
//        Logger.log("==============================================");
//        Logger.log("GROU_NAME: " + group_name);
//        Logger.log("TABLE_TITLE: " + String(col2_data[table["BEGIN"]+1]));
//        Logger.log("TABLE_ID: " + table_id);
//        Logger.log("TABLE COUNTERS:");
//        Logger.log(table_counters);
//        Logger.log("==============================================");        
        if (table["TEMPLATE"]["IS_TEMPLATE"])
        {
          table["TEMPLATE"]["LANDMARKS"]["TABLE_ID"] = table_id;
          //var template_landmarks = table_landmarks(
          insertPTTable(table["TEMPLATE"]["LANDMARKS"], table_counters, template_col2_data, template_col3_data, group_sheet, unique_items_counter[table_id]);
        }
        else
          insertPTTable(table, table_counters, col2_data, col3_data, group_sheet, unique_items_counter[table_id]);
        
        // jump after the end of the table
        j = table["END"];
        
      } // if (String(col2_data[j]) === "DOCS")

    } // for scanning col2_data
  } // for scanning all sheets
  
  var group_sheets = new_ss.getSheets();
  var total_sum = 0;
  for (k=0; k<group_sheets.length; k++)
  {
    var gn = group_sheets[k].getName();
    group_sheets[k].getRange(2,2).setValue(Group_Budget[gn]);
    total_sum += Number(Group_Budget[gn]);
  }
  Logger.log(Group_Budget);
  Logger.log("SSSSSSSSSSSSSSSSSSSSS: TOTAL_SUM: " + total_sum)
  moveSpreadsheetToOutputDir(new_ss);
  BOM_per_group(BOM);
} // function createPTTables

function sheetBudget()
{
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var active_sheet = active_spreadsheet.getActiveSheet();
  var sheet_data = active_sheet.getDataRange();
  var lastRow = sheet_data.getLastRow();
  
  var col2_data = active_sheet.getRange(1,2,lastRow,1).getValues();
  var col3_data = active_sheet.getRange(1,3,lastRow,1).getValues();
  var sum = 0;
  for (var i=1; i<=lastRow; i++)
  {
    //Logger.log(String(col2_data[i]));
    if (String(col2_data[i]).indexOf("ΚΟΣΤΟΣ") > -1)
    {
      var cost_per_item = Number(col3_data[i]);
      var num_of_items = Number(col3_data[i+1]);
      
      sum += cost_per_item * num_of_items;
    }
  }
  active_sheet.getRange(3,4).setValue(sum);
  //return sum;
}


// Create a new blank sheet
// Να μη γίνει button. Ο καθένας να κάνει copy από ένα υπάρχον sheet.
function newSheet()
{
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var master_sheet = active_spreadsheet.getSheetByName("MasterSheet");
  
  var new_sheet_name = Browser.inputBox("Δώστε το όνομα του νέου sheet: ΤΜΗΜΑ-ΚΑΕ (ένα φύλλο εργασίας ανά Τμήμα και ανά ΚΑΕ.");
  var new_sheet = master_sheet.copyTo(active_spreadsheet);
  new_sheet.setName(new_sheet_name);
  new_sheet.getRange(1, 1).setValue(new_sheet_name);
  new_sheet.getRange(2, 1).setValue(Utilities.formatDate(new Date(), "Europe/Athens", "yyyy-MM-dd-H:m"));
  active_spreadsheet.setActiveSheet(new_sheet);
  SpreadsheetApp.flush();
}

function set_auditor_conditional_format(active_sheet, ranges, F, T)
{
  var frule1 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextEqualTo(F)
  .setBackground("red")
  .setRanges(ranges)
  .build();
  var frule2 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextEqualTo(T)
  .setBackground("green")
  .setRanges(ranges)
  .build();
  var frules = active_sheet.getConditionalFormatRules();
  frules.push(frule1);
  frules.push(frule2);
  active_sheet.setConditionalFormatRules(frules);
}

// New table
function createTable()
{
  // Creates template for a conformity table, read to fill in specs and docs
  // Always fist line of each spec will be number of items.
  
  // Κάθε πίνακας (δηλ. υλικό) θα έχει ένα μοναδικό διακριτικό: ένα timestamp μέχρι εκατοστό του δευτερολέπτου
  
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var active_sheet = active_spreadsheet.getActiveSheet();
  var sheet_data = active_sheet.getDataRange();
  var lastRow = sheet_data.getLastRow();
  
  var timestamp = Utilities.formatDate(new Date(), "Europe/Athens", "yyyyMMddHHmmss");
  //Logger.log(timestamp);
  
  var doc_strings = ["DOCS", "Ενδεικτικό Μοντέλο", "Datasheet", "PDF Προσοφοράς", "Ιστοσελίδα Κοστολόγησης", "Σχόλιο", "ΚΟΣΤΟΣ ΑΝΑ ΤΕΜΑΧΙΟ", "ΤΕΜΑΧΙΑ", "GROUP"];
  
  var rowPointer = lastRow+4;
  for ( var i=0; i<doc_strings.length; i++ )
  {
    var cell = active_sheet.getRange(rowPointer,2);
    cell.setValue(doc_strings[i]);
    
    switch (doc_strings[i])
    {
      case "DOCS":
        active_sheet.getRange(rowPointer,2).setFontColor("red");
        active_sheet.getRange(rowPointer,2).setFontWeight("bold");
        break;
      case "Ενδεικτικό Μοντέλο": 
        active_sheet.getRange(rowPointer,3).setNote("Περιγραφή του μοντέλοου, π.χ. SHARP MX-M363N");
        active_sheet.getRange(rowPointer,2).setFontColor("#000000");
        break;
      case "Datasheet":
      case "PDF Προσοφοράς":
        active_sheet.getRange(rowPointer,3).setNote("Ανεβάστε το PDF στο drive και από εκεί εισάγετε link με την εντολή =hyperlink");
        active_sheet.getRange(rowPointer,2).setFontColor("#000000");
        break;
      case "Ιστοσελίδα Κοστολόγησης":
        active_sheet.getRange(rowPointer,3).setNote("Eισάγετε link με την εντολή =hyperlink");
        active_sheet.getRange(rowPointer,2).setFontColor("#000000");
        break
        case "Σχόλιο":
        active_sheet.getRange(rowPointer,2).setFontColor("#000000");
        break;
      case "ΚΟΣΤΟΣ ΑΝΑ ΤΕΜΑΧΙΟ":
        active_sheet.getRange(rowPointer,3).setNote("Κόστος ανά τεμάχιο (όχι όλα μαζί!).");
        active_sheet.getRange(rowPointer,2).setFontColor("red");
        active_sheet.getRange(rowPointer,2).setFontWeight("bold");
        active_sheet.getRange(rowPointer,3).setFontWeight("bold");
        active_sheet.getRange(rowPointer,3).setNumberFormat("€ 00.00");
        break;
      case "ΤΕΜΑΧΙΑ":
        active_sheet.getRange(rowPointer,3).setNote("Αριθμός Τεμαχίων");
        active_sheet.getRange(rowPointer,2).setFontColor("red");
        active_sheet.getRange(rowPointer,2).setFontWeight("bold");
        active_sheet.getRange(rowPointer,3).setFontWeight("bold");
        break;
      case "GROUP":
        active_sheet.getRange(rowPointer,3).setNote("Επιλέξτε ένα από τα διαθέσιμα groups");
        active_sheet.getRange(rowPointer,2).setFontColor("red");
        active_sheet.getRange(rowPointer,2).setFontWeight("bold");
        
        var vsheet = active_spreadsheet.getSheetByName("Groups");
        var vrange = vsheet.getRange("B2:B100");
        var rule = SpreadsheetApp.newDataValidation().requireValueInRange(vrange).build();
        active_sheet.getRange(rowPointer,3).setDataValidation(rule);
        
        break;
        
      default:
        SpreadsheetApp.getUi().alert("Χέσε μέσα Πολυχρόνη");
    }
    rowPointer++;
  }
  active_sheet.getRange(rowPointer,2).setValue("BEGIN");
  active_sheet.getRange(rowPointer,2).setFontColor("red");
  active_sheet.getRange(rowPointer,2).setFontWeight("bold");
  
  //active_sheet.getRange(rowPointer,4).setValue("Departmental Veto:");
  //active_sheet.getRange(rowPointer,4).setNote("Για τους Υπεύθυνους Τμημάτων/Υπηρεσιών το αποκάτω checkbox");
  active_sheet.getRange(rowPointer,4).setValue("");
  active_sheet.getRange(rowPointer,4).setNote("");
  active_sheet.getRange(rowPointer,4).setFontWeight("bold");
  active_sheet.getRange(rowPointer,5).setValue("Approval:");
  active_sheet.getRange(rowPointer,5).setNote("Για τους Auditors το αποκάτω checkbox");
  active_sheet.getRange(rowPointer,5).setFontWeight("bold");
  active_sheet.getRange(rowPointer,6).setValue("Approval:");
  active_sheet.getRange(rowPointer,6).setNote("Για τους Auditors το αποκάτω checkbox");
  active_sheet.getRange(rowPointer,6).setFontWeight("bold");
  active_sheet.getRange(rowPointer,7).setValue("Approval:");
  active_sheet.getRange(rowPointer,7).setNote("Για τους Auditors το αποκάτω checkbox");
  active_sheet.getRange(rowPointer,7).setFontWeight("bold");
  active_sheet.getRange(rowPointer,8).setValue("Coordinator Approval:");
  active_sheet.getRange(rowPointer,8).setNote("Για τον Coordinator το αποκάτω checkbox");
  active_sheet.getRange(rowPointer,8).setFontWeight("bold");
  
  rowPointer++;
  active_sheet.getRange(rowPointer,2).activateAsCurrentCell();
  active_sheet.getRange(rowPointer,2).setValue("ΤΙΤΛΟΣ ΠΙΝΑΚΑ");
  active_sheet.getRange(rowPointer,2).setNote("Αλλάξτε τον τίτλο του πίνακα");
  active_sheet.getRange(rowPointer,2).setFontColor("#000000");
  
  active_sheet.getRange(rowPointer,3).setValue(timestamp);
  active_sheet.getRange(rowPointer,3).setNote("Μην πειράξετε αυτό το κελί. Είναι το (σχεδόν) μοναδικό ID του Πίνακα.");
  
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox();
  //active_sheet.getRange(rowPointer,4).setDataValidation(rule);
  //active_sheet.getRange(rowPointer,4).setNote("Υπεύθυνος Τμήματος");
  active_sheet.getRange(rowPointer,4).setValue(false);
  
//  var protection = active_sheet.getRange(rowPointer,4).protect();
//  var me = Session.getEffectiveUser();
//  protection.addEditor(me);
//  protection.removeEditors(protection.getEditors());
//  if (protection.canDomainEdit()) 
//  {
//      protection.setDomainEdit(false);
//  }
  
  active_sheet.getRange(rowPointer,5).setDataValidation(rule);
  active_sheet.getRange(rowPointer,5).setNote("Μάνος Σταυρακάκης");
  active_sheet.getRange(rowPointer,6).setDataValidation(rule);
  active_sheet.getRange(rowPointer,6).setNote("Μανώλης Σαλδάρης");
  active_sheet.getRange(rowPointer,7).setDataValidation(rule);
  active_sheet.getRange(rowPointer,7).setNote("Νεκτάριος Παπαδάκης");
  active_sheet.getRange(rowPointer,8).setDataValidation(rule);
  active_sheet.getRange(rowPointer,8).setNote("Δημήτρης Καλοψικάκης");
  
  var ranges = [active_sheet.getRange(rowPointer,4)];
  set_auditor_conditional_format(active_sheet, ranges, "TRUE", "FALSE");
  
  var ranges = [active_sheet.getRange(rowPointer,5), active_sheet.getRange(rowPointer,6), 
                active_sheet.getRange(rowPointer,7), active_sheet.getRange(rowPointer,8)];
  set_auditor_conditional_format(active_sheet, ranges, "FALSE", "TRUE");
  
  active_sheet.getRange(rowPointer,2,1,7).setBackground("lightgray");
  active_sheet.getRange(rowPointer,2).setFontWeight("bold");
  active_sheet.getRange(rowPointer,3).setFontWeight("bold");
  
  rowPointer++;
  active_sheet.getRange(rowPointer,2).setValue("Προδιαγραφή");
  active_sheet.getRange(rowPointer,2).setFontColor("#000000");
  active_sheet.getRange(rowPointer,3).setValue("Απαίτηση");
  active_sheet.getRange(rowPointer,4).setValue("Σχόλιο Υπευθύνου");
  active_sheet.getRange(rowPointer,5).setValue("Σχόλιο Auditor 1");
  active_sheet.getRange(rowPointer,6).setValue("Σχόλιο Auditor 2");
  active_sheet.getRange(rowPointer,7).setValue("Σχόλιο Auditor 3");
  active_sheet.getRange(rowPointer,8).setValue("Σχόλιο Coordinator");
  active_sheet.getRange(rowPointer,2,1,7).setBackground("lightgray");
  active_sheet.getRange(rowPointer,2,1,7).setFontWeight("bold");  
  rowPointer++;
  
  var values = [];
  
  for ( var i=1; i<=50; i++)
  {
    values.push(["Προδιαγραφή " + i]);
    //active_sheet.getRange(rowPointer, 2).setValue("Προδιαγραφή " + i);
    //rowPointer++;
  }
  
  Logger.log("BEFORE: " + rowPointer);
  active_sheet.getRange(rowPointer,2,50,1).setValues(values);
  active_sheet.getRange(rowPointer,2,50,1).clearFormat();
  rowPointer += i-1;
  Logger.log("AFTER: " + rowPointer);
  
  active_sheet.getRange(rowPointer,2).setValue("END");
  active_sheet.getRange(rowPointer,2).setFontColor("red");
  active_sheet.getRange(rowPointer,2).setFontWeight("bold");
  rowPointer++;
  rows = rowPointer - (lastRow+4);
  var table_range = active_sheet.getRange(lastRow+4,2,rows,7);
  table_range.setBorder(true, true, true, true, true, true);
  table_range.setWrap(true);
  var table_title = "ΤΙΤΛΟΣ ΠΙΝΑΚΑ"
  table_title = table_title.replace(/[!"#$%&\'`()*+,-\.\/:;<=>?@\[\\\]^\{\|\}~]/g, "");
  var table_name_range = active_sheet.getName() + "_" + table_title + "_" + timestamp;
  table_name_range = table_name_range.replace(/[ -]/g, "_");
  active_spreadsheet.setNamedRange(table_name_range, table_range);
  //active_spreadsheet.setNamedRange("AAAA", table_range);
  //table_range.shiftRowGroupDepth(1);
  active_sheet.getRange(lastRow+4,3,rows,7).setHorizontalAlignment("center");
  
} // createTable

function insertTemplate()
{
  // Names of Templates must be unique.
  
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var active_sheet = active_spreadsheet.getActiveSheet();
  var sheet_data = active_sheet.getDataRange();
  var lastRow = sheet_data.getLastRow();
  
  var template_name = active_sheet.getRange(2,8).getValue();
  var template_id = "";
  var model = "";
  var datasheet = "";
  var pdf_offer = "";
  var cost_site = "";
  var comment = "";
  var cost_per_item = "";
  var num_of_items = "";
  var group = "";
  
  var template_specs_sheet = active_spreadsheet.getSheetByName("TemplateSpecs");
  var template_specs_last_row = template_specs_sheet.getDataRange().getLastRow();
  var col2_data = template_specs_sheet.getRange(1,2,template_specs_last_row,1).getValues();
  var col3_data = template_specs_sheet.getRange(1,3,template_specs_last_row,1).getValues();
  for (var i=0; i<col2_data.length; i++)
  {
    //Logger.log(col2_data[i]);
    if ( String(col2_data[i]) === template_name )
    {
      template_id = String(col3_data[i]);
      //Logger.log(template_id);
      group = col3_data[i-2];
      num_of_items = col3_data[i-3];
      cost_per_item = col3_data[i-4];
      //Logger.log(cost_per_item);
      comment = col3_data[i-5];
      cost_site = col3_data[i-6];
      pdf_offer = col3_data[i-7];
      datasheet = col3_data[i-8];
      model = col3_data[i-9];
      break;
    }
  }
  
  var doc_strings = ["DOCS", "Ενδεικτικό Μοντέλο", "Datasheet", "PDF Προσοφοράς", "Ιστοσελίδα Κοστολόγησης", "Σχόλιο", "ΚΟΣΤΟΣ ΑΝΑ ΤΕΜΑΧΙΟ", "ΤΕΜΑΧΙΑ", "GROUP"];
  
  rowPointer = lastRow+4;
  for ( var i=0; i<doc_strings.length; i++ )
  {
    var cell = active_sheet.getRange(rowPointer,2);
    cell.setValue(doc_strings[i]);
    
    switch (doc_strings[i])
    {
      case "DOCS":
        active_sheet.getRange(rowPointer,2).setFontColor("red");
        active_sheet.getRange(rowPointer,2).setFontWeight("bold");
        break;
      case "Ενδεικτικό Μοντέλο": 
        active_sheet.getRange(rowPointer,3).setNote("Περιγραφή του μοντέλοου, π.χ. SHARP MX-M363N");
        active_sheet.getRange(rowPointer,3).setValue(model);
        break;
      case "Datasheet":
        active_sheet.getRange(rowPointer,3).setValue(datasheet);
        break;
      case "PDF Προσοφοράς":
        active_sheet.getRange(rowPointer,3).setNote("Ανεβάστε το PDF στο drive και από εκεί εισάγετε link με την εντολή =hyperlink");
        active_sheet.getRange(rowPointer,3).setValue(pdf_offer);
        break;
      case "Ιστοσελίδα Κοστολόγησης":
        active_sheet.getRange(rowPointer,3).setNote("Eισάγετε link με την εντολή =hyperlink");
        active_sheet.getRange(rowPointer,3).setValue(cost_site);
        break
        case "Σχόλιο":
        active_sheet.getRange(rowPointer,3).setValue("TEMPLATE: " + comment);
        break;
      case "ΚΟΣΤΟΣ ΑΝΑ ΤΕΜΑΧΙΟ":
        active_sheet.getRange(rowPointer,3).setValue(cost_per_item);
        active_sheet.getRange(rowPointer,3).setNote("Κόστος ανά τεμάχιο (όχι όλα μαζί!).");
        active_sheet.getRange(rowPointer,2).setFontColor("red");
        active_sheet.getRange(rowPointer,2).setFontWeight("bold");
        active_sheet.getRange(rowPointer,3).setFontWeight("bold");
        active_sheet.getRange(rowPointer,3).setNumberFormat("€ 00.00");
        break;
      case "ΤΕΜΑΧΙΑ":
        active_sheet.getRange(rowPointer,3).setValue(num_of_items);
        active_sheet.getRange(rowPointer,3).setNote("Αριθμός Τεμαχίων");
        active_sheet.getRange(rowPointer,2).setFontColor("red");
        active_sheet.getRange(rowPointer,2).setFontWeight("bold");
        active_sheet.getRange(rowPointer,3).setFontWeight("bold");
        break;
      case "GROUP":
        active_sheet.getRange(rowPointer,3).setValue(group);
        active_sheet.getRange(rowPointer,3).setNote("Επιλέξτε ένα από τα διαθέσιμα groups");
        active_sheet.getRange(rowPointer,2).setFontColor("red");
        active_sheet.getRange(rowPointer,2).setFontWeight("bold");
        
        var vsheet = active_spreadsheet.getSheetByName("Groups");
        var vrange = vsheet.getRange("B2:B100");
        var rule = SpreadsheetApp.newDataValidation().requireValueInRange(vrange).build();
        active_sheet.getRange(rowPointer,3).setDataValidation(rule);
        
        break;
        
      default:
        SpreadsheetApp.getUi().alert("Χέσε μέσα Πολυχρόνη");
    }
    rowPointer++;
  }
  active_sheet.getRange(rowPointer,2).setValue("BEGIN");
  active_sheet.getRange(rowPointer,2).setFontColor("red");
  active_sheet.getRange(rowPointer,2).setFontWeight("bold");
  
  //active_sheet.getRange(rowPointer,4).setValue("Departmental Veto:");
  //active_sheet.getRange(rowPointer,4).setNote("Για τους Υπεύθυνους Τμημάτων/Υπηρεσιών το αποκάτω checkbox");
  //active_sheet.getRange(rowPointer,4).setFontWeight("bold");
  active_sheet.getRange(rowPointer,5).setValue("Approval:");
  active_sheet.getRange(rowPointer,5).setNote("Για τους Auditors το αποκάτω checkbox");
  active_sheet.getRange(rowPointer,5).setFontWeight("bold");
  active_sheet.getRange(rowPointer,6).setValue("Approval:");
  active_sheet.getRange(rowPointer,6).setNote("Για τους Auditors το αποκάτω checkbox");
  active_sheet.getRange(rowPointer,6).setFontWeight("bold");
  active_sheet.getRange(rowPointer,7).setValue("Approval:");
  active_sheet.getRange(rowPointer,7).setNote("Για τους Auditors το αποκάτω checkbox");
  active_sheet.getRange(rowPointer,7).setFontWeight("bold");
  active_sheet.getRange(rowPointer,8).setValue("Coordinator Approval:");
  active_sheet.getRange(rowPointer,8).setNote("Για τον Coordinator το αποκάτω checkbox");
  active_sheet.getRange(rowPointer,8).setFontWeight("bold");
  
  rowPointer++;
  active_sheet.getRange(rowPointer,2).activateAsCurrentCell();
  active_sheet.getRange(rowPointer,2).setValue(template_name);
  //active_sheet.getRange(rowPointer,2).setNote("Αλλάξτε τον τίτλο του πίνακα");
  
  active_sheet.getRange(rowPointer,3).setValue(template_id);
  active_sheet.getRange(rowPointer,3).setNote("Μην πειράξετε αυτό το κελί. Είναι το (σχεδόν) μοναδικό ID του Πίνακα.");
  
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox();
  //active_sheet.getRange(rowPointer,4).setDataValidation(rule);
  //active_sheet.getRange(rowPointer,4).setNote("Υπεύθυνος Τμήματος");
  active_sheet.getRange(rowPointer,4).setValue(false);
  
  active_sheet.getRange(rowPointer,5).setDataValidation(rule);
  active_sheet.getRange(rowPointer,5).setNote("Μάνος Σταυρακάκης");
  active_sheet.getRange(rowPointer,6).setDataValidation(rule);
  active_sheet.getRange(rowPointer,6).setNote("Μανώλης Σαλδάρης");
  active_sheet.getRange(rowPointer,7).setDataValidation(rule);
  active_sheet.getRange(rowPointer,7).setNote("Νεκτάριος Παπαδάκης");
  active_sheet.getRange(rowPointer,8).setDataValidation(rule);
  active_sheet.getRange(rowPointer,8).setNote("Δημήτρης Καλοψικάκης");
  
  var ranges = [active_sheet.getRange(rowPointer,4)];
  set_auditor_conditional_format(active_sheet, ranges, "TRUE", "FALSE");
  
  var ranges = [active_sheet.getRange(rowPointer,5), active_sheet.getRange(rowPointer,6), 
                active_sheet.getRange(rowPointer,7), active_sheet.getRange(rowPointer,8)];
  set_auditor_conditional_format(active_sheet, ranges, "FALSE", "TRUE");
  
  active_sheet.getRange(rowPointer,2,1,7).setBackground("lightgray");
  active_sheet.getRange(rowPointer,2).setFontWeight("bold");
  active_sheet.getRange(rowPointer,3).setFontWeight("bold");
  
  rowPointer++;
  //active_sheet.getRange(rowPointer,2).setValue("Προδιαγραφή");
  //active_sheet.getRange(rowPointer,3).setValue("Απαίτηση");
  //active_sheet.getRange(rowPointer,4).setValue("Σχόλιο Υπευθύνου");
  //active_sheet.getRange(rowPointer,5).setValue("Σχόλιο Auditor 1");
  //active_sheet.getRange(rowPointer,6).setValue("Σχόλιο Auditor 2");
  //active_sheet.getRange(rowPointer,7).setValue("Σχόλιο Auditor 3");
  //active_sheet.getRange(rowPointer,8).setValue("Σχόλιο Coordinator");
  //active_sheet.getRange(rowPointer,2,1,7).setBackground("lightgray");
  //active_sheet.getRange(rowPointer,2,1,7).setFontWeight("bold");  
  //rowPointer++;
  //for ( var i=1; i<=50; i++)
  //{
  //  active_sheet.getRange(rowPointer, 2).setValue("Προδιαγραφή " + i);
  //  rowPointer++;
  //}
  active_sheet.getRange(rowPointer,2).setValue("END");
  active_sheet.getRange(rowPointer,2).setFontColor("red");
  active_sheet.getRange(rowPointer,2).setFontWeight("bold");
  rowPointer++;
  rows = rowPointer - (lastRow+4);
  var table_range = active_sheet.getRange(lastRow+4,2,rows,7);
  table_range.setBorder(true, true, true, true, true, true);
  table_range.setWrap(true);
  //var table_title = "ΤΙΤΛΟΣ ΠΙΝΑΚΑ"
  var table_title = template_name;
  table_title = table_title.replace(/[!"#$%&\'`()*+,-\.\/:;<=>?@\[\\\]^\{\|\}~]/g, "");
  var table_name_range = active_sheet.getName() + "_" + table_title + "_" + template_id;
  table_name_range = table_name_range.replace(/[ -]/g, "_");
  active_spreadsheet.setNamedRange(table_name_range, table_range);

  //active_spreadsheet.setNamedRange("AAAA", table_range);
  //table_range.shiftRowGroupDepth(1);
  active_sheet.getRange(lastRow+4,3,rows,7).setHorizontalAlignment("center");
} //insertTemplate

function makeTemplateList()
{
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var template_specs_sheet = active_spreadsheet.getSheetByName("TemplateSpecs");
  var template_list_sheet = active_spreadsheet.getSheetByName("TemplateList");
  var sheet_data = template_specs_sheet.getDataRange();
  var lastRow = sheet_data.getLastRow();
  var col2_data = template_specs_sheet.getRange(1,2,lastRow,1).getValues();
  var col3_data = template_specs_sheet.getRange(1,3,lastRow,1).getValues();
  
  template_list_sheet.getRange(1, 1).setValue("TemplateName");
  template_list_sheet.getRange(1, 1).setFontWeight("bold");
  template_list_sheet.getRange(1, 2).setValue("TemplateID");
  template_list_sheet.getRange(1, 2).setFontWeight("bold");
  var list_row = 2;
  for (var i=1; i<=lastRow; i++)
  {
    if ( String(col2_data[i]) === "BEGIN" )
    {
      var template_name = String(col2_data[i+1]);
      var template_id   = String(col3_data[i+1]);
      template_list_sheet.getRange(list_row, 1).setValue(template_name);
      template_list_sheet.getRange(list_row, 2).setValue(template_id);
      list_row++;
    }
  }
} // makeTemplateList


function addBUdget(cell)
{
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sum = 0;
  
  for (var i=0; i < sheets.length; i++ )
  {
    var sheet = sheets[i];
    var value = sheet.getRange(cell).getValue();
    
    if (typeof(value) == 'number')
    {
      sum += value
    }
    
  }
  
  return sum;
}


function copyCurrentActiveSpreadsheet()
{
  var active_ss = SpreadsheetApp.getActiveSpreadsheet();
  var active_ss_sheets = active_ss.getSheets(); 
  var active_ss_name  =  active_ss.getName();
  
  var new_ss = SpreadsheetApp.create(active_ss_name+Utilities.formatDate(new Date(), "UTC+8", "yyyy-MM-dd-H:m"));
  
  for ( var i=0; i<active_ss_sheets.length; i++ )
  {
    var c_sheet = active_ss_sheets[i];
    var c_data = c_sheet.getDataRange();
    var lastRow = c_data.getLastRow();
    var lastColumn = c_data.getLastColumn();
    var searchRange = c_sheet.getRange(1,1,lastRow, lastColumn);
    
    // create a sheet
    name = c_sheet.getName();
    new_ss_sheet = new_ss.insertSheet(name);
    //new_ss.getRange("A1").setValue("AAA");
    
    for ( var j=1; j<lastRow; j++)
    {
      for (var k=1; k<lastColumn; k++)
      {
        var cell = c_sheet.getRange(j,k);
        //var cell_value = cell.getValue();
        //var cell_format = cell.getNumberFormat();
        var cell_formula = cell.getFormula();
        var new_cell = new_ss_sheet.getRange(j,k);
        //new_cell.setNumberFormat(cell_format);
        new_cell.setValue(Utilities.formatDate(new Date(), "UTC+8", "yyyy-MM-dd-H:m"));
        //new_cell.setValue(cell_value);
        //if (cell_formula)
        //{
        //  new_cell.setValue(cell_formula);
        //}
        //else
        //{
        //  new_cell.setValue(cell_value);
        //}
      }
    }
  }
}

//
// Later:
//
function is_inside_table(cell)
{
    // returns Object. 
    // First value is TRUE/FALSE weather cell is inside table or not
    // Second value is the range of the talbe
    
    var retv = new Object();
    
    var active_ss = SpreadsheetApp.getActiveSpreadsheet();
    var active_sheet = active_ss.getActiveSheet();
    var last_row = active_sheet.getDataRange().getLastRow();
    var col2_data = active_sheet.getRange(1, 2, last_row, 1).getValues();
    //var col = cell.getColumn(); // I don't care about this. I know that landmarks are on col 2
    var row = cell.getRow();
    Logger.log("Row: " + row);
    var i = 0;
    var k = row-1;
    while ( String(col2_data[k]) != "DOCS" && String(col2_data[k]) != "END" && k>=0)
    {
        Logger.log(col2_data[k]);
        k--;
        
        i++;
        if (i > 10000)
        {
            break;
        }
    }
    if (k>=0 && String(col2_data[k]) === "DOCS")
    {
        var j = row-1;
        while (String(col2_data[j]) != "END" && j<=last_row)
        {
            j++;
        }
        if ( j > last_row)
        {
            retv["IN"] = false;
            retv["LANDMARKS"] = null;
        }
        else
        {
            retv["IN"] = true;
            retv["LANDMARKS"] = new Object();
            retv["LANDMARKS"]["DOCS"] = k+1;
            retv["LANDMARKS"]["END"] = j+1;
        }
    }
    else // if k==-1 or col2_data[k] === "END"
    {
        retv["IN"] = false;
        retv["LANDMARKS"] = null;
    }
    
    return retv;
}

function duplicateTable()
{
    // Create a new table that is copy of the table of the currently active cell.
    var active_ss = SpreadsheetApp.getActiveSpreadsheet();
    var active_sheet = active_ss.getActiveSheet();
    var active_cell = active_sheet.getActiveCell();
    
    var F = is_inside_table(active_cell);
    Logger.log(F);
    return F["IN"];
    
    // Check whether this cell is in a table or not.
    // If yes, find its "borders" ==> calulate its dimensions ("area")
    // insert lines below the current table
    // copy cells of current table
    // update ID (with a new timestamp).
}

function removeTable()
{
    // Check if this is called from within a table.
    // Calulate table's topography
    // "move" table to Trash Sheet; i.e. copy table to Trash Sheet and delete lines.
    
    var active_ss = SpreadsheetApp.getActiveSpreadsheet();
    var active_sheet = active_ss.getActiveSheet();
    var active_cell = active_sheet.getActiveCell();
    
    Logger.log("Sheet ID: " + active_sheet.getSheetId());
    
    var trash_sheet = active_ss.getSheetByName("Trash"); 
    var trash_last_row = trash_sheet.getDataRange().getLastRow();
    
    var F = is_inside_table(active_cell);
    
    if ( F["IN"] )
    {
        var docs_row = F["LANDMARKS"]["DOCS"];
        var num_of_rows = F["LANDMARKS"]["END"] - F["LANDMARKS"]["DOCS"]+1;
        var num_of_cols = 7;
        // get values from range
        // jst values is not sufficient.
        // I need copy of rows with everythin (formatting, etc)
        active_sheet.getRange(docs_row,2,num_of_rows, num_of_cols).copyTo(trash_sheet.getRange(trash_last_row+4, 2, num_of_rows, num_of_cols));
        //active_sheet.deleteRows(docs_row, num_of_rows);
    }
    else
    {
        return "Not in a table";
    }
}

function moveUPTable(N)
{
    // move current table up N positions, until top
    Logger.log("moveUpTable: " + N);
}

function moveDownTable(N)
{
    // move current table down N positions, until bottom
    Logger.log("moveDownTable: " + N);
}

function nameTableRanges()
{
    var active_ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = active_ss.getSheets();
    
    var namedRanges = active_ss.getNamedRanges();
    for (k=0; k<namedRanges.length; k++)
    {
        namedRanges[k].remove();
    }
    
    for (var i=0; i<sheets.length; i++)
    {
        var sheet_name = sheets[i].getName();
        
        if ( is_not_dept_specs_sheet(sheets[i]) )
          continue;
          
//        if ( ["Depts", "Items", "Groups", "MasterSheet", "ΚΑΕ", "ΚΑΕ-Είδη", "TemplateList", "TemplateSpecs", "SandBox", "ToDo", "Admin", "BudgetCheck", "Trash"].indexOf(sheet_name) != -1 )
//        {
//            continue;
//        }
    
        //var active_sheet = active_ss.getActiveSheet();
        //var active_cell = active_sheet.getActiveCell();
        
        var last_row = sheets[i].getDataRange().getLastRow();
        var col2_data = sheets[i].getRange(1, 2, last_row, 1).getValues();
        var col3_data = sheets[i].getRange(1, 3, last_row, 1).getValues();
        
        for (var k=0; k<col2_data.length; k++)
        {
            if (String(col2_data[k]) === "DOCS")
            {
              Logger.log(";a;a");
              var tl = table_landmarks(k, col2_data);
              Logger.log(tl);
              var numRows = tl["END"] - tl["DOCS"] + 1;
              //var range_name = String(col2_data[k+10]) + "_" + String(col3_data[k+10]);
              var table_title = String(col2_data[k+10]).replace(/[ -]/g, "_");
              table_title = table_title.replace(/[!"#$%&\'`()*+,-\.\/:;<=>?@\[\\\]^\{\|\}~]/g, "");
              Logger.log("Table Title: " + table_title);
              var table_id = String(col3_data[k+10]);
              var range_name = sheet_name + "_" + table_title + "_" + table_id;
              range_name = range_name.replace(/[ -]/g, "_");
              Logger.log(range_name);
              var range = sheets[i].getRange(k+1, 2, numRows, 7);
              active_ss.setNamedRange(range_name, range);
              //var table_landmarks = 
            }
        }
    }
    
//    var F = is_inside_table(active_cell);
//    
//    if ( F["IN"] )
//    {
//        var docs_row = F["LANDMARKS"]["DOCS"];
//        var num_of_rows = F["LANDMARKS"]["END"] - F["LANDMARKS"]["DOCS"]+1;
//        var num_of_cols = 7;
//        // get values from range
//        // jst values is not sufficient.
//        // I need copy of rows with everythin (formatting, etc)
//        //active_sheet.getRange(docs_row,2,num_of_rows, num_of_cols).copyTo(trash_sheet.getRange(trash_last_row+4, 2, num_of_rows, num_of_cols));
//        var range = active_sheet.getRange(docs_row,2,num_of_rows, num_of_cols);
//        active_ss.setNamedRange("LALA", range);
//        //active_sheet.deleteRows(docs_row, num_of_rows);
//    }
//    else
//    {
//        Logger.log("Nothing");
//        //return "Not in a table";
//    }

}

////////////////// TESTS //////////////////////////////////
// All test functions have the prefix "test_"

// Templates must be discriminated

// Bad case is two tables with different specs but with the same id

// Checks all table ids are unique
// It must return a two value vector. The first value must be
// true/false, the second value must some useful info (an Object)
// with any useful information. In this case it must return 
// an array of the duplicated ids.
function test_table_ids()
{
    var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = active_spreadsheet.getSheets();
    
    for (i=0; i<sheets.length; i++)
    {
        var sheet_name = sheets[i].getName();
        
        if ( is_not_dept_specs_sheet(sheets[i]) )
          continue;
      
//        if ( ["Depts", "Items", "Groups", "MasterSheet", "ΚΑΕ", "ΚΑΕ-Είδη", "TemplateList", "TemplateSpecs", "SandBox", "ToDo", "Admin", "BudgetCheck", "Trash"].indexOf(sheet_name) != -1 )
//        { 
//            continue;
//        }
        
    }
}

function test_table_in_multiple_groups()
{
    // checks uif there are tables (table_ids) that have been
    // enlisted in multiple groups.
    Logger.log("test_table_in_multiple_groups");
    
    var active_ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = active_ss.getSheets();
    
    // In which group does each table belong? 
    var id_groups = new Object();
    var multiple_group_ids = [];
    
    for (var i=0; i<sheets.length; i++)
    {
        if ( is_not_dept_specs_sheet(sheets[i]) )
            continue;
        
        var last_row = sheets[i].getDataRange().getLastRow();
        var col2_data = sheets[i].getRange(1, 2, last_row, 1).getValues();
        var col3_data = sheets[i].getRange(1, 3, last_row, 1).getValues();
        for (var j=0; j< col2_data.length; j++)
        {
            if (String(col2_data[j]) === "GROUP" )
            {
                var group_descr = String(col3_data[j]);
                var table_id = String(col3_data[j+2]);

                if (table_id in id_groups)
                {
                    if (id_groups[table_id].indexOf(group_descr) == -1 )
                        id_groups[table_id].push(group_descr);
                    
                    if ( id_groups[table_id].length > 1 )
                        multiple_group_ids.push(table_id);
                }
                else
                {
                    id_groups[table_id] = [];
                    id_groups[table_id].push(group_descr);
                }
            }
        } // loop through col2
    } // loop over sheets
    for (var key in id_groups)
    {
        Logger.log(key + ": " + id_groups[key] + " : " + id_groups[key].length);
    }
    Browser.msgBox("Tables that belong to more than one Group: " + multiple_group_ids);
    return multiple_group_ids;
}

function test_email()
{
  MailApp.sendEmail("kalopsik@uoc.gr", "TEST SUBJECT", "THIS IS THE BODY");
}