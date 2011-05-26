/*
*     Copyright (C) <2011> by <EDINA, University of Edinburgh>
*
*     Permission is hereby granted, free of charge, to any person obtaining a copy
*     of this software and associated documentation files (the "Software"), to deal
*     in the Software without restriction, including without limitation the rights
*     to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*     copies of the Software, and to permit persons to whom the Software is
*     furnished to do so, subject to the following conditions:
*
*     The above copyright notice and this permission notice shall be included in
*     all copies or substantial portions of the Software.
* 
*     THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*     IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*     FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*     AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*     LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*     OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
* 	  THE SOFTWARE.
*
*    Simple Excel export script.
*    HSSFWorkbookBuilder class is taken from http://skepticalhumorist.blogspot.com/2010/12/groovy-dslbuilders-poi-spreadsheets.html 
*    
*    This script is designed to read a list of queries from queries.groovy, execute each in the specified database and output the
*    results of each in a separate sheet in an Excel spreadsheet.
*    
*    ***********************************************************************************************************************************
*    BEFORE RUNNING:
*    ***********************************************************************************************************************************
*    
*    Before running this script, edit queries.groovy and add any further queries you wish to include in the spreadsheet in the format:
*
*    <query name>="<query SQL>";
*
*    where <query name> will become the name of an Excel sheet which will contain the results of the <query SQL> query.  
*
*    e.g:    Item_licences="SELECT * FROM v_jopen_cc_licence";
*     
*    Include the database driver the in the class path
*    e.g put postgresql-8.1-408.jdbc3.jar into $GROOVY_HOME/lib/
*
*    Complete the properties below.
*
*    ***********************************************************************************************************************************
*
*    To generate the stats spreadsheet, run:
*
*    groovy export_stats.groovy
*
*    This will create a symlink to the latest spreadsheet
*
*
*    @author colingormley
*/

// *********** Please complete the following properties ******************
def dbUrl = "jdbc:postgresql://localhost:5432/databasename";
def dbUser = "dbuser";
def dbPass = "dbpass";
def dbDriver ="org.postgresql.Driver";
def outputDir =  "/path/to/dir";
def queriesFile = "/path/to/queries.groovy";
// Note - don't include .xls suffix - code will append this
def spreadsheetName = "StatsSpreadsheet";
//************************************************************************


import groovy.sql.Sql

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

class HSSFWorkbookBuilder {
 
  private Workbook workbook = new HSSFWorkbook()
  private Sheet sheet
  private int rows
 
  Workbook workbook(Closure closure) {
    closure.delegate = this
    closure.call()
    workbook
  }
 
  void sheet(String name, Closure closure) {
    sheet = workbook.createSheet(name)
    rows = 0
    closure.delegate = this
    closure.call()
  }
 
  void row(values) {
    Row row = sheet.createRow(rows++ as int)
    values.eachWithIndex {value, col ->
      Cell cell = row.createCell(col)
      switch (value) {
          default: cell.setCellValue(new HSSFRichTextString("" + value)); break
      }
    }
  }
 
}





// Connect to the database
def postgres = Sql.newInstance(dbUrl, dbUser, dbPass, dbDriver)

// Read the queries
def queries = new ConfigSlurper().parse(new File(queriesFile).toURL())
println "About to run the following queries:"
queries.each{println it}

// Run query and 
def workbook = new HSSFWorkbookBuilder().workbook {
    queries.each{k,v->
        sheet(k) {
        postgres.eachRow(
          v,
          {meta -> row(meta*.columnName)},           // header row with columns names from ResultSetMetaData
          {rs -> row(rs.toRowResult().values())}     // data row for each ResultSet row
        )
    }
  }
}

postgres.close();


def dir = new File(outputDir)
if(!dir.exists()){
     dir.mkdir();
}

// Output data
def fileName = "${spreadsheetName}_${new Date().format('ddMMyy')}.xls"
new File(dir, fileName).withOutputStream{ os ->
    workbook.write(os);
}

spreadsheetName="${spreadsheetName}.xls";

// Remove existing spreadsheetName
new File(dir, spreadsheetName).delete()

// Create new spreadsheetName
"ln -s ${dir.absolutePath}/${fileName} ${dir.absolutePath}/${spreadsheetName}".execute()

println "Finished.  Please see ${dir.canonicalPath}/${spreadsheetName} for results."