Stats to Excel Export Script
============================

About
=====

Using two simple Groovy scripts, it's easy to export the results of any SQL query into a new sheet in an Excel spreadsheet.

Prerequisites
=============

* Mac/Unix/Solaris.  (Not tested, but may work on Windows with "Cygwin":http://www.cygwin.com/)
* Groovy 1.77 (Should work with "Groovy 1.8":http://groovy.codehaus.org/Download, but not tested).

Pre-run
=======

Before running this script, please ensure you do the following:
* Add your your JDBC jar file to your classpath.  Easiest way is to copy it into into $GROOVY_HOME/lib/ 
* Edit queries.groovy and add any further queries you wish to include in the spreadsheet in the format:

<query name>="<query SQL>";
where <query name> will become the name of an Excel sheet which will contain the results of the <query SQL> query, e.g:

Users="SELECT * FROM eperson";

* The following properties in export_stats.groovy should be updated:

def dbUrl = "jdbc:postgresql://localhost:5432/databasename";
def dbUser = "dbuser";
def dbPass = "dbpass";
def dbDriver ="org.postgresql.Driver";
def outputDir =  "/path/to/dir";
def queriesFile = "/path/to/queries.groovy";
// Note - don't include .xls suffix - code will append this
def spreadsheetName = "StatsSpreadsheet";

Running
=======

In order to run the script, open a terminal and run:

groovy export_stats.groovy

This will parse queries.groovy, query the database and put the results of each query in a new sheet in an Excel spreadsheet.

Note that the if you rerun the script on the same day, the spreadsheet will be overwritten, but if if you run again on a different date, the symlink to the latest copy of the spreadsheet will be updated.
