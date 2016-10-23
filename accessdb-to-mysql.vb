' Connection String required to connect to MS Access database
' PLEASE CHANGE data source= to the path where your ACCESS DATABASE file is on your C:\ drive
connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Smartlaunch\Server\Data\DB\Smartlaunch.mdb;"

'Create todays date for use in the Access query
dim dt
dt = Date()

sql = "SELECT SUM(Transactions) AS Sales FROM FinancialTransactions WHERE Date > #"&dt&"# "

' Create ADO Connection/Command objects
SET cn = CREATEOBJECT("ADODB.Connection")
SET cmd = CREATEOBJECT("ADODB.Command")

' Open connection
cn.open connectionString
' Associate connection object with command object
cmd.ActiveConnection = cn
' Set the SQL statement of the command object
cmd.CommandText = sql

' Execute query and save the results into the rs array for use later in the script i.e rs("Sales")
SET rs = cmd.EXECUTE

' Connection String required to connect to MySQL database
set objConn = CreateObject("ADODB.Connection")
objConn.Open "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=YOUR-MYSQL-SERVER-IP-GOES-HERE; DATABASE=YOUR-MYSQL-DATABASE-GOES-HERE; " &_
"UID=YOUR-MYSQL-USERNAME-GOES-HERE;PASSWORD=YOUR-MYSQL-PASSWORD-GOES-HERE; OPTION=3"
set objRS = CreateObject("ADODB.Recordset")

'what is dim store = I have a number of different stores i need to collect the finances from each
' night so I have them stored according to the shops location in the mysql database.
' change accoring to your own needs.

dim store
store = "STORE-IN-DUBLIN"

'we saved the results from the ACCESS query into Sales and now we are going to assign it to the variable DailyTakings and insert it and the store into our new mysql database
Dim DailyTakings
DailyTakings = rs("Sales")

commandText = "INSERT INTO dailyreporting(store,gross)VALUES('"+store+"','"+DailyTakings+"')"

objConn.Execute commandText
set objRS = Nothing
objConn.Close
set objConn = Nothing