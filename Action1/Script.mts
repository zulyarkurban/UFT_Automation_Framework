On Error Resume Next

' Browser can be set as chrome, edge
BROWSER_TYPE = "ie"
'DATA_ROW = 3 means it will read the data from third row from excel.
DATA_ROW = 3 

'Test case

Call ReadExcel()
Call OpenBrowser()
Call SignIn()
Call Registration()
Call SelectTShirtAndCheckout()
Call Payment()
Call SignOut()
Call OpenResultFolder()
