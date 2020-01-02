'*****************************************************************************************************************
'@Author: Sumithra HP
'@Creation Date: 25th June 2018
'@Test Script Name: Add_Component_CPI_cancel_Reverse

'*****************************************************************************************************************

'########################### Mandatory #########################
'Identifying and Loading the Required OR & Fun Lib at Run Time
rootFolderPath = Split(Environment.Value("TestDir"),"TestScripts\")(0)
call LoadFunctionLibrary(rootFolderPath & "FunctionLibrary\CommonFunctions.qfl")

'Loading the required Function Library
Call loadAllFunctionLibrary()

'Loading the required Object Repository
call loadMainObjectRepository()

'########################### Mandatory #########################

'Get Test Name
ReportName = Environment.Value("TestName")&".html"
'Environment.Value("ReportName") = ReportName
Environment.Value("ReportStep") = 0
Environment.Value("TestStatus") = True

'*****************************************************************************************************************
'Login into branch+ application and getting the State Name

'Fetching Current State name from DB
currentState = getCurrentStateName_From_DB()
Environment.Value("currentState") = currentState
Environment.Value("StateName") = currentState

'Custom Report FIle Creation
startTime = initializeCustomReport(currentState)
Environment.Value("startTime") = startTime

Call BranchPlusLogin()

'*****************************************************************************************************************
'Generating Test Data Path
testDataPath = Environment.Value("BranchPlusSharedPath") & currentState & "\TestData\CPI_TestData.xlsx"

'Importing Test Data from Excel Shet(specific sheet)
Call importSpecificSheet(testDataPath,"CPI")

'*****************************************************************************************************************
' Reading Data which is required for Test Script
'*****************************************************************************************************************
ChargeAmt = DataTable.Value("ChargeAmount")
PayMethodAfterCPIAdded = DataTable.Value("PayMethodForCPIAdded")
PayMethodAfterCPICancelled = DataTable.Value("PayMethodForCPICancelled")
PayMethodAfterCPIReversed = DataTable.Value("PayMethodForCPIReversed")

'*****************************************************************************************************************
'Open F2 Screeen Account Activity
Call openF2Screen()

'open accounts
AcctNum = get_AcctNum_From_DB_To_Add_CPI()

'search customer by entering loan number
Call customer_Search_Using_LoanNumber_F2(AcctNum)

'**********************************************************************************************************************
'Scenario #1 - Add Component (CPI Insurance) for Active account
'**********************************************************************************************************************
Call writeReportLog("<b>Scenario 1: Add Component (CPI Insurance) for Active account </b>" ,"Passed",FailComments,"")
'**********************************************************************************************************************

'select account to add CPI
Call select_Account_ToAdd_CPI_ComponentTab()
 
startDate = Date 
startDate = Right("0" & Month(startDate), 2) & "/" & Right("0" & Day(startDate), 2) & "/" & year(startDate)

'get maturity date
MaturityDate = get_maturityDate_AccountTab()
 
endDate=trim(DateAdd("d",45,startDate)) 
'get end date
endDate = get_CPI_endDate(MaturityDate,startDate,endDate)
endDate = Right("0" & Month(endDate), 2) & "/" & Right("0" & Day(endDate), 2) & "/" & year(endDate)
 
'navigate to component tab
Call navigateToTab("Component")

'Click on New button
Call clickNewButton()

'select New CPI insurance
Call selectValue_From_ChooseAnItem_Popup_ComponentTab("New CPI Insurance")

'add CPI
Call add_CPI_Insurance_Component_Tab(startDate,endDate,ChargeAmt)

'*******************************************************************************************************
'Scenario #2 - Check Payment Tab after CPI added
'*******************************************************************************************************
Call writeReportLog("<b>Scenario 2: Check Payment Tab after CPI added </b>" ,"Passed",FailComments,"")
'*******************************************************************************************************
'check CPI added in payment tab
call checkPaymentTab_CPI(startDate,PayMethodAfterCPIAdded,ChargeAmt)

'******************************************************************************************************
'Scenario #3 - Cancel CPI
'*******************************************************************************************************
Call writeReportLog("<b>Scenario 3: Cancel CPI </b>" ,"Passed",FailComments,"")
'*******************************************************************************************************
'cancel CPI
Call cancel_CPI_ComponentTab()

'*********************************************************************************************************
'Scenario #4 - Check Payment Tab after CPI cancelled
'*********************************************************************************************************
Call writeReportLog("<b>Scenario 4: Check Payment Tab after CPI cancelled </b>" ,"Passed",FailComments,"")
'*********************************************************************************************************
'check CPI cancelled in payment tab
call checkPaymentTab_CPI(startDate,PayMethodAfterCPICancelled,ChargeAmt)

'*********************************************************************************************************
'Scenario #5 - Reverse CPI Refund
'*********************************************************************************************************
Call writeReportLog("<b>Scenario 5: Reverse CPI Refund </b>" ,"Passed",FailComments,"")
'***********************************************************************************************************
'reverse CPI 
Call reverse_CPI_Refund_PaymentTab()

'check CPI reversed in payment tab
call checkPaymentTab_CPI(startDate,PayMethodAfterCPIReversed,ChargeAmt)

'verify cpi in component tab after reversed
Call verify_CPIAdded_afterReversed_ComponentTab()

'close current window
Call closeCurrentWindow()

'Closing Branch Plus Appllication
Call closeBranchApplication()

'Capturing End Time
endTime = Time()

'********************************************************************************************************************************************************
'Test Execution REPORTS - 'OnTest Execution Complete
'********************************************************************************************************************************************************
Call updateSummaryReport(startTime, endTime)

'**********************************************************************************************************************************************************

