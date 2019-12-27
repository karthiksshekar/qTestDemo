'******************************************************************************************************************
'@Autor Sumithra
'@Creation Date 15th June 2018
'@Test Name - verify Welcome Packets
'******************************************************************************************************************

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

'Fetching Current State name from DB
currentState = getCurrentStateName_From_DB()
Environment.Value("currentState") = currentState
Environment.Value("StateName") = currentState

'Custom Report FIle Creation
startTime = initializeCustomReport(currentState)
Environment.Value("startTime") = startTime
AccountStatus = "ACTIVE"
AccountType_Code = "55"

'********************************************************************************************************************
'Login into branch+ applications
Call BranchPlusLogin()

'*********************************************************************************************************************
'Generating Test Data Path
testDataPath =Environment.Value("BranchPlusSharedPath") & currentState & "\TestData\Add_Change_TextOpt_TestData.xlsx"

'Importing Test Data from Excel Shet(specific sheet)
Call importSpecificSheet(testDataPath,"Add_Change_Text")

'**********************************************************************************************************************
' Reading Data which is required for Test Script
'**********************************************************************************************************************

AccountStatus = DataTable.Value("AccountStatus")
AccountType_Code = DataTable.Value("AccountTypeCode")
'**********************************************************************************************************************

'Open F2 Screeen Account Activity
Call openF2Screen()

'open accounts
Loan_Number = get_LoanNumber_From_DB(AccountStatus,AccountType_Code)

'search customer by entering loan number
Call customer_Search_Using_LoanNumber_F2(Loan_Number)

'open Customer forms-F6
Call openCustomerFormsWindow()

'open Call Work Authorization Package
Call openReportFromPBTree_CustomerForms("English;Welcome Package;Call Work Authorization")
Call writeReportLog("Capturing screenshot: Call Work Authorization Package","Passed","","Y")

'open Notice of Insurance Requirement Package
Call openReportFromPBTree_CustomerForms("English;Welcome Package;Notice Of Insurance Requirement")
Call writeReportLog("Capturing screenshot: Notice Of Insurance Requirement Package","Passed","","Y")

'open Privacy statements package
Call openReportFromPBTree_CustomerForms("English;Welcome Package;Privacy Statement")
Call writeReportLog("Capturing screenshot: Privacy statement Package","Passed","","Y")

'open Text message disclosure package
Call openReportFromPBTree_CustomerForms("English;Welcome Package;Text Msg Disclosure")
Call writeReportLog("Capturing screenshot: Text message disclosure","Passed","","Y")
'Call closeCurrentWindow()

'open welcome letter package
Call openReportFromPBTree_CustomerForms("English;Welcome Package;Welcome Letter")
Call writeReportLog("Capturing screenshot: Welcome Letter","Passed","","Y")

'open welcome call check list package
Call openReportFromPBTree_CustomerForms("English;Welcome Package;Welcome Call Check List")
Call writeReportLog("Capturing screenshot: Welcome Call Check List","Passed","","Y")

'close current window
Call closeCurrentWindow()

'Capturing End Time
endTime = Time()

'Closing Window
Call closeCurrentWindow()

'Closing Branch Plus Appllication
Call closeBranchApplication()
'********************************************************************************************************************************************************
'Test Execution REPORTS - 'OnTest Execution Complete
'********************************************************************************************************************************************************
Call updateSummaryReport(startTime, endTime)

'********************************************************************************************************************************************************** @@ hightlight id_;_5900086_;_script infofile_;_ZIP::ssf77.xml_;_


