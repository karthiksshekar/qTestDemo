'******************************************************************************************************************
'@Autor Sumithra
'@Creation Date 15th June 2018
'@Test Name - verify Welcome Packets
'******************************************************************************************************************

'Loading the required Object Repository
call loadMainObjectRepository()

'Loading the required Function Library
Call loadAllFunctionLibrary()

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



'****************************************************************************************
'@MethodName loadSpecificFunctionLibrary
'@Author Karthik.SHekar
'@Date 27 Dec 2019
'@Description this Function will Load the Function Library at Run time for the TEst Script
'@param	libFilePath --> Complete Path of the File where Function Libary is Located
'		libName ---> Library File Name along with it's extension

'Ex: loadSpecificFunctionLibrary("c:\FunctionLib\","BranchPlusFieldLevelValidation.qfl")
'*****************************************************************************************
Function loadSpecificFunctionLibrary(libFilePath,libName)
	LoadFunctionLibrary libFilePath & libName
End Function


'****************************************************************************************
'@MethodName loadAllFunctionLibrary
'@Author Karthik.SHekar
'@Date 27 Dec 2019
'@Description this Function will Load all Function Library at Run time for the TEst Script expect "Field Level Library"

'Ex: loadAllFunctionLibrary()
'*****************************************************************************************
Function loadAllFunctionLibrary()
	rtFldPath = getProjectRootFolderPath()
	libFoldPath = rtFldPath & "FunctionLibrary\"
	call loadSpecificFunctionLibrary(libFoldPath,"BranchPlusSQLFunctions.qfl")
	call loadSpecificFunctionLibrary(libFoldPath,"BranchPlusReportsFunctions.qfl")
	call loadSpecificFunctionLibrary(libFoldPath,"BranchPlusFunctions.qfl")
	call loadSpecificFunctionLibrary(libFoldPath,"CommonFunctions.qfl")
	Reporter.ReportEvent micDone,"loadAllFunctionLibrary", "Loading of all Function Library Completed" 
End Function

'****************************************************************************************
'@MethodName loadSpecificFunctionLibrary
'@Author Karthik.SHekar
'@Date 27 Dec 2019
'@Description this Function will Load the Object Repositry at Run time for the TEst Script
'@param	ORfilePath --> Complete Path of the File where OR is Located
'		ORName ---> OR File Name along with it's extension

'EX: loadSpecificObjectRepository(getRootFolderPath(),"BranchPlusObjectRepo.tsr")
'*****************************************************************************************
Function loadSpecificObjectRepository(ORfilePath,ORName)
	 call RepositoriesCollection.Add(ORfilePath  & ORName)
End Function

'****************************************************************************************
'@MethodName loadMainObjectRepository
'@Author Karthik.SHekar
'@Date 27 Dec 2019
'@Description this Function will Load Shared Branch Plus Object Repository at Run time 

'Ex: loadMainObjectRepository()
'*****************************************************************************************
Function loadMainObjectRepository()
	rtFldPath = getProjectRootFolderPath()
	Call loadSpecificObjectRepository(rtFldPath & "ObjectRepositories\","BranchPlusObjectRepo.tsr")	
End Function

'****************************************************************************************
'@MethodName getProjectRootFolderPath
'@Author Karthik.SHekar
'@Date 27 Dec 2019
'@Description this Function will get the RootFolder Path as per the GIT Folder Structure

''msgbox getProjectRootFolderPath() 
'*****************************************************************************************
Function getProjectRootFolderPath()
	testCasePath = Environment.Value("TestDir")
	rootFolderPath = Split(testCasePath,"TestScripts\")(0)
	Reporter.ReportEvent micDone,"Root Folder Path",rootFolderPath 
	getProjectRootFolderPath = rootFolderPath
End Function



