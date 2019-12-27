'*****************************************************************************************************************
'@Author: Sumithra HP
'@Creation Date: 28th June 2018
'@Test Script Name: Add_CPI_Approved

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

'******************************************************************************************************************
'Login into branch+ application and getting the State Name

'Fetching Current State name from DB
currentState = getCurrentStateName_From_DB()
Environment.Value("currentState") = currentState
Environment.Value("StateName") = currentState

'Custom Report FIle Creation
startTime = initializeCustomReport(currentState)
Environment.Value("startTime") = startTime

'Login to Branch Plus Application
Call BranchPlusLogin()

'******************************************************************************************************************
'Generating Test Data Path
testDataPath =Environment.Value("BranchPlusSharedPath") & currentState & "\TestData\CPI_Approved_TestData.xlsx"

'Importing Test Data from Excel Shet(specific sheet)
Call importSpecificSheet(testDataPath,"CPI")

'*******************************************************************************************************************
' Reading Data which is required for Test Script
'*******************************************************************************************************************
AccountStatus = DataTable.Value("AccountStatus")
AccountType_Code = DataTable.Value("AccountTypeCode")
DealerComm = DataTable.Value("DealerCOMM")
RepoComm = DataTable.Value("RepoCOMM")
SurAmount =DataTable.Value("SurAmount")
OtherFee = DataTable.Value("OtherFee")
PFSChargeAmt =DataTable.Value("GapChargeAmount")
VSCchargeAmount =DataTable.Value("VSCChargeAmount")
CompDescOne = DataTable.Value("CompDescOne")
CompDescTwo = DataTable.Value("CompDescTwo")
CompDescThree= DataTable.Value("CompDescThree")

'******************************************************************************************************************

wait 5
'get account number from db
acctNum = get_Account_Id_From_DB(AccountStatus,AccountType_Code)

'Open F2 Screeen Account Activity
Call openF2Screen()

wait 5
'search customer by entering account number
Call customer_Search_Using_Account_ID_F2(acctNum)
wait 10
'get reference id
refId =getReferenceID_AccountTab()

'delete component descriptions
Call delete_ComponentDescription_ComponentTab_Approved()

'**************************
'STATE VARIATION
'**************************
If Environment.Value("currentState") = "FL" Then
  'get DocStamp value 
   DocStampChargeBfr = get_ChargeAmount_Component_Tab(1)
End If
 @@ hightlight id_;_2688856_;_script infofile_;_ZIP::ssf103.xml_;_
'Navigate to Account tab
Call navigateToTab("Account")
	
'get DealerCheck in Acount tab
 DealerCheck = get_DealerCheck_AccountTab()
 
'get AmtFinance in Acount tab
AmtFinance = get_TotalFinance_AccountTab()	

'get TotalFinance in Acount tab
TotalFinance= get_TotalRepayable_AccountTab()

'Click Button
Call clickOnSaveButon()

'Navigate to Component tab
Call navigateToTab("Component")

wait 15

''*****************
''Windows 10 Variation
''*****************
'If Environment.Value("OS") = "Windows 10" and Environment.Value("currentState") = "FL" Then
'	CompDescOne = "GAP"
'End If

'add component description value
Call add_ComponentDescription_Component_Tab(1,CompDescOne) @@ hightlight id_;_7145182_;_script infofile_;_ZIP::ssf15.xml_;_
wait(3)
'enter charge amount in component tab
Call enter_ChargeAmount_Component_Tab(1,PFSChargeAmt)

 
'add component description value
Call add_ComponentDescription_Component_Tab(2,CompDescTwo)
'enter charge amount in component tab
Call enter_ChargeAmount_Component_Tab(2,VSCchargeAmount)
'enter surcharge amount in component tab
Call enter_SurchargeAmount_Component_Tab(2,SurAmount)
'enter repo comm in component tab
Call enter_RepoComm_Component_Tab(2,RepoComm)
'enter dealer comm in component tab
Call enter_DealerComm_Component_Tab(2,DealerComm)

Call add_ComponentDescription_Component_Tab(3,CompDescThree)
Call enter_ChargeAmount_Component_Tab(3,OtherFee)

'click on save button
Call clickOnSaveButon()

'***********************
'WIndows Variation 
'***********************
'If Environment.Value("OS") <> "Windows 10" and Environment.Value("currentState") <> "FL" Then
	''update PFS GAP
	Call UpdateComponentDescription_ComponentTab(refId)
	'click on refresh button
	Call DoRefresh_F5()
'End If

Call writeReportLog("Component description : PFS GAP is added successfully","Passed",FailComments,"")

'enter charge amount in component tab
Call enter_ChargeAmount_Component_Tab(1,PFSChargeAmt)
'enter repo comm in component tab
Call enter_RepoComm_Component_Tab(1,RepoComm)
'enter dealer comm in component tab
Call enter_DealerComm_Component_Tab(1,DealerComm)

'click on save button
Call clickOnSaveButon()	
	
If Environment.Value("currentState") = "FL" Then
    'get doc stamp value
    DocStampCharge = get_ChargeAmount_Component_Tab(1)
    'get total other charges
	ExpTotalotherCharge = ccur(PFSChargeAmt)+ccur(VSCchargeAmount)+ccur(OtherFee)+ccur(SurAmount)+ccur(DocStampCharge)
	'get amount financed
	ExpAmtFinanced = ccur(AmtFinance)+ccur(ExpTotalotherCharge)-ccur(DocStampChargeBfr)
else
    'get total other charges 
	ExpTotalotherCharge = ccur(PFSChargeAmt)+ccur(VSCchargeAmount)+ccur(OtherFee)+ccur(SurAmount)
	'get amount financed
	ExpAmtFinanced = ccur(AmtFinance)+ccur(ExpTotalotherCharge)
End If

ExpRepoComm = ccur(RepoComm)+ccur(RepoComm)
ExpDealerCheck = ccur(DealerCheck)+ccur(DealerComm)+ccur(OtherFee)+ccur(DealerComm)

'Navigate to Account tab
Call navigateToTab("Account")

'get DealerCheck in Acount tab
ActDealerCheck = get_DealerCheck_AccountTab()
'get RepoCheck in Acount tab
ActRepoCheck =  get_RepoCheck_AccountTab()
'get AmtFinance in Acount tab
ActAmtFinance = get_TotalFinance_AccountTab()
'get TotalFinance in Acount tab
ActTotalFinance= get_TotalRepayable_AccountTab()
'get total other charges
ActTotalOtherChg =get_TotalOtherCharges_AccountTab()
'get total interset in account tab
TotalInterest = get_TotalInterest_AccountTab()
'get expected total finance
ExpTotalFinance = ccur(TotalInterest)+ccur(ActAmtFinance)

'validate Amount financed
Call validateDataAndReport(ccur(ActAmtFinance),ccur(ExpAmtFinanced),"Validation of Amount financed: ActAmtFinance: $" & ccur(ActAmtFinance) & "  ExpAmtFinanced: $ " & ccur(ExpAmtFinanced))
'validate total other charge
Call validateDataAndReport(ccur(ActTotalOtherChg),ccur(ExpTotalotherCharge),"Validation of Total other charges: ActTotalotherChg: $" & ccur(ActTotalOtherChg) & "  ExpTotalotherCharge: $ " & ccur(ExpTotalotherCharge))
'validate Repo comm
Call validateDataAndReport(ccur(ActRepoCheck),ccur(ExpRepoComm),"Validation of Repo Commission Check: ActRepoCheck: $" & ccur(ActRepoCheck) & "  ExpRepoCheck: $ " & ccur(ExpRepoComm))
'validate dealer check
Call validateDataAndReport(ccur(ActDealerCheck),ccur(ExpDealerCheck),"Validation of Dealer Check Amount: ActDealerCheck: $" & ccur(ActDealerCheck) & "  expDealerCheck: $ " & ccur(ExpDealerCheck))
'validtae total financed
Call validateDataAndReport(ccur(ActTotalFinance),ccur(ExpTotalFinance),"Validation of TotalFinance Amount: ActTotalFinance: $" & ccur(ActTotalFinance) & "  ExpTotalFinance: $ " & ccur(ExpTotalFinance))

'Closing Window
 Call closeCurrentWindow()
 
'Closing Branch Plus Appllication
'Call closeBranchApplication()

'Capturing End Time
 endTime = Time()
 
 '************************************************************************************************************************************************
'Test Execution REPORTS - 'OnTest Execution Complete
'************************************************************************************************************************************************
Call updateSummaryReport(startTime, endTime)

'************************************************************************************************************************************************

