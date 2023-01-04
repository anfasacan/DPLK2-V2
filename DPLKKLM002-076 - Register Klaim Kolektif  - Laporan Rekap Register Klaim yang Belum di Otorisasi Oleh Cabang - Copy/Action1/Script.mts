Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dt_Username,preperation

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKKLM002-076 - Klaim - Transaksi.xlsx", "DPLKKLM002")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
preperation = Split(DataTable.Value("PREPERATION",dtlocalsheet),",")
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, preperation)
Iteration = Environment.Value("ActionIteration")
REM ------- DPLK


Call DA_Login()
Call AC_GoTo_Menu()

Set objkey = CreateObject("WScript.Shell")
wait 2
Call CaptureImageUFTV2(Browser("DPLK").Page("Register Klaim Kolektif").WebButton("Btn Tambah"),"Klik Button Tambah ", " ",  compatibilityMode.Desktop, ReportStatus.Passed)
wait 2
Browser("DPLK").Page("Register Klaim Kolektif").WebButton("Btn Tambah").Click
wait 2
Browser("DPLK").Page("Register Klaim Kolektif").WebEdit("Field Tanggal Pengajuan").click
objkey.SendKeys (DataTable.Value("TANGGAL_TERIMA_DOKUMEN",dtlocalsheet))

Call CaptureImageUFTV2(Browser("DPLK").Page("Register Klaim Kolektif").WebButton("Btn Search No Kolektif"),"Klik Button Cari Data Dan Pilih Data Yang Ingin Digunakan ", " ",  compatibilityMode.Desktop, ReportStatus.Passed)
wait 2
Browser("DPLK").Page("Register Klaim Kolektif").WebButton("Btn Search No Kolektif").Click
Browser("DPLK").Page("Register Klaim Kolektif").WebEdit("Field Search Kolektif").Set DataTable.Value("NO_KOLEKTIF",dtlocalsheet)
wait 3
Browser("DPLK").Page("Register Klaim Kolektif").WebElement("Object Table No Kolektif").Click
wait 60
Jumlah_Row = Browser("DPLK").Page("Register Klaim Kolektif").WebTable("Table Tambah Kolektif").RowCount()
For Iterator = 2 To Jumlah_Row Step 1
	No_Rek_DPLK = Browser("DPLK").Page("Register Klaim Kolektif").WebTable("Table Tambah Kolektif").GetCellData(Iterator,2)
	If No_Rek_DPLK = DataTable.Value("NO_REK_DPLK",dtlocalsheet) Then
		set CekBox = Browser("DPLK").Page("Register Klaim Kolektif").WebTable("Table Tambah Kolektif").ChildItem(iterator,1,"WebCheckBox",0)
		CekBox.click
		wait 2
		Call CaptureImageUFTV2(Browser("DPLK").Page("Register Klaim Kolektif").WebButton("Btn Search No Kolektif"),"Pilih Data Yang Digunakan", "Disini Menggunakan Data Norek PDLK " & No_Rek_DPLK ,  compatibilityMode.Desktop, ReportStatus.Passed)
		wait 2		
		Exit for 
	End If
Next

Browser("DPLK").Page("Register Klaim Kolektif").WebCheckBox("Check Box St Rek Perusahaan").Click
wait 2
Call CaptureImageUFTV2(Browser("DPLK").Page("Register Klaim Kolektif"),"Isi Semua Field yang dibutuhkan", " ",  compatibilityMode.Desktop, ReportStatus.Done)
wait 2
Call CaptureImageUFTV2(Browser("DPLK").Page("Register Klaim Kolektif").WebButton("Btn Kirim Ke Calculate"),"Klik Button Kirim Ke Calculate", " ",  compatibilityMode.Desktop, ReportStatus.Done)
wait 2

'Browser("DPLK").Page("Register Klaim Kolektif").WebButton("Btn Kirim Ke Calculate").Click



'Call DA_Logout("0")


Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDPLK, LibReport, LibRepo, objSysInfo
	Dim tempDPLKPath, tempDPLKPath2, PathDPLK
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempDPLKPath 	= Environment.Value("TestDir")
	tempDPLKPath2 	= InStrRev(tempDPLKPath, "\")
	PathDPLK 		= Left(tempDPLKPath, tempDPLKPath2)
	
	LibPathDPLK	= PathDPLK & "Lib_Repo_Excel\LibDPLK\"
	LibReport			= PathDPLK & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDPLK & "Lib_Repo_Excel\Repo\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- DPLK lib
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Klaim_Laporan.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_LogMenu.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Dashboard.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Log.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Function.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Klaim_Transaksi.tsr")
	
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_Username					= DataTable.Value("USERID",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
End Sub
