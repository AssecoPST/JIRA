AlterarRunMode

'###Project = "Levantamento ATM sem Cartão (EMIS)" #######

'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "CT.xlsx", "Test Sets", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)

'cria os requsitos
while DataTable.Value("Titulo","Global") <>""

	Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").Link("btnCreate").Click
	
	Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebEdit("cboIssueType").WaitProperty "visible", true
	Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebEdit("cboIssueType").Set "Test Set"

	Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebEdit("txtSummary").Set datatable.value("Titulo", "Global")
	
	Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebEdit("txtDescription").Set datatable.value("Descricao", "Global")

	'Seleciona a release	
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebElement("cboRelease").Object.setactive
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebElement("cboMoreRelease").Click
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebElement("cboRelease").Click

	Select Case datatable.value("Release", "Global")
		Case "R23"
			'Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").Link("ddlReleaseR23").Click
			Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebList("ddlRelease").Select 0

		Case "3G" 
			'Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").Link("ddlRelease3G").Click
			Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebList("ddlRelease").Select 1

	End Select

	
	'Valida se a checkBox Create Another está selecionado
	checked=Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebCheckBox("chkCreateAnother").GetROProperty ("checked")
	If checked=1 Then
		Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebCheckBox("chkCreateAnother").Click
	End If	
	
	'submeter o requisito
	Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebButton("btnCreateSubmit").Click

	i=i+1
    DataTable.SetCurrentRow(i)  
wend
