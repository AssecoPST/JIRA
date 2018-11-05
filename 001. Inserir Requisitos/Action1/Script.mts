AlterarRunMode

'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "Requisitos", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)
 @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebElement("More")_;_script infofile_;_ZIP::ssf5.xml_;_
'cria os requsitos
while DataTable.Value("Requisito","Global") <>""

	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").Link("btnCreate").Click @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").Link("Create")_;_script infofile_;_ZIP::ssf3.xml_;_
	
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebEdit("cboIssueType").WaitProperty "visible", true
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebEdit("cboIssueType").Set "Requirement"

	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebEdit("txtSummary").Set datatable.value("Requisito", "Global") & " - " & datatable.value("Nome", "Global") @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebEdit("summary")_;_script infofile_;_ZIP::ssf6.xml_;_
	
	'Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").Link("tabText").Click
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebEdit("txtDescription").Set datatable.value("Descricao", "Global")
	
	'Seleciona a release	
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebList("lstRelease").Select "3G"
 @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebList("customfield 13102-textarea")_;_script infofile_;_ZIP::ssf8.xml_;_
	'valida a posição do tipo de requisito na lista
	Select Case datatable.value("Tipo", "Global")
		Case "Funcional"
			ReqTipoID = 1
		Case "Não Funcional"
			ReqTipoID = 2
	End Select
	
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebList("cboRequirementType").Click
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebList("cboRequirementType").Select ReqTipoID
	
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebEdit("txtLinkNotes").Set datatable.value("Lotus", "Global") @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebEdit("customfield 11006")_;_script infofile_;_ZIP::ssf12.xml_;_
	
	'Valida se a checkBox Create Another está selecionado
	checked=Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebCheckBox("chkCreateAnother").GetROProperty ("checked")
	If checked=1 Then
		Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebCheckBox("chkCreateAnother").Click
	End If	
	
	'submeter o requisito
	Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebButton("btnCreateSubmit").Click @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebButton("Create")_;_script infofile_;_ZIP::ssf14.xml_;_

	i=i+1
    DataTable.SetCurrentRow(i)  
wend


