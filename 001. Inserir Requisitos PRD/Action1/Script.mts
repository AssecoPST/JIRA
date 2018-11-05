AlterarRunMode

'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "Requisitos", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)
 @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebElement("More")_;_script infofile_;_ZIP::ssf5.xml_;_
'cria os requsitos
while DataTable.Value("Requisito","Global") <>""
	
	Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").Link("btnCreate").Click
	
	Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebEdit("cboIssueType").WaitProperty "visible", true
	Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebEdit("cboIssueType").Set "Requirement"

	Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebEdit("txSummary").Set datatable.value("Requisito", "Global") & " - " & datatable.value("Nome", "Global")

	Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebEdit("txtDescription").Set datatable.value("Descricao", "Global")


	'Seleciona a release	
	Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebList("cboRelease").Select datatable.value("Release", "Global")

	'Seleciona o Tipo de Requisito
	Select Case trim(datatable.value("Tipo", "Global"))
		Case "Funcional"
			ReqTipoID = 1
		Case "Não Funcional"
			ReqTipoID = 2
	End Select
	
	Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebList("cboReqType1").Select "#"&ReqTipoID
	
	' Se é requisto "Não funcional" então preenche o subtipo
	If ReqTipoID = 2 Then 
	
		'valida subtipo 
		Select Case trim(datatable.value("SubTipo", "Global"))
			Case "Usabilidade"
				ReqTipoID = 1
			Case "Performance"
				ReqTipoID = 2
			Case "Segurança"
				ReqTipoID = 3
		End Select
		
		Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebList("cboReqType2").Select "#"&ReqTipoID	
			
	End If
	
	'##### FALTA :: Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebEdit("txtLinkNotes").Set datatable.value("Lotus", "Global") @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebEdit("customfield 11006")_;_script infofile_;_ZIP::ssf12.xml_;_
	

	'Valida se a checkBox Create Another está selecionado
	'Se tiver, retira a seleção, porque o JIRA bloqueia se forem efetuados vários registos sem fechar a janela.
	checked=Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebCheckBox("chkCreateAnother").GetROProperty ("checked")
	If checked=1 Then
		Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebCheckBox("chkCreateAnother").Click
	End If	
	
	'submeter o requisito
	Browser("Débitos Directos AGO -").Page("Débitos Directos AGO -").WebButton("btnSubmitCreate").Click
	
	'avança para a linha seguinte do excel
	i=i+1
    DataTable.SetCurrentRow(i)  
wend


