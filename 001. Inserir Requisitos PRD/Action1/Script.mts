
'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "Requisitos", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)
 @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebElement("More")_;_script infofile_;_ZIP::ssf5.xml_;_
Browser("micClass:=Browser").Page("micClass:=Page").Link("html id:=create_link").Click
 
'cria os requsitos
while DataTable.Value("Requisito","Global") <>""
	
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=issuetype-field", "role:=combobox").WaitProperty "disabled", 0
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=issuetype-field", "role:=combobox").Set "Requirement"

	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=summary").Set datatable.value("Requisito", "Global") & " - " & datatable.value("Nome", "Global")
	'Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=summary").Set datatable.value("Requisito", "Global") 'no âmbito de importar info ja do JIRA
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=description").Set datatable.value("Descricao", "Global")

	
	' """"""""""""""""""""" Identificação da Release """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
	If datatable.value("Release", "Global") = "R3G" Then
		Browser("micClass:=Browser").Page("micClass:=Page").WebCheckBox("html id:=customfield_13600-2").Set "ON"
	Else
		Browser("micClass:=Browser").Page("micClass:=Page").WebCheckBox("html id:=customfield_13600-1").Set "ON"		
	End If	
	' """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""	
	

	'Seleciona o Tipo de Requisito
	Select Case trim(datatable.value("Tipo", "Global"))
		Case "Funcional"
			ReqTipoID = 1
		Case "Não Funcional"
			ReqTipoID = 2
	End Select
	
	Browser("micClass:=Browser").Page("micClass:=Page").WebList("html id:=customfield_10112").Select "#"&ReqTipoID
	
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
		
		Browser("micClass:=Browser").Page("micClass:=Page").WebList("hmtl id:=customfield_10112:1").Select "#"&ReqTipoID	
			
	End If
	
	'##### FALTA :: Browser("PTPRMOB1171646I board").Page("Create Issue - JIRA (QA)").WebEdit("txtLinkNotes").Set datatable.value("Lotus", "Global") @@ hightlight id_;_Browser("PTPRMOB1171646I board").Page("Levantamento ATM sem Cartão").WebEdit("customfield 11006")_;_script infofile_;_ZIP::ssf12.xml_;_
	

	'Valida se a checkBox Create Another está selecionado
	'Se 0 não está selecionado
	'Se 1 está selecionado
	checked=Browser("micClass:=Browser").Page("micClass:=Page").WebCheckBox("html id:=qf-create-another", "visible:=true", "value:=on").GetROProperty ("checked")
	If checked=0 Then
		Browser("micClass:=Browser").Page("micClass:=Page").WebCheckBox("hmtl id:=qf-create-another", "visible:=true", "value:=on").Click
	End If	
	
	'submeter o requisito
	Browser("micClass:=Browser").Page("micClass:=Page").WebButton("html id:=create-issue-submit").Click
	
	'avança para a linha seguinte do excel
	i=i+1
    DataTable.SetCurrentRow(i)  
wend

Browser("micClass:=Browser").Page("micClass:=Page").Link("class:=cancel", "visible:=true").Click

