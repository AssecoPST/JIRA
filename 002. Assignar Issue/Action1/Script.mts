AlterarRunMode

'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "Assignee", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)

while DataTable.Value("Requisito","Global") <>""
		
	Browser("[PTDDAGO1181737I-158]").Page("[PTDDAGO1181737I-158]").WebEdit("txtsearchString").Set trim(datatable.value("REQ_KEY", "Global"))
	Browser("[PTDDAGO1181737I-158]").Page("[PTDDAGO1181737I-158]").WebEdit("txtsearchString").Submit
	

	'valida se a consulta foi efetuada corretamente
	Issue = trim(Browser("[PTDDAGO1181737I-158]").Page("[PTDDAGO1181737I-158]").Link("lnkIssueType").GetROProperty ("innertext"))
	
	If Issue = datatable.value("REQ_KEY", "Global") Then
	
		Browser("[PTDDAGO1181737I-158]").Page("[PTDDAGO1181737I-158]").Link("btnAssign").Click
		
		Setting.WebPackage("ReplayType") = 2 'alterar modo de inserção do texto. caso contrário o JIRA não assume que o campo está preenchido
		Browser("[PTDDAGO1181737I-158]").Page("[PTDDAGO1181737I-158]").WebEdit("txtAssignee").Click
		Browser("[PTDDAGO1181737I-158]").Page("[PTDDAGO1181737I-158]").WebEdit("txtAssignee").Set trim(datatable.value("Assign", "Global"))
		Browser("[PTDDAGO1181737I-158]").Page("Assign: PTDDAGO1181737I-1").WebElement("btnMore").Click
		
		If Browser("[PTDDAGO1181737I-158]").Page("Assign: PTDDAGO1181737I-1").WebList("lstAssigneeSuggestions").GetROProperty("visible") = False Then
			Browser("[PTDDAGO1181737I-158]").Page("Assign: PTDDAGO1181737I-1").WebElement("btnMore").Click
		End If
	
		
		itemToSelect = trim(datatable.value("Assign", "Global")) 'escolher o assignee
		arrWeblistItems = split(Browser("[PTDDAGO1181737I-158]").Page("Assign: PTDDAGO1181737I-1").WebList("lstAssigneeSuggestions").GetROProperty("all items"),";")
		
		Setting.WebPackage("ReplayType") = 1
		For itemCounter =0 to ubound(arrWeblistItems)
			var = instr(1, trim(arrWeblistItems(itemCounter)), "-")	
			ValorLista = left(trim(arrWeblistItems(itemCounter)), var-2)			
			If ValorLista = trim(itemToSelect) Then 
				Browser("[PTDDAGO1181737I-158]").Page("Assign: PTDDAGO1181737I-1").WebList("lstAssigneeSuggestions").Select itemCounter
		   		Exit for
		   End If		
		Next                                            
				
					
		'submeter 
		Browser("[PTDDAGO1181737I-158]").Page("[PTDDAGO1181737I-158]").WebButton("btnSubmitAssign").Click

		
	End If

	
	i=i+1
    DataTable.SetCurrentRow(i)  
wend
