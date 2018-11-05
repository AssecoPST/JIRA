AlterarRunMode

'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "link", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)

while DataTable.Value("Requisito","Global") <>""

	'valida se existe dependencias entre requisitos
	If trim(DataTable.Value("Dependencia","Global")) <> "" Then
		
		Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebEdit("txtSearchString").Set datatable.value("REQ_KEY", "Global")
		Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebEdit("txtSearchString").Submit
		
		'valida se a consulta foi efetuada corretamente
		Issue = Browser("[TS-113] REQ054 - Manuntenção").Page("[TS-113] REQ054 - Manuntenção").Link("lnkIssueType").GetROProperty ("text")
		
		If Issue = datatable.value("REQ_KEY", "Global") Then
		
			Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").Link("mnuMore").Click
			
			Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").Link("mnuLink").Click
			
			'Identificar a relação entre os issues
			Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebList("ddlRelacaoIssue").WaitProperty "visible", true
			
			itemToSelect = "depends on" 'escolher a relação
			arrWeblistItems = split(Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebList("ddlRelacaoIssue").GetROProperty("all items"),";")
			
			For itemCounter =0 to ubound(arrWeblistItems)		                     
				If trim(arrWeblistItems(itemCounter)) = trim(itemToSelect) Then 
					Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebList("ddlRelacaoIssue").Select itemCounter
			   		Exit for
			   End If		
			Next                                            
				
			
			'procurar issue a associar
			Browser("[TS-33] NEC001 - Login").Page("Link - JIRA-TST").Link("lnkSearchIssue").Click
			
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebEdit("txtSearchIssue").WaitProperty "visible", true
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebEdit("txtSearchIssue").Set trim(datatable.value("LINK_REQ", "Global"))
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebButton("btnSearchIssue").Click
			
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebCheckBox("chkIssue").Click	
			
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebButton("btnAddIssue").Click
			
			'submeter link
			Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebButton("btnLink").Click
		
		
		End If
	End if 
	
	' valida se existem impactos que afectam o requisito
	If trim(DataTable.Value("Impacto","Global")) <> "" Then
		
		Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebEdit("txtSearchString").Set datatable.value("REQ_KEY", "Global")
		Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebEdit("txtSearchString").Submit
	
		'valida se a consulta foi efetuada corretamente
		Issue = Browser("[TS-113] REQ054 - Manuntenção").Page("[TS-113] REQ054 - Manuntenção").Link("lnkIssueType").GetROProperty ("text")
		
		If Issue = datatable.value("REQ_KEY", "Global") Then
	
			Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").Link("mnuMore").Click
			
			Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").Link("mnuLink").Click
			
			'Identificar a relação entre os issues
			Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebList("ddlRelacaoIssue").WaitProperty "visible", true
			
			itemToSelect = "affects" 'escolher a relação
			arrWeblistItems = split(Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebList("ddlRelacaoIssue").GetROProperty("all items"),";")
			
			For itemCounter =0 to ubound(arrWeblistItems)		                     
				If trim(arrWeblistItems(itemCounter)) = trim(itemToSelect) Then 
					Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebList("ddlRelacaoIssue").Select itemCounter
			   		Exit for
			   End If		
			Next                                            
				
			
			'procurar issue a associar
			Browser("[TS-33] NEC001 - Login").Page("Link - JIRA-TST").Link("lnkSearchIssue").Click
			
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebEdit("txtSearchIssue").WaitProperty "visible", true
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebEdit("txtSearchIssue").Set trim(datatable.value("LINK_IMP", "Global"))
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebButton("btnSearchIssue").Click
			
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebCheckBox("chkIssue").Click	
			
			Browser("[TS-33] NEC001 - Login").Page("Find JIRA issues - JIRA-TST").WebButton("btnAddIssue").Click
			
			'submeter link
			Browser("[TS-33] NEC001 - Login").Page("[TS-33] NEC001 - Login").WebButton("btnLink").Click
				
		End If
	End If

	i=i+1
    DataTable.SetCurrentRow(i)  
wend
