'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "Assignee", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)

while DataTable.Value("Requisito","Global") <>""
	
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=quickSearchInput").WaitProperty "disabled", 0	
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=quickSearchInput").Set trim(datatable.value("REQ_KEY", "Global"))
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=quickSearchInput").Submit
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=quickSearchInput").WaitProperty "disabled", 0

	'valida se a consulta foi efetuada corretamente
	Issue = trim(Browser("micClass:=Browser").Page("micClass:=Page").Link("html id:=key-val", "class:=issue-link").GetROProperty ("innertext"))
	
	If Issue = datatable.value("REQ_KEY", "Global") Then
	
		Browser("micClass:=Browser").Page("micClass:=Page").Link("html id:=assign-issue").Click
		
		Setting.WebPackage("ReplayType") = 2 'alterar modo de inserção do texto. caso contrário o JIRA não assume que o campo está preenchido
		Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=assignee-field").Click
		Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=assignee-field").Set trim(datatable.value("Assign", "Global"))
		Browser("micClass:=Browser").Page("micClass:=Page").WebElement("innertext:=More", "html tag:=SPAN", "index:=6").Click
		
		If Browser("micClass:=Browser").Page("micClass:=Page").WebList("html id:=assignee-suggestions").GetROProperty("visible") = False Then
			Browser("micClass:=Browser").Page("micClass:=Page").WebElement("innertext:=More", "html tag:=SPAN", "index:=6").Click
		End If
		
		itemToSelect = trim(datatable.value("Assign", "Global")) 'escolher o assignee
		arrWeblistItems = split(Browser("micClass:=Browser").Page("micClass:=Page").WebList("html id:=assignee-suggestions").GetROProperty("all items"),";")
		
		Setting.WebPackage("ReplayType") = 1
		For itemCounter =0 to ubound(arrWeblistItems)
			var = instr(1, trim(arrWeblistItems(itemCounter)), "-")	
			ValorLista = left(trim(arrWeblistItems(itemCounter)), var-2)			
			If ValorLista = trim(itemToSelect) Then 
				Browser("micClass:=Browser").Page("micClass:=Page").WebList("html id:=assignee-suggestions").Select itemCounter
		   		Exit for
		   End If		
		Next                                            
				
					
		'submeter 
		Browser("micClass:=Browser").Page("micClass:=Page").WebButton("html id:=assign-issue-submit").Click
	
	End If

	
	i=i+1
    DataTable.SetCurrentRow(i)  
wend

'msgbox Browser("micClass:=Browser").Page("micClass:=Page").WebList("html id:=assignee-suggestions").GetROProperty("visible")
'
'Setting.WebPackage("ReplayType") = 2 'alterar modo de inserção do texto. caso contrário o JIRA não assume que o campo está preenchido
'Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=assignee-field").Click
'Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=assignee-field").Set "João Lopes"
'Browser("micClass:=Browser").Page("micClass:=Page").WebElement("innertext:=More", "html tag:=SPAN", "index:=6").Click
'
'msgbox Browser("micClass:=Browser").Page("micClass:=Page").WebList("html id:=assignee-suggestions").GetROProperty("visible")
'Browser("micClass:=Browser").Page("micClass:=Page").WebElement("innertext:=More", "html tag:=SPAN", "index:=6").Click
