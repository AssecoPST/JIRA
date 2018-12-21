'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "link", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)

while DataTable.Value("Requisito","Global") <>""

	'valida se existe dependencias entre requisitos
	If trim(DataTable.Value("Dependencia","Global")) <> "" Then
		
		Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:= quickSearchInput").Set datatable.value("REQ_KEY", "Global")
		Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=quickSearchInput").Submit
		
		'valida se a consulta foi efetuada corretamente
		Issue = Browser("micClass:=Browser").Page("micClass:=Page").Link("html id:=key-val", "class:=issue-link").GetROProperty ("text")
		
		If Issue = datatable.value("REQ_KEY", "Global") Then
		
			Browser("micClass:=Browser").Page("micClass:=Page").Link("html id:=opsbar-operations_more").Click
			
			Browser("micClass:=Browser").Page("micClass:=Page").Link("html id:=link-issue").Click
			
			'Identificar a relação entre os issues
			Browser("micClass:=Browser").Page("micClass:=Page").WebList("html id:=link-type").WaitProperty "visible", true
			
			itemToSelect = "depends on" 'escolher a relação
			arrWeblistItems = split(Browser("micClass:=Browser").Page("micClass:=Page").WebList("html id:=link-type").GetROProperty("all items"),";")
			
			For itemCounter =0 to ubound(arrWeblistItems)		                     
				If trim(arrWeblistItems(itemCounter)) = trim(itemToSelect) Then 
					Browser("micClass:=Browser").Page("micClass:=Page").WebList("html id:=link-type").Select itemCounter
			   		Exit for
			   End If		
			Next                                            
				
			
			'procurar issue a associar
			Browser("micClass:=Browser").Page("micClass:=Page").Link("html id:=remote-jira-issue-search").Click
			
			Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=link-search-text").WaitProperty "visible", true
			Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=link-search-text").Set trim(datatable.value("LINK_REQ", "Global"))
			Browser("micClass:=Browser").Page("micClass:=Page").WebButton("html id:=simple-search-panel-button").Click
			
			Browser("micClass:=Browser").Page("micClass:=Page").WebCheckBox("html id:=linkjiraissue-select-all").Click	
			
			Browser("micClass:=Browser").Page("micClass:=Page").WebButton("html id:=linkjiraissue-add-selected").Click
			
			'submeter link
			Browser("micClass:=Browser").Page("micClass:=Page").WebButton("name:=Link", "type:=submit").Click
		
		
		End If
	End if 
	
	' valida se existem impactos que afectam o requisito
'	If trim(DataTable.Value("Impacto","Global")) <> "" Then
'		
'		Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("txtSearchString").Set datatable.value("REQ_KEY", "Global")
'		Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("txtSearchString").Submit
'	
'		'valida se a consulta foi efetuada corretamente
'		Issue = Browser("micClass:=Browser").Page("micClass:=Page").Link("lnkIssueType").GetROProperty ("text")
'		
'		If Issue = datatable.value("REQ_KEY", "Global") Then
'	
'			Browser("micClass:=Browser").Page("micClass:=Page").Link("mnuMore").Click
'			
'			Browser("micClass:=Browser").Page("micClass:=Page").Link("mnuLink").Click
'			
'			'Identificar a relação entre os issues
'			Browser("micClass:=Browser").Page("micClass:=Page").WebList("ddlRelacaoIssue").WaitProperty "visible", true
'			
'			itemToSelect = "affects" 'escolher a relação
'			arrWeblistItems = split(Browser("micClass:=Browser").Page("micClass:=Page").WebList("ddlRelacaoIssue").GetROProperty("all items"),";")
'			
'			For itemCounter =0 to ubound(arrWeblistItems)		                     
'				If trim(arrWeblistItems(itemCounter)) = trim(itemToSelect) Then 
'					Browser("micClass:=Browser").Page("micClass:=Page").WebList("ddlRelacaoIssue").Select itemCounter
'			   		Exit for
'			   End If		
'			Next                                            
'				
'			
'			'procurar issue a associar
'			Browser("micClass:=Browser").Page("micClass:=Page").Link("lnkSearchIssue").Click
'			
'			Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("txtSearchIssue").WaitProperty "visible", true
'			Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("txtSearchIssue").Set trim(datatable.value("LINK_IMP", "Global"))
'			Browser("micClass:=Browser").Page("micClass:=Page").WebButton("btnSearchIssue").Click
'			
'			Browser("micClass:=Browser").Page("micClass:=Page").WebCheckBox("chkIssue").Click	
'			
'			Browser("micClass:=Browser").Page("micClass:=Page").WebButton("btnAddIssue").Click
'			
'			'submeter link
'			Browser("micClass:=Browser").Page("micClass:=Page").WebButton("btnLink").Click
'				
'		End If
'	End If

	i=i+1
    DataTable.SetCurrentRow(i)  
wend



