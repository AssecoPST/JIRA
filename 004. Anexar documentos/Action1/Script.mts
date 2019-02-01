'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "Anexos", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)

while DataTable.Value("Requisito","Global") <>""
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=quickSearchInput").WaitProperty "disabled", 0
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=quickSearchInput").Set datatable.value("REQ_KEY", "Global")
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=quickSearchInput").Submit
	Browser("micClass:=Browser").Page("micClass:=Page").WebEdit("html id:=quickSearchInput").WaitProperty "disabled", 0
	
	'valida se a consulta foi efetuada corretamente
	Issue = Browser("micClass:=Browser").Page("micClass:=Page").Link("html id:=key-val", "class:=issue-link").GetROProperty ("text")
	
	If Issue = datatable.value("REQ_KEY", "Global") Then
	
		Browser("micClass:=Browser").Page("micClass:=Page").Link("html id:=opsbar-operations_more").Click
		
		Window("Google Chrome").WinObject("Chrome Legacy Window").Click 592,376 'botão/link "Attach files"
		
		Window("Google Chrome").Window("Open").WaitProperty "visible", true
		Window("Google Chrome").Window("Open").WinObject("micClass:=WinEdit", "attached text:=File &name:").Click  @@ hightlight id_;_527208_;_script infofile_;_ZIP::ssf6.xml_;_
		Window("Google Chrome").Window("Open").WinObject("micClass:=WinEdit", "attached text:=File &name:").Type datatable.value("Ficheiro", "Global") @@ hightlight id_;_527208_;_script infofile_;_ZIP::ssf7.xml_;_
		Window("Google Chrome").Window("Open").WinObject("micClass:=WinEdit", "attached text:=File &name:").Type  micReturn  @@ hightlight id_;_527878_;_script infofile_;_ZIP::ssf2.xml_;_
		
	End If

	i=i+1
    DataTable.SetCurrentRow(i)  
wend @@ hightlight id_;_527878_;_script infofile_;_ZIP::ssf2.xml_;_

	

