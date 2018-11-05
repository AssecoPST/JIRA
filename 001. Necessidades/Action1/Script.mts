'DEFINIR INPUT
Dim Project, IssueType

Project = "Teste - Scrum (TS)"
IssueType = "Need"

'repositório do ficheiro a ler
localFicheiro = "P:\JIRA\"

'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "Necessidades", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)

'prime o botão Create
Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").Link("btnCreate").Click 

while DataTable.Value("Necessidade","Global") <>""
	Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").WebEdit("cboProject").WaitProperty "disabled", false, 20
	Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").WebEdit("cboProject").Set Project
	Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").WebEdit("txtIssueType").Set IssueType
	Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").WebEdit("txtsummary").Set datatable.value("Necessidade", "Global") & " - " & datatable.value("Nome", "Global")
	Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").WebList("cboOrigem").Select 1	
	
	'Valida se a checkBox Create Another está selecionado
	checked=Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").WebCheckBox("chkCreateAnother").GetROProperty ("checked")
	If checked=0 Then
		Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").WebCheckBox("chkCreateAnother").Click
	End If	
	
	Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").WebButton("btnCreateIssue").Click
	
	i=i+1
    DataTable.SetCurrentRow(i)  
wend

'sair
Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").Link("btnCancelIssue").Click
