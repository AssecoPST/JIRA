'DEFINIR INPUT
Dim Project, IssueType

Project = "Teste - Scrum (TS)"
IssueType = "Requirement"

'repositório do ficheiro a ler
localFicheiro = "P:\JIRA\"

'importa da sheet Necessidades para a GlobalSheet 
datatable.ImportSheet localFicheiro & "datapool.xlsx", "Requisitos", "Global"

'Posicionar na primeira linha 
i=1
DataTable.SetCurrentRow(i)

'prime o botão Create
Browser("[TS-32] Integração Routing").Page("[TS-32] Integração Routing").Link("btnCreate").Click 

Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebEdit("cboProject").Set Project
Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebEdit("cboIssueType").Set IssueType

while DataTable.Value("Requisito","Global") <>""

	Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebEdit("txtSummary").WaitProperty "disabled", 0
	Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebEdit("txtSummary").Set datatable.value("Requisito", "Global") & " - " & datatable.value("Nome", "Global")
	
	'Origem = Cliente Interno
	Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebList("cboOrigem").Select 1
	
	'indicar o link para o Lotus Note
	Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebEdit("txtLinkNote").Set datatable.value("Lotus", "Global")

	'seleciona as releases R23 e 3G
	If not Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebButton("btnAffectsVersionsR23").Exist Then
		Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebElement("cboAffectsVersions").Click
		Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").Link("lnkAffectsVersionsR23").Click
		
		Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebElement("cboAffectsVersions").Click
		Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").Link("lnkAffectsVersions3G").Click
	End If 
	
	'Valida se a checkBox Create Another está selecionado
	checked=Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebCheckBox("chkCreateAnother").GetROProperty ("checked")
	If checked=0 Then
		Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebCheckBox("chkCreateAnother").Click
	End If	
	
	Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").WebButton("btnCreateIssue").Click
	
	i=i+1
    DataTable.SetCurrentRow(i)  
wend

'sair
Browser("[TS-32] Integração Routing").Page("Create Issue - JIRA-TST").Link("btnCancel").Click

If Browser("[TS-32] Integração Routing").Dialog("Message from webpage").WinButton("btnMsgOK").Exist Then
	Browser("[TS-32] Integração Routing").Dialog("Message from webpage").WinButton("btnMsgOK").Click
End If
