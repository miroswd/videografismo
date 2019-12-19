''' Gravador BDBR
bloco = ActiveCanvas.Scene.SelectedTemplate(2).text
lines = ActiveCanvas.Scene.SelectedTemplate(2).RowCount

''' Se o bloco estiver vazio
If bloco = empty then  
	For num = 10 to 17 ''' Número do template
		ActiveCanvas.Scene.SelectedTemplate(num).text = empty ''' Zera os templates
	Next
End If

''' Copiandos as linhas do Bloco, para os templates respectivos da linha
For numberLine = 1 to lines and bloco <> empty	''' Loop de 1 até o número de linhas do Bloco, se o bloco não estiver vazio
	ActiveCanvas.Scene.Select "BLOCO1"
	ActiveCanvas.Selection.Home scCtrl
	ActiveCanvas.Selection.Home
	
	If numberLine > 1 then ''' Se o número de linhas for maior que 1, deverá fazer a seleção das linhas posteriores
		For downCount = 1 to numberLine - 1 ''' Conta quantas linhas abaixo da primeira precisa ser copiada
			ActiveCanvas.Selection.Down
		Next
	End If

	ActiveCanvas.Selection.End scShift
	ActiveCanvas.Selection.Execute("Copy")

	ActiveCanvas.Scene.Select "LINHA" & numberLine '''Seleciono o template
	ActiveCanvas.Selection.Execute("Paste")
	
	For n = numberLine + 1 to 8 ''' Esvazia os templates
		ActiveCanvas.Scene.GetTemplate("LINHA" & n).text = empty
	Next
Next
