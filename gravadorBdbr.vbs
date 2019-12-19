bloco = ActiveCanvas.Scene.SelectedTemplate(2).text
lines = ActiveCanvas.Scene.SelectedTemplate(2).RowCount

If bloco = empty then
	For num = 10 to 17
		ActiveCanvas.Scene.SelectedTemplate(num).text = empty
	Next
End If

For numberLine = 1 to 8
	If lines = numberLine then
		ActiveCanvas.Scene.Select "BLOCO1"
		ActiveCanvas.Selection.Home scCtrl
		ActiveCanvas.Selection.Home
		If numberLine > 1 then
			For downCount = 1 to numberLine
				ActiveCanvas.Selection.Down
			Next
		End If
		ActiveCanvas.Selection.End scShift
		ActiveCanvas.Selection.Execute("Copy")

		ActiveCanvas.Scene.Select "LINHA" & numberLine
		ActiveCanvas.Selection.Execute("Paste")
		
		For n = numberLine + 1 to 8
			ActiveCanvas.Scene.GetTemplate("LINHA" & n).text = empty
		Next
	End If
Next
