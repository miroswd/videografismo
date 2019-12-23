''' Cria pausa e ativação na aba FUNDO_1 - para segundos e frames

input = inputbox("Digite o momento da pausa, separando por dois pontos." & VbCrLF & VbCrLF & "Exemplo » 02:23") 


If len(input) <> 5 and InStr(input,":") <> 3 then
	call validation()
Else
	keyFrames = (segundos * 30) + frames 

	ActiveCanvas.Scene.Transition("FUNDO_1").Activate
	ActiveCanvas.Scene.Select "pause"
	With ActiveCanvas.Scene.ActiveObject
	   .SuppressUpdates true
	   .KeyFrame kfPauseEnabled, keyFrames, 5.000000, kfJump
	   .KeyFrame kfPauseType, keyFrames, 1.000000, kfJump
	   .KeyFrame kfPauseTimeout, keyFrames, 0.000000, kfJump
	   .SuppressUpdates false
	End With
	ActiveCanvas.Scene.Select "f1_f2"
	ActiveCanvas.Scene.ActiveTransition.AddEvent 1, keyFrames + 1 '' Ativando outra aba de transição
	ActiveCanvas.Scene.ActiveTransition.GetEvent(1, keyFrames + 1).SetParam "Add.Activation", "FUNDO_2"
End If




'''' Validação - tentei criar com while, mas não deu

Sub validation()
	''' Validando se tem 5 caracteres e se ':' está na posição 3
	If len(input) <> 5 and InStr(input,":") = empty and InStr(input,":") <> 3 then
		input = inputbox("Valor inválido." & VbCrLF & VbCrLF & "Para segundos digite entre: 00 e 59" & VbCrLF & "Para frames digite entre: 00 e 29" ) 
		call validation()
	End If

	timePause = split(input,":") '' Separando os segundos e frames
	segundos = timePause(0)
	frames = timePause(1)
	
	''' Validando limite de segundos e frames
	If segundos > 59 and frames > 29 then
		input = inputbox("Valor inválido." & VbCrLF & VbCrLF & "Para segundos digite entre: 00 e 59" & VbCrLF & "Para frames digite entre: 00 e 29" ) 
		call validation()
	End if
End Sub
