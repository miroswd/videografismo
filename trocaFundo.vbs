''' Cria pausa e ativação na aba FUNDO_1 - para segundos e frames

input = inputbox("Digite o momento da pausa, separando por dois pontos." & VbCrLF & VbCrLF & "Exemplo » 02:23") 

While len(input) <> 5 or InStr(input,":") <> 3
	input = inputbox("Valor inválido." & VbCrLF & VbCrLF & "Para segundos digite entre: 00 e 59" & VbCrLF & "Para frames digite entre: 00 e 29" ) 
Wend

Sub createKeyframes()
	''' Cria os keyframes de pausa e ativação - Só será chamado se passar pela validação
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
	ActiveCanvas.Scene.ActiveTransition.AddEvent 1, keyFrames + 1 '' Ativando outra aba de transição (1 frame a mais em relação ao da pausa)
	ActiveCanvas.Scene.ActiveTransition.GetEvent(1, keyFrames + 1).SetParam "Add.Activation", "FUNDO_2"
End Sub


Sub validation()
	''' Verificar se possui 5 caracteres
	''' Verificar se segundos e frames são numéricos
	''' Verificar se o 3º caracter são os 2 pontos
	''' Verificar se segundos é menor do que 60
	''' Verificar se frames é menor do que 30

	While len(input) <> 5 or InStr(input,":") = empty or InStr(input,":") <> 3 then
		input = inputbox("Valor inválido." & VbCrLF & VbCrLF & "Para segundos digite entre: 00 e 59" & VbCrLF & "Para frames digite entre: 00 e 29" ) 
	Wend

	timePause = split(input,":") '' Separando os segundos e frames
	segundos = timePause(0)
	frames = timePause(1)
	
	While segundos > 59 or frames > 29 then
		input = inputbox("Valor inválido." & VbCrLF & VbCrLF & "Para segundos digite entre: 00 e 59" & VbCrLF & "Para frames digite entre: 00 e 29" ) 
	Wend
End Sub
