''' Cria pausa e ativação na aba FUNDO_1 - para segundos e frames

pause = inputbox("Digite o momento da pausa, separando por dois pontos." & VbCrLF & VbCrLF & "Exemplo » 02:23","Defina a pausa", "00:00") 

If not IsEmpty(pause) then ''' Caso pressione o ok
	call validation()
End if 

Sub validation()
	If len(pause) <> 5 or InStr(pause,":") = null or InStr(pause,":") <> 3 then
		'Se o input for diferente de 5 ou não existir dois pontos ou os dois pontos não estiver na posição 3
		call inputTime()
	Else
		'Se passou da primeira validação
		timePause = split(pause,":")
		seconds = timePause(0)
		frames = timePause(1)
 			
 		If IsNumeric(seconds) = False or IsNumeric(frames) = False then
 			call inputTime()
		Else
 			call timeDefine(seconds,frames)
 		End If
	End If
End Sub

Function timeDefine(seconds,frames)
	If seconds > 59 or frames > 29 then 
		call inputTime()
	Else
		keyFrames = (seconds * 30) + frames
		call createKeyframes(keyFrames)
	End If
End Function


Sub inputTime()
	''' Será chamado caso dê erro na validação	
	pause = inputbox("Valor inválido." & VbCrLF & VbCrLF & "Para segundos digite entre: 00 e 59" & VbCrLF & "Para frames digite entre: 00 e 29", "ERRO!")

	If not IsEmpty(pause) then ''' Se não for cancelado
		call validation()
	End if 
End Sub


Function createKeyframes(keyFrames)
	''' Cria os keyframes de pausa e ativação - Só será chamado se passar pela validação
	ActiveCanvas.Scene.Transition("FUNDO_1").Activate
	ActiveCanvas.Scene.Select "pause"
	ActiveCanvas.Scene.ActiveObject.OutTime keyFrames
	With ActiveCanvas.Scene.ActiveObject
	   .SuppressUpdates true
	   .KeyFrame kfPauseEnabled, keyFrames, 5.000000, kfJump
	   .KeyFrame kfPauseType, keyFrames, 1.000000, kfJump
	   .KeyFrame kfPauseTimeout, keyFrames, 0.000000, kfJump
	   .SuppressUpdates false
	End With
	ActiveCanvas.Scene.Select "f1_f2"
	ActiveCanvas.Scene.ActiveObject.OutTime keyFrames + 1
	ActiveCanvas.Scene.ActiveTransition.AddEvent 1, keyFrames + 1 '' Ativando outra aba de transição (1 frame a mais em relação ao da pausa)
	ActiveCanvas.Scene.ActiveTransition.GetEvent(1, keyFrames + 1).SetParam "Add.Activation", "FUNDO_2"

End Function
