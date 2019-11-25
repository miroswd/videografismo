''' Ativando os pontos de A->B ou B->A

If ActiveCanvas.Scene.getTemplate("ROTA").text = "AB" then
ActiveCanvas.Scene("BA").visible = true
ActiveCanvas.Scene("AB").visible = false
ActiveCanvas.Scene("CA").visible = false

'''POSICIONA
With ActiveCanvas.Scene.ActiveObject
   .SuppressUpdates true
   .KeyFrame kfXPos, 0, -0.311600
   .KeyFrame kfYPos, 0, -0.759100
   .KeyFrame kfZRot, 0, 42.850000
   .KeyFrame kfXCenter, 0, 0.000000
   .KeyFrame kfYCenter, 0, 0.500000
   .KeyFrame kfZCenter, 0, 0.000000
   .SuppressUpdates false
End With

'''ANIMA
With ActiveCanvas.Scene.ActiveObject
   .SuppressUpdates true
   .KeyFrame kfXScale, 60, 0.840000
   .KeyFrame kfYScale, 1, 0.020000
   .SuppressUpdates false
End With

ElseIf ActiveCanvas.Scene.getTemplate("ROTA").text = "BA" then
ActiveCanvas.Scene("AB").visible = true
ActiveCanvas.Scene("BA").visible = false
ActiveCanvas.Scene("CA").visible = false

'''POSICIONA
ActiveCanvas.Scene.ActiveObject.KeyFrame kfZRot, 0, -135.950000
With ActiveCanvas.Scene.ActiveObject
   .SuppressUpdates true
   .KeyFrame kfXPos, 0, 0.296800
   .KeyFrame kfYPos, 0, -0.183200
   .KeyFrame kfZRot, 0, -135.950000
   .SuppressUpdates false
End With

'''ANIMA
ActiveCanvas.Scene.ActiveObject.KeyFrame kfYScale, 0, 0.020000
With ActiveCanvas.Scene.ActiveObject
   .SuppressUpdates true
   .KeyFrame kfXScale, 0, 0.000000
   .KeyFrame kfXScale, 60, 0.840000
   .KeyFrame kfYScale, 0, 0.020000
   .SuppressUpdates false
End With

ElseIf ActiveCanvas.Scene.getTemplate("ROTA").text = "CA" then
ActiveCanvas.Scene("CA").visible = true
ActiveCanvas.Scene("AB").visible = false
ActiveCanvas.Scene("BA").visible = false

With ActiveCanvas.Scene.ActiveObject
   .SuppressUpdates true
   .KeyFrame kfXPos, 0, -0.308200
   .KeyFrame kfYPos, 0, -0.764300
   .KeyFrame kfZRot, 0, 90.000000
   .SuppressUpdates false
End With

ActiveCanvas.Scene.ActiveObject.KeyFrame kfYScale, 1, 0.020000
With ActiveCanvas.Scene.ActiveObject
   .SuppressUpdates true
   .KeyFrame kfXScale, 60, 0.360000
   .KeyFrame kfYScale, 1, 0.020000
   .SuppressUpdates false
End With


Else
msgbox "rota errada"
End IF




