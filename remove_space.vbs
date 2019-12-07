'''Remove space
template1 = ActiveCanvas.Scene.GetTemplate("Template Name 1").Text
template2 = ActiveCanvas.Scene.GetTemplate("Template Name 2").Text

While InStr(template1, "  ")    ''' While exist more than 1 space
template1 = Replace(template1, "  ", " ") ''' Replaces 2 spaces for 1
ActiveCanvas.Scene.GetTemplate("Template Name 1").Text = template1
Wend

While InStr(template2,"  ")
template2 = Replace(template2, "  ", " ")
ActiveCanvas.Scene.GetTemplate("Template Name 2").Text = template2
Wend
