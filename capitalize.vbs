''' Capitalize

N = 1  ''' Number of Templates
For T = 1 to N 
template = LCase(ActiveCanvas.Scene.SelectedTemplate(T).text)
ActiveCanvas.Scene.SelectedTemplate(T).text = Ucase(Mid(template,1,1))  & Mid(template,2)
Next
