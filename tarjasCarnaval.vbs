textFile = "texto"
localFile = "I:\TESTE\Automacoes\" & textFile & ".txt"

set fso = CreateObject("Scripting.FileSystemObject")
set oFso = fso.OpenTextFile(localFile)

msg = 12889
Do Until oFso.AtEndOfStream
msg = msg + 1
ActiveCanvas.Scene.SelectedTemplate(0).text = oFso.ReadLine
Lyric.Message msg
Lyric.Record

Loop
