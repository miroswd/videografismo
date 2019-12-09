Set fso = CreateObject("Scripting.FileSystemObject")

'''''Define a pasta com os arquivos'''''
Set pasta = fso.GetFolder("I:\SPTV2\Movies\ARTES\")

Call rename
'''''Arquivos da pasta'''''
Sub rename

    For each arquivo in pasta.Files
        arquivo = arquivo.name

        '''''Verifica se o arquivo possui loop no nome'''''
        If InSTR(arquivo,"loop") then
            underscore = InSTR(arquivo,"_loop") - 1                       ''Acha a posição do underscore
            Set local = fso.GetFile("I:\SPTV2\Movies\ARTES\" & arquivo)   ''Local do arquivo
            arquivo = Mid(arquivo,1,underscore) & ".mov"                  ''Slicing e novo nome do arquivo
            
            '''''Renomeia o arquivo'''''
            local.Move ("I:\SPTV2\Movies\ARTES\" & arquivo)

        End IF    
    Next

End Sub
