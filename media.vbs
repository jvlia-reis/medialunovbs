Dim n1, n2, n3, media, situacao, resp
call entrada_notas 
sub entrada_notas()
n1=cdbl(inputbox("Digite a nota 01", "ATEN��O"))
n2=cdbl(inputbox("Digite a nota 02", "ATEN��O"))
n3=cdbl(inputbox("Digite a nota 03", "ATEN��O"))

'PROCESSAMENTO
media=round((n1+n2+n3)/3,1)
if media < 4 then 'Se a m�dia menor que 4 ent�o 
   situacao="Reprovado"
elseif media >=4 and media < 6 then 
   situacao="Recupera��o"
else
   situacao="Aprovado com louvor" 
end if 

'Sa�da de Dados 
resp=msgbox ("M�dia do aluno: "& media &"" + vbnewline &_
    "Situa��o do aluno: "& situacao &"" + vbnewline &_
	"Novo c�lculo?",vbinformation+vbyesno,"Rendimento do Aluno")
if resp=vbyes then 
    call entrada_notas
else 
    wscript.quit
end if
end sub 