Dim n1, n2, n3, media, situacao, resp
call entrada_notas 
sub entrada_notas()
n1=cdbl(inputbox("Digite a nota 01", "ATENÇÃO"))
n2=cdbl(inputbox("Digite a nota 02", "ATENÇÃO"))
n3=cdbl(inputbox("Digite a nota 03", "ATENÇÃO"))

'PROCESSAMENTO
media=round((n1+n2+n3)/3,1)
if media < 4 then 'Se a média menor que 4 então 
   situacao="Reprovado"
elseif media >=4 and media < 6 then 
   situacao="Recuperação"
else
   situacao="Aprovado com louvor" 
end if 

'Saída de Dados 
resp=msgbox ("Média do aluno: "& media &"" + vbnewline &_
    "Situação do aluno: "& situacao &"" + vbnewline &_
	"Novo cálculo?",vbinformation+vbyesno,"Rendimento do Aluno")
if resp=vbyes then 
    call entrada_notas
else 
    wscript.quit
end if
end sub 