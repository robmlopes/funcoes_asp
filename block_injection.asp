<%
'Autor: Robson Martins Lopes
'E-mail: robmlopes@yahoo.com.br
'Data.: 11/09/2015
'Objetivo: Pesquisar o valor de vindo de um request e remover uma lista de injection
function block_injection(valor, lista, delimitador, trocar)
 array_list =  split(lista, delimitador)
 for i=0 to ubound(array_list)
  valor = replace(valor, array_list(i), trocar)
 next
 block_injection = valor
end function

lista_injection = "insert,update,delete,script,sys,objects,*, from ,select ,where,order by,group by,exec ,xp_,injection,'"
valor = "Meu nome é robson, moro no parque selecta martins lopes script ' delete * from sys.objects "

response.write "valor original: "& block_injection(valor, lista_injection, ",", "") &"<br>"&"<br>"
response.write "valor sem injection: "& block_injection(valor, lista_injection, ",", "")&"<br>"
%>