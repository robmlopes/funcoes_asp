<%
'Autor: Robson Martins Lopes
'E-mail: robmlopes@yahoo.com.br
'Data.: 04/09/2015
'Objetivo: Pesquisar um item em uma lista
valor = "1"
lista = "1,2,3,4,5"

if  in_array(valor, lista) then 
	Response.write "A lista ("&replace(lista,valor,"<font color='green'><b>"&valor&"</b></font>")&") cont�m o item: "&valor&"<br>"
else
	Response.write "A lista ("&lista&")<br> N�o Cont�m o item: <font color='red'><b>"&valor&"</br></font><br>"
end if

Function in_array(element, texto)
  in_array = False
  arr  = split(texto, ",")
  For i=0 To Ubound(arr)
     If Trim(arr(i)) = Trim(element) Then
        in_array = True
        Exit Function      
     End If
  Next
End Function
%>