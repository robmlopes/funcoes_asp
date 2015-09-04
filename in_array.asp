<%
'Autor: Robson Martins Lopes
<<<<<<< HEAD
'E-mail: robmlopes@yahoo.com.br
=======
>>>>>>> ade905160022b48d5228e01e30efe1f258f2b862
'Data.: 04/09/2015
'Objetivo: Pesquisar um item em uma lista
valor = "1"
lista = "1,2,3,4,5"

if  in_array(valor, lista) then 
<<<<<<< HEAD
	Response.write "A lista ("&replace(lista,valor,"<font color='green'><b>"&valor&"</b></font>")&") contém o item: "&valor&"<br>"
else
	Response.write "A lista ("&lista&")<br> Não Contém o item: <font color='red'><b>"&valor&"</br></font><br>"
=======
	Response.write "A lista ("&replace(lista,valor,"<font color='red'><b>"&valor&"</br></font>")&")<br> ContÃ©m o item: "&valor&"<br>"
else
	Response.write "A lista ("&lista&")<br> NÃ£o ContÃ©m o item: <font color='red'><b>"&valor&"</br></font><br>"
>>>>>>> ade905160022b48d5228e01e30efe1f258f2b862
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