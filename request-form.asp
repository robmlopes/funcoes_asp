<%
Dim item
For Each item In Request.Form
  Response.Write item
  Response.Write " = "
  Response.Write Request.Form(item)
  Response.Write "<br>"
Next
Set item = Nothing
%>
