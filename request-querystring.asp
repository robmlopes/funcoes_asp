<%
Dim item
For Each item In Request.QueryString
  Response.Write item
  Response.Write " = "
  Response.Write Request.QueryString(item)
  Response.Write "<br>"
Next
Set item = Nothing
%>
