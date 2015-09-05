<% 
Set objCDOSYSMail = Server.CreateObject("CDO.Message") 
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 

'objCDOSYSCon.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.seudominio" 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "usuario@dominio.com" 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "senha" 
objCDOSYSCon.Fields.update 

Set objCDOSYSMail.Configuration = objCDOSYSCon 

objCDOSYSMail.From = "Remetente <usuário@dominio.com>" 
objCDOSYSMail.To = "NOME <destino@dominio.com>"
'objCDOSYSMail.cc = "Copia <copia@dominio.com>" 
'objCDOSYSMail.bcc = "Copia Oculta <copiaoculta@dominio.com>" 

objCDOSYSMail.Subject = " Envio autenticado - CDOSYS Apps" 
objCDOSYSMail.TextBody = "Teste do componente CDOSYS - Texto sem HTML" 
'objCDOSYSMail.HtmlBody = "<html> <head><meta http-equiv=""Content-Type"" content=""text/html;charset=utf-8""></head><body></body></html>" 
objCDOSYSMail.Send 

Set objCDOSYSMail = Nothing 
Set objCDOSYSCon = Nothing 

response.write "Email enviado com sucesso" 
'Response.Redirect "Enviado.asp" 
%>