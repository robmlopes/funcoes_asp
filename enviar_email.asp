<%
'Instanciar objetos
Set Mail = Server.CreateObject("CDO.Message") 'E-mail
Set MailConfig = Server.CreateObject ("CDO.Configuration") 'Configuração

'Parâmetros de configuração do e-mail
'MailConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True 'SMTP SSL
MailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Método de envio
MailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'Autenticação
MailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 'Tempo limite
MailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.seudominio" 'SMTP
MailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587 'Porta SMTP
MailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "usuario@dominio.com" 'Usuário
MailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "senha" 'Senha usuário
MailConfig.Fields.update 

'Salvar parâmetros de configuração do e-mail
Set Mail.Configuration = MailConfig

'Parâmetros do e-mail
Mail.From = "Remetente <usuário@dominio.com>" 'Remetente
Mail.To = "NOME <destino@dominio.com>" 'Destinatário
'Mail.cc = "Copia <copia@dominio.com>" 'Com cópia
'Mail.bcc = "Copia Oculta <copiaoculta@dominio.com>" 'Com cópia oculta

'Assunto do e-mail
Mail.Subject = " Envio autenticado - CDOSYS Apps"

'Usar se o corpo do e-mail for texto
Mail.TextBody = "Teste do componente CDOSYS - Texto sem HTML"

'Usar se o corpo do e-mail for HTML
Mail.HtmlBody = "<html> <head><meta http-equiv=""Content-Type"" content=""text/html;charset=utf-8""></head><body></body></html>" 
Mail.BodyPart.ContentTransferEncoding = "quoted-printable"
Mail.BodyPart.Charset = "UTF-8"

'Enviar e-mail
Mail.Send 

Set Mail = Nothing
Set MailConfig = Nothing

response.write "E-mail enviado com sucesso" 
'Response.Redirect "Enviado.asp" 
%>
