Function EnviaEmail()

Dim linha As Integer

linha = 8

Sheets("Cadastro de Item").Unprotect "cpd*2487"
ActiveSheet.Shapes.Item("btn").Visible = False
ActiveSheet.Shapes.Item("btn_remover").Visible = False
ActiveSheet.Shapes.Item("btn_add").Visible = False

While Sheets("Cadastro de Item").Range("B" & linha).Value <> "fim"
linha = linha + 1
Wend

Range("B" & linha).Select
Selection.ClearContents
Selection.Value = ""

Criar
converterpdf

Dim data
Dim nome As String
Dim iMsg, iConf, Flds

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields



data = Now

schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
'Configura o smtp
Flds.Item(schema & "smtpserver") = "zimbramail.penso.com.br"
'Configura a porta de envio de email
Flds.Item(schema & "smtpserverport") = 25
Flds.Item(schema & "smtpauthenticate") = 1
'Configura o email do remetente
            Flds.Item(schema & "sendusername") = "teste@gmail.com.br"  'CONTA PARA ENTRAR NO SERVIDOR  IMPORTANTE !!!!
'Configura a senha do email remetente
            Flds.Item(schema & "sendpassword") = "teste22"  'SENHA PARA ENTRAR NO SERVIDOR  IMPORTANTE !!!!
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

With iMsg
   'Email do destinatário
              .To = "teste2@gmail.com.br"
   'Seu email
   .From = "Sistemas@gmail.com.br"  'REMETENTE (PODE SER INVENTADO)
   'Título do email
   .Subject = "Cadastro de Item"
   'Mensagem do e-mail, você pode enviar formatado em HTML
   .HTMLBody = "Segue como anexo o formulario para o cadastro de item, enviado pelo usuario (Terminal): " & Application.UserName & " as: " & data
   'Seu nome ou apelido
   '.Sender = "Nathalia"
   'Nome da sua organização
   '.Organization = "Empresarial"
   'e-mail de responder para
   '.ReplyTo = "teste@gmail.com.br"
   'Anexo a ser enviado na mensagem. Retire a aspa da linha abaixo e coloque o endereço do arquivo
   .AddAttachment ("c:\Cadastro de Item.xlsx")
   .AddAttachment ("c:\cadastro de item.pdf")
   Set .Configuration = iConf
   .Send
End With

Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing

Deletar

disparar

ActiveSheet.Shapes.Item("btn").Visible = True
ActiveSheet.Shapes.Item("btn_remover").Visible = True
ActiveSheet.Shapes.Item("btn_add").Visible = True
Range("B" & linha).Select
Selection.ClearContents
Selection.Value = "fim"
Sheets("Cadastro de Item").Protect "cpd*2487"
End Function

