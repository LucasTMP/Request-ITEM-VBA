Sub deleta_linha()

Sheets("Cadastro de Item").Unprotect "cpd*2487"


Dim linha As Integer
linha = 8

While Sheets("Cadastro de Item").Range("B" & linha).Value <> "fim"

linha = linha + 1

Wend

    If linha = 12 Then

    MsgBox "Número mínimo de linhas atingido.", vbInformation, "Erro ao remover linhas"

    Else
    
    Rows(linha - 1).Delete

    End If

    Rows("1:" & linha + 20).EntireRow.Hidden = False
    
Sheets("Cadastro de Item").Protect "cpd*2487"

End Sub
