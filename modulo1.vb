Sub add_linha()
Sheets("Cadastro de Item").Unprotect "cpd*2487"


Dim linha As Integer
linha = 8

While Sheets("Cadastro de Item").Range("B" & linha).Value <> "fim"

linha = linha + 1

Wend

    If linha = 28 Then

    MsgBox "Número máximo de linhas atingido.", vbInformation, "Erro ao adicionar linhas"

    Else
    
    Range("B" & linha - 1 & ":" & "U" & linha - 1).Select
    Selection.Copy
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows(linha).RowHeight = 33
    Range("B" & linha & ":" & "U" & linha).Select
    Selection.ClearContents
    Range("B" & linha).Select
    Selection.Value = linha - 7

    Application.CutCopyMode = False
    Rows("1:" & linha + 20).EntireRow.Hidden = False

End If

Sheets("Cadastro de Item").Protect "cpd*2487"

End Sub
