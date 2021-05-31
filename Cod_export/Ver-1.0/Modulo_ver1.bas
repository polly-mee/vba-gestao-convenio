Attribute VB_Name = "Modulo_ver1"
Sub Cadastro_contrato2()
'
' Cadastro_contrato2 Macro
' Insere as informações do contrato na aba Contratos
'
    Application.ScreenUpdating = False
    
'''''''' Checagem do preenchimento de todos os campos do formulário
    
    If Sheets("Cadastro").Range("C6").Value = "" Then
        MsgBox ("Preencha o campo de Processo")
    ElseIf Sheets("Cadastro").Range("f6").Value = "" Then
         MsgBox ("Preencha o campo de Razão Social do fornecedor")
    ElseIf Sheets("Cadastro").Range("K6").Value = "" Then
        MsgBox ("Preencha o campo de CNPJ")
    ElseIf Sheets("Cadastro").Range("C14").Value = "" Then
        MsgBox ("Preencha o campo de Data do contrato")
    ElseIf Sheets("Cadastro").Range("C10").Value = "" Then
        MsgBox ("Preencha o campo de Nº do contrato")
    ElseIf Sheets("Cadastro").Range("K14").Value = "" Then
        MsgBox ("Preencha o campo de Valor contratado")
    ElseIf Sheets("Cadastro").Range("g14").Value = "" Then
        MsgBox ("Preencha o campo de Vigência")
    ElseIf Sheets("Cadastro").Range("f10").Value = "" Then
        MsgBox ("Preencha o campo de Rubrica")
    ElseIf Sheets("Cadastro").Range("i18").Value = "" Then
        MsgBox ("Preencha o campo de Objeto de contratação")
    ElseIf Sheets("Cadastro").Range("f22").Value = "" Then
        MsgBox ("Preencha o campo de Execução física")

'''''''' Códigos para salvar nas abas
    Else
    
' Início da cópia dos dados preenchidos no formulário para a aba de Contratos
    
        Range("C6:D6").Select
            Selection.Copy
            Sheets("Contratos").Select
                Range("B1").Select
                    Selection.End(xlDown).Select
                    Selection.End(xlDown).Select
                    Selection.End(xlDown).Select
                    Selection.End(xlUp).Select
                    ActiveCell.Offset(1, 0).Select
                        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("F6:I6").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("K6").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("C14:E14").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("C10:D10").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("K14").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("G14:I14").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 4).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("C18:G18").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("F10").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("I18:K18").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("F22").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Contratos").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
    
        Application.CutCopyMode = False
    
' Limpa as informações inseridas no formulário após a inserção dele na aba de Contratos
' e seleciona a primeira célula para iniciar o preenchimento
    
        Range("I18:K18,C18:G18,K14,G14:I14,C14:E14,C10:D10,K6,F6:I6,C6:D6,A56").Select
            Selection.ClearContents
        Range("C6:D6").Select
    
        Application.ScreenUpdating = True
    
' Mensagem de confirmação
    
        MsgBox "Contrato cadastrado com sucesso", vbOKOnly, "Concluído"
      
    End If
        
End Sub
Sub Cadastro_doc_liquidacao()
'
' Cadastro_doc_liquidacao Macro
' Insere as informações acerca do documento de liquidação na aba Despesas. Não inclui pagamento.
'
'''''''' Checagem do preenchimento de todos os campos do formulário
    
    If Sheets("Cadastro").Range("Y14").Value = "" Then
        MsgBox ("Preencha o campo de Ano de pagamento")
    ElseIf Sheets("Cadastro").Range("o6").Value = "" Then
         MsgBox ("Preencha o campo de Processo")
    ElseIf Sheets("Cadastro").Range("o14").Value = "" Then
        MsgBox ("Preencha o campo de Rubrica")
    ElseIf Sheets("Cadastro").Range("o18").Value = "" Then
        MsgBox ("Preencha o campo de Nº do documento fiscal")
    ElseIf Sheets("Cadastro").Range("t18").Value = "" Then
        MsgBox ("Preencha o campo de Data de emissão")
    ElseIf Sheets("Cadastro").Range("y18").Value = "" Then
        MsgBox ("Preencha o campo de Valor do documento (bruto)")
    ElseIf Sheets("Cadastro").Range("o22").Value = "" Then
        MsgBox ("Preencha o campo de Descrição do produto pago (conforme Validação)")

''''' Códigos para salvar nas abas
    Else
'
    Application.ScreenUpdating = False

' Início da cópia dos dados preenchidos no formulário para a aba de Contratos

        Range("R6:W6").Select
            Selection.Copy
            Sheets("Despesas").Select
                Range("B1").Select
                    Selection.End(xlDown).Select
                    Selection.End(xlDown).Select
                    Selection.End(xlDown).Select
                    Selection.End(xlUp).Select
                    ActiveCell.Offset(1, 0).Select
                        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("Y6:AB6").Select
            Application.CutCopyMode = False
                Selection.Copy
            Sheets("Despesas").Select
            ActiveCell.Offset(0, 1).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("Y14:z14").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("O6:P6").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("R10").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("Y10").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("O14").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("O18:R18").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 2).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("T18:W18").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("Y18:z18").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Sheets("Cadastro").Select
    ' O abaixo insere as informações de valor na coluna "Valor CH/OB"
        Range("Y18:z18").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 3).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
        Range("T22:AB22").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Despesas").Select
                ActiveCell.Offset(0, 3).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Sheets("Cadastro").Select
            
' Limpa as informações inseridas no formulário após a inserção dele na aba de Despesas
' e seleciona a primeira célula para iniciar o preenchimento
    
        Range("O6:P6,Y14:Z14,O18:R18,T18:W18,Y18:Z18,T22:AB22,O22:P22,D56,C56,B56"). _
        Select
        Selection.ClearContents
    Range("O6:P6").Select
    
        Application.ScreenUpdating = True
    
' Mensagem de confirmação
    
        MsgBox "Documento de liquidação cadastrado com sucesso", vbOKOnly, "Concluído"
        
    End If
  
End Sub
Sub Limpar_contratos()
'
' Limpar_contratos Macro
' Limpa as informações preenchidas no formulário

'
    Range("C6:D6,F6:I6,K6,C10:D10,C14:E14,G14:I14,K14,C18:G18,I18:K18,A56").Select
        Selection.ClearContents
    Range("C6:D6").Select
End Sub
Sub Limpar_doc_liq()
'
' Limpar_doc_liq Macro
' Limpa as informações preenchidas no formulário

'
    Range("O6:P6,Y14:Z14,O18:R18,T18:W18,Y18:Z18,T22:AB22,O22:P22,D56,C56,B56"). _
        Select
        Selection.ClearContents
    Range("O6:P6").Select
End Sub
