VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadastro 
   Caption         =   "Cadastro geral"
   ClientHeight    =   8520.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   OleObjectBlob   =   "frmCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblCadastrarContr_Click()
' Insere as informações do contrato na aba Contratos
'
    Application.ScreenUpdating = False
    
'''''''' Checagem do preenchimento de todos os campos do formulário

    If Me.txtProcessoContr.Value = "" Then
        MsgBox ("Preencha o campo de Processo")

        ElseIf Me.txtForneContr.Value = "" Then
             MsgBox ("Preencha o campo de Razão Social do fornecedor")

        ElseIf Me.txtCNPJContr.Value = "" Then
            MsgBox ("Preencha o campo de CNPJ")

        ElseIf Me.txtDtContr.Value = "" Then
            MsgBox ("Preencha o campo de Data do contrato")

        ElseIf Me.txtNumContr.Value = "" Then
            MsgBox ("Preencha o campo de Nº do contrato")

        ElseIf Me.txtVlrContr.Value = "" Then
            MsgBox ("Preencha o campo de Valor contratado")

        ElseIf Me.txtVigenciaContr.Value = "" Then
            MsgBox ("Preencha o campo de Vigência")

        ElseIf Me.cboRubricaContr.Value = "" Then
            MsgBox ("Preencha o campo de Rubrica")

        ElseIf Me.txtObjContr.Value = "" Then
            MsgBox ("Preencha o campo de Objeto de contratação")

        ElseIf Me.txtExecContr.Value = "" Then
            MsgBox ("Preencha o campo de Execução física")

        Else
        
        ' Alcance da última célula preenchida
        
            Sheets("Contratos").Select
                Range("B1").Select
                Selection.End(xlDown).Select
                Selection.End(xlDown).Select
                ActiveCell.Offset(1, 0).Select
            'O código abaixo insere as informações do formulário nas células da aba
                ActiveCell.Offset(0, 0).Value = Me.txtProcessoContr.Text
                ActiveCell.Offset(0, 1).Value = Me.txtForneContr.Text
                ActiveCell.Offset(0, 2).Value = Me.txtCNPJContr.Text
                ActiveCell.Offset(0, 3).Value = Me.txtDtContr.Text
                ActiveCell.Offset(0, 4).Value = Me.txtNumContr.Text
                ActiveCell.Offset(0, 5).Value = Me.txtVlrContr.Text
                ActiveCell.Offset(0, 9).Value = Me.txtVigenciaContr.Text
                ActiveCell.Offset(0, 10).Value = Me.txtObsContr.Text
                ActiveCell.Offset(0, 11).Value = Me.cboRubricaContr.Text
                ActiveCell.Offset(0, 12).Value = Me.txtObjContr.Text
        
    End If

    Application.ScreenUpdating = True
    
    MsgBox "Contrato cadastrado com sucesso", vbOKOnly, "Concluído"
    
End Sub

Private Sub lblCadastrarValidacao_Click()

' Insere as informações acerca do documento de liquidação na aba Despesas. Não inclui pagamento.
'
    Application.ScreenUpdating = False

'''''''' Checagem do preenchimento de todos os campos do formulário
    
    If Me.txtAnoValid.Value = "" Then
        MsgBox ("Preencha o campo de Ano de pagamento")
        
        ElseIf Me.txtProcessoValid.Value = "" Then
             MsgBox ("Preencha o campo de Processo")
        
        ElseIf Me.cboRubricaValid.Value = "" Then
            MsgBox ("Preencha o campo de Rubrica")
        
        ElseIf Me.txtNumDocValid.Value = "" Then
            MsgBox ("Preencha o campo de Nº do documento fiscal")
        
        ElseIf Me.txtDtemissaoValid.Value = "" Then
            MsgBox ("Preencha o campo de Data de emissão")
        
        ElseIf Me.txtVlrValid.Value = "" Then
            MsgBox ("Preencha o campo de Valor do documento (bruto)")
        
        ElseIf Me.txtProdutoValid.Value = "" Then
            MsgBox ("Preencha o campo de Descrição do produto pago (conforme Validação)")
    
        Else

' Alcance da última célula preenchida
        
            Sheets("Despesas").Select
                Range("B1").Select
                Selection.End(xlDown).Select
                Selection.End(xlDown).Select
                ActiveCell.Offset(1, 0).Select
            'O código abaixo insere as informações do formulário nas células da aba
                ActiveCell.Offset(0, 0).Value = Me.txtForneValid.Text
                ActiveCell.Offset(0, 1).Value = Me.txtCNPJValid.Text
                ActiveCell.Offset(0, 2).Value = Me.txtAnoValid.Text
                ActiveCell.Offset(0, 3).Value = Me.txtProcessoValid.Text
                ActiveCell.Offset(0, 4).Value = Me.cboMetaValid.Text
                ActiveCell.Offset(0, 5).Value = Me.cboEtapaValid.Text
                ActiveCell.Offset(0, 6).Value = Me.cboRubricaValid.Text
                ActiveCell.Offset(0, 8).Value = Me.txtNumDocValid.Text
                ActiveCell.Offset(0, 9).Value = Me.txtDtemissaoValid.Text
                ActiveCell.Offset(0, 10).Value = Me.txtVlrValid.Text
                ActiveCell.Offset(0, 17).Value = Me.txtProdutoValid.Text

    End If
    
    Application.ScreenUpdating = True

    MsgBox "Documento de liquidação cadastrado com sucesso", vbOKOnly, "Concluído"

End Sub

Private Sub lblProcurar_Click()

    Dim Lin As Integer
    
    Application.ScreenUpdating = False
    
    'Posiciona na linha
    Lin = 4
    'Zera a busca da NF
    Me.cboNF.Clear
    
    With Sheets("Despesas")
    'Faz a busca pelas nfs do processo
    Do Until .Range("E" & Lin).Value = ""
        
        If .Range("E" & Lin).Value = Me.txtProcesso.Text Then
            Me.cboNF.AddItem .Range("j" & Lin).Value
        End If
        
        Lin = Lin + 1
    Loop
                
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub lblGravarpag_Click()

    Dim wdesp As Worksheet
    Dim Lin As Integer
             
    Application.ScreenUpdating = False
             
    Lin = 4
        
    Set wdesp = Sheets("Despesas")
    wdesp.Select
    Do While wdesp.Range("E" & Lin).Value <> ""
        If wdesp.Range("E" & Lin).Text = Me.txtProcesso.Text Then
            If wdesp.Range("J" & Lin).Text = Me.cboNF.Text Then
                wdesp.Range("N" & Lin).Value = Me.txtComprovante.Value
                wdesp.Range("O" & Lin).Value = Me.txtDtpag.Value
                wdesp.Range("P" & Lin).Value = Me.txtValorliq.Value
            End If
        
        End If
        
        Lin = Lin + 1
                
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Comprovante inserido com sucesso!", vbOKOnly, "Processo concluído"
      
End Sub

Private Sub MultiPage1_Change()

End Sub
