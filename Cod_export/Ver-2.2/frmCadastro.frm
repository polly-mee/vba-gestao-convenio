VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadastro 
   Caption         =   "Cadastro geral"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7545
   OleObjectBlob   =   "frmCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

Dim lin As Integer
Dim senha As String
Dim resultado As VbMsgBoxResult
Dim planilha As Worksheet

    Application.ScreenUpdating = False
    
    ''Desbloqueia as planilhas para edi��o

    'Contador para refer�ncia de linha
    lin = 2
    'Zera a busca das comboboxes
    Me.cboMetaValid.Clear
    Me.cboEtapaValid.Clear
    Me.cboRubricaValid.Clear
    
    With Sheets("Listas")
    'Faz a busca pelas Metas, Etapas e Rubricas do processo conforme a quantidade de rubricas
    Do Until .Range("A" & lin).Value = ""
        
        Me.cboMetaValid.AddItem .Range("E" & lin).Value
        Me.cboEtapaValid.AddItem .Range("C" & lin).Value
        Me.cboRubricaValid.AddItem .Range("A" & lin).Value
        Me.cboRubricaContr.AddItem .Range("A" & lin).Value
        
        lin = lin + 1
    Loop
                
    End With
  
  'Solicita��o de senha para desbloqueio de planilha
        
resposta_sim:
    senha = InputBox("Digite a senha para desbloquear a planilha:", "Desbloquear a planilha")
    
    For Each planilha In Sheets
        On Error GoTo trata_erro
        
        planilha.Unprotect Password:=senha
    Next
    
    Application.ScreenUpdating = True

Exit Sub

resposta_nao:
    MsgBox "As planilhas n�o foram desbloqueadas. Senha inv�lida.", vbOKOnly, "Processo conclu�do"
    Exit Sub

trata_erro:

    resultado = MsgBox("Senha inv�lida. Tentar novamente?", vbYesNo, "Senha inv�lida")
    
    If resultado = vbYes Then
        Resume resposta_sim
        Else
            Resume resposta_nao
    End If

End Sub

Private Sub spbExecContr_Change()
    ''Spinbutton do cadastro de CONTRATO
    Application.ScreenUpdating = False
    
    'Definindo m�nimo e m�ximo
    Me.spbExecContr.Min = 1
    Me.spbExecContr.Max = 1000
    
    'Velocidade
    Me.spbExecContr.Delay = 20
    
    'Atribuindo � textbox
    Me.txtExecContr.Value = Me.spbExecContr.Value
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub lblCadastrarContr_Click()
' Insere as informa��es do processo na aba Contratos
'
    Dim Descricao As String
    Dim lin As Integer
    Dim Check As VbMsgBoxResult
    
    Application.ScreenUpdating = False
    
'''''''' Checagem do preenchimento de todos os campos do formul�rio

    If Me.txtProcessoContr.Value = "" Then
        MsgBox ("Preencha o campo de Processo")

        ElseIf Me.txtForneContr.Value = "" Then
             MsgBox ("Preencha o campo de Raz�o Social do fornecedor")

        ElseIf Me.txtCNPJContr.Value = "" Then
            MsgBox ("Preencha o campo de CNPJ")

        ElseIf Me.txtVlrContr.Value = "" Then
            MsgBox ("Preencha o campo de Valor contratado")

        ElseIf Me.cboRubricaContr.Value = "" Then
            MsgBox ("Preencha o campo de Rubrica")

        Else
        
        ' Alcance da �ltima c�lula preenchida
        
            Sheets("Contratos").Select
                Range("B1").Select
                Selection.End(xlDown).Select
                Selection.End(xlDown).Select
                ActiveCell.Offset(1, 0).Select
            'O c�digo abaixo insere as informa��es do formul�rio nas c�lulas da aba
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
                ActiveCell.Offset(0, 13).Value = Me.txtExecContr.Value
        
            'Mensagem de confirma��o
            
            MsgBox "Contrato cadastrado com sucesso", vbOKOnly, "Conclu�do"
    End If

    'Limpar caixas de texto para nova inser��o sem fechar a caixa de formul�rio
    
    Me.txtProcessoContr = Empty
    Me.txtForneContr = Empty
    Me.txtCNPJContr = Empty
    Me.txtDtContr = Empty
    Me.txtNumContr = Empty
    Me.txtVlrContr = Empty
    Me.txtVigenciaContr = Empty
    Me.txtObsContr = Empty
    Me.cboRubricaContr = Empty
    Me.txtObjContr = Empty
    Me.txtExecContr = Empty

    Application.ScreenUpdating = True
            
End Sub

Private Sub txtProcessoValid_Change()

Dim linpr As Integer

    Application.ScreenUpdating = False
 ''''''''Busca as informa��es automaticamente de Raz�o social, CNPJ e contrato da aba de contratos
    'Contador para refer�ncia de linha
    linpr = 4
                   
    With Sheets("Contratos")
    'Faz a busca pelo processo
        Do Until .Range("B" & linpr).Value = ""
            
            If .Range("B" & linpr).Value = Me.txtProcessoValid.Text Then
                Me.txtCNPJValid.Value = .Range("D" & linpr).Value
                Me.txtForneValid.Value = .Range("C" & linpr).Value
                Me.txtContrValid.Value = .Range("F" & linpr).Value
            End If
            
            linpr = linpr + 1
        Loop
                
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub lblCadastrarValidacao_Click()

' Insere as informa��es acerca do documento de liquida��o na aba Despesas. N�o inclui pagamento.

    Application.ScreenUpdating = False

'''''''' Checagem do preenchimento de todos os campos do formul�rio
    
    If Me.txtAnoValid.Value = "" Then
        MsgBox ("Preencha o campo de Ano de pagamento")
        
        ElseIf Me.txtProcessoValid.Value = "" Then
             MsgBox ("Preencha o campo de Processo")
        
        ElseIf Me.cboRubricaValid.Value = "" Then
            MsgBox ("Preencha o campo de Rubrica")
        
        ElseIf Me.txtNumDocValid.Value = "" Then
            MsgBox ("Preencha o campo de N� do documento fiscal")
        
        ElseIf Me.txtDtemissaoValid.Value = "" Then
            MsgBox ("Preencha o campo de Data de emiss�o")
        
        ElseIf Me.txtVlrValid.Value = "" Then
            MsgBox ("Preencha o campo de Valor do documento (bruto)")
        
        ElseIf Me.txtProdutoValid.Value = "" Then
            MsgBox ("Preencha o campo de Descri��o do produto pago (conforme Valida��o)")
    
        Else
        
            ' Alcance da �ltima c�lula preenchida
            Sheets("Despesas").Select
                Range("B1").Select
                Selection.End(xlDown).Select
                Selection.End(xlDown).Select
                ActiveCell.Offset(1, 0).Select
            'O c�digo abaixo insere as informa��es do formul�rio nas c�lulas da aba
                ActiveCell.Offset(0, 0).Value = Me.txtForneValid.Text
                ActiveCell.Offset(0, 1).Value = Me.txtCNPJValid.Text
                ActiveCell.Offset(0, 2).Value = Me.txtAnoValid.Text
                ActiveCell.Offset(0, 3).Value = Me.txtProcessoValid.Text
                ActiveCell.Offset(0, 4).Value = Me.cboMetaValid.Value
                ActiveCell.Offset(0, 5).Value = Me.cboEtapaValid.Value
                ActiveCell.Offset(0, 6).Value = Me.cboRubricaValid.Value
                ActiveCell.Offset(0, 8).Value = Me.txtNumDocValid.Text
                ActiveCell.Offset(0, 9).Value = Me.txtDtemissaoValid.Text
                ActiveCell.Offset(0, 10).Value = Me.txtVlrValid.Text
                ActiveCell.Offset(0, 14).Value = Me.txtVlrValid.Text
                ActiveCell.Offset(0, 17).Value = Me.txtProdutoValid.Text
            
            'Mensagem de confirma��o
            MsgBox "Documento de liquida��o cadastrado com sucesso", vbOKOnly, "Conclu�do"

    End If
    
    'Limpar caixas de texto para nova inser��o sem fechar a caixa de formul�rio
    
    Me.txtForneValid = Empty
    Me.txtCNPJValid = Empty
    Me.txtAnoValid = Empty
    Me.txtProcessoValid = Empty
    Me.cboMetaValid = Empty
    Me.cboEtapaValid = Empty
    Me.cboRubricaValid = Empty
    Me.txtNumDocValid = Empty
    Me.txtDtemissaoValid = Empty
    Me.txtVlrValid = Empty
    Me.txtProdutoValid = Empty
    Me.txtContrValid = Empty
            
    Application.ScreenUpdating = True

End Sub

Private Sub lblProcurar_Click()

    'Bot�o para procurar as nfs do processo descrito
    
    Dim linpr As Integer
    
    Application.ScreenUpdating = False
    
    'Contador para refer�ncia de linha
    linpr = 4
    'Zera a busca da NF
    Me.cboNFComp.Clear
    
    With Sheets("Despesas")
    'Faz a busca pelas nfs do processo
    Do Until .Range("E" & linpr).Value = ""
        
        If .Range("E" & linpr).Value = Me.txtProcessoComp.Text Then
            Me.cboNFComp.AddItem .Range("J" & linpr).Value
        End If
        
        linpr = linpr + 1
    Loop
                
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub lblGravarpag_Click()

    'Registra as informa��es de Comprovante na aba de despesas de acordo com a nf solicitada

    Dim wdesp As Worksheet
    Dim Lingr As Integer
             
    Application.ScreenUpdating = False
    'Contador para refer�ncia de linha
    Lingr = 4
        
    Set wdesp = Sheets("Despesas")
    wdesp.Select
    
    'In�cio do registro das informa��es na aba
    
    Do While wdesp.Range("E" & Lingr).Value <> ""
        If wdesp.Range("E" & Lingr).Text = Me.txtProcessoComp.Text Then
            If wdesp.Range("J" & Lingr).Text = Me.cboNFComp.Text Then
                wdesp.Range("N" & Lingr).Value = Me.txtComprovanteComp.Value
                wdesp.Range("O" & Lingr).Value = Me.txtDtpagComp.Value
                wdesp.Range("P" & Lingr).Value = Me.txtValorliqComp.Value
            End If
        
        End If
        
        Lingr = Lingr + 1
                
    Loop
    
    Me.txtProcesso = Empty
    Me.cboNFComp = Empty
    Me.txtDtpagComp = Empty
    Me.txtValorliqComp = Empty
    Me.txtComprovanteComp = Empty
    
    Application.ScreenUpdating = True
    
    MsgBox "Comprovante inserido com sucesso!", vbOKOnly, "Processo conclu�do"
      
End Sub

Private Sub UserForm_Terminate()

'Bloqueia as abas ao fechar o formul�rio de cadastro

Dim senha As String
Dim planilha As Worksheet

    Application.ScreenUpdating = False
    
    'Solicita��o da senha para o usu�rio
    senha = InputBox("Digite a senha para bloquear a planilha:", "Bloquear a planilha")
    
    'Bloqueio de cada aba da planilha
    
    For Each planilha In Sheets
        planilha.Protect Password:=senha, AllowFiltering:=True
    Next
    
    Application.ScreenUpdating = True

End Sub
