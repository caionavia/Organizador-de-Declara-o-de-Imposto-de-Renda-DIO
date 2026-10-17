Attribute VB_Name = "ModAdicionarBanco"
Option Explicit

' ============================================================
' Botao "ADICIONAR BANCO" - aba INFORMES
' Replica o ultimo bloco (titulo, BANCO, VALOR ATUAL, ANEXO),
' incrementa o numero do titulo, limpa a coluna D (mantendo
' formatacao e validacao) e reposiciona o botao abaixo.
' ============================================================
Sub AdicionarBanco()

    ' ---- Configuracao ----
    Const NOME_ABA As String = "INFORMES"
    Const NOME_BOTAO_PADRAO As String = "Rectangle 8"   ' usado so se rodar fora do botao (F5)
    Const ATUALIZAR_TOTAL As Boolean = True             ' False = nao mexe na formula do TOTAL
    Const CEL_TOTAL As String = "C8"
    ' -----------------------

    Dim ws As Worksheet
    Dim shp As Shape
    Dim ultLinha As Long, iniBloco As Long
    Dim novoIni As Long, novoFim As Long
    Dim txt As String, novoTxt As String
    Dim nDig As Long
    Dim offBotao As Double
    Dim f As String, refNova As String

    Set ws = ThisWorkbook.Worksheets(NOME_ABA)

    ' Localiza o botao (o que foi clicado; se nao houver, pelo nome padrao)
    On Error Resume Next
    Set shp = ws.Shapes(Application.Caller)
    If shp Is Nothing Then Set shp = ws.Shapes(NOME_BOTAO_PADRAO)
    On Error GoTo 0

    ' Ultimo bloco = ultima linha usada na coluna C (e as 3 anteriores)
    ultLinha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    iniBloco = ultLinha - 3
    If iniBloco < 1 Then GoTo Falha

    ' Titulo do bloco precisa comecar com numero (ex.: "4o Banco")
    txt = CStr(ws.Cells(iniBloco, "C").Value)
    nDig = 0
    Do While nDig < Len(txt)
        If Mid$(txt, nDig + 1, 1) Like "#" Then
            nDig = nDig + 1
        Else
            Exit Do
        End If
    Loop
    If nDig = 0 Then GoTo Falha
    novoTxt = CStr(CLng(Left$(txt, nDig)) + 1) & Mid$(txt, nDig + 1)

    ' Posicao do botao em relacao a primeira linha abaixo da tabela
    If Not shp Is Nothing Then offBotao = shp.Top - ws.Rows(ultLinha + 1).Top

    Application.ScreenUpdating = False
    On Error GoTo Erro

    ' Copia linha em branco + bloco (5 linhas) para logo abaixo
    ws.Unprotect
    ws.Rows((ultLinha - 4) & ":" & ultLinha).Copy Destination:=ws.Rows(ultLinha + 1)
    Application.CutCopyMode = False

    novoIni = ultLinha + 2          ' titulo do novo bloco
    novoFim = ultLinha + 5          ' ultima linha do novo bloco

    ' Coluna C: so o titulo muda (4o -> 5o)
    ws.Cells(novoIni, "C").Value = novoTxt

    ' Coluna D: remove conteudo, mantem formatacao
    With ws.Range(ws.Cells(novoIni, "D"), ws.Cells(novoFim, "D"))
        .Hyperlinks.Delete
        .ClearContents
    End With

    ' Inclui o novo VALOR ATUAL na formula do TOTAL
    If ATUALIZAR_TOTAL Then
        f = ws.Range(CEL_TOTAL).Formula
        refNova = "D" & (novoIni + 2)
        If Right$(f, 1) = ")" And InStr(f, refNova) = 0 Then
            ws.Range(CEL_TOTAL).Formula = Left$(f, Len(f) - 1) & "," & refNova & ")"
        End If
    End If

    ' Reposiciona o botao abaixo da nova tabela (mesma distancia de antes)
    If Not shp Is Nothing Then shp.Top = ws.Rows(novoFim + 1).Top + offBotao
    ws.Protect

    Application.ScreenUpdating = True
    Exit Sub

Falha:
    MsgBox "Nao foi possivel identificar o ultimo bloco de banco na coluna C." & vbCrLf & _
           "Esperado: titulo iniciando com numero (ex.: 4o Banco) 3 linhas acima da ultima linha usada.", _
           vbExclamation, "Adicionar banco"
    Exit Sub

Erro:
    Application.CutCopyMode = False
    On Error Resume Next
    ws.Protect
    Application.ScreenUpdating = True
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Adicionar banco"
End Sub
