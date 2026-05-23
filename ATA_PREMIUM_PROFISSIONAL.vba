Option Explicit

' ==========================================
' ATA DE REUNIÃO PREMIUM - DESIGN MODERNO
' ==========================================
' Autor: Sistema de Atas Profissional
' Versão: 2.0
' Descrição: Cria Ata de Reunião com design premium em azul/cinza
' ==========================================

Sub CriarAtaPremium()
    Dim wb As Workbook
    Dim wsAta As Worksheet
    Dim caminhoExportacao As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook
    
    ' Remove planilha antiga se existir
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("ATA DE REUNIÃO").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Cria nova planilha
    Set wsAta = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsAta.Name = "ATA DE REUNIÃO"

    ' Preenche conteúdo
    ConstruirAtaPremium wsAta

    ' Exportar PDF (opcional)
    caminhoExportacao = wb.Path
    If caminhoExportacao <> "" Then
        On Error Resume Next
        wsAta.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminhoExportacao & "\ATA_REUNIAO.pdf", _
                                  Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False
        On Error GoTo 0
    End If

    wsAta.Activate
    wsAta.Range("A1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Ata Premium criada com sucesso!", vbInformation
End Sub

' ==========================================
Sub ConstruirAtaPremium(ws As Worksheet)
' ==========================================
    
    ' === PALETA DE CORES PROFISSIONAL ===
    Dim azulEscuro As Long:      azulEscuro = RGB(13, 39, 70)        ' #0D2746
    Dim azulPrimario As Long:    azulPrimario = RGB(25, 75, 137)     ' #194B89
    Dim azulSecundario As Long:  azulSecundario = RGB(41, 117, 199)  ' #2975C7
    Dim azulClaro As Long:       azulClaro = RGB(220, 237, 254)      ' #DCEDFE
    Dim azulHover As Long:       azulHover = RGB(200, 224, 245)      ' #C8E0F5
    
    Dim cinzaEscuro As Long:     cinzaEscuro = RGB(48, 52, 59)       ' #303C3B
    Dim cinzaMedio As Long:      cinzaMedio = RGB(119, 124, 137)     ' #777C89
    Dim cinzaClaro As Long:      cinzaClaro = RGB(230, 232, 236)     ' #E6E8EC
    Dim cinzaMuitoClaro As Long: cinzaMuitoClaro = RGB(245, 246, 248)' #F5F6F8
    
    Dim branco As Long:          branco = RGB(255, 255, 255)         ' #FFFFFF
    Dim preto As Long:           preto = RGB(0, 0, 0)                ' #000000
    Dim ouro As Long:            ouro = RGB(184, 134, 11)            ' #B8860B
    Dim verdeAprovado As Long:   verdeAprovado = RGB(52, 130, 80)    ' #348250
    Dim laranjaAviso As Long:    laranjaAviso = RGB(228, 126, 25)    ' #E47E19
    Dim vermelhoAlerta As Long:  vermelhoAlerta = RGB(192, 57, 43)   ' #C0392B
    
    ' Configura colunas
    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 32
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 2.5

    ' === CABEÇALHO PRINCIPAL ===
    CriarCabecalhoPrincipal ws, azulEscuro, branco, ouro
    
    ' === SEÇÃO IDENTIFICAÇÃO ===
    CriarSecaoIdentificacao ws, azulPrimario, azulClaro, cinzaMuitoClaro, cinzaEscuro, branco
    
    ' === SEÇÃO PARTICIPANTES ===
    CriarSecaoParticipantes ws, azulSecundario, azulHover, cinzaClaro, cinzaEscuro, branco
    
    ' === SEÇÃO PAUTA ===
    CriarSecaoPauta ws, azulPrimario, azulClaro, cinzaMuitoClaro, cinzaEscuro, branco
    
    ' === SEÇÃO DISCUSSÕES ===
    CriarSecaoDiscussoes ws, azulSecundario, azulHover, cinzaClaro, verdeAprovado, laranjaAviso, vermelhoAlerta, cinzaEscuro, branco
    
    ' === SEÇÃO PENDÊNCIAS ===
    CriarSecaoPendencias ws, azulPrimario, azulClaro, cinzaMuitoClaro, verdeAprovado, laranjaAviso, vermelhoAlerta, cinzaEscuro, branco
    
    ' === SEÇÃO OBSERVAÇÕES ===
    CriarSecaoObservacoes ws, azulSecundario, azulHover, cinzaClaro, cinzaEscuro, branco
    
    ' === SEÇÃO ASSINATURAS ===
    CriarSecaoAssinaturas ws, azulPrimario, azulClaro, cinzaEscuro, branco
    
    ' === RODAPÉ ===
    CriarRodape ws, azulEscuro, cinzaMedio, branco
    
    ' Configuração de página
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .LeftMargin = 25
        .RightMargin = 25
        .TopMargin = 25
        .BottomMargin = 25
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2
    End With
    
    ActiveWindow.Zoom = 85
End Sub

' ==========================================
Sub CriarCabecalhoPrincipal(ws As Worksheet, corFundo As Long, corTexto As Long, corOuro As Long)
' ==========================================
    Dim i As Integer
    
    ' Linha de topo decorativa
    ws.Range("A1:H1").RowHeight = 8
    ws.Range("A1:H1").Interior.Color = corOuro
    
    ' Linha 2: Empresa/Logo
    ws.Range("A2:H2").RowHeight = 20
    ws.Range("A2:H2").Interior.Color = corFundo
    With ws.Range("B2:G2")
        .Merge
        .Value = "EMPRESA / CONSTRUTORA"
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Linha 3: Título Principal
    ws.Range("A3:H3").RowHeight = 50
    ws.Range("A3:H3").Interior.Color = corFundo
    With ws.Range("B3:G3")
        .Merge
        .Value = "ATA DE REUNIÃO"
        .Font.Bold = True
        .Font.Size = 24
        .Font.Color = corTexto
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Linha 4: Subtítulo
    ws.Range("A4:H4").RowHeight = 22
    ws.Range("A4:H4").Interior.Color = RGB(25, 75, 137)
    With ws.Range("B4:G4")
        .Merge
        .Value = "Reunião de Obra  |  Acompanhamento Técnico e Planejamento"
        .Font.Size = 10
        .Font.Color = RGB(200, 220, 245)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Linha separadora
    ws.Range("A5:H5").RowHeight = 6
    ws.Range("A5:H5").Interior.Color = corOuro
End Sub

' ==========================================
Sub CriarSecaoIdentificacao(ws As Worksheet, corTitulo As Long, corTituloFundo As Long, corFundoPar As Long, corTexto As Long, corBranco As Long)
' ==========================================
    Dim linhaAtual As Integer: linhaAtual = 6
    Dim i As Integer
    
    ' Título da seção
    Call CriarTituloSecao(ws, linhaAtual, "IDENTIFICAÇÃO DA REUNIÃO", corTitulo, corTituloFundo, corTexto)
    linhaAtual = linhaAtual + 1
    
    ' Dados de identificação
    Dim dados() As Variant
    ReDim dados(1 To 8, 1 To 4)
    
    dados(1, 1) = "Nº da Ata:":        dados(1, 2) = "001/2026"
    dados(1, 3) = "Data:":             dados(1, 4) = "22/05/2026 (Sexta-feira)"
    
    dados(2, 1) = "Obra/Projeto:":     dados(2, 2) = "________________________________"
    dados(2, 3) = "Hora Início:":      dados(2, 4) = "________________________________"
    
    dados(3, 1) = "Local:":            dados(3, 2) = "________________________________"
    dados(3, 3) = "Hora Término:":     dados(3, 4) = "________________________________"
    
    dados(4, 1) = "Responsável Ata:":  dados(4, 2) = "________________________________"
    dados(4, 3) = "Próxima Reunião:":  dados(4, 4) = "________________________________"
    
    dados(5, 1) = "Eng. Responsável:": dados(5, 2) = "________________________________"
    dados(5, 3) = "Contrato Nº:":      dados(5, 4) = "________________________________"
    
    dados(6, 1) = "Fiscal/Gestor:":    dados(6, 2) = "________________________________"
    dados(6, 3) = "Etapa da Obra:":    dados(6, 4) = "Fase Final"
    
    dados(7, 1) = "Cliente:":          dados(7, 2) = "________________________________"
    dados(7, 3) = "% Concluído:":      dados(7, 4) = "________________________________"
    
    dados(8, 1) = "Contratada:":       dados(8, 2) = "________________________________"
    dados(8, 3) = "Revisão:":          dados(8, 4) = "00"
    
    For i = 1 To 8
        ws.Rows(linhaAtual).RowHeight = 22
        
        ws.Cells(linhaAtual, 1).Interior.Color = corTitulo
        ws.Cells(linhaAtual, 8).Interior.Color = corTitulo
        
        ' Label (coluna B)
        ws.Cells(linhaAtual, 2).Value = dados(i, 1)
        ws.Cells(linhaAtual, 2).Font.Bold = True
        ws.Cells(linhaAtual, 2).Font.Size = 9
        ws.Cells(linhaAtual, 2).Font.Color = corTexto
        ws.Cells(linhaAtual, 2).VerticalAlignment = xlCenter
        
        ' Valor (coluna C)
        ws.Cells(linhaAtual, 3).Value = dados(i, 2)
        ws.Cells(linhaAtual, 3).Font.Size = 9
        ws.Cells(linhaAtual, 3).VerticalAlignment = xlCenter
        
        ' Label 2 (coluna E)
        ws.Cells(linhaAtual, 5).Value = dados(i, 3)
        ws.Cells(linhaAtual, 5).Font.Bold = True
        ws.Cells(linhaAtual, 5).Font.Size = 9
        ws.Cells(linhaAtual, 5).Font.Color = corTexto
        ws.Cells(linhaAtual, 5).VerticalAlignment = xlCenter
        
        ' Valor 2 (coluna F)
        ws.Cells(linhaAtual, 6).Value = dados(i, 4)
        ws.Cells(linhaAtual, 6).Font.Size = 9
        ws.Cells(linhaAtual, 6).VerticalAlignment = xlCenter
        
        ' Cores alternadas
        Dim corFundo As Long
        If i Mod 2 = 0 Then
            corFundo = corFundoPar
        Else
            corFundo = RGB(255, 255, 255)
        End If
        
        ws.Range(ws.Cells(linhaAtual, 2), ws.Cells(linhaAtual, 7)).Interior.Color = corFundo
        
        linhaAtual = linhaAtual + 1
    Next i
End Sub

' ==========================================
Sub CriarSecaoParticipantes(ws As Worksheet, corTitulo As Long, corTituloFundo As Long, corFundoPar As Long, corTexto As Long, corBranco As Long)
' ==========================================
    Dim linhaAtual As Integer: linhaAtual = ws.UsedRange.Rows.Count + 2
    Dim i As Integer, c As Integer
    
    ' Separador
    ws.Rows(linhaAtual).RowHeight = 10
    ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 8)).Interior.Color = RGB(240, 242, 245)
    linhaAtual = linhaAtual + 1
    
    ' Título
    Call CriarTituloSecao(ws, linhaAtual, "PARTICIPANTES DA REUNIÃO", corTitulo, corTituloFundo, corTexto)
    linhaAtual = linhaAtual + 1
    
    ' Cabeçalhos das colunas
    Dim cabecalhos As Variant
    cabecalhos = Array("", "NOME COMPLETO", "EMPRESA / FUNÇÃO", "CARGO", "CONTATO", "ASSINATURA", "", "")
    
    ws.Rows(linhaAtual).RowHeight = 24
    For c = 0 To 7
        With ws.Cells(linhaAtual, c + 1)
            .Value = cabecalhos(c)
            .Font.Bold = True
            .Font.Size = 9
            .Font.Color = corBranco
            .Interior.Color = corTitulo
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
    linhaAtual = linhaAtual + 1
    
    ' Linhas de participantes
    For i = 1 To 6
        ws.Rows(linhaAtual).RowHeight = 22
        
        ws.Cells(linhaAtual, 1).Interior.Color = corTitulo
        ws.Cells(linhaAtual, 8).Interior.Color = corTitulo
        
        ws.Cells(linhaAtual, 2).Value = IIf(i = 1, "Rodrigo Silva", "____________________")
        ws.Cells(linhaAtual, 3).Value = IIf(i = 1, "Setor Elétrico", "________________")
        ws.Cells(linhaAtual, 4).Value = IIf(i = 1, "Eletricista Resp.", "________________")
        ws.Cells(linhaAtual, 5).Value = "________________"
        ws.Cells(linhaAtual, 6).Value = "________________"
        
        Dim corFundo As Long
        If i Mod 2 = 0 Then
            corFundo = corFundoPar
        Else
            corFundo = corBranco
        End If
        
        For c = 2 To 7
            With ws.Cells(linhaAtual, c)
                .Font.Size = 9
                .Interior.Color = corFundo
                .VerticalAlignment = xlCenter
            End With
        Next c
        
        linhaAtual = linhaAtual + 1
    Next i
End Sub

' ==========================================
Sub CriarSecaoPauta(ws As Worksheet, corTitulo As Long, corTituloFundo As Long, corFundoPar As Long, corTexto As Long, corBranco As Long)
' ==========================================
    Dim linhaAtual As Integer: linhaAtual = ws.UsedRange.Rows.Count + 2
    Dim i As Integer
    
    ' Separador
    ws.Rows(linhaAtual).RowHeight = 10
    ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 8)).Interior.Color = RGB(240, 242, 245)
    linhaAtual = linhaAtual + 1
    
    ' Título
    Call CriarTituloSecao(ws, linhaAtual, "PAUTA DA REUNIÃO", corTitulo, corTituloFundo, corTexto)
    linhaAtual = linhaAtual + 1
    
    ' Itens de pauta
    Dim pauta() As Variant
    pauta = Array( _
        "1.  Análise do cronograma geral da obra - fase final", _
        "2.  Status das instalações elétricas - previsão de conclusão", _
        "3.  Testes de pressurização da tubulação - resultados e pendências", _
        "4.  Andamento dos isolamentos e bengalas dos evaporadores")
    
    For i = 0 To UBound(pauta)
        ws.Rows(linhaAtual).RowHeight = 20
        
        ws.Cells(linhaAtual, 1).Interior.Color = corTitulo
        ws.Cells(linhaAtual, 8).Interior.Color = corTitulo
        
        With ws.Range(ws.Cells(linhaAtual, 2), ws.Cells(linhaAtual, 7))
            .Merge
            .Value = pauta(i)
            .Font.Size = 9
            .VerticalAlignment = xlCenter
            If i Mod 2 = 0 Then
                .Interior.Color = corBranco
            Else
                .Interior.Color = corFundoPar
            End If
        End With
        
        linhaAtual = linhaAtual + 1
    Next i
End Sub

' ==========================================
Sub CriarSecaoDiscussoes(ws As Worksheet, corTitulo As Long, corTituloFundo As Long, corFundoPar As Long, _
                         corVerde As Long, corLaranja As Long, corVermelho As Long, corTexto As Long, corBranco As Long)
' ==========================================
    Dim linhaAtual As Integer: linhaAtual = ws.UsedRange.Rows.Count + 2
    Dim i As Integer, c As Integer
    
    ' Separador
    ws.Rows(linhaAtual).RowHeight = 10
    ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 8)).Interior.Color = RGB(240, 242, 245)
    linhaAtual = linhaAtual + 1
    
    ' Título
    Call CriarTituloSecao(ws, linhaAtual, "DISCUSSÕES E DECISÕES", corTitulo, corTituloFundo, corTexto)
    linhaAtual = linhaAtual + 1
    
    ' Cabeçalhos
    Dim cab As Variant
    cab = Array("", "Nº", "DISCIPLINA", "DESCRIÇÃO DETALHADA", "RESPONSÁVEL", "PRAZO", "STATUS", "")
    
    ws.Rows(linhaAtual).RowHeight = 24
    For c = 0 To 7
        With ws.Cells(linhaAtual, c + 1)
            .Value = cab(c)
            .Font.Bold = True
            .Font.Size = 9
            .Font.Color = corBranco
            .Interior.Color = corTitulo
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
    linhaAtual = linhaAtual + 1
    
    ' Dados de discussões
    Dim disc() As Variant
    ReDim disc(1 To 8, 1 To 6)
    
    disc(1, 1) = "01": disc(1, 2) = "ELÉTRICA"
    disc(1, 3) = "Rodrigo informou que as instalações elétricas estão na fase final."
    disc(1, 4) = "Rodrigo": disc(1, 5) = "25/05/2026": disc(1, 6) = "EM ANDAMENTO"
    
    disc(2, 1) = "02": disc(2, 2) = "PRESSURIZAÇÃO"
    disc(2, 3) = "Iniciados testes de pressurização. Identificados vazamentos."
    disc(2, 4) = "Eq. Mecânica": disc(2, 5) = "27/05/2026": disc(2, 6) = "EM ANDAMENTO"
    
    disc(3, 1) = "03": disc(3, 2) = "PRESSURIZAÇÃO"
    disc(3, 3) = "Prazo limite para conclusão dos testes: 27/05/2026."
    disc(3, 4) = "Eq. Mecânica": disc(3, 5) = "27/05/2026": disc(3, 6) = "PENDENTE"
    
    disc(4, 1) = "04": disc(4, 2) = "ISOLAMENTO"
    disc(4, 3) = "Isolamentos térmicos com 80% concluídos."
    disc(4, 4) = "Eq. Isolamento": disc(4, 5) = "23/05/2026": disc(4, 6) = "EM ANDAMENTO"
    
    For i = 5 To 8
        disc(i, 1) = Format(i, "00")
        disc(i, 2) = "________________________"
        disc(i, 3) = "________________________"
        disc(i, 4) = "____________"
        disc(i, 5) = "____________"
        disc(i, 6) = "____________"
    Next i
    
    For i = 1 To 8
        ws.Rows(linhaAtual).RowHeight = 48
        
        ws.Cells(linhaAtual, 1).Interior.Color = corTitulo
        ws.Cells(linhaAtual, 8).Interior.Color = corTitulo
        
        ws.Cells(linhaAtual, 2).Value = disc(i, 1)
        ws.Cells(linhaAtual, 2).HorizontalAlignment = xlCenter
        ws.Cells(linhaAtual, 2).Font.Size = 9
        
        ws.Cells(linhaAtual, 3).Value = disc(i, 2)
        ws.Cells(linhaAtual, 3).Font.Bold = True
        ws.Cells(linhaAtual, 3).Font.Size = 9
        
        ws.Cells(linhaAtual, 4).Value = disc(i, 3)
        ws.Cells(linhaAtual, 4).WrapText = True
        ws.Cells(linhaAtual, 4).Font.Size = 9
        
        ws.Cells(linhaAtual, 5).Value = disc(i, 4)
        ws.Cells(linhaAtual, 5).HorizontalAlignment = xlCenter
        ws.Cells(linhaAtual, 5).Font.Size = 9
        
        ws.Cells(linhaAtual, 6).Value = disc(i, 5)
        ws.Cells(linhaAtual, 6).HorizontalAlignment = xlCenter
        ws.Cells(linhaAtual, 6).Font.Size = 9
        
        ws.Cells(linhaAtual, 7).Value = disc(i, 6)
        ws.Cells(linhaAtual, 7).HorizontalAlignment = xlCenter
        ws.Cells(linhaAtual, 7).Font.Bold = True
        ws.Cells(linhaAtual, 7).Font.Size = 9
        
        Select Case disc(i, 6)
            Case "EM ANDAMENTO"
                ws.Cells(linhaAtual, 7).Interior.Color = corLaranja
                ws.Cells(linhaAtual, 7).Font.Color = corBranco
            Case "PENDENTE"
                ws.Cells(linhaAtual, 7).Interior.Color = RGB(255, 224, 178)
                ws.Cells(linhaAtual, 7).Font.Color = RGB(150, 60, 0)
            Case "CONCLUÍDO"
                ws.Cells(linhaAtual, 7).Interior.Color = corVerde
                ws.Cells(linhaAtual, 7).Font.Color = corBranco
            Case "ATRASADO"
                ws.Cells(linhaAtual, 7).Interior.Color = corVermelho
                ws.Cells(linhaAtual, 7).Font.Color = corBranco
        End Select
        
        Dim corFundo As Long
        If i Mod 2 = 0 Then
            corFundo = corFundoPar
        Else
            corFundo = corBranco
        End If
        
        For c = 2 To 6
            ws.Cells(linhaAtual, c).Interior.Color = corFundo
            ws.Cells(linhaAtual, c).VerticalAlignment = xlCenter
        Next c
        
        linhaAtual = linhaAtual + 1
    Next i
End Sub

' ==========================================
Sub CriarSecaoPendencias(ws As Worksheet, corTitulo As Long, corTituloFundo As Long, corFundoPar As Long, _
                         corVerde As Long, corLaranja As Long, corVermelho As Long, corTexto As Long, corBranco As Long)
' ==========================================
    Dim linhaAtual As Integer: linhaAtual = ws.UsedRange.Rows.Count + 2
    Dim i As Integer, c As Integer
    
    ' Separador
    ws.Rows(linhaAtual).RowHeight = 10
    ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 8)).Interior.Color = RGB(240, 242, 245)
    linhaAtual = linhaAtual + 1
    
    ' Título
    Call CriarTituloSecao(ws, linhaAtual, "PENDÊNCIAS E AÇÕES DEFINIDAS", corTitulo, corTituloFundo, corTexto)
    linhaAtual = linhaAtual + 1
    
    ' Cabeçalhos
    Dim cab As Variant
    cab = Array("", "ITEM", "AÇÃO / PENDÊNCIA", "RESPONSÁVEL", "PRAZO", "PRIORIDADE", "STATUS", "")
    
    ws.Rows(linhaAtual).RowHeight = 24
    For c = 0 To 7
        With ws.Cells(linhaAtual, c + 1)
            .Value = cab(c)
            .Font.Bold = True
            .Font.Size = 9
            .Font.Color = corBranco
            .Interior.Color = corTitulo
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
    linhaAtual = linhaAtual + 1
    
    ' Dados
    Dim pend() As Variant
    ReDim pend(1 To 8, 1 To 5)
    
    pend(1, 1) = "P-01": pend(1, 2) = "Concluir instalações elétricas"
    pend(1, 3) = "Rodrigo": pend(1, 4) = "25/05/2026": pend(1, 5) = "ALTA"
    
    pend(2, 1) = "P-02": pend(2, 2) = "Corrigir vazamentos nos testes de pressurização"
    pend(2, 3) = "Eq. Mecânica": pend(2, 4) = "23/05/2026": pend(2, 5) = "CRÍTICA"
    
    pend(3, 1) = "P-03": pend(3, 2) = "Aprovação final dos testes de pressurização"
    pend(3, 3) = "Eq. Mecânica": pend(3, 4) = "27/05/2026": pend(3, 5) = "ALTA"
    
    For i = 4 To 8
        pend(i, 1) = Format(i, "00")
        pend(i, 2) = "________________________"
        pend(i, 3) = "____________"
        pend(i, 4) = "____________"
        pend(i, 5) = "______"
    Next i
    
    For i = 1 To 8
        ws.Rows(linhaAtual).RowHeight = 22
        
        ws.Cells(linhaAtual, 1).Interior.Color = corTitulo
        ws.Cells(linhaAtual, 8).Interior.Color = corTitulo
        
        ws.Cells(linhaAtual, 2).Value = pend(i, 1)
        ws.Cells(linhaAtual, 2).HorizontalAlignment = xlCenter
        ws.Cells(linhaAtual, 2).Font.Size = 9
        
        ws.Cells(linhaAtual, 3).Value = pend(i, 2)
        ws.Cells(linhaAtual, 3).Font.Size = 9
        
        ws.Cells(linhaAtual, 4).Value = pend(i, 3)
        ws.Cells(linhaAtual, 4).HorizontalAlignment = xlCenter
        ws.Cells(linhaAtual, 4).Font.Size = 9
        
        ws.Cells(linhaAtual, 5).Value = pend(i, 4)
        ws.Cells(linhaAtual, 5).HorizontalAlignment = xlCenter
        ws.Cells(linhaAtual, 5).Font.Size = 9
        
        ws.Cells(linhaAtual, 6).Value = pend(i, 5)
        ws.Cells(linhaAtual, 6).HorizontalAlignment = xlCenter
        ws.Cells(linhaAtual, 6).Font.Bold = True
        ws.Cells(linhaAtual, 6).Font.Size = 9
        
        Select Case pend(i, 5)
            Case "CRÍTICA"
                ws.Cells(linhaAtual, 6).Interior.Color = corVermelho
                ws.Cells(linhaAtual, 6).Font.Color = corBranco
            Case "ALTA"
                ws.Cells(linhaAtual, 6).Interior.Color = corLaranja
                ws.Cells(linhaAtual, 6).Font.Color = corBranco
            Case "MÉDIA"
                ws.Cells(linhaAtual, 6).Interior.Color = RGB(255, 248, 225)
                ws.Cells(linhaAtual, 6).Font.Color = RGB(130, 80, 0)
        End Select
        
        ws.Cells(linhaAtual, 7).Value = "PENDENTE"
        ws.Cells(linhaAtual, 7).HorizontalAlignment = xlCenter
        ws.Cells(linhaAtual, 7).Font.Bold = True
        ws.Cells(linhaAtual, 7).Font.Size = 9
        ws.Cells(linhaAtual, 7).Interior.Color = RGB(255, 224, 178)
        ws.Cells(linhaAtual, 7).Font.Color = RGB(150, 60, 0)
        
        Dim corFundo As Long
        If i Mod 2 = 0 Then
            corFundo = corFundoPar
        Else
            corFundo = corBranco
        End If
        
        For c = 2 To 5
            ws.Cells(linhaAtual, c).Interior.Color = corFundo
            ws.Cells(linhaAtual, c).VerticalAlignment = xlCenter
        Next c
        
        linhaAtual = linhaAtual + 1
    Next i
End Sub

' ==========================================
Sub CriarSecaoObservacoes(ws As Worksheet, corTitulo As Long, corTituloFundo As Long, corFundoPar As Long, corTexto As Long, corBranco As Long)
' ==========================================
    Dim linhaAtual As Integer: linhaAtual = ws.UsedRange.Rows.Count + 2
    Dim i As Integer
    
    ' Separador
    ws.Rows(linhaAtual).RowHeight = 10
    ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 8)).Interior.Color = RGB(240, 242, 245)
    linhaAtual = linhaAtual + 1
    
    ' Título
    Call CriarTituloSecao(ws, linhaAtual, "OBSERVAÇÕES GERAIS", corTitulo, corTituloFundo, corTexto)
    linhaAtual = linhaAtual + 1
    
    ' Observações
    Dim obs() As Variant
    obs = Array( _
        "• Testes de pressurização devem ser documentados com relatório fotográfico e registro de pressão.", _
        "• Confirmar com Rodrigo lista completa de pontos elétricos pendentes para acompanhamento diário.", _
        "• Verificar disponibilidade de material para tubulação da área externa antes do início.", _
        "• ________________________________________________________________________________________", _
        "• ________________________________________________________________________________________", _
        "• ________________________________________________________________________________________", _
        "• ________________________________________________________________________________________", _
        "• ________________________________________________________________________________________")
    
    For i = 0 To UBound(obs)
        ws.Rows(linhaAtual).RowHeight = 20
        
        ws.Cells(linhaAtual, 1).Interior.Color = corTitulo
        ws.Cells(linhaAtual, 8).Interior.Color = corTitulo
        
        With ws.Range(ws.Cells(linhaAtual, 2), ws.Cells(linhaAtual, 7))
            .Merge
            .Value = obs(i)
            .Font.Size = 9
            .WrapText = True
            .VerticalAlignment = xlCenter
            If i Mod 2 = 0 Then
                .Interior.Color = corBranco
            Else
                .Interior.Color = corFundoPar
            End If
        End With
        
        linhaAtual = linhaAtual + 1
    Next i
End Sub

' ==========================================
Sub CriarSecaoAssinaturas(ws As Worksheet, corTitulo As Long, corTituloFundo As Long, corTexto As Long, corBranco As Long)
' ==========================================
    Dim linhaAtual As Integer: linhaAtual = ws.UsedRange.Rows.Count + 2
    Dim i As Integer
    
    ' Separador
    ws.Rows(linhaAtual).RowHeight = 10
    ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 8)).Interior.Color = RGB(240, 242, 245)
    linhaAtual = linhaAtual + 1
    
    ' Título
    Call CriarTituloSecao(ws, linhaAtual, "APROVAÇÃO E ASSINATURAS", corTitulo, corTituloFundo, corTexto)
    linhaAtual = linhaAtual + 1
    
    ' Espaço de assinatura
    ws.Rows(linhaAtual).RowHeight = 50
    ws.Cells(linhaAtual, 1).Interior.Color = corTitulo
    ws.Cells(linhaAtual, 8).Interior.Color = corTitulo
    linhaAtual = linhaAtual + 1
    
    ' Cargos de assinatura
    Dim cargos() As Variant
    cargos = Array("Engenheiro Responsável", "Gestor / Fiscal de Obra", "Responsável pela Ata")
    
    Dim colAssinatura() As Integer
    ReDim colAssinatura(0 To 2)
    colAssinatura(0) = 2: colAssinatura(1) = 4: colAssinatura(2) = 6
    
    For i = 0 To 2
        ws.Rows(linhaAtual).RowHeight = 18
        ws.Cells(linhaAtual, 1).Interior.Color = corTitulo
        ws.Cells(linhaAtual, 8).Interior.Color = corTitulo
        
        With ws.Range(ws.Cells(linhaAtual, colAssinatura(i)), ws.Cells(linhaAtual, colAssinatura(i) + 1))
            .Merge
            .Value = cargos(i)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 9
            .Font.Color = corTitulo
            .Interior.Color = corTituloFundo
        End With
    Next i
End Sub

' ==========================================
Sub CriarRodape(ws As Worksheet, corFundo As Long, corTexto As Long, corBranco As Long)
' ==========================================
    Dim linhaAtual As Integer: linhaAtual = ws.UsedRange.Rows.Count + 2
    
    ' Separador
    ws.Rows(linhaAtual).RowHeight = 8
    ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 8)).Interior.Color = RGB(184, 134, 11)
    linhaAtual = linhaAtual + 1
    
    ' Rodapé
    ws.Rows(linhaAtual).RowHeight = 18
    ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 8)).Interior.Color = corFundo
    
    With ws.Range(ws.Cells(linhaAtual, 2), ws.Cells(linhaAtual, 7))
        .Merge
        .Value = "Documento gerado em " & Format(Date, "dd/mm/yyyy") & "  |  Ata Nº 001/2026  |  Versão 2.0  |  CONFIDENCIAL"
        .Font.Size = 8
        .Font.Color = corTexto
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' ==========================================
Sub CriarTituloSecao(ws As Worksheet, linha As Integer, titulo As String, corFundo As Long, corFundoTitulo As Long, corTexto As Long)
' ==========================================
    ws.Rows(linha).RowHeight = 24
    
    ws.Cells(linha, 1).Interior.Color = corFundo
    ws.Cells(linha, 8).Interior.Color = corFundo
    
    With ws.Range(ws.Cells(linha, 2), ws.Cells(linha, 7))
        .Merge
        .Value = "■  " & titulo
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = corTexto
        .Interior.Color = corFundoTitulo
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub
