Attribute VB_Name = "modExportadorCDR"
Option Explicit

'==============================================================================
' EXPORTADOR DE ARQUIVOS CDR  -  CorelDRAW VBA
'------------------------------------------------------------------------------
' Percorre uma pasta (e TODAS as subpastas), localiza os arquivos .CDR e os
' exporta para outra pasta, reproduzindo exatamente a mesma estrutura de pastas.
'
' REGRA DE OURO: o .CDR original NUNCA e aberto, alterado ou salvo.
' Cada arquivo e COPIADO para uma pasta temporaria; a copia e que e aberta,
' (opcionalmente) tem os textos convertidos em curvas, e e exportada.
' Depois a copia e fechada sem salvar e apagada.
'
' Para iniciar, execute a macro:  ExportadorCDR
'==============================================================================

' ----- Formatos de exportacao (usados tambem pelo UserForm) -----
Public Const FMT_JPEG As Long = 1
Public Const FMT_PNG As Long = 2
Public Const FMT_PDF As Long = 3

' ----- Ajustes que podem ser alterados facilmente -----
Private Const JPEG_DPI As Long = 300
Private Const PNG_DPI As Long = 96                       ' padrao WEB
Private Const PNG_FUNDO_TRANSPARENTE As Boolean = True   ' padrao WEB (PNG-24)

' 0 = NAO forcar o tipo de imagem do JPEG (comportamento padrao do Corel).
' Se precisar forcar, troque o 0 por cdrRGBColorImage ou cdrCMYKColorImage.
Private Const JPEG_TIPO_IMAGEM_FORCADO As Long = 0

' ----- Comportamento quando o destino ja existe -----
Private Const MODO_IGNORAR As Long = 0
Private Const MODO_SUBSTITUIR As Long = 1

' ----- Resultado da execucao -----
Public Type ResultadoExportacao
    Encontrados As Long
    Exportados As Long
    Ignorados As Long
    Erros As Long
    Cancelado As Boolean
    ListaErros As String
    ListaAvisos As String
End Type

Private mFSO As Object          ' FileSystemObject
Private mPastaTemp As String    ' pasta temporaria desta execucao (com "\" final)
Private mContadorTemp As Long


'==============================================================================
' PONTO DE ENTRADA
'==============================================================================
Public Sub ExportadorCDR()
    frmExportador.Show
End Sub


'==============================================================================
' PROCESSAMENTO PRINCIPAL
' Retorna True se o processamento rodou (mesmo com erros em arquivos isolados).
' Retorna False se foi interrompido antes de comecar (nenhum .cdr, usuario
' cancelou a pergunta de substituicao, erro inesperado).
'==============================================================================
Public Function ProcessarPasta(ByVal pastaOrigem As String, ByVal pastaDestino As String, _
        ByVal formato As Long, ByVal paginasSeparadas As Boolean, _
        ByVal converterCurvas As Boolean, ByVal frm As Object, _
        ByRef res As ResultadoExportacao) As Boolean

    Dim vazio As ResultadoExportacao
    Dim origem As String, destino As String
    Dim arquivos As Collection
    Dim i As Long, total As Long
    Dim caminho As String, rel As String
    Dim pastaSaida As String, nomeBase As String
    Dim existentes As Long, modo As Long
    Dim resp As VbMsgBoxResult
    Dim msgErro As String, aviso As String
    Dim docAtivo As Document
    Dim otimizado As Boolean

    On Error GoTo FalhaGeral

    res = vazio
    Set mFSO = CreateObject("Scripting.FileSystemObject")

    origem = NormalizarPasta(mFSO.GetFolder(pastaOrigem).Path)
    destino = NormalizarPasta(pastaDestino)

    '--- 1) BUSCA RECURSIVA de todos os .cdr ---------------------------------
    Set arquivos = New Collection
    ColetarCDR mFSO.GetFolder(origem), arquivos, destino, res
    total = arquivos.Count
    res.Encontrados = total

    If total = 0 Then
        MsgBox "Nenhum arquivo .CDR foi encontrado na pasta de origem" & vbCrLf & _
               "(nem nas suas subpastas)." & _
               IIf(Len(res.ListaAvisos) > 0, vbCrLf & vbCrLf & "Avisos:" & vbCrLf & res.ListaAvisos, ""), _
               vbInformation, "Exportador de CDR"
        Exit Function
    End If

    '--- 2) DESTINOS JA EXISTENTES: pergunta UMA vez, antes de comecar -------
    For i = 1 To total
        caminho = arquivos(i)
        CalcularSaida caminho, origem, destino, pastaSaida, nomeBase
        If SaidaJaExiste(pastaSaida, nomeBase, formato, paginasSeparadas) Then
            existentes = existentes + 1
        End If
    Next i

    modo = MODO_SUBSTITUIR
    If existentes > 0 Then
        resp = MsgBox("Já existem exportações no destino para " & existentes & _
                      " de " & total & " arquivo(s)." & vbCrLf & vbCrLf & _
                      "SIM  = Substituir os existentes" & vbCrLf & _
                      "NAO  = Ignorar os existentes (manter como estão)" & vbCrLf & _
                      "CANCELAR = Não exportar nada", _
                      vbQuestion + vbYesNoCancel + vbDefaultButton2, "Arquivos já existentes")
        Select Case resp
            Case vbYes: modo = MODO_SUBSTITUIR
            Case vbNo: modo = MODO_IGNORAR
            Case Else: Exit Function
        End Select
    End If

    '--- 3) PREPARACAO ---------------------------------------------------------
    On Error Resume Next
    Set docAtivo = ActiveDocument            ' para reativar no final
    On Error GoTo FalhaGeral

    CriarPastaRecursiva destino
    CriarPastaTemp
    Application.Optimization = True
    otimizado = True

    '--- 4) LOOP PRINCIPAL -----------------------------------------------------
    For i = 1 To total

        If frm.CancelamentoSolicitado Then
            res.Cancelado = True
            Exit For
        End If

        caminho = arquivos(i)
        rel = Mid$(caminho, Len(origem) + 1)           ' caminho relativo
        CalcularSaida caminho, origem, destino, pastaSaida, nomeBase

        frm.AtualizarProgresso i - 1, total, _
            "Processando " & i & " de " & total & ": " & Abreviar(mFSO.GetFileName(caminho), 55)

        If modo = MODO_IGNORAR And SaidaJaExiste(pastaSaida, nomeBase, formato, paginasSeparadas) Then
            res.Ignorados = res.Ignorados + 1
        Else
            msgErro = "": aviso = ""
            If ExportarArquivo(caminho, pastaSaida, nomeBase, formato, _
                               paginasSeparadas, converterCurvas, msgErro, aviso) Then
                res.Exportados = res.Exportados + 1
                If Len(aviso) > 0 Then
                    res.ListaAvisos = res.ListaAvisos & "- " & rel & ": " & LinhaUnica(aviso) & vbCrLf
                End If
            Else
                res.Erros = res.Erros + 1
                res.ListaErros = res.ListaErros & "- " & rel & ": " & LinhaUnica(msgErro) & vbCrLf
            End If
        End If
    Next i

    frm.AtualizarProgresso res.Exportados + res.Ignorados + res.Erros, total, _
        IIf(res.Cancelado, "Cancelado pelo usuário.", "Concluído.")

    ProcessarPasta = True

Finalizar:
    On Error Resume Next
    LimparPastaTemp
    If otimizado Then Application.Optimization = False
    If Not docAtivo Is Nothing Then docAtivo.Activate
    Application.Refresh
    Exit Function

FalhaGeral:
    ProcessarPasta = False
    MsgBox "Erro inesperado durante o processamento:" & vbCrLf & vbCrLf & _
           Err.Description & " (erro " & Err.Number & ")", vbCritical, "Exportador de CDR"
    Resume Finalizar
End Function


'==============================================================================
' EXPORTA UM UNICO .CDR  (copia -> abre copia -> [curvas] -> exporta -> descarta)
'==============================================================================
Private Function ExportarArquivo(ByVal caminhoCDR As String, ByVal pastaSaida As String, _
        ByVal nomeBase As String, ByVal formato As Long, ByVal paginasSeparadas As Boolean, _
        ByVal converterCurvas As Boolean, ByRef msgErro As String, ByRef aviso As String) As Boolean

    Dim copiaTemp As String
    Dim doc As Document
    Dim sucesso As Boolean

    On Error GoTo Falha

    CriarPastaRecursiva pastaSaida                       ' cria a estrutura no destino
    copiaTemp = CriarCopiaTemporaria(caminhoCDR)         ' o original nunca e aberto
    Set doc = Application.OpenDocument(copiaTemp)

    ' Seguranca: confirma que o documento aberto e MESMO a copia temporaria
    If StrComp(doc.FileName, mFSO.GetFileName(copiaTemp), vbTextCompare) <> 0 Then
        Set doc = Nothing    ' nao mexe em nenhum outro documento
        Err.Raise vbObjectError + 1001, , "O CorelDRAW retornou um documento inesperado. " & _
                                          "Operação abortada por segurança."
    End If

    If converterCurvas Then ConverterTextosEmCurvas doc, aviso
    ExportarDocumento doc, pastaSaida, nomeBase, formato, paginasSeparadas, aviso
    sucesso = True

Limpeza:
    On Error Resume Next
    FecharDescartando doc, mFSO.GetFileName(copiaTemp)
    ApagarArquivo copiaTemp
    ExportarArquivo = sucesso
    Exit Function

Falha:
    msgErro = Err.Description & " (erro " & Err.Number & ")"
    sucesso = False
    Resume Limpeza
End Function


'==============================================================================
' EXPORTACAO (documento inteiro ou pagina a pagina)
'==============================================================================
Private Sub ExportarDocumento(ByVal doc As Document, ByVal pastaSaida As String, _
        ByVal nomeBase As String, ByVal formato As Long, _
        ByVal paginasSeparadas As Boolean, ByRef aviso As String)

    Dim ext As String, arq As String, subpasta As String
    Dim i As Long

    ext = ExtensaoFormato(formato)

    If paginasSeparadas Then
        ' Subpasta com o nome base do arquivo + um arquivo por pagina: nome_1-1, nome_2-2...
        subpasta = pastaSaida & nomeBase & "\"
        CriarPastaRecursiva subpasta
        For i = 1 To doc.Pages.Count
            arq = subpasta & nomeBase & "_" & i & "-" & i & ext
            ExportarPagina doc, i, arq, formato
        Next i
    Else
        arq = pastaSaida & nomeBase & ext
        If formato = FMT_PDF Then
            ConfigurarIntervaloPDF doc, 0            ' documento inteiro (padrao atual)
            doc.PublishToPDF arq
        Else
            ' JPEG/PNG so comportam 1 pagina por arquivo
            If doc.Pages.Count > 1 Then
                aviso = "documento com " & doc.Pages.Count & " páginas: em " & _
                        UCase$(Mid$(ext, 2)) & " somente a 1ª página foi exportada " & _
                        "(use 'Salvar cada página como um arquivo separado')."
            End If
            ExportarPagina doc, 1, arq, formato
        End If
    End If
End Sub

Private Sub ExportarPagina(ByVal doc As Document, ByVal pagina As Long, _
        ByVal caminho As String, ByVal formato As Long)

    If formato = FMT_PDF Then
        ConfigurarIntervaloPDF doc, pagina
        doc.PublishToPDF caminho
    Else
        doc.Pages(pagina).Activate
        ExportarBitmap doc, caminho, formato
    End If
End Sub

' Define quais paginas o PublishToPDF vai gerar (pagina = 0 -> documento inteiro)
Private Sub ConfigurarIntervaloPDF(ByVal doc As Document, ByVal pagina As Long)
    If pagina = 0 Then
        On Error Resume Next        ' no modo "inteiro" o padrao do Corel ja serve
        doc.PDFSettings.PublishRange = pdfWholeDocument
        On Error GoTo 0
    Else
        With doc.PDFSettings        ' aqui NAO ignoramos erros: nao pode sair PDF errado
            .PublishRange = pdfPageRange
            .PageRange = CStr(pagina)
        End With
    End If
End Sub

' JPEG (300 DPI, modo de cor padrao do Corel) e PNG (padrao WEB) da pagina ATIVA
Private Sub ExportarBitmap(ByVal doc As Document, ByVal caminho As String, ByVal formato As Long)
    Dim opts As StructExportOptions
    Dim filtro As cdrFilter

    Set opts = CreateStructExportOptions
    With opts
        .Overwrite = True      ' a decisao de substituir ja foi tomada antes
        .AntiAliasingType = cdrNormalAntiAliasing
        If formato = FMT_PNG Then
            filtro = cdrPNG
            .ResolutionX = PNG_DPI
            .ResolutionY = PNG_DPI
            .ImageType = cdrRGBColorImage
            .Transparent = PNG_FUNDO_TRANSPARENTE
        Else
            filtro = cdrJPEG
            .ResolutionX = JPEG_DPI
            .ResolutionY = JPEG_DPI
            ' ImageType NAO e definido (a menos que a constante la em cima seja alterada)
            If JPEG_TIPO_IMAGEM_FORCADO <> 0 Then .ImageType = JPEG_TIPO_IMAGEM_FORCADO
        End If
    End With

    ' Export e um comando (Sub): nao retorna objeto e nao precisa de Finish
    doc.Export caminho, filtro, cdrCurrentPage, opts
End Sub


'==============================================================================
' TEXTOS EM CURVAS  (sempre na COPIA temporaria)
'==============================================================================
Private Sub ConverterTextosEmCurvas(ByVal doc As Document, ByRef aviso As String)
    Dim pg As Page
    Dim lyr As Layer

    ' Todas as paginas, incluindo as camadas da pagina mestre (AllLayers)
    For Each pg In doc.Pages
        For Each lyr In pg.AllLayers
            ConverterCamada lyr
        Next lyr
    Next pg

    ' Verificacao: se sobrou texto, avisa no relatorio
    If ContarTextosDocumento(doc) > 0 Then
        aviso = aviso & "alguns textos não puderam ser convertidos em curvas (objetos travados?). "
    End If
End Sub

Private Sub ConverterCamada(ByVal lyr As Layer)
    On Error Resume Next
    lyr.Editable = True           ' camada travada (so na copia temporaria)
    ConverterShapes lyr.Shapes
End Sub

' Converte os textos de um conjunto de shapes, inclusive dentro de PowerClips
Private Sub ConverterShapes(ByVal shps As Shapes)
    Dim sr As ShapeRange
    Dim s As Shape

    On Error Resume Next

    Set sr = shps.FindShapes(Type:=cdrTextShape)
    If Not sr Is Nothing Then
        If sr.Count > 0 Then sr.ConvertToCurves
    End If

    Set sr = shps.FindShapes
    If Not sr Is Nothing Then
        For Each s In sr
            If Not s.PowerClip Is Nothing Then ConverterShapes s.PowerClip.Shapes
        Next s
    End If
End Sub

Private Function ContarTextosDocumento(ByVal doc As Document) As Long
    Dim pg As Page
    Dim lyr As Layer
    Dim n As Long

    On Error Resume Next
    For Each pg In doc.Pages
        For Each lyr In pg.AllLayers
            n = n + ContarTextosShapes(lyr.Shapes)
        Next lyr
    Next pg
    ContarTextosDocumento = n
End Function

Private Function ContarTextosShapes(ByVal shps As Shapes) As Long
    Dim sr As ShapeRange
    Dim s As Shape
    Dim n As Long

    On Error Resume Next

    Set sr = shps.FindShapes(Type:=cdrTextShape)
    If Not sr Is Nothing Then n = sr.Count

    Set sr = shps.FindShapes
    If Not sr Is Nothing Then
        For Each s In sr
            If Not s.PowerClip Is Nothing Then n = n + ContarTextosShapes(s.PowerClip.Shapes)
        Next s
    End If
    ContarTextosShapes = n
End Function


'==============================================================================
' DOCUMENTOS / ARQUIVOS TEMPORARIOS
'==============================================================================
' Fecha o documento SEM salvar. So fecha se for a copia temporaria.
Private Sub FecharDescartando(ByRef doc As Document, ByVal nomeTemp As String)
    On Error Resume Next
    If doc Is Nothing Then Exit Sub
    If StrComp(doc.FileName, nomeTemp, vbTextCompare) = 0 Then
        doc.Dirty = False          ' evita o aviso "salvar alteracoes?"
        doc.Close
    End If
    Set doc = Nothing
End Sub

Private Sub CriarPastaTemp()
    Dim p As String
    p = NormalizarPasta(mFSO.GetSpecialFolder(2).Path) & "ExportadorCDR_" & Format$(Now, "yyyymmdd_hhnnss")
    mFSO.CreateFolder p
    mPastaTemp = p & "\"
    mContadorTemp = 0
End Sub

Private Function CriarCopiaTemporaria(ByVal origem As String) As String
    Dim copia As String
    mContadorTemp = mContadorTemp + 1
    copia = mPastaTemp & "doc_" & Format$(mContadorTemp, "0000") & ".cdr"
    mFSO.CopyFile origem, copia, True
    On Error Resume Next
    mFSO.GetFile(copia).Attributes = 0     ' remove "somente leitura" herdado
    CriarCopiaTemporaria = copia
End Function

Private Sub ApagarArquivo(ByVal caminho As String)
    On Error Resume Next
    If Len(caminho) > 0 Then mFSO.DeleteFile caminho, True
End Sub

Private Sub LimparPastaTemp()
    On Error Resume Next
    If Len(mPastaTemp) > 0 Then
        If mFSO.FolderExists(mPastaTemp) Then mFSO.DeleteFolder SemBarraFinal(mPastaTemp), True
    End If
    mPastaTemp = ""
End Sub


'==============================================================================
' BUSCA RECURSIVA E CAMINHOS
'==============================================================================
' Entra em TODAS as subpastas, em qualquer nivel. Ignora a pasta de destino
' (caso esteja dentro da origem) e pastas internas do Windows.
Private Sub ColetarCDR(ByVal pasta As Object, ByVal lista As Collection, _
        ByVal destino As String, ByRef res As ResultadoExportacao)

    Dim arq As Object
    Dim subpasta As Object

    On Error GoTo Falha

    For Each arq In pasta.Files
        If StrComp(mFSO.GetExtensionName(arq.Name), "cdr", vbTextCompare) = 0 Then
            lista.Add arq.Path
        End If
    Next arq

    For Each subpasta In pasta.SubFolders
        If StrComp(NormalizarPasta(subpasta.Path), destino, vbTextCompare) <> 0 Then
            Select Case LCase$(subpasta.Name)
                Case "$recycle.bin", "system volume information"
                    ' pastas internas do Windows: ignoradas
                Case Else
                    ColetarCDR subpasta, lista, destino, res     ' RECURSAO
            End Select
        End If
    Next subpasta
    Exit Sub

Falha:
    res.ListaAvisos = res.ListaAvisos & "- Pasta inacessível (ignorada): " & pasta.Path & vbCrLf
End Sub

' Calcula a pasta de saida e o nome base, mantendo o caminho RELATIVO.
' Ex.: origem  G:\Arquivos\   arquivo G:\Arquivos\Clientes\Joao\Cartao.cdr
'      destino G:\Exportados\ -> pastaSaida G:\Exportados\Clientes\Joao\  nomeBase Cartao
Private Sub CalcularSaida(ByVal caminho As String, ByVal origem As String, _
        ByVal destino As String, ByRef pastaSaida As String, ByRef nomeBase As String)

    Dim rel As String
    Dim p As Long

    rel = Mid$(caminho, Len(origem) + 1)          ' "Clientes\Joao\Cartao.cdr"
    p = InStrRev(rel, "\")
    If p > 0 Then
        pastaSaida = destino & Left$(rel, p)      ' "...\Exportados\Clientes\Joao\"
    Else
        pastaSaida = destino
    End If
    nomeBase = mFSO.GetBaseName(caminho)
End Sub

Private Function SaidaJaExiste(ByVal pastaSaida As String, ByVal nomeBase As String, _
        ByVal formato As Long, ByVal paginasSeparadas As Boolean) As Boolean

    Dim ext As String, subpasta As String, prefixo As String
    Dim f As Object

    ext = ExtensaoFormato(formato)

    If paginasSeparadas Then
        subpasta = pastaSaida & nomeBase
        If mFSO.FolderExists(subpasta) Then
            prefixo = LCase$(nomeBase & "_")
            For Each f In mFSO.GetFolder(subpasta).Files
                If LCase$(Left$(f.Name, Len(prefixo))) = prefixo Then
                    If StrComp(Right$(f.Name, Len(ext)), ext, vbTextCompare) = 0 Then
                        SaidaJaExiste = True
                        Exit Function
                    End If
                End If
            Next f
        End If
    Else
        SaidaJaExiste = mFSO.FileExists(pastaSaida & nomeBase & ext)
    End If
End Function

' Cria a pasta e todas as pastas-pai que faltarem (CreateFolder do FSO nao e recursivo)
Private Sub CriarPastaRecursiva(ByVal caminho As String)
    Dim pai As String
    caminho = SemBarraFinal(caminho)
    If Len(caminho) = 0 Then Exit Sub
    If mFSO.FolderExists(caminho) Then Exit Sub
    pai = mFSO.GetParentFolderName(caminho)
    If Len(pai) > 0 Then
        If Not mFSO.FolderExists(pai) Then CriarPastaRecursiva pai
    End If
    mFSO.CreateFolder caminho
End Sub

Private Function ExtensaoFormato(ByVal formato As Long) As String
    Select Case formato
        Case FMT_JPEG: ExtensaoFormato = ".jpg"
        Case FMT_PNG: ExtensaoFormato = ".png"
        Case Else: ExtensaoFormato = ".pdf"
    End Select
End Function

Private Function NormalizarPasta(ByVal p As String) As String
    p = Trim$(p)
    If Len(p) > 0 Then
        If Right$(p, 1) <> "\" Then p = p & "\"
    End If
    NormalizarPasta = p
End Function

Private Function SemBarraFinal(ByVal p As String) As String
    p = Trim$(p)
    If Len(p) > 3 And Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    SemBarraFinal = p
End Function

Private Function Abreviar(ByVal texto As String, ByVal n As Long) As String
    If Len(texto) > n Then Abreviar = Left$(texto, n - 3) & "..." Else Abreviar = texto
End Function

Private Function LinhaUnica(ByVal texto As String) As String
    texto = Replace(texto, vbCrLf, " ")
    texto = Replace(texto, vbCr, " ")
    texto = Replace(texto, vbLf, " ")
    LinhaUnica = Trim$(texto)
End Function


'==============================================================================
' RELATORIO FINAL
'==============================================================================
Public Sub MostrarRelatorio(ByRef res As ResultadoExportacao, ByVal pastaDestino As String)
    Dim msg As String
    Dim log As String
    Dim estilo As VbMsgBoxStyle

    If res.Cancelado Then
        msg = "Exportação cancelada pelo usuário."
    Else
        msg = "Exportação concluída!"
    End If

    msg = msg & vbCrLf & vbCrLf & _
          "Arquivos encontrados: " & res.Encontrados & vbCrLf & _
          "Arquivos exportados: " & res.Exportados & vbCrLf
    If res.Ignorados > 0 Then
        msg = msg & "Arquivos ignorados (já existiam): " & res.Ignorados & vbCrLf
    End If
    msg = msg & "Arquivos com erro: " & res.Erros

    If res.Erros > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Arquivos com erro:" & vbCrLf & ResumirLista(res.ListaErros, 8)
    End If
    If Len(res.ListaAvisos) > 0 Then
        msg = msg & vbCrLf & "Avisos:" & vbCrLf & ResumirLista(res.ListaAvisos, 5)
    End If

    If res.Erros > 0 Or Len(res.ListaAvisos) > 0 Then
        log = SalvarLog(res, pastaDestino)
        If Len(log) > 0 Then
            msg = msg & vbCrLf & "Relatório detalhado salvo em:" & vbCrLf & log
        End If
    End If

    If res.Erros > 0 Or res.Cancelado Or Len(res.ListaAvisos) > 0 Then
        estilo = vbExclamation
    Else
        estilo = vbInformation
    End If
    MsgBox msg, estilo, "Exportador de CDR"
End Sub

Private Function ResumirLista(ByVal lista As String, ByVal maximo As Long) As String
    Dim linhas() As String
    Dim i As Long, n As Long
    Dim saida As String

    linhas = Split(lista, vbCrLf)
    For i = LBound(linhas) To UBound(linhas)
        If Len(linhas(i)) > 0 Then
            n = n + 1
            If n <= maximo Then saida = saida & linhas(i) & vbCrLf
        End If
    Next i
    If n > maximo Then
        saida = saida & "... e mais " & (n - maximo) & " (veja o relatório detalhado)." & vbCrLf
    End If
    ResumirLista = saida
End Function

' Grava o relatorio completo (somente quando ha erros ou avisos)
Private Function SalvarLog(ByRef res As ResultadoExportacao, ByVal pastaDestino As String) As String
    Dim caminho As String
    Dim ts As Object

    On Error GoTo Falha

    caminho = NormalizarPasta(pastaDestino) & "Relatorio_Exportacao_" & _
              Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    Set ts = mFSO.CreateTextFile(caminho, True, True)     ' Unicode
    ts.WriteLine "Relatório de exportação - " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    ts.WriteLine String$(60, "-")
    ts.WriteLine "Arquivos encontrados: " & res.Encontrados
    ts.WriteLine "Arquivos exportados: " & res.Exportados
    ts.WriteLine "Arquivos ignorados: " & res.Ignorados
    ts.WriteLine "Arquivos com erro: " & res.Erros
    If res.Cancelado Then ts.WriteLine "(Processo cancelado pelo usuário)"
    If Len(res.ListaErros) > 0 Then
        ts.WriteLine ""
        ts.WriteLine "ERROS:"
        ts.Write res.ListaErros
    End If
    If Len(res.ListaAvisos) > 0 Then
        ts.WriteLine ""
        ts.WriteLine "AVISOS:"
        ts.Write res.ListaAvisos
    End If
    ts.Close
    SalvarLog = caminho
    Exit Function

Falha:
    SalvarLog = ""
End Function
