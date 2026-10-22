Option Explicit

'==============================================================================
' EXPORTADOR DE ARQUIVOS CDR  -  CorelDRAW VBA
'------------------------------------------------------------------------------

' ----- Formatos de exportacao (usados tambem pelo UserForm) -----
Public Const FMT_JPEG As Long = 1
Public Const FMT_PNG As Long = 2
Public Const FMT_PDF As Long = 3

' ----- Ajustes que podem ser alterados facilmente -----
Private Const JPEG_DPI As Long = 300
Private Const PNG_DPI As Long = 300
Private Const PNG_FUNDO_TRANSPARENTE As Boolean = True   ' padrao WEB (PNG-24)

' Modo de cor do JPEG. Sem forcar, o Corel exportava em RGB mesmo com documento CMYK.
' cdrCMYKColorImage = CMYK (documentos de grafica);  cdrRGBColorImage = RGB;
' 0 = nao forcar (comportamento padrao do Corel, que saiu em RGB).
Private Const JPEG_TIPO_IMAGEM_FORCADO As Long = cdrCMYKColorImage

' Qualidade do JPEG. No Corel o ajuste interno e "compressao" (0 = menor compressao
' = qualidade maxima, 100% na tela). Se o resultado sair com qualidade errada,
' troque este valor (ex.: 100) e teste.
Private Const JPEG_COMPRESSAO As Long = 0

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
    ' --- campos novos (grupos / documento aberto) ---
    ModoDocAberto As Boolean
    UsouGrupos As Boolean
    NomeDocumento As String
    ItensGerados As Long
    ItensFalhos As Long
    ListaFalhasItens As String
End Type

' Um grupo exportavel (grupo de nivel superior de uma camada de uma pagina)
Private Type ItemGrupo
    pagina As Long
    Indice As Long          ' numero do grupo dentro da pagina (1..n)
    nome As String          ' nome dado pelo usuario (ja sanitizado) ou vazio
    Forma As Shape
End Type

' Margem (em mm) somada em volta da arte ao exportar GRUPOS em PDF. 0 = tamanho exato.
Private Const PDF_MARGEM_MM As Double = 0

Private mFSO As Object          ' FileSystemObject
Private mPastaTemp As String    ' pasta temporaria desta execucao (com "\" final)
Private mContadorTemp As Long
Private mFrm As Object          ' formulario (progresso / cancelamento)
Private mUsados As Object       ' Dictionary: caminhos de saida ja usados nesta execucao
Private mItensGerados As Long
Private mItensFalhos As Long
Private mListaFalhas As String
Private mRotuloDoc As String    ' arquivo/documento em processamento (para mensagens)
Private mProgAtual As Long
Private mProgTotal As Long


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
        ByVal gruposSeparados As Boolean, _
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
    res.UsouGrupos = gruposSeparados
    Set mFSO = CreateObject("Scripting.FileSystemObject")
    IniciarExecucao frm

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
        If SaidaJaExiste(pastaSaida, nomeBase, formato, paginasSeparadas, gruposSeparados) Then
            existentes = existentes + 1
        End If
    Next i

    modo = MODO_SUBSTITUIR
    If existentes > 0 Then
        resp = MsgBox("Jï¿½ existem exportaï¿½ï¿½es no destino para " & existentes & _
                      " de " & total & " arquivo(s)." & vbCrLf & vbCrLf & _
                      "SIM  = Substituir os existentes" & vbCrLf & _
                      "NAO  = Ignorar os existentes (manter como estï¿½o)" & vbCrLf & _
                      "CANCELAR = Nï¿½o exportar nada", _
                      vbQuestion + vbYesNoCancel + vbDefaultButton2, "Arquivos jï¿½ existentes")
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
        mRotuloDoc = rel
        mProgAtual = i - 1
        mProgTotal = total
        CalcularSaida caminho, origem, destino, pastaSaida, nomeBase

        frm.AtualizarProgresso i - 1, total, _
            "Processando " & i & " de " & total & ": " & Abreviar(mFSO.GetFileName(caminho), 55)

        If modo = MODO_IGNORAR And SaidaJaExiste(pastaSaida, nomeBase, formato, paginasSeparadas, gruposSeparados) Then
            res.Ignorados = res.Ignorados + 1
        Else
            msgErro = "": aviso = ""
            If ExportarArquivo(caminho, pastaSaida, nomeBase, formato, _
                               paginasSeparadas, gruposSeparados, converterCurvas, msgErro, aviso) Then
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
        IIf(res.Cancelado, "Cancelado pelo usuï¿½rio.", "Concluï¿½do.")

    res.ItensGerados = mItensGerados
    res.ItensFalhos = mItensFalhos
    res.ListaFalhasItens = mListaFalhas

    ProcessarPasta = True

Finalizar:
    On Error Resume Next
    LimparPastaTemp
    If otimizado Then Application.Optimization = False
    If Not docAtivo Is Nothing Then docAtivo.Activate
    Application.Refresh
    Set mFrm = Nothing
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
        ByVal gruposSeparados As Boolean, _
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
                                          "Operaï¿½ï¿½o abortada por seguranï¿½a."
    End If

    If converterCurvas Then ConverterTextosEmCurvas doc, aviso
    ' (os textos ja foram convertidos na copia acima; os documentos temporarios
    '  dos grupos em PDF herdam as curvas, entao nao precisam converter de novo)
    ExportarDocumento doc, pastaSaida, nomeBase, formato, paginasSeparadas, _
                      gruposSeparados, False, False, aviso
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
        ByVal paginasSeparadas As Boolean, ByVal gruposSeparados As Boolean, _
        ByVal curvasNoTemporario As Boolean, ByVal tolerante As Boolean, _
        ByRef aviso As String)

    Dim ext As String, arq As String, subpasta As String, msg As String
    Dim i As Long

    ' NOVO: um arquivo por grupo (percorre TODAS as paginas; combina com "paginas")
    If gruposSeparados Then
        ExportarGruposDocumento doc, pastaSaida, nomeBase, formato, curvasNoTemporario, aviso
        Exit Sub
    End If

    ext = ExtensaoFormato(formato)

    If paginasSeparadas Then
        ' Subpasta com o nome base do arquivo + um arquivo por pagina: nome_1-1, nome_2-2...
        subpasta = pastaSaida & nomeBase & "\"
        CriarPastaRecursiva subpasta
        For i = 1 To doc.Pages.Count
            arq = subpasta & nomeBase & "_" & i & "-" & i & ext
            If tolerante Then
                ' modo documento aberto: uma pagina com erro nao derruba as demais
                If mFrm.CancelamentoSolicitado Then Exit For
                msg = ""
                If ExportarPaginaSegura(doc, i, arq, formato, msg) Then
                    mItensGerados = mItensGerados + 1
                Else
                    RegistrarFalhaItem "Pï¿½gina " & i & " (" & mFSO.GetFileName(arq) & ")", msg
                End If
            Else
                ExportarPagina doc, i, arq, formato
                mItensGerados = mItensGerados + 1
            End If
        Next i
    Else
        arq = pastaSaida & nomeBase & ext
        If formato = FMT_PDF Then
            ConfigurarIntervaloPDF doc, 0            ' documento inteiro (padrao atual)
            doc.PublishToPDF arq
        Else
            ' JPEG/PNG so comportam 1 pagina por arquivo
            If doc.Pages.Count > 1 Then
                aviso = "documento com " & doc.Pages.Count & " pï¿½ginas: em " & _
                        UCase$(Mid$(ext, 2)) & " somente a 1ï¿½ pï¿½gina foi exportada " & _
                        "(use 'Salvar cada pï¿½gina como um arquivo separado')."
            End If
            ExportarPagina doc, 1, arq, formato
        End If
        mItensGerados = mItensGerados + 1
    End If
End Sub

Private Function ExportarPaginaSegura(ByVal doc As Document, ByVal pagina As Long, _
        ByVal caminho As String, ByVal formato As Long, ByRef msg As String) As Boolean
    On Error GoTo Falha
    ExportarPagina doc, pagina, caminho, formato
    ExportarPaginaSegura = True
    Exit Function
Falha:
    msg = Err.Description & " (erro " & Err.Number & ")"
End Function

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
Private Sub ExportarBitmap(ByVal doc As Document, ByVal caminho As String, ByVal formato As Long, _
        Optional ByVal somenteSelecao As Boolean = False)
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
            ' modo de cor: definido por JPEG_TIPO_IMAGEM_FORCADO (CMYK)
            If JPEG_TIPO_IMAGEM_FORCADO <> 0 Then .ImageType = JPEG_TIPO_IMAGEM_FORCADO
        End If
    End With

    ' "Manter tamanho": preserva as proporcoes/tamanho original da arte
    On Error Resume Next
    CallByName opts, "MaintainAspect", VbLet, True
    Err.Clear
    On Error GoTo 0

    ' somenteSelecao = exporta apenas o que esta selecionado (o Corel recorta a area
    ' nos limites da selecao, incluindo contorno/efeitos); senao, a pagina ativa

    ' Chamada original (comprovadamente funcional nesta versao do Corel)
    If somenteSelecao Then
        doc.Export caminho, filtro, cdrSelection, opts
    Else
        doc.Export caminho, filtro, cdrCurrentPage, opts
    End If

    ' Confere: o arquivo precisa existir, senao nao pode contar como exportado
    If Not mFSO.FileExists(caminho) Then
        Err.Raise vbObjectError + 1040, , "O CorelDRAW nï¿½o gravou o arquivo: " & mFSO.GetFileName(caminho)
    End If
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
        aviso = aviso & "alguns textos nï¿½o puderam ser convertidos em curvas (objetos travados?). "
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
    res.ListaAvisos = res.ListaAvisos & "- Pasta inacessï¿½vel (ignorada): " & pasta.Path & vbCrLf
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
        ByVal formato As Long, ByVal paginasSeparadas As Boolean, _
        ByVal gruposSeparados As Boolean) As Boolean

    Dim ext As String, subpasta As String, prefixo As String
    Dim f As Object

    ext = ExtensaoFormato(formato)

    If paginasSeparadas Or gruposSeparados Then
        subpasta = pastaSaida & nomeBase
        If mFSO.FolderExists(subpasta) Then
            prefixo = LCase$(nomeBase & IIf(gruposSeparados, "_Pagina-", "_"))
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
    Dim falhas As Boolean

    falhas = (res.Erros > 0 Or res.ItensFalhos > 0)

    If res.Cancelado Then
        msg = "Exportaï¿½ï¿½o cancelada pelo usuï¿½rio."
    ElseIf res.ItensFalhos > 0 Then
        msg = "Exportaï¿½ï¿½o concluï¿½da com falhas."
    Else
        msg = "Exportaï¿½ï¿½o concluï¿½da!"
    End If

    If res.ModoDocAberto Then
        msg = msg & vbCrLf & vbCrLf & _
              "Documento: " & res.NomeDocumento & vbCrLf & _
              "Arquivos gerados: " & res.ItensGerados & vbCrLf & _
              "Itens com falha: " & res.ItensFalhos
        If res.Erros > 0 Then
            msg = msg & vbCrLf & vbCrLf & "Erro ao processar o documento:" & vbCrLf & _
                  ResumirLista(res.ListaErros, 8)
        End If
    Else
        msg = msg & vbCrLf & vbCrLf & _
              "Arquivos encontrados: " & res.Encontrados & vbCrLf & _
              "Arquivos exportados: " & res.Exportados & vbCrLf
        If res.Ignorados > 0 Then
            msg = msg & "Arquivos ignorados (jï¿½ existiam): " & res.Ignorados & vbCrLf
        End If
        msg = msg & "Arquivos com erro: " & res.Erros
        If res.UsouGrupos Then
            msg = msg & vbCrLf & "Arquivos gerados (por grupo): " & res.ItensGerados & vbCrLf & _
                  "Grupos/itens com falha: " & res.ItensFalhos
        End If

        If res.Erros > 0 Then
            msg = msg & vbCrLf & vbCrLf & "Arquivos com erro:" & vbCrLf & ResumirLista(res.ListaErros, 8)
        End If
    End If

    If res.ItensFalhos > 0 Then
        msg = msg & vbCrLf & "Itens que falharam:" & vbCrLf & ResumirLista(res.ListaFalhasItens, 8)
    End If
    If Len(res.ListaAvisos) > 0 Then
        msg = msg & vbCrLf & "Avisos:" & vbCrLf & ResumirLista(res.ListaAvisos, 5)
    End If

    If falhas Or Len(res.ListaAvisos) > 0 Then
        log = SalvarLog(res, pastaDestino)
        If Len(log) > 0 Then
            msg = msg & vbCrLf & "Relatï¿½rio detalhado salvo em:" & vbCrLf & log
        End If
    End If

    If falhas Or res.Cancelado Or Len(res.ListaAvisos) > 0 Then
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
        saida = saida & "... e mais " & (n - maximo) & " (veja o relatï¿½rio detalhado)." & vbCrLf
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
    ts.WriteLine "Relatï¿½rio de exportaï¿½ï¿½o - " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    ts.WriteLine String$(60, "-")
    ts.WriteLine "Arquivos encontrados: " & res.Encontrados
    ts.WriteLine "Arquivos exportados: " & res.Exportados
    ts.WriteLine "Arquivos ignorados: " & res.Ignorados
    ts.WriteLine "Arquivos com erro: " & res.Erros
    If res.ModoDocAberto Then ts.WriteLine "Documento aberto: " & res.NomeDocumento
    If res.ModoDocAberto Or res.UsouGrupos Then
        ts.WriteLine "Arquivos gerados (itens): " & res.ItensGerados
        ts.WriteLine "Itens com falha: " & res.ItensFalhos
    End If
    If res.Cancelado Then ts.WriteLine "(Processo cancelado pelo usuï¿½rio)"
    If Len(res.ListaErros) > 0 Then
        ts.WriteLine ""
        ts.WriteLine "ERROS:"
        ts.Write res.ListaErros
    End If
    If Len(res.ListaFalhasItens) > 0 Then
        ts.WriteLine ""
        ts.WriteLine "ITENS COM FALHA:"
        ts.Write res.ListaFalhasItens
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

'==============================================================================
' ESTADO DA EXECUCAO / FALHAS POR ITEM
'==============================================================================
Private Sub IniciarExecucao(ByVal frm As Object)
    Set mFrm = frm
    Set mUsados = CreateObject("Scripting.Dictionary")
    mItensGerados = 0
    mItensFalhos = 0
    mListaFalhas = ""
    mRotuloDoc = ""
    mProgAtual = 0
    mProgTotal = 1
End Sub

Private Sub RegistrarFalhaItem(ByVal item As String, ByVal msg As String)
    mItensFalhos = mItensFalhos + 1
    If Len(mRotuloDoc) > 0 Then item = mRotuloDoc & " > " & item
    mListaFalhas = mListaFalhas & "- " & item & ": " & LinhaUnica(msg) & vbCrLf
End Sub


'==============================================================================
' EXPORTACAO POR GRUPOS
' - "Grupo" = shape do tipo cdrGroupShape no NIVEL SUPERIOR de uma camada da
'   pagina. Subgrupos ficam dentro do grupo pai e NAO geram arquivo proprio.
' - JPEG/PNG: o grupo e selecionado e exportado com "somente selecao"; o Corel
'   recorta a area nos limites da arte (nada de pagina vazia, so aquele grupo).
' - PDF: o grupo e copiado para um documento TEMPORARIO cuja pagina tem o
'   tamanho exato da arte; o documento original nao e alterado.
' - Nome do arquivo: <Doc>_Pagina-<p>_Grupo-<n>[_<NomeDoGrupo>].<ext>
'   dentro da subpasta <Doc>\ (mesma ideia da exportacao por paginas).
'==============================================================================
Private Sub ExportarGruposDocumento(ByVal doc As Document, ByVal pastaSaida As String, _
        ByVal nomeBase As String, ByVal formato As Long, _
        ByVal curvasNoTemporario As Boolean, ByRef aviso As String)

    Dim itens() As ItemGrupo
    Dim n As Long, k As Long
    Dim ext As String, subpasta As String, arq As String
    Dim msg As String, rotulo As String

    ext = ExtensaoFormato(formato)
    n = PlanejarGrupos(doc, itens, aviso)
    If n = 0 Then
        aviso = aviso & "nenhum grupo exportï¿½vel foi encontrado. "
        Exit Sub
    End If

    subpasta = pastaSaida & nomeBase & "\"
    CriarPastaRecursiva subpasta

    For k = 1 To n
        If mFrm.CancelamentoSolicitado Then Exit For

        rotulo = "Pï¿½gina " & itens(k).pagina & ", Grupo " & itens(k).Indice
        mFrm.AtualizarProgresso mProgAtual * n + (k - 1), mProgTotal * n, _
            "Exportando " & Abreviar(nomeBase, 35) & " - " & rotulo & " (" & k & " de " & n & ")"

        arq = NomeArquivoGrupo(subpasta, nomeBase, itens(k), ext)
        msg = ""
        If Len(arq) = 0 Then
            RegistrarFalhaItem rotulo, "caminho de destino longo demais para o Windows."
        ElseIf ExportarUmGrupo(doc, itens(k), arq, formato, curvasNoTemporario, msg) Then
            mItensGerados = mItensGerados + 1
        Else
            RegistrarFalhaItem rotulo & " (" & mFSO.GetFileName(arq) & ")", msg
        End If
    Next k
End Sub

' Lista os grupos de todas as paginas (somente camadas visiveis e grupos visiveis).
' Ordem: pagina a pagina; dentro da pagina, do fundo para o topo (a numeracao
' dos grupos antigos nao muda quando novos grupos sao criados por cima).
Private Function PlanejarGrupos(ByVal doc As Document, ByRef itens() As ItemGrupo, _
        ByRef aviso As String) As Long

    Dim pg As Page
    Dim lyr As Layer
    Dim shp As Shape
    Dim p As Long, l As Long, s As Long
    Dim n As Long, cont As Long
    Dim gruposOcultos As Long, camadasOcultas As Long

    ReDim itens(1 To 16)

    For p = 1 To doc.Pages.Count
        Set pg = doc.Pages(p)
        cont = 0
        For l = pg.Layers.Count To 1 Step -1
            Set lyr = pg.Layers(l)
            If lyr.Visible Then
                For s = lyr.Shapes.Count To 1 Step -1
                    Set shp = lyr.Shapes(s)
                    If shp.Type = cdrGroupShape Then
                        If ShapeVisivel(shp) Then
                            cont = cont + 1
                            n = n + 1
                            If n > UBound(itens) Then ReDim Preserve itens(1 To n * 2)
                            itens(n).pagina = p
                            itens(n).Indice = cont
                            itens(n).nome = NomeGrupoPersonalizado(shp)
                            Set itens(n).Forma = shp
                        Else
                            gruposOcultos = gruposOcultos + 1
                        End If
                    End If
                Next s
            Else
                camadasOcultas = camadasOcultas + 1
            End If
        Next l
    Next p

    If gruposOcultos > 0 Or camadasOcultas > 0 Then
        aviso = aviso & gruposOcultos & " grupo(s) oculto(s) e " & camadasOcultas & _
                " camada(s) oculta(s) nï¿½o foram exportados. "
    End If
    PlanejarGrupos = n
End Function

' Exporta UM grupo. Nunca deixa erro escapar: devolve False + mensagem.
' Restaura o bloqueio da camada / do objeto, se precisou liberar.
Private Function ExportarUmGrupo(ByVal doc As Document, ByRef it As ItemGrupo, _
        ByVal caminho As String, ByVal formato As Long, _
        ByVal converterCurvas As Boolean, ByRef msg As String) As Boolean

    Dim lyr As Layer
    Dim oShp As Object
    Dim camadaAlterada As Boolean, travadoAlterado As Boolean, travado As Boolean
    Dim ok As Boolean

    On Error GoTo Falha

    doc.Pages(it.pagina).Activate

    Set lyr = it.Forma.Layer
    If Not lyr.Editable Then            ' camada travada: libera so durante a exportacao
        lyr.Editable = True
        camadaAlterada = True
    End If

    Set oShp = it.Forma
    On Error Resume Next
    travado = oShp.Locked               ' (late binding: nem toda versao expoe Locked)
    If Err.Number <> 0 Then travado = False: Err.Clear
    If travado Then
        oShp.Locked = False
        If Err.Number = 0 Then travadoAlterado = True
        Err.Clear
    End If
    On Error GoTo Falha

    If formato = FMT_PDF Then
        ExportarGrupoPDF doc, it.Forma, caminho, converterCurvas
    Else
        SelecionarShape doc, it.Forma
        ExportarBitmap doc, caminho, formato, True
    End If
    ok = True

Limpeza:
    On Error Resume Next
    If travadoAlterado Then oShp.Locked = True
    If camadaAlterada Then lyr.Editable = False
    ExportarUmGrupo = ok
    Exit Function

Falha:
    msg = Err.Description & " (erro " & Err.Number & ")"
    ok = False
    Resume Limpeza
End Function

' PDF de um grupo: copia para um documento temporario, ajusta a pagina ao tamanho
' da arte, publica e descarta o temporario (inclusive em caso de erro).
Private Sub ExportarGrupoPDF(ByVal doc As Document, ByVal shp As Shape, _
        ByVal caminho As String, ByVal converterCurvas As Boolean)

    Dim tmp As Document
    Dim sr As Object, oPg As Object
    Dim larg As Double, alt As Double
    Dim avisoTmp As String
    Dim errNum As Long, errDesc As String

    On Error GoTo Falha

    shp.Copy                                    ' usa a area de transferencia
    Set tmp = Application.CreateDocument
    If tmp Is doc Then Err.Raise vbObjectError + 1011, , "Nï¿½o foi possï¿½vel criar o documento temporï¿½rio."
    tmp.Unit = cdrMillimeter
    tmp.ActiveLayer.Paste

    If converterCurvas Then ConverterTextosEmCurvas tmp, avisoTmp

    Set sr = tmp.ActivePage.Shapes.All
    If sr.Count = 0 Then Err.Raise vbObjectError + 1012, , "O grupo nï¿½o foi colado no documento temporï¿½rio."

    larg = sr.SizeWidth + 2 * PDF_MARGEM_MM
    alt = sr.SizeHeight + 2 * PDF_MARGEM_MM
    If larg <= 0 Or alt <= 0 Then Err.Raise vbObjectError + 1013, , "O grupo tem tamanho invï¿½lido."

    Set oPg = tmp.ActivePage
    On Error Resume Next
    oPg.SetSize larg, alt
    If Err.Number <> 0 Then
        Err.Clear
        oPg.SizeWidth = larg
        oPg.SizeHeight = alt
    End If
    Err.Clear
    On Error GoTo Falha
    ' confere: nao pode sair um PDF do tamanho da pagina original
    If Abs(oPg.SizeWidth - larg) > 0.05 Or Abs(oPg.SizeHeight - alt) > 0.05 Then
        Err.Raise vbObjectError + 1014, , "Nï¿½o foi possï¿½vel ajustar o tamanho da pï¿½gina temporï¿½ria."
    End If

    sr.Move oPg.CenterX - sr.CenterX, oPg.CenterY - sr.CenterY

    ConfigurarIntervaloPDF tmp, 0
    tmp.PublishToPDF caminho

Limpeza:
    On Error Resume Next
    If Not tmp Is Nothing Then
        tmp.Dirty = False
        tmp.Close
    End If
    Set tmp = Nothing
    doc.Activate                                ' volta para o documento de trabalho
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, , errDesc
    Exit Sub

Falha:
    errNum = Err.Number
    errDesc = Err.Description
    If errNum = 0 Then errNum = vbObjectError + 1019
    Resume Limpeza
End Sub

' Seleciona SOMENTE o shape (para exportar "so selecao") e confere.
Private Sub SelecionarShape(ByVal doc As Document, ByVal shp As Shape)
    Dim o As Object
    Dim sel As Object
    Dim n As Long

    LimparSelecao doc
    Set o = shp
    On Error Resume Next
    o.CreateSelection
    If Err.Number <> 0 Then
        Err.Clear
        o.Select True
    End If
    Err.Clear
    Set sel = ObterSelecao(doc)
    If Not sel Is Nothing Then n = sel.Count
    On Error GoTo 0
    If n <> 1 Then Err.Raise vbObjectError + 1030, , "Nï¿½o foi possï¿½vel selecionar o grupo para exportaï¿½ï¿½o."
End Sub

Private Function ObterSelecao(ByVal doc As Document) As Object
    Dim o As Object, a As Object
    On Error Resume Next
    Set o = doc
    Set ObterSelecao = o.SelectionRange
    If ObterSelecao Is Nothing Then
        Err.Clear
        Set a = Application
        Set ObterSelecao = a.ActiveSelectionRange
    End If
End Function

Private Sub LimparSelecao(ByVal doc As Document)
    Dim o As Object
    On Error Resume Next
    Set o = doc
    o.ClearSelection
End Sub

Private Sub RestaurarSelecao(ByVal doc As Document, ByVal selOrig As Object)
    On Error Resume Next
    LimparSelecao doc
    If selOrig Is Nothing Then Exit Sub
    If selOrig.Count > 0 Then selOrig.CreateSelection
End Sub

Private Function ShapeVisivel(ByVal shp As Shape) As Boolean
    Dim o As Object
    Dim v As Boolean
    ShapeVisivel = True
    On Error Resume Next
    Set o = shp
    v = o.Visible                     ' late binding: se nao existir, considera visivel
    If Err.Number = 0 Then ShapeVisivel = v
    Err.Clear
End Function

' Nome dado pelo usuario ao grupo; vazio quando e o nome padrao do Corel
' ("Group of 3 Objects", "Grupo de 3 objetos"...) ou quando nao esta acessivel.
Private Function NomeGrupoPersonalizado(ByVal shp As Shape) As String
    Dim nome As String, l As String

    On Error Resume Next
    nome = Trim$(shp.Name)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If Len(nome) = 0 Then Exit Function

    l = LCase$(nome)
    If l Like "group of #*" Or l Like "grupo de #*" Or l Like "groupe de #*" _
       Or l Like "gruppe von #*" Or l Like "gruppo di #*" Then Exit Function

    NomeGrupoPersonalizado = NomeSeguro(nome, 40)
End Function

' Monta o caminho final do arquivo do grupo (unico nesta execucao).
' Devolve "" se, mesmo sem o nome do grupo, o caminho passar do limite do Windows.
Private Function NomeArquivoGrupo(ByVal subpasta As String, ByVal nomeBase As String, _
        ByRef it As ItemGrupo, ByVal ext As String) As String
    Dim corpo As String, cam As String

    corpo = nomeBase & "_Pagina-" & it.pagina & "_Grupo-" & it.Indice
    cam = subpasta & corpo & IIf(Len(it.nome) > 0, "_" & it.nome, "") & ext
    If Len(cam) > 250 Then cam = subpasta & corpo & ext
    If Len(cam) > 250 Then Exit Function
    NomeArquivoGrupo = NomeUnico(cam)
End Function

' Garante que dois itens da mesma execucao nunca gerem o mesmo arquivo:
' se o caminho ja foi usado, acrescenta _2, _3...
Private Function NomeUnico(ByVal caminho As String) As String
    Dim p As Long, n As Long
    Dim corpo As String, ext As String, cand As String

    p = InStrRev(caminho, ".")
    corpo = Left$(caminho, p - 1)
    ext = Mid$(caminho, p)
    cand = caminho
    n = 1
    Do While mUsados.Exists(LCase$(cand))
        n = n + 1
        cand = corpo & "_" & n & ext
    Loop
    mUsados.Add LCase$(cand), True
    NomeUnico = cand
End Function

' Remove/substitui caracteres invalidos no Windows, pontos/espacos finais e
' nomes reservados (CON, NUL, COM1...).
Private Function NomeSeguro(ByVal texto As String, ByVal maxLen As Long) As String
    Dim i As Long, cod As Long
    Dim c As String, r As String

    For i = 1 To Len(texto)
        c = Mid$(texto, i, 1)
        cod = AscW(c)
        If (cod >= 0 And cod < 32) Or InStr("\/:*?""<>|", c) > 0 Then
            r = r & "-"
        Else
            r = r & c
        End If
    Next i

    r = Trim$(r)
    If Len(r) > maxLen Then r = Left$(r, maxLen)
    Do While Len(r) > 0
        If Right$(r, 1) = "." Or Right$(r, 1) = " " Then
            r = Left$(r, Len(r) - 1)
        Else
            Exit Do
        End If
    Loop

    Select Case UCase$(r)
        Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", _
             "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", _
             "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            r = "_" & r
    End Select
    NomeSeguro = r
End Function


'==============================================================================
' EXPORTACAO DO DOCUMENTO ABERTO NO CORELDRAW  (nao usa pasta de origem)
' Trabalha no documento ativo, no estado atual (inclusive alteracoes nao salvas).
' NAO salva, NAO fecha, NAO altera objetos. Restaura: pagina ativa, selecao,
' bloqueios, ajustes de PDF e o estado "modificado" do documento.
'==============================================================================
Public Function VerificarDocumentoAberto() As Boolean
    Dim n As Long
    On Error Resume Next
    n = Application.Documents.Count
    On Error GoTo 0
    If n = 0 Then
        MsgBox "Nï¿½o hï¿½ nenhum documento aberto no CorelDRAW." & vbCrLf & _
               "Abra o documento que deseja exportar e tente novamente.", _
               vbExclamation, "Exportador de CDR"
    End If
    VerificarDocumentoAberto = (n > 0)
End Function

Public Function ExportarDocumentoAberto(ByVal pastaDestino As String, ByVal formato As Long, _
        ByVal paginasSeparadas As Boolean, ByVal gruposSeparados As Boolean, _
        ByVal converterCurvas As Boolean, ByVal frm As Object, _
        ByRef res As ResultadoExportacao) As Boolean

    Dim vazio As ResultadoExportacao
    Dim doc As Document
    Dim destino As String, nomeBase As String, aviso As String
    Dim pgOrig As Page
    Dim selOrig As Object
    Dim sujoOrig As Boolean
    Dim pdfS As Object
    Dim pdfRange As Variant, pdfPaginas As Variant
    Dim pdfSalvo As Boolean
    Dim otimizado As Boolean, preparado As Boolean
    Dim resp As VbMsgBoxResult

    On Error GoTo FalhaGeral

    res = vazio
    res.ModoDocAberto = True
    res.UsouGrupos = gruposSeparados
    Set mFSO = CreateObject("Scripting.FileSystemObject")
    IniciarExecucao frm

    destino = NormalizarPasta(pastaDestino)
    If Len(destino) = 0 Then
        MsgBox "Selecione a pasta de destino para exportar o documento aberto.", _
               vbExclamation, "Exportador de CDR"
        Exit Function
    End If
    If Not VerificarDocumentoAberto() Then Exit Function

    Set doc = ActiveDocument
    nomeBase = NomeSeguro(mFSO.GetBaseName(doc.FileName), 100)
    If Len(nomeBase) = 0 Then nomeBase = "Documento"
    res.NomeDocumento = doc.FileName
    res.Encontrados = 1
    mRotuloDoc = ""

    ' --- saidas ja existentes: pergunta antes de mexer em qualquer coisa ---
    If SaidaJaExiste(destino, nomeBase, formato, paginasSeparadas, gruposSeparados) Then
        resp = MsgBox("Jï¿½ existem exportaï¿½ï¿½es deste documento na pasta de destino." & vbCrLf & vbCrLf & _
                      "SIM  = Substituir os arquivos de mesmo nome" & vbCrLf & _
                      "NAO  = Nï¿½o exportar nada", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Arquivos jï¿½ existentes")
        If resp <> vbYes Then Exit Function
    End If

    ' --- guarda o estado original do documento ---
    Set pgOrig = doc.ActivePage
    Set selOrig = ObterSelecao(doc)
    sujoOrig = doc.Dirty
    On Error Resume Next
    Err.Clear
    Set pdfS = doc.PDFSettings
    pdfRange = pdfS.PublishRange
    pdfPaginas = pdfS.PageRange
    pdfSalvo = (Err.Number = 0)
    Err.Clear
    On Error GoTo FalhaGeral

    CriarPastaRecursiva destino
    Application.Optimization = True
    otimizado = True
    preparado = True

    If converterCurvas And Not gruposSeparados Then
        aviso = "textos nï¿½o convertidos em curvas para nï¿½o alterar o documento aberto " & _
                "(a conversï¿½o sï¿½ ï¿½ aplicada nos documentos temporï¿½rios da exportaï¿½ï¿½o por grupos). "
    End If

    mProgAtual = 0
    mProgTotal = 1
    frm.AtualizarProgresso 0, 1, "Exportando o documento aberto..."

    On Error GoTo FalhaDoc
    ExportarDocumento doc, destino, nomeBase, formato, paginasSeparadas, gruposSeparados, _
                      converterCurvas, True, aviso
    res.Exportados = 1
    If frm.CancelamentoSolicitado Then res.Cancelado = True
    ExportarDocumentoAberto = True

Finalizar:
    On Error Resume Next
    If preparado Then
        If otimizado Then Application.Optimization = False
        If pdfSalvo Then
            pdfS.PageRange = pdfPaginas
            pdfS.PublishRange = pdfRange
        End If
        If Not doc Is Nothing Then
            doc.Activate
            If Not pgOrig Is Nothing Then pgOrig.Activate
            RestaurarSelecao doc, selOrig
            If Not sujoOrig Then doc.Dirty = False
        End If
        Application.Refresh
    End If
    res.ItensGerados = mItensGerados
    res.ItensFalhos = mItensFalhos
    res.ListaFalhasItens = mListaFalhas
    If Len(aviso) > 0 Then res.ListaAvisos = res.ListaAvisos & "- " & res.NomeDocumento & ": " & LinhaUnica(aviso) & vbCrLf
    frm.AtualizarProgresso 1, 1, IIf(res.Cancelado, "Cancelado pelo usuï¿½rio.", "Concluï¿½do.")
    Set mFrm = Nothing
    Exit Function

FalhaDoc:
    res.Erros = 1
    res.ListaErros = res.ListaErros & "- " & res.NomeDocumento & ": " & _
                     LinhaUnica(Err.Description & " (erro " & Err.Number & ")") & vbCrLf
    ExportarDocumentoAberto = True          ' o relatorio sera exibido
    Resume Finalizar

FalhaGeral:
    ExportarDocumentoAberto = False
    MsgBox "Erro inesperado durante a exportaï¿½ï¿½o do documento aberto:" & vbCrLf & vbCrLf & _
           Err.Description & " (erro " & Err.Number & ")", vbCritical, "Exportador de CDR"
    Resume Finalizar
End Function