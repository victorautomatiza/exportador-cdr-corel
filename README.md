# CDR Batch Exporter para CorelDRAW

![Licença MIT](https://img.shields.io/badge/licen%C3%A7a-MIT-green.svg)
![Plataforma](https://img.shields.io/badge/CorelDRAW-VBA-blue.svg)
![Sistema](https://img.shields.io/badge/Windows-0078D6.svg)

Macro VBA com interface gráfica que exporta **em lote** todos os arquivos `.CDR` de uma pasta **e de todas as suas subpastas** para **JPEG, PNG ou PDF**, reproduzindo exatamente a mesma estrutura de pastas no destino.

> **Os arquivos `.CDR` originais nunca são abertos, alterados ou salvos.** Todo o trabalho é feito em uma cópia temporária que é descartada ao final.

---

## Sumário

- [Recursos](#recursos)
- [Como funciona](#como-funciona)
- [Instalação](#instalação)
- [Como usar](#como-usar)
- [Formatos de exportação](#formatos-de-exportação)
- [Arquivos com várias páginas](#arquivos-com-várias-páginas)
- [Arquivos já existentes no destino](#arquivos-já-existentes-no-destino)
- [Relatório final](#relatório-final)
- [Configurações avançadas](#configurações-avançadas)
- [Solução de problemas](#solução-de-problemas)
- [Estrutura do repositório](#estrutura-do-repositório)
- [Contribuindo](#contribuindo)
- [Créditos](#créditos)
- [Licença](#licença)

---

## Recursos

- Busca **recursiva** de `.CDR` em qualquer nível de subpastas.
- **Preserva a estrutura de pastas** da origem no destino (as pastas são criadas automaticamente).
- Exporta para **JPEG** (300 DPI), **PNG** (padrão web) ou **PDF** (`PublishToPDF`).
- Opção de **salvar cada página como um arquivo separado**, em uma subpasta com o nome do arquivo.
- Opção de **converter textos em curvas** (somente na cópia temporária, sem tocar no original).
- Arquivos com o mesmo nome em pastas diferentes **não se sobrescrevem**.
- Pergunta **uma única vez** o que fazer quando já existem arquivos exportados no destino.
- Barra de progresso, botão **Cancelar** durante o processamento e relatório final com erros e avisos.
- Lembra as últimas escolhas entre um uso e outro.

---

## Como funciona

Exemplo com origem `Arquivos` e destino `Exportados`:

```
Arquivos/                                  Exportados/
├── Cartões de Visita/                     ├── Cartões de Visita/
│   ├── Cartão - Cliente-1.cdr     ───►    │   ├── Cartão - Cliente-1.jpg
│   └── Cartão - Cliente-2.cdr             │   └── Cartão - Cliente-2.jpg
├── Panfletos/                             ├── Panfletos/
│   └── Panfleto - Cliente-1.cdr           │   └── Panfleto - Cliente-1.jpg
└── Clientes/                              └── Clientes/
    └── 2026/                                  └── 2026/
        └── Janeiro/                               └── Janeiro/
            └── arquivo.cdr                            └── arquivo.jpg
```

Para cada `.CDR` encontrado, a ferramenta:

1. calcula o **caminho relativo** do arquivo em relação à pasta de origem e o recria no destino;
2. **copia** o arquivo para uma pasta temporária (`%TEMP%\ExportadorCDR_...`);
3. abre **somente a cópia** no CorelDRAW;
4. se marcado, converte os textos em curvas (na cópia);
5. exporta no formato escolhido;
6. fecha a cópia **sem salvar** e a apaga.

Como o original nunca é aberto, não há risco de alterá-lo, mesmo que ele já esteja aberto no CorelDRAW. Antes de qualquer alteração, a macro também confere se o documento aberto é mesmo a cópia temporária; se não for, aborta sem mexer em nada.

---

## Instalação

### Requisitos

- Windows
- CorelDRAW com o **suporte a VBA instalado** (em algumas versões o VBA é um componente opcional do instalador)
- Testado no CorelDRAW: **24.0.0.301**

### Passo a passo

1. Baixe ou clone este repositório.
2. Abra o CorelDRAW e pressione **`Alt + F11`** para abrir o editor VBA.
3. No painel de projetos, selecione **`GlobalMacros`** (projeto global do Corel).
4. Menu **Arquivo > Importar arquivo...** e escolha `src/modExportadorCDR.bas`.
5. Menu **Inserir > UserForm**. Um formulário vazio será criado. **Não desenhe nada nele.**
6. Na janela **Propriedades** (`F4`), altere o campo **`(Name)`** do formulário para:

   ```
   frmExportador
   ```

7. Dê dois cliques no formulário (ou pressione `F7`) para abrir o código dele. **Apague tudo** o que estiver lá e cole o conteúdo de `src/frmExportador_codigo.txt`. Você pode copiar a partir da linha `Option Explicit`; a caixa de comentário no topo é só instrução.
8. Salve (`Ctrl + S`) e feche o editor.

> Todos os controles da interface (campos, botões, opções, barra de progresso) são criados por código, por isso o formulário precisa estar vazio.

### Executando a macro

- **Pelo gerenciador de macros:** `Alt + F8` (ou **Ferramentas > Macros > Gerenciador de macros**), escolha `ExportadorCDR` e execute.
- **Pelo editor VBA:** clique dentro da sub `ExportadorCDR` e pressione `F5`.
- **Opcional:** em **Ferramentas > Personalização**, você pode criar um botão ou atalho de teclado para a macro `ExportadorCDR`.

---

## Como usar

A janela tem os seguintes campos:

| Campo | Descrição |
|---|---|
| **Arquivos** | Pasta de origem. Todos os `.CDR` dela e das subpastas serão processados. |
| **Destino** | Pasta onde os arquivos exportados serão salvos. Precisa ser diferente da origem. |
| **Formato** | JPEG, PNG ou PDF. |
| **Salvar cada página como um arquivo separado** | Exporta página a página, em uma subpasta com o nome do arquivo. |
| **Converter textos em curvas** | Converte os textos em curvas antes de exportar (somente na cópia temporária). |

Clique em **Salvar** para iniciar. Durante o processamento, o botão **Sair** vira **Cancelar**: o arquivo atual é concluído e o processo para.

---

## Formatos de exportação

| Formato | Configuração |
|---|---|
| **JPEG** | 300 DPI. O tipo de imagem (RGB/CMYK) não é forçado pela macro. |
| **PNG** | 96 DPI, RGB, antialiasing normal e fundo transparente (padrão de exportação web do Corel). |
| **PDF** | `PublishToPDF` com as configurações padrão do documento. |

---

## Arquivos com várias páginas

**Opção desmarcada (padrão)**

- **PDF:** um único arquivo com todas as páginas.
- **JPEG / PNG:** esses formatos comportam uma página por arquivo, então somente a **1ª página** é exportada. O relatório final registra um aviso quando isso acontece.

**Opção marcada**

Cada página vira um arquivo, dentro de uma subpasta com o nome base do `.CDR`:

```
Panfleto-Cliente.cdr   (3 páginas)

Panfleto-Cliente/
├── Panfleto-Cliente_1-1.jpg
├── Panfleto-Cliente_2-2.jpg
└── Panfleto-Cliente_3-3.jpg
```

Mesmo um arquivo de uma página só ganha a subpasta, para que o comportamento seja sempre o mesmo.

---

## Arquivos já existentes no destino

A ferramenta **nunca sobrescreve em silêncio**. Antes de começar, ela conta quantas exportações já existem e faz **uma única pergunta**:

- **Sim:** substitui os existentes.
- **Não** (padrão): ignora os existentes e mantém como estão.
- **Cancelar:** não exporta nada.

---

## Relatório final

Ao terminar, uma janela mostra o resultado:

```
Exportação concluída!

Arquivos encontrados: 25
Arquivos exportados: 24
Arquivos com erro: 1

Arquivos com erro:
- Clientes\João\Cartão.cdr: <mensagem do erro>
```

Quando há erros ou avisos, um arquivo `Relatorio_Exportacao_AAAAMMDD_HHMMSS.txt` com o detalhe completo é gravado na raiz da pasta de destino. Sem erros nem avisos, nenhum arquivo extra é criado.

---

## Configurações avançadas

No topo de `modExportadorCDR.bas` há constantes que podem ser ajustadas:

| Constante | Padrão | Efeito |
|---|---|---|
| `JPEG_DPI` | `300` | Resolução do JPEG. |
| `PNG_DPI` | `96` | Resolução do PNG. |
| `PNG_FUNDO_TRANSPARENTE` | `True` | Use `False` para PNG com fundo branco. |
| `JPEG_TIPO_IMAGEM_FORCADO` | `0` | `0` não força nada. Troque por `cdrRGBColorImage` ou `cdrCMYKColorImage` para forçar o modo de cor do JPEG. |

---

## Solução de problemas

**"Nenhum arquivo .CDR foi encontrado"**
Confira a pasta de origem. Se alguma subpasta não pôde ser lida (permissão, por exemplo), a mensagem mostra um aviso com o caminho.

**Erro de compilação "Variável não definida"**
Algum nome de constante do CorelDRAW pode variar entre versões. Anote qual palavra ficou destacada e abra o **Pesquisador de Objetos** (`F2`) para conferir o nome correto na sua versão.

**Caracteres acentuados quebrados após colar ou importar**
Os arquivos do projeto usam codificação ANSI (Windows-1252), que é a que o editor VBA espera. Se o seu editor de texto converter para UTF-8, os acentos podem quebrar.

**Janela de fontes ausentes ou de links**
O CorelDRAW pode exibir seus avisos normais de fontes ausentes ao abrir a cópia. Isso aconteceria também com o arquivo original.

**Imagens vinculadas não aparecem**
Como o arquivo é aberto a partir de uma cópia em pasta temporária, links externos definidos por caminho relativo podem não ser encontrados. Links por caminho absoluto funcionam normalmente.

**Caminhos muito longos**
O Windows limita caminhos a cerca de 260 caracteres. Se um arquivo falhar por isso, ele aparece na lista de erros do relatório.

---

## Estrutura do repositório

```
corel-cdr-batch-exporter/
├── src/
│   ├── modExportadorCDR.bas        # Módulo com toda a lógica
│   └── frmExportador_codigo.txt    # Código do UserForm (colar no formulário)
├── .gitattributes
├── .gitignore
├── LICENSE
└── README.md
```

---

## Contribuindo

Contribuições são bem-vindas. Abra uma *issue* descrevendo o problema ou a sugestão, ou envie um *pull request*. Ao reportar um erro, informe a **versão do CorelDRAW** e a **mensagem de erro exata**.

---

## Créditos

Projeto idealizado e mantido por **Victor Nascimento Feitosa**.

Desenvolvido com o auxílio de:

- **[Claude](https://www.anthropic.com/claude)** (Anthropic)
- **[ChatGPT](https://chatgpt.com)** (OpenAI)

---

## Licença

Distribuído sob a licença **MIT**. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.
