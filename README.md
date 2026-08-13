# Conversor Word para PDF em Lote

Utilitário para Windows que converte automaticamente todos os arquivos `.doc` e `.docx` de uma pasta para PDF.

O conversor utiliza o Microsoft Word instalado no computador e salva cada PDF na mesma pasta e com o mesmo nome do documento original.

## Funcionalidades

- Conversão de vários documentos em uma única execução;
- Compatibilidade com arquivos `.doc` e `.docx`;
- Suporte a caminhos com espaços e caracteres acentuados;
- Seleção da pasta digitando ou colando o caminho;
- Execução por arrastar e soltar uma pasta sobre o arquivo `.bat`;
- Exibição do progresso durante a conversão;
- Continuidade do processamento quando um documento apresenta erro;
- Fechamento automático do Microsoft Word ao finalizar;
- Exclusão automática dos arquivos temporários do Word iniciados por `~$`.

## Requisitos

- Windows 10 ou Windows 11;
- Microsoft Word instalado e ativado;
- Windows PowerShell 5.1 ou superior;
- Permissão de leitura e gravação na pasta dos documentos.

> O Microsoft Word é necessário porque o conversor utiliza a automação COM do próprio aplicativo para preservar a formatação, as fontes, as margens, os cabeçalhos e a paginação dos documentos.

## Arquivo principal

```text
Converter_Word_para_PDF.bat
```

## Como instalar

1. Baixe o arquivo `Converter_Word_para_PDF.bat` deste repositório.
2. Salve o arquivo em uma pasta de sua preferência.
3. Não é necessário instalar bibliotecas adicionais.

Caso o Windows exiba um alerta de segurança, confirme a execução somente se o arquivo tiver sido baixado deste repositório ou de outra fonte confiável.

## Como usar

### Opção 1 — Informar o caminho da pasta

1. Clique duas vezes em `Converter_Word_para_PDF.bat`.
2. Digite ou cole o caminho completo da pasta que contém os documentos.
3. Pressione `Enter`.
4. Aguarde a mensagem de conclusão.

Exemplo:

```text
C:\Users\usuario\OneDrive - Empresa\Área de Trabalho\Documentos
```

Não é necessário adicionar aspas ao caminho, mesmo quando existirem espaços.

### Opção 2 — Arrastar e soltar

1. Localize a pasta que contém os arquivos Word.
2. Arraste a pasta sobre `Converter_Word_para_PDF.bat`.
3. Solte a pasta sobre o arquivo.
4. Aguarde a conclusão da conversão.

## Resultado esperado

Antes da execução:

```text
Documentos/
├── Memorial_de_Calculo.docx
├── Relatorio_Tecnico.docx
└── Procedimento.doc
```

Depois da execução:

```text
Documentos/
├── Memorial_de_Calculo.docx
├── Memorial_de_Calculo.pdf
├── Relatorio_Tecnico.docx
├── Relatorio_Tecnico.pdf
├── Procedimento.doc
└── Procedimento.pdf
```

Os documentos originais são mantidos. Somente os arquivos PDF são criados ou atualizados.

## Atenção aos arquivos existentes

Se já existir um PDF com o mesmo nome na pasta de destino, ele poderá ser substituído pela nova conversão. Faça uma cópia de segurança quando precisar preservar versões anteriores.

## Como funciona

O arquivo em lote executa uma rotina PowerShell que:

1. Valida o caminho informado;
2. Localiza os arquivos `.doc` e `.docx` da pasta selecionada;
3. Ignora arquivos temporários do Microsoft Word;
4. Abre o Word de forma invisível;
5. Abre cada documento no modo somente leitura;
6. Exporta o documento no formato PDF;
7. Fecha cada documento sem realizar alterações;
8. Encerra o Word e informa a quantidade de PDFs criados.

## Solução de problemas

### `'}' de fechamento ausente` ou `$ErrorActionPreference não é reconhecido`

Esse erro ocorre em versões antigas do arquivo nas quais o comando PowerShell foi dividido incorretamente em várias linhas. Baixe a versão mais recente de `Converter_Word_para_PDF.bat`.

### `A pasta informada não existe`

Confira se:

- O caminho foi copiado por completo;
- A pasta ainda existe;
- A unidade de rede ou o OneDrive está conectado;
- Você possui acesso à pasta.

### `Nenhum arquivo .doc ou .docx foi encontrado`

O programa pesquisa somente a pasta selecionada. Ele não percorre automaticamente as subpastas.

### Erro ao criar o objeto COM do Word

Verifique se o Microsoft Word está instalado, ativado e abrindo normalmente. Feche processos travados do Word pelo Gerenciador de Tarefas e tente novamente.

### O arquivo não foi convertido

As causas mais comuns são:

- Documento corrompido;
- Documento protegido por senha;
- Arquivo bloqueado por outro usuário;
- Falta de permissão para gravar na pasta;
- PDF de destino aberto ou bloqueado por outro programa;
- Fonte, vínculo ou componente do documento exigindo uma confirmação manual.

O conversor registra a falha na tela e tenta processar os demais documentos.

### O Word permaneceu aberto

Normalmente o Word é encerrado automaticamente. Caso a execução seja interrompida à força, pode permanecer um processo `WINWORD.EXE`. Feche-o pelo Gerenciador de Tarefas antes de executar novamente.

## Limitações

- Funciona somente no Windows;
- Requer o Microsoft Word instalado;
- Não converte arquivos localizados em subpastas;
- Não converte arquivos `.odt`, `.rtf`, `.xls`, `.xlsx` ou outros formatos;
- Documentos protegidos por senha podem exigir abertura manual;
- A velocidade depende da quantidade, do tamanho e da complexidade dos documentos.

## Segurança e privacidade

Todo o processamento acontece localmente no computador. O conversor não envia os documentos para serviços externos e não altera os arquivos Word originais.

Antes de executar qualquer arquivo `.bat`, revise o conteúdo e confirme que ele foi obtido de uma fonte confiável.

## Estrutura sugerida do repositório

```text
batch-word-para-pdf/
├── Converter_Word_para_PDF.bat
└── README.md
```

## Créditos

Rotina desenvolvida a partir da ideia do projeto [batch-docx2pdf](https://github.com/matthansen0/batch-docx2pdf), com adaptação para arquivo `.bat`, mensagens em português, tratamento individual de erros, suporte a arrastar e soltar e encerramento controlado do Microsoft Word.

Adaptação e documentação: **Jovani Kaczalla**.

## Contribuições

Sugestões, correções e melhorias podem ser enviadas por meio de *issues* ou *pull requests* neste repositório.

