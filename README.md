# SAP B1 — RPA Baixa de Ativos em Lote

Automação desktop para execução em lote do processo de **baixa de ativos fixos** no SAP Business One, via **Service Layer (REST API)**. Desenvolvida em Python com interface gráfica Tkinter.

---

## Visão Geral

O processo de baixa de ativos no SAP B1 (transação `AFAB` / endpoint `AssetRetirement`) pode envolver centenas de registros. Esta ferramenta automatiza esse processo, lendo uma lista de ativos a partir de um arquivo CSV e enviando as requisições em lotes configuráveis para a API SAP Service Layer, exibindo progresso e log em tempo real.

---

## Funcionalidades

- Conexão com SAP B1 Service Layer (HTTPS)
- Autenticação por sessão (`B1SESSION`)
- Leitura de ativos a partir de arquivo **CSV** (suporte a separador `,` e `;`)
- Processamento em **lotes** (tamanho configurável)
- Configuração de **datas** independentes: Document Date, Posting Date e Asset Value Date
- Configuração de **filial** (BPLId)
- **Parada controlada** da execução pelo usuário
- **Log em tempo real** com codificação por cores (OK, Erro, Aviso, Info)
- **Exportação do log** para arquivo `.txt`
- **Persistência de configurações** em `config.json` (senha ofuscada em Base64)
- Interface responsiva com painel de progresso (total, lotes, OK, erros)

---

## Tecnologias

| Tecnologia | Uso |
|---|---|
| Python 3.x | Linguagem principal |
| Tkinter / ttk | Interface gráfica desktop |
| Requests | Comunicação HTTP com SAP Service Layer |
| Threading | Execução assíncrona sem travar a UI |
| PyInstaller | Geração de executável `.exe` standalone |

---

## Estrutura do Projeto

```
RPA BaixaAtivos/
├── rpa_baixa_ativos.py       # Script principal (lógica + GUI)
├── config.json               # Configurações persistidas (gerado em runtime)
├── RPA_Baixa_Ativos.spec     # Spec do PyInstaller para build do .exe
├── grok_logs.csv             # Exemplo de arquivo CSV de ativos
└── Documentação/
    ├── 1.FS_SAP_AssetRetirement_v2.1.docx   # Functional Specification
    ├── 2.TD_SAP_AssetRetirement_v2.0.docx   # Technical Design
    ├── 3.BL_SAP_AssetRetirement_v2.0.docx   # Business Logic
    ├── 4.TP_SAP_AssetRetirement_v2.0.docx   # Test Plan
    └── UI_Ativo_01.png                       # Screenshot da interface
```

---

## Formato do CSV

O arquivo CSV deve conter uma coluna chamada `AssetNumber`:

```csv
AssetNumber
10001234
10001235
10001236
```

Suporta tanto `,` quanto `;` como separador. A codificação recomendada é **UTF-8**.

---

## Configuração

Na interface, preencha os campos na aba **Conexão SAP**:

| Campo | Descrição |
|---|---|
| Service Layer | URL base da API (ex: `https://servidor:50000/b1s/v1`) |
| Company DB | Nome do banco de dados da empresa no SAP |
| Usuário | Login do usuário SAP |
| Senha | Senha do usuário SAP |
| BPLId | ID da filial (Branch) no SAP |
| Tamanho do lote | Quantidade de ativos por requisição (padrão: 100) |

As configurações são salvas automaticamente em `config.json` ao executar. A senha é armazenada em Base64 (ofuscação, não criptografia).

> **Atenção:** Não versione o `config.json` com credenciais reais em ambientes compartilhados.

---

## Instalação e Execução

### Pré-requisitos

```bash
pip install requests
```

### Executar via Python

```bash
python rpa_baixa_ativos.py
```

### Gerar executável (.exe)

```bash
pyinstaller RPA_Baixa_Ativos.spec
```

O executável será gerado em `dist/RPA_Baixa_Ativos.exe` — standalone, sem necessidade de Python instalado.

---

## Como Usar

1. Abra o aplicativo
2. Preencha os dados de conexão SAP e clique em **Salvar Configurações**
3. Selecione o arquivo **CSV** com os números dos ativos (`AssetNumber`)
4. Defina as datas e o tamanho do lote
5. Clique em **EXECUTAR**
6. Acompanhe o progresso e o log em tempo real
7. Ao finalizar, exporte o log se necessário

---

## Endpoint SAP Utilizado

```
POST /b1s/v1/AssetRetirement
```

Payload enviado por lote:

```json
{
  "DocumentDate": "2024-01-01",
  "PostingDate": "2024-01-01",
  "AssetValueDate": "2024-01-01",
  "BPLId": 60,
  "AssetDocumentLineCollection": [
    { "AssetNumber": "10001234", "Quantity": 1, "TotalLC": 0 }
  ]
}
```

---

## Documentação

A pasta `Documentação/` contém os documentos do ciclo de vida do projeto:

- **FS** — Functional Specification: requisitos funcionais do processo
- **TD** — Technical Design: especificação técnica da solução
- **BL** — Business Logic: regras de negócio aplicadas
- **TP** — Test Plan: plano e casos de teste

---

## Licença

Uso interno. Todos os direitos reservados.
