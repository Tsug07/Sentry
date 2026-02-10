# Sentry - Controle de Certidoes

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![CustomTkinter](https://img.shields.io/badge/CustomTkinter-GUI-blue?style=for-the-badge)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Ativo-brightgreen?style=for-the-badge)

Dashboard para monitoramento e controle de Certidoes Negativas de Debito (CNDs) de empresas. Identifica certidoes vencidas, faltantes e positivas de forma automatizada.

## Funcionalidades

- **Verificar Positiva**: Escaneia PDFs de CNDs em subpastas e identifica se alguma certidao contem "Certidao Positiva de Debitos - CPD"
- **Verificar Vencimento**: Extrai a data de validade do nome dos arquivos PDF (formato `dd.mm.aaaa`) e classifica como VALIDA ou VENCIDA
- **Dashboard visual**: Cards de estatisticas clicaveis (Total, Completas, Vencidas, Faltantes) que filtram a tabela
- **Grafico de distribuicao**: Grafico de pizza mostrando a proporcao de empresas completas, incompletas e com erros
- **Busca e ordenacao**: Busca por nome de empresa e ordenacao por qualquer coluna
- **Exportacao Excel**: Gera relatorio `.xlsx` formatado com cores, formatacao condicional e validacao de dados
- **Processamento paralelo**: Usa ThreadPoolExecutor (ate 8 threads) para processar multiplas pastas simultaneamente
- **Persistencia de configuracao**: Salva ultima pasta selecionada e modo de operacao em `cnd_config.json`

## Estrutura esperada de pastas

```
Pasta Principal/
  Empresa A/
    CND MUNICIPAL xxxx.pdf
    CND RFB xxxx.pdf
    CND FGTS xxxx.pdf
    CND PROC xxxx.pdf
    CND ESTADUAL xxxx.pdf
  Empresa B/
    CND MUNICIPAL 01.06.2025.pdf
    CND RFB 15.07.2025.pdf
    ...
```

Cada subpasta representa uma empresa. Os arquivos PDF devem conter no nome o tipo da certidao (`CND MUNICIPAL`, `CND RFB`, `CND FGTS`, `CND PROC`, `CND ESTADUAL`). No modo de vencimento, o nome do arquivo deve incluir a data no formato `dd.mm.aaaa`.

## CNDs monitoradas

| Sigla | Certidao |
|-------|----------|
| CND MUNICIPAL | Certidao Negativa de Debitos Municipal |
| CND RFB | Certidao Negativa de Debitos - Receita Federal |
| CND FGTS | Certidao de Regularidade do FGTS |
| CND PROC | Certidao Negativa de Debitos Trabalhistas |
| CND ESTADUAL | Certidao Negativa de Debitos Estadual |

## Requisitos

- Python 3.10+

### Bibliotecas

```
customtkinter
PyPDF2
openpyxl
matplotlib
```

### Instalacao das dependencias

```bash
pip install customtkinter PyPDF2 openpyxl matplotlib
```

## Como usar

```bash
python Sentry.py
```

1. Selecione a pasta principal que contem as subpastas das empresas
2. Escolha o modo de operacao: **Verificar Positiva** ou **Verificar Vencimento**
3. Clique em **Processar**
4. Analise os resultados na tabela e nos cards de estatisticas
5. Clique nos cards para filtrar por status (Completas, Vencidas, Faltantes)
6. Use a busca para localizar empresas especificas
7. Exporte o relatorio em Excel clicando em **Exportar**

## Status das empresas

| Status | Descricao |
|--------|-----------|
| COMPLETO | Todas as 5 CNDs esperadas foram encontradas |
| INCOMPLETO | Uma ou mais CNDs estao faltando |
| ERRO | Ocorreu um erro ao processar a pasta da empresa |

## Cores na tabela

| Cor | Significado |
|-----|-------------|
| Vermelho | Certidao vencida (prioridade mais alta) |
| Laranja | Certidao faltante (NAO encontrada) |
| Amarelo | Status incompleto |
| Roxo | Erro no processamento |
| Verde | Completo / Valida |

## Configuracao

O arquivo `cnd_config.json` armazena:

- `expected_files`: Lista das CNDs esperadas
- `target_line`: Texto que identifica uma certidao positiva
- `last_folder`: Ultima pasta processada
- `mode`: Ultimo modo selecionado
- `ignored_folders`: Pastas ignoradas durante o processamento

## Logs

O arquivo `cnd_dashboard.log` registra todas as operacoes realizadas para auditoria e depuracao.
