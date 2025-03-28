
#  Gerador de Documentos em Power BI 

![Power BI Logo](https://upload.wikimedia.org/wikipedia/commons/thumb/c/cf/Power_bi_logo_black.svg/1200px-Power_bi_logo_black.svg.png)

Automação para geração de documentação detalhada de relatórios Power BI a partir de arquivos `.pbit`.

> 🔍 **Baseado no trabalho original de [Julia Lira](https://github.com/data-ju/Power_BI_Documentation)**  
> Este projeto foi adaptado a partir da solução inicial desenvolvida por Julia Lira para extração de conteúdo de arquivos PBIT.

## 📌 Visão Geral

Este projeto automatiza a criação de documentação técnica para relatórios Power BI, extraindo informações diretamente dos arquivos `.pbit` (Power BI Template). O sistema gera documentos Word contendo:

- Listagem de todas as páginas do relatório
- Detalhes dos visuais (gráficos, tabelas, etc.)
- Estrutura de tabelas e colunas
- Medidas e suas expressões DAX
- Fontes de dados utilizadas
- Relacionamentos entre tabelas

## 🛠️ Pré-requisitos

- Python 3.8 ou superior
- Pacotes Python:
  - `python-docx`
  - `pywin32`

Instale as dependências com:
```bash
pip install python-docx pywin32
```

## ⚙️ Configuração
Clone este repositório

Edite o arquivo config.py para definir:

Caminho base do seu OneDrive/arquivos

Localização do modelo Word (modelo.docx)

Pasta onde estão os arquivos .pbit

## 🚀 Como Usar
Coloque seus arquivos .pbit na pasta configurada

Execute o script principal:
```
python main.py
```

O sistema irá:

Processar cada arquivo .pbit encontrado

Criar uma pasta para cada relatório

Gerar um documento Word com a documentação completa

## 🔄 Processo de Geração

Converte .pbit para .zip (temporariamente)

Extrai os metadados do relatório

Analisa a estrutura do arquivo

Coleta informações sobre:

Páginas e visuais

Modelo de dados

Medidas DAX

Relacionamentos

Gera documento Word formatado

Organiza em pastas nomeadas conforme os relatórios

## ✨ Features

✅ Processamento em lote de múltiplos arquivos

✅ Substituição automática de versões anteriores

✅ Modelo Word customizável

✅ Extração completa de metadados

✅ Gerenciamento de versões de documentos

## 📝 Modelo Word

O arquivo modelo.docx deve conter os seguintes marcadores (que serão substituídos):


Nome do Relatório:
Data da documentação:
Páginas
Tabelas
Medidas
Visuais
Fontes
Relacionamentos 

## 🙏 Agradecimentos
[Julia Lira](https://github.com/data-ju/Power_BI_Documentation) pelo código original de extração de PBIT


## 📄 Licença
Este projeto está licenciado sob a MIT License - veja o arquivo [LICENSE](https://github.com/adilsonssdev/Gerando_DocPBIX_Via_PYTHON/edit/main/License) para detalhes.
