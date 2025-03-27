
#  Generador de Documentos em Power BI 

Automação para geração de documentação detalhada de relatórios Power BI a partir de arquivos `.pbit`.

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

##  🚀 Como Usar
Coloque seus arquivos .pbit na pasta configurada

Execute o script principal:
```
python main.py
```
Processar cada arquivo .pbit encontrado

Criar uma pasta para cada relatório

Gerar um documento Word com a documentação completa

##  🔄 Processo de Geração
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

Copy
Nome do Relatório:
Data da documentação:
Páginas
Tabelas
Medidas
Visuais
Fontes
Relacionamentos

## 📄 Licença
Este projeto está licenciado sob a MIT License - veja o arquivo LICENSE para detalhes.








