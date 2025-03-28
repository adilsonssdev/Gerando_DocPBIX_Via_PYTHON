#### código correto para o arquivo config.py

import os
from pathlib import Path

# Configurações
nome_modelo_word = 'modelo.docx'

# Caminho base - usando Path para melhor compatibilidade
one_drive_path = Path.home() / 'OneDrive - ' / 'o-nome-do-seu-modelo-word-aqui(com a extensao do arquivo)' #Exemplo: r'C:\Users\user\Downloads'

# Caminhos corrigidos (ajuste conforme sua estrutura real)
caminho_BI = one_drive_path / 'Pastas do arquivo' / 'Documentacao'
caminho_modelo_word = caminho_BI / 'Modelos'  # Pasta Modelos está dentro de Contas a Receber
caminho_documentacao = caminho_BI  # Mesmo diretório do BI

# Lista para armazenar os caminhos dos arquivos pbit encontrados
arquivos_pbit = []

# Encontrar todos os arquivos .pbit no diretório e subdiretórios
print("\nProcurando arquivos .pbit...")
for root, dirs, files in os.walk(caminho_BI):
    for file in files:
        if file.endswith('.pbit'):
            arquivo_pbit = Path(root) / file
            arquivos_pbit.append(arquivo_pbit)
            print(f"Encontrado: {arquivo_pbit}")

# Verificar se foram encontrados arquivos
if not arquivos_pbit:
    print("Nenhum arquivo .pbit encontrado!")
else:
    print(f"\nTotal de arquivos .pbit encontrados: {len(arquivos_pbit)}")

# Processar cada arquivo pbit encontrado
for arquivo_pbit in arquivos_pbit:
    nome_BI = arquivo_pbit.stem  # Obtém o nome do arquivo sem extensão
    print(f"\nProcessando arquivo: {nome_BI}")
    
    # Definir caminhos para este arquivo específico
    arquivo_zip = arquivo_pbit.parent / f'{nome_BI}.zip'
    arquivo_documentacao = arquivo_pbit.parent / f'{nome_BI}_doc.docx'
    
    # Aqui você pode adicionar o código para processar cada arquivo pbit
    # e gerar a documentação correspondente
    
    print(f"Caminho do ZIP: {arquivo_zip}")
    print(f"Caminho da documentação: {arquivo_documentacao}")

# Verificação de caminhos adicionais (opcional)
print("\nVerificação de caminhos:")
print(f"Modelo Word: {caminho_modelo_word / nome_modelo_word}")
print(f"Pasta Modelos existe? {caminho_modelo_word.exists()}")
print(f"Conteúdo da pasta Modelos: {list(caminho_modelo_word.glob('*')) if caminho_modelo_word.exists() else 'Pasta não existe'}")

