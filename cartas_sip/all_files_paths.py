import os

# Defina o caminho para a pasta
folder_path = r'C:\Users\s1002322\Syngenta\Brazil HROps - Comp Cycle - Zip_SIP'

# Percorre a árvore de diretórios
for root, dirs, files in os.walk(folder_path):
    for filename in files:
        # Constrói o caminho completo para o arquivo
        file_path = os.path.join(root, filename)
        print(file_path) # Isso imprimirá o caminho completo do arquivo
