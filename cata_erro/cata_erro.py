import pandas as pd
import re
import sys

# Força UTF-8 no terminal (Windows)
sys.stdout.reconfigure(encoding='utf-8')

# Linhas de ruído que queremos ignorar
trechos_para_remover = [
    "OBS",
    "Início Criar conta",
    "Informação adicional",
    "Docs.que",
    "Empr.:",
    "Operação (Empresa",
    "Energia Compensada Positiva",
    "_________________________________________________________________________",
    "Faturamento residencial",
    "Instalação(ões)",
    "Erro interno:",
    "Erro durante leitura na tabela",
    "No total",
    "Fim    Criar conta:"
]

# Carrega o Excel
df = pd.read_excel("erros.xlsx")
coluna = df.columns[0].strip()

# Regex para identificar contas no formato (conta: 123456789012)
regex_conta = re.compile(r'\(conta:\s*(\d{12})\)')

resultados = []
conta_atual = None

# Varre de baixo para cima
for linha in reversed(df[coluna].tolist()):
    linha_str = str(linha)

    # Ignora linhas de ruído
    if any(t.lower() in linha_str.lower() for t in trechos_para_remover):
        continue

    # Procura conta no formato correto
    match_conta = regex_conta.search(linha_str)
    if match_conta:
        conta_atual = match_conta.group(1)  # <-- aqui captura apenas o número
        continue

    # Associa erros à conta atual
    if conta_atual:
        resultados.append({'CC': conta_atual, 'ERRO': linha_str})

# Mantém ordem original
resultados.reverse()

# Cria DataFrame final
df_final = pd.DataFrame(resultados)

# Salva em Excel
df_final.to_excel("erros_separados.xlsx", index=False)

print(f"Arquivo 'erros_separados.xlsx' criado com {len(df_final)} linhas!")
