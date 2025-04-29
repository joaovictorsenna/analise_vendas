import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Caminhos
CAMINHO_EXCEL = "./dados/vendas.xlsx"
CAMINHO_SAIDA = "/relatorios"

# Criar pasta de relatórios se não existir
os.makedirs(CAMINHO_SAIDA, exist_ok=True)

# 1. Ler os dados
try:
    df = pd.read_excel(CAMINHO_EXCEL)
except FileNotFoundError:
    print(
        f"Erro: O arquivo {CAMINHO_EXCEL} não foi encontrado. Verifique o caminho.")
    exit(1)  # Encerra o script se o arquivo não for encontrado

# 2. Verificar dados ausentes
if df.isnull().any().any():
    print("Atenção: Existem valores ausentes nos dados!")

# 3. Verificar as colunas essenciais
required_columns = ["Quantidade", "Preço Unitário", "Data"]
for col in required_columns:
    if col not in df.columns:
        print(f"Erro: A coluna '{col}' está ausente no arquivo.")
        exit(1)  # Encerra o script se alguma coluna necessária estiver ausente

# 4. Criar coluna de valor total
df["Valor Total"] = df["Quantidade"] * df["Preço Unitário"]
# Converte a data com tratamento de erro
df["Data"] = pd.to_datetime(df["Data"], errors='coerce')

# 5. Verificar se a conversão de data foi bem-sucedida
if df["Data"].isnull().any():
    print("Atenção: Algumas datas não foram convertidas corretamente.")
    # Você pode lidar com esses casos conforme necessário, como remover ou preencher valores nulos

# 6. Agrupamentos
vendas_por_vendedor = df.groupby(
    "Vendedor")["Valor Total"].sum().sort_values(ascending=False)
vendas_por_dia = df.groupby("Data")["Valor Total"].sum()

# 7. Gráfico: Total por Vendedor
plt.figure(figsize=(10, 6))
sns.barplot(x=vendas_por_vendedor.values,
            y=vendas_por_vendedor.index, palette="Blues_d")
plt.title("Total de Vendas por Vendedor")
plt.xlabel("Total em R$")
plt.ylabel("Vendedor")
plt.tight_layout()

# Exibir o gráfico antes de salvar
plt.show()
# Salvar o gráfico
plt.savefig(f"{CAMINHO_SAIDA}/vendas_por_vendedor.png")
plt.close()

# 8. Gráfico: Vendas por Dia
plt.figure(figsize=(10, 6))
vendas_por_dia.plot(marker='o', linestyle='-', color='tab:green')
plt.title("Evolução das Vendas por Dia")
plt.xlabel("Data")
plt.ylabel("Total em R$")
plt.grid(True)
plt.xticks(rotation=45)  # Rotaciona as datas para melhor leitura
plt.tight_layout()

# Exibir o gráfico antes de salvar
plt.show()
# Salvar o gráfico
plt.savefig(f"{CAMINHO_SAIDA}/vendas_por_dia.png")
plt.close()

print("Relatórios gerados com sucesso na pasta:", CAMINHO_SAIDA)
