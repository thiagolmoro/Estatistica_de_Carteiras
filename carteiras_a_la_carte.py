# pip install selenium webdriver-manager pandas xlsxwriter tqdm

# Código que calcula a soma total e mostra a média

import pandas as pd
import time
import locale
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from tqdm import tqdm

# Define o locale para português do Brasil
locale.setlocale(locale.LC_NUMERIC, "pt_BR.UTF-8")

# Configuração do Selenium para rodar sem abrir o navegador (modo headless)
chrome_options = Options()
chrome_options.add_argument("--headless")  # Roda sem abrir o navegador
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Instala e configura o WebDriver automaticamente
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Carregar o CSV
df = pd.read_csv("Carteiras_A_La_Carte_Recomendadas.csv", delimiter=",")

# Converter a data para datetime
df['Data_Recomendada'] = pd.to_datetime(df['Data_Recomendada'], format="%d/%m/%Y")

def obter_valor_carteira(numero_carteira, data_recomendada):
    url = f"https://tradergrafico.com.br/carteiras/?Simu={numero_carteira}"
    try:
        driver.get(url)
        time.sleep(25)  # Espera a página carregar completamente

        data_formatada = data_recomendada.strftime("%d/%m/%y")
        linhas = driver.find_elements(By.TAG_NAME, "tr")

        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            if colunas and data_formatada in colunas[0].text.strip():
                resultado_texto = colunas[1].text.strip()
                resultado_texto = resultado_texto.replace('.', '')  # Remove pontos separadores de milhar
                resultado_texto = resultado_texto.replace(',', '.')  # Substitui vírgula decimal por ponto
                
                try:
                    resultado_ajustado = int(float(locale.atof(resultado_texto)))  # Converte para inteiro
                except ValueError:
                    resultado_ajustado = 0  # Caso ocorra erro na conversão, atribui 0
                
                return resultado_ajustado, url

        return 0, url  # Retorna 0 caso não encontre um valor
    except Exception:
        return 0, url

# Criar barra de progresso personalizada
barra_de_progresso = tqdm(
    total=len(df['Carteira']),
    position=0,
    leave=True,
    desc="🔄 Processando carteiras",
    bar_format="{l_bar}{bar} {n_fmt}/{total_fmt} - {elapsed} ⏳",
    colour="cyan"
)

# Criar colunas vazias antes do loop
df["Resultado_Ajustado"] = 0
df["URL_Carteira"] = ""

# Usar loop for ao invés de apply()
for index, row in df.iterrows():
    df.at[index, "Resultado_Ajustado"], df.at[index, "URL_Carteira"] = obter_valor_carteira(
        row['Carteira'], row['Data_Recomendada']
    )
    barra_de_progresso.update(1)  # Atualiza corretamente a barra a cada carteira processada

barra_de_progresso.close()  # Fecha a barra de progresso após terminar

# Converter 'Resultado_Ajustado' para inteiro
df['Resultado_Ajustado'] = df['Resultado_Ajustado'].astype(int)

# Calcular soma total e média
total_resultado = df['Resultado_Ajustado'].sum()
media_resultado = df['Resultado_Ajustado'].mean()

df.loc['Total'] = ['', '', total_resultado, '']
df.loc['Média'] = ['', '', media_resultado, '']

# Formatar os números para padrão brasileiro (milhar com ponto, sem decimais)
df['Resultado_Ajustado'] = df['Resultado_Ajustado'].apply(lambda x: f"{int(x):,}".replace(',', '.'))
df.loc['Total', 'Resultado_Ajustado'] = f"{int(total_resultado):,}".replace(',', '.')
df.loc['Média', 'Resultado_Ajustado'] = f"{int(media_resultado):,}".replace(',', '.')

# Certificar-se de que 'Data_Recomendada' é datetime antes de formatar
df['Data_Recomendada'] = pd.to_datetime(df['Data_Recomendada'], errors='coerce', dayfirst=True)
df['Data_Recomendada'] = df['Data_Recomendada'].dt.strftime('%d/%m/%Y')

# Salvar em CSV
df.to_csv("Carteiras_Com_Valores_Ajustados_ALC.csv", index=False, sep=';')

# Salvar em Excel com formatação
with pd.ExcelWriter("Carteiras_Com_Valores_Ajustados_ALC.xlsx", engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Resultados", index=False)
    workbook = writer.book
    worksheet = writer.sheets["Resultados"]
    
    for i, col in enumerate(df.columns):
        column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, column_width)

# Fechar o navegador
driver.quit()

print("✅ Processo concluído! Resultados salvos com soma e média.")
