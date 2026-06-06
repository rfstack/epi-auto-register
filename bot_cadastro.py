import os
import pyautogui
import pandas as pd
import pyperclip  # pip install pyperclip
from dotenv import load_dotenv

# ===== CONFIG =====
load_dotenv()
login = os.getenv("LOGIN")
senha = os.getenv("SENHA")
site  = os.getenv("SITE")

caminho = "Relação.xlsx"
aba = "Base_Normalizada"

COL_FUNCAO = "Codigo_Função"
COL_DESC   = "Descrição_EPI"
COL_CA     = "CA"
COL_TEMPO  = "Tempo"

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.15


# ===== FUNÇÕES =====
def colar(valor: str, limpar=True):
    if limpar:
        pyautogui.hotkey("ctrl", "a")
        pyautogui.sleep(0.05)
    pyperclip.copy("" if valor is None else str(valor))
    pyautogui.hotkey("ctrl", "v")
    pyautogui.sleep(0.15)

def normalizar_num_str(s: str) -> str:
    s = "" if s is None else str(s).strip()
    return s[:-2] if s.endswith(".0") else s


# ===== LÊ PLANILHA =====
df = pd.read_excel(caminho, sheet_name=aba, dtype=str)

faltando = [c for c in [COL_FUNCAO, COL_DESC, COL_CA, COL_TEMPO] if c not in df.columns]
if faltando:
    raise ValueError(f"Colunas faltando: {faltando}. Colunas encontradas: {list(df.columns)}")

for c in [COL_FUNCAO, COL_DESC, COL_CA, COL_TEMPO]:
    df[c] = df[c].fillna("").astype(str).str.strip()

df = df[(df[COL_FUNCAO] != "") & (df[COL_DESC] != "")].copy()
df[COL_CA]    = df[COL_CA].apply(normalizar_num_str)
df[COL_TEMPO] = df[COL_TEMPO].apply(normalizar_num_str)

# importante: manter ordem da planilha (groupby com sort=False)
# opcional: se preferir garantir agrupamento contíguo por EPI:
# df = df.sort_values(by=[COL_DESC, COL_FUNCAO], kind="stable")


# ===== ABRIR SITE / LOGIN =====
pyautogui.press("win"); pyautogui.sleep(1)
pyautogui.write("edge"); pyautogui.sleep(1)
pyautogui.press("enter"); pyautogui.sleep(2)
pyautogui.write(site); pyautogui.sleep(1)
pyautogui.press("enter"); pyautogui.sleep(8)

pyautogui.click(-395,390)  # login
colar(login)
pyautogui.press("tab")
colar(senha)
pyautogui.press("enter")
pyautogui.sleep(10)

# ===== NAVEGAÇÃO ATÉ A TELA DE CADASTRO EPI (SEU FLUXO) =====
pyautogui.click(-1341,206)
pyautogui.sleep(1)
pyautogui.press("tab", 2)
pyautogui.sleep(1)
pyautogui.press("enter")
pyautogui.sleep(1)
pyautogui.press("tab", 2)
pyautogui.sleep(1)
pyautogui.press("enter")
pyautogui.sleep(1)
pyautogui.click(-1206,419)
pyautogui.sleep(8)


# ===== AÇÕES DO SITE (SEPARADAS) =====
def abrir_epi(desc_epi: str):
    """
    Busca o EPI pela descrição e entra nas propriedades do EPI.
    (Seu fluxo atual: buscar -> enter -> entrar nas propriedades)
    """
    # campo de pesquisa/descrição
    pyautogui.click(-968,170)
    pyautogui.sleep(0.2)
    colar(desc_epi)
    print(desc_epi)
    pyautogui.press("enter")

    # entrar nas propriedades do EPI (seu fluxo)
    pyautogui.sleep(2)
    pyautogui.click(-41,329)
    pyautogui.sleep(1)
    pyautogui.click(-79,400)
    pyautogui.sleep(5)

    # clicar/ativar área do formulário
    pyautogui.click(-495,157)
    pyautogui.sleep(2)

    # ir para aba/seção onde cadastra função (seu clique)
    pyautogui.click(-150,249)
    pyautogui.sleep(2)


def cadastrar_uma_funcao(cod_funcao: str, tempo: str):
    """
    Preenche os campos da função (Código + Tempo + outras opções) e salva.
    Aqui é onde ele volta pra tela de "cadastrar função" após salvar.
    """

    # código da função (você já faz Tab e cola)
    pyautogui.click(-1041,277)
    colar(cod_funcao)
    print(cod_funcao)
    pyautogui.sleep(3)
    pyautogui.click(-1040,317)
    pyautogui.sleep(2)

    # (seu fluxo de selecionar/confirmar função)
    pyautogui.click(-668,280)
    pyautogui.sleep(1)
    pyautogui.click(-1056,319)
    pyautogui.sleep(2)
    pyautogui.write("1")
    pyautogui.sleep(1)

    # tempo
    pyautogui.click(-1115,371)
    pyautogui.sleep(1)
    pyautogui.press("tab")
    colar(tempo)
    pyautogui.sleep(1)

    # campo "troca" (tempo - 10)
    pyautogui.click(-1112,460)
    pyautogui.sleep(1)
    pyautogui.press("tab")
    pyautogui.sleep(1)

    try:
        valor_troca = int(str(tempo).strip()) - 10
    except (ValueError, TypeError):
        valor_troca = 0

    pyautogui.write(str(valor_troca))
    pyautogui.sleep(2)
    pyautogui.click(-389,575)
    pyautogui.sleep(4)  # espera salvar e voltar pra tela de cadastrar função
    

def sair_do_epi_para_buscar_outro():
    pyautogui.click(-258,210)
    pyautogui.sleep(2)
    pyautogui.scroll(-1000)
    pyautogui.sleep(3)
    pyautogui.click(-984,167)
    pyautogui.sleep(2)
    pyautogui.click(-66,545)
    pyautogui.sleep(8)
    print("Falta salvar")

    # opcional: limpar campo de busca para o próximo EPI
    pyautogui.click(-968,170)
    pyautogui.sleep(2)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("backspace")
    pyautogui.sleep(2)


# ===== LOOP POR EPI (GRUPOS) =====
for desc_epi, grupo in df.groupby(COL_DESC, sort=False):
    # 1) abre 1x o EPI
    abrir_epi(desc_epi)

    # 2) cadastra todas as funções desse EPI
    for _, row in grupo.iterrows():
        cod_funcao = row[COL_FUNCAO]
        tempo      = row[COL_TEMPO]
        cadastrar_uma_funcao(cod_funcao, tempo)

    # 3) terminou esse EPI -> sai e volta pra buscar outro
    sair_do_epi_para_buscar_outro()