"""
Conciliador Banco x Corporativo
--------------------------------
Como usar:
  1. Edite o config.yaml com os arquivos e o mês corretos.
  2. Coloque os dois arquivos Excel na mesma pasta.
  3. Abra o terminal nessa pasta e rode:  python conciliador.py
  4. Abra o arquivo de saída gerado.
"""
import pandas as pd, re, unicodedata, sys, os
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings; warnings.filterwarnings("ignore")

# ── lê config.yaml se existir, senão usa defaults ──────────────────────────
def carregar_config():
    defaults = {
        "arquivo_corporativo": "Base_Corporativo.xlsx",
        "aba_corporativo": "ControlePagamentos",
        "arquivo_banco": "Base_BTG_-_Londrina.xlsx",
        "aba_banco": "Extrato",
        "unidade": "Londrina",
        "banco": "BTG",
        "mes": 5,
        "saida": "Conciliacao_Londrina_Maio.xlsx",
        "limiar_nome": 60,
    }
    try:
        import yaml
        with open("config.yaml", encoding="utf-8") as f:
            dados = yaml.safe_load(f)
        defaults.update({k: v for k, v in dados.items() if v is not None})
    except ImportError:
        print("Aviso: PyYAML não instalado — usando config interno. Rode: pip install pyyaml")
    except FileNotFoundError:
        print("Aviso: config.yaml não encontrado — usando config interno.")
    return defaults

# ── normalização ────────────────────────────────────────────────────────────
def sa(s):
    return ''.join(c for c in unicodedata.normalize('NFKD', str(s)) if not unicodedata.combining(c))

def nn(s):
    s = sa(s).lower()
    s = re.sub(r'\-\s*qrcode.*', '', s)
    s = re.sub(r'\b[a-z0-9]{12,}\b', '', s)
    s = re.sub(r'\(\d+/\d+\)', '', s)
    s = re.sub(r'\bltda\b|\bs\.?a\.?\b|\bepp\b|\bme\b|\bmei\b', '', s)
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z ]', ' ', s)).strip()

def v2(x):
    try: return round(abs(float(x)), 2)
    except: return None

# ── leitura das bases ────────────────────────────────────────────────────────
def carregar_corporativo(cfg):
    try:
        c = pd.read_excel(cfg["arquivo_corporativo"], sheet_name=cfg["aba_corporativo"], header=0)
    except FileNotFoundError:
        sys.exit(f"ERRO: arquivo '{cfg['arquivo_corporativo']}' não encontrado na pasta.")
    cl = list(c.columns)
    c = c.rename(columns={cl[1]:"data_lanc", cl[6]:"rec", cl[7]:"serv",
                           cl[8]:"forma", cl[9]:"filial", cl[12]:"valor", cl[15]:"mes"})
    c["fn"] = c["filial"].apply(lambda x: sa(x).lower().strip() if pd.notna(x) else "")
    c["pn"] = c["forma"].apply(lambda x: sa(x).lower().strip() if pd.notna(x) else "")
    unidade_n = sa(cfg["unidade"]).lower().strip()
    banco_n   = sa(cfg["banco"]).lower().strip()
    sub = c[(c["fn"] == unidade_n) & (c["pn"] == banco_n) & (c["mes"] == cfg["mes"])].copy()
    if sub.empty:
        sys.exit(f"ERRO: nenhum registro encontrado para {cfg['unidade']} / {cfg['banco']} / Mês {cfg['mes']}. Verifique o config.yaml.")
    sub["va"] = sub["valor"].apply(v2)
    sub["rn"] = sub["rec"].apply(nn)
    return sub.reset_index(drop=True)

def carregar_banco(cfg):
    try:
        raw = pd.read_excel(cfg["arquivo_banco"], sheet_name=cfg["aba_banco"], header=None)
    except FileNotFoundError:
        sys.exit(f"ERRO: arquivo '{cfg['arquivo_banco']}' não encontrado na pasta.")
    h = next((i for i in range(len(raw))
               if raw.iloc[i].astype(str).str.contains("Data de lançamento", na=False).any()), None)
    if h is None:
        sys.exit("ERRO: coluna 'Data de lançamento' não encontrada no extrato. Verifique se o arquivo é o extrato correto.")
    b = pd.read_excel(cfg["arquivo_banco"], sheet_name=cfg["aba_banco"], header=h)
    b = b.loc[:, ~b.columns.astype(str).str.startswith("Unnamed")].dropna(how="all")
    b.columns = ["data", "desc", "valor", "saldo"][:len(b.columns)]
    b["valor"] = pd.to_numeric(b["valor"], errors="coerce")
    s = b[b["valor"] < 0].copy()
    pref = ["pagamento de boleto enviado para", "devolucao do pix enviado para",
            "pix enviado para", "ted enviado para", "pagamento de conta / tributo -"]
    def en(d):
        for p in pref:
            if sa(d).lower().startswith(p): return d[len(p):].strip()
        return str(d)
    s["rb"] = s["desc"].apply(en)
    s["rn"] = s["rb"].apply(nn)
    s["va"] = s["valor"].apply(v2)
    unidade_n = sa(cfg["unidade"]).lower().replace(" ", "")
    s["interna"] = s["rn"].str.contains("clinica capilar " + sa(cfg["unidade"]).lower(), regex=False)
    s["estorno"]  = s["desc"].apply(lambda d: sa(str(d)).lower().startswith("devolucao"))
    return s.reset_index(drop=True)

# ── motor de match 1:1 ───────────────────────────────────────────────────────
def conciliar(corp, banco, cfg):
    pares = []
    for i, cr in corp.iterrows():
        for j, br in banco.iterrows():
            if br["interna"] or br["estorno"]: continue
            vex = cr["va"] == br["va"]
            vsc = int(cr["va"] or 0) == int(br["va"] or 0)
            if not (vex or vsc): continue
            sim = fuzz.token_set_ratio(cr["rn"], br["rn"])
            if not vex and sim < cfg["limiar_nome"]: continue
            pares.append(((2 if vex else 1) * 1000 + sim, vex, sim, i, j))
    pares.sort(reverse=True)
    uc, ub, M = {}, {}, []
    for _, vex, sim, i, j in pares:
        if i in uc or j in ub: continue
        uc[i] = ub[j] = True
        M.append((i, j, vex, sim))
    conc    = [(i, j, sim) for i, j, vx, sim in M if corp.loc[i,"va"] == banco.loc[j,"va"]]
    dv      = [(i, j, sim) for i, j, vx, sim in M if corp.loc[i,"va"] != banco.loc[j,"va"]]
    corp_nm = [i for i in corp.index if i not in uc]
    banco_nm= [j for j in banco.index if j not in ub and not banco.loc[j,"interna"] and not banco.loc[j,"estorno"]]
    return conc, dv, corp_nm, banco_nm

# ── geração da planilha ──────────────────────────────────────────────────────
AZUL="1F4E79"; VERM="C00000"; VERDE_C="375623"

def eh(ws, ncol, titulo, cor=AZUL):
    ws.append([titulo])
    ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=ncol)
    c = ws.cell(ws.max_row, 1)
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=cor)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[ws.max_row].height = 22

def lc(ws, cols):
    ws.append(cols)
    for k in range(1, len(cols)+1):
        c = ws.cell(ws.max_row, k)
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center", wrap_text=True)

def aba(wb, nome, titulo, cabec, linhas, cor=AZUL):
    w = wb.create_sheet(nome)
    eh(w, len(cabec), titulo, cor)
    w.append([])
    lc(w, cabec)
    for ln in linhas:
        w.append(ln)
        r = w.max_row
        for k in range(1, len(cabec)+1):
            w.cell(r,k).font = Font(name="Arial")
            w.cell(r,k).alignment = Alignment(vertical="center", wrap_text=True)
    for k in range(1, len(cabec)+1):
        maxw = max([len(str(cabec[k-1]))] + [len(str(w.cell(rr,k).value or "")) for rr in range(4, w.max_row+1)])
        w.column_dimensions[get_column_letter(k)].width = min(max(maxw+2, 10), 50)
    return w

def inferir_metodo(desc):
    d = sa(str(desc)).lower()
    if d.startswith("pagamento de boleto"): return "Boleto"
    return "PIX"

def aba_para_lancar(wb, cfg, banco, banco_nm):
    """Aba no padrão exato do corporativo — copia e cola direto na base."""
    AMARELO = "FFF2CC"   # células que precisam de preenchimento manual
    NAVY_H  = "1F4E79"

    # Colunas na mesma ordem do corporativo (cols 0-15)
    COLS = [
        "",                                      # 0  Unnamed (fórmula — deixar vazio)
        "Data lançam.",                          # 1  auto: data do banco
        "Lançado por:",                          # 2  manual
        "Metodo PG",                             # 3  auto: inferido da descrição
        "Classificação Antiga",                  # 4  manual
        "Classificação\nPnL & FCF (Proposta RCT)", # 5  manual
        "Nome da empresa recebedora",            # 6  auto: recebedor do banco
        "Serviço prestado",                      # 7  manual
        "Forma de pagamento",                    # 8  auto: cfg banco
        "Filial pagadora",                       # 9  auto: cfg unidade
        "Será rateado para demais unidades?",    # 10 auto: Não
        "Data Vencimento",                       # 11 auto: data do banco
        "Valor item",                            # 12 auto: valor
        "Valor NF",                              # 13 auto: valor (igual)
        "Status Aprovação\n(Birro)",             # 14 manual
        "MÊS",                                  # 15 auto: cfg mes
    ]
    # índices que precisam de preenchimento manual (destacar em amarelo)
    MANUAL = {2, 4, 5, 7, 14}

    w = wb.create_sheet("Para Lançar no Corp", 1)  # segunda aba, logo após Resumo

    # Banner explicativo
    w.merge_cells("A1:P1")
    c = w["A1"]
    c.value = ("✅  Copie as linhas abaixo e cole na aba 'ControlePagamentos' da base do corporativo.  "
               "Células em AMARELO precisam ser preenchidas antes de colar.")
    c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY_H)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    w.row_dimensions[1].height = 24

    # Linha de cabeçalho (idêntica ao corporativo)
    w.append([""] * len(COLS))  # linha 2 = cabeçalho
    r_hdr = w.max_row
    for k, col in enumerate(COLS, 1):
        c = w.cell(r_hdr, k)
        c.value = col
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY_H)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    w.row_dimensions[r_hdr].height = 30

    thin = Side(style="thin", color="D9D9D9")
    borda = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Linhas de dados
    for j in banco_nm:
        br = banco.loc[j]
        row_vals = [
            "",                          # 0  fórmula — vazio
            br["data"],                  # 1  data lançamento = data do banco
            "",                          # 2  lançado por — MANUAL
            inferir_metodo(br["desc"]),  # 3  método PG
            "",                          # 4  classificação antiga — MANUAL
            "",                          # 5  classificação PnL — MANUAL
            br["rb"],                    # 6  nome recebedor
            "",                          # 7  serviço prestado — MANUAL
            cfg["banco"],               # 8  forma de pagamento
            cfg["unidade"],             # 9  filial
            "Não",                       # 10 será rateado
            br["data"],                  # 11 data vencimento = data do banco
            br["va"],                    # 12 valor item
            br["va"],                    # 13 valor NF
            "",                          # 14 status aprovação — MANUAL
            cfg["mes"],                 # 15 mês
        ]
        w.append(row_vals)
        r = w.max_row
        for k in range(1, len(COLS) + 1):
            cell = w.cell(r, k)
            cell.font = Font(name="Arial", size=11)
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            cell.border = borda
            if (k - 1) in MANUAL:
                cell.fill = PatternFill("solid", fgColor=AMARELO)
            # formatos
            if k in (2, 12):   # datas
                cell.number_format = "DD/MM/YYYY"
            if k in (13, 14):  # valores
                cell.number_format = '#,##0.00'

    # Larguras de coluna
    larguras = [4, 14, 14, 10, 32, 28, 30, 24, 14, 12, 10, 14, 13, 13, 12, 7]
    for k, w_col in enumerate(larguras, 1):
        w.column_dimensions[get_column_letter(k)].width = w_col

    # Totalizador de valor
    if banco_nm:
        r_tot = w.max_row + 1
        w.cell(r_tot, 12).value = f"TOTAL: {len(banco_nm)} lançamentos"
        w.cell(r_tot, 12).font = Font(name="Arial", bold=True, color=NAVY_H)
        w.cell(r_tot, 13).value = round(banco.loc[banco_nm, "va"].sum(), 2)
        w.cell(r_tot, 13).font = Font(name="Arial", bold=True, color=NAVY_H)
        w.cell(r_tot, 13).number_format = '#,##0.00'

    return w

def gerar_relatorio(cfg, corp, banco, conc, dv, corp_nm, banco_nm):
    wb = Workbook()
    ws = wb.active; ws.title = "Resumo"
    eh(ws, 4, f"Conciliação — {cfg['unidade']} / {cfg['banco']} / Mês {cfg['mes']}")
    ws.append([])
    lc(ws, ["Categoria", "Qtd", "Valor (R$)", "O que fazer"])
    linhas_res = [
        ("✓  Conciliados — valor e nome batem",
         len(conc), round(sum(corp.loc[i,"va"] for i,_,_ in conc), 2),
         "Nenhuma ação necessária"),
        ("~  Divergência de valor — mesmo recebedor, centavos diferem",
         len(dv), round(sum(banco.loc[j,"va"] for _,j,_ in dv), 2),
         "Conferir valor correto com o corporativo"),
        ("⚠  A REPORTAR — pago no banco, sem registro no corporativo",
         len(banco_nm), round(banco.loc[banco_nm,"va"].sum(), 2),
         "Enviar ao corporativo para lançamento"),
        ("⚠  Pendente — registrado no corp., sem pagamento no extrato",
         len(corp_nm), round(corp.loc[corp_nm,"va"].sum(), 2),
         "Verificar se foi pago via cartão / outra conta / ainda não pago"),
        ("—  Transferências internas (excluídas da conciliação)",
         int(banco["interna"].sum()), round(banco.loc[banco["interna"],"va"].sum(), 2),
         "Movimentação entre contas próprias — ignorada"),
    ]
    cores = {2: "FFF2CC", 3: "FFF2CC"}
    for idx, (cat, q, val, acao) in enumerate(linhas_res):
        ws.append([cat, q, val, acao])
        r = ws.max_row
        ws.cell(r, 3).number_format = '#,##0.00'
        for k in range(1, 5):
            ws.cell(r,k).font = Font(name="Arial")
            ws.cell(r,k).alignment = Alignment(wrap_text=True, vertical="center")
            if idx in cores:
                ws.cell(r,k).fill = PatternFill("solid", fgColor=cores[idx])
    for col, w in zip("ABCD", [50, 8, 16, 46]):
        ws.column_dimensions[col].width = w

    # ABA NO PADRÃO DO CORPORATIVO
    aba_para_lancar(wb, cfg, banco, banco_nm)

    # A REPORTAR
    linhas = [[str(banco.loc[j,"data"]), banco.loc[j,"rb"], banco.loc[j,"va"],
               banco.loc[j,"desc"], "A investigar", ""] for j in banco_nm]
    w = aba(wb, "A Reportar",
            "PAGOS NO BANCO SEM REGISTRO NO CORPORATIVO  →  reportar ao corporativo",
            ["Data", "Recebedor (banco)", "Valor (R$)", "Descrição original", "Status", "Observação"],
            linhas, VERM)
    for r in range(4, w.max_row+1): w.cell(r, 3).number_format = '#,##0.00'

    # DIVERGÊNCIA DE VALOR
    linhas = [[corp.loc[i,"rec"], corp.loc[i,"va"], banco.loc[j,"va"],
               round(banco.loc[j,"va"]-corp.loc[i,"va"], 2),
               str(banco.loc[j,"data"]), banco.loc[j,"rb"], "A investigar"] for i,j,_ in dv]
    w = aba(wb, "Divergencia de Valor",
            "MESMO RECEBEDOR, VALOR DIFERENTE  →  conferir",
            ["Recebedor (corp)", "Valor corp (R$)", "Valor banco (R$)", "Diferença (R$)",
             "Data banco", "Recebedor (banco)", "Status"],
            linhas)
    for r in range(4, w.max_row+1):
        for k in (2,3,4): w.cell(r,k).number_format = '#,##0.00'

    # PENDENTE
    linhas = [[corp.loc[i,"rec"], corp.loc[i,"serv"], corp.loc[i,"va"],
               "A investigar", ""] for i in corp_nm]
    aba(wb, "Pendente no Banco",
        "REGISTRADO NO CORPORATIVO SEM PAGAMENTO NO EXTRATO",
        ["Recebedor (corp)", "Serviço", "Valor (R$)", "Status", "Observação"],
        linhas)

    # CONCILIADOS
    linhas = [[corp.loc[i,"rec"], corp.loc[i,"va"],
               str(banco.loc[j,"data"]), banco.loc[j,"rb"], int(sim)] for i,j,sim in conc]
    aba(wb, "Conciliados",
        "CONFERIDOS — valor e recebedor batem (auditoria)",
        ["Recebedor (corp)", "Valor (R$)", "Data banco", "Recebedor (banco)", "Confiança nome %"],
        linhas, VERDE_C)

    # TRANSFERÊNCIAS INTERNAS
    linhas = [[str(banco.loc[j,"data"]), banco.loc[j,"rb"],
               banco.loc[j,"va"], banco.loc[j,"desc"]]
              for j in banco.index if banco.loc[j,"interna"]]
    aba(wb, "Transf Internas",
        "IGNORADAS — movimentação entre contas próprias",
        ["Data", "Recebedor", "Valor (R$)", "Descrição original"],
        linhas)

    wb.save(cfg["saida"])
    return cfg["saida"]

# ── main ─────────────────────────────────────────────────────────────────────
def main():
    cfg = carregar_config()
    print(f"\nConciliando: {cfg['unidade']} / {cfg['banco']} / Mês {cfg['mes']}")
    print(f"Corporativo : {cfg['arquivo_corporativo']}")
    print(f"Banco       : {cfg['arquivo_banco']}\n")
    corp  = carregar_corporativo(cfg)
    banco = carregar_banco(cfg)
    conc, dv, corp_nm, banco_nm = conciliar(corp, banco, cfg)
    saida = gerar_relatorio(cfg, corp, banco, conc, dv, corp_nm, banco_nm)
    print("=" * 55)
    print(f"  Conciliados          : {len(conc):>4}")
    print(f"  Divergência de valor : {len(dv):>4}")
    print(f"  A REPORTAR           : {len(banco_nm):>4}  (R$ {banco.loc[banco_nm,'va'].sum():,.2f})")
    print(f"  Pendente no banco    : {len(corp_nm):>4}  (R$ {corp.loc[corp_nm,'va'].sum():,.2f})")
    print("=" * 55)
    print(f"\nArquivo gerado: {saida}\n")

if __name__ == "__main__":
    main()