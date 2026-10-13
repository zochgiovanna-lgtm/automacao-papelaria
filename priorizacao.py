import os
import json
import gspread
import pandas as pd
import numpy as np

# 1. Autenticação segura
credenciais_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
gc = gspread.service_account_from_dict(credenciais_json)

# 2. Abrir a planilha (Lendo a linha 2 como cabeçalho)
planilha = gc.open('Pedidos_Papelaria')
aba_origem = planilha.worksheet('Página1')

try:
    dados = aba_origem.get_all_records(head=2)
except gspread.exceptions.GSpreadException:
    # Normalmente acontece quando existem colunas sem cabeçalho (título em
    # branco) ou cabeçalhos repetidos na linha 2 da Página1. Em vez de supor
    # uma ordem fixa de colunas, lemos o cabeçalho real da planilha (seja
    # qual for a ordem que ela estiver) e só corrigimos o que está vazio ou
    # duplicado — preservando a leitura correta de todas as colunas normais,
    # como "Concluido", no lugar certo.
    cabecalho_bruto = aba_origem.row_values(2)
    cabecalho_corrigido = []
    contador_vazios = 0
    vistos = {}
    for titulo in cabecalho_bruto:
        titulo_limpo = titulo.strip()
        if titulo_limpo == '':
            contador_vazios += 1
            titulo_limpo = f'_coluna_vazia_{contador_vazios}'
        if titulo_limpo in vistos:
            vistos[titulo_limpo] += 1
            titulo_limpo = f'{titulo_limpo}_{vistos[titulo_limpo]}'
        else:
            vistos[titulo_limpo] = 0
        cabecalho_corrigido.append(titulo_limpo)

    print(f"[DIAGNÓSTICO] Cabeçalho da Página1 tinha célula(s) vazia/duplicada. Corrigido para: {cabecalho_corrigido}")
    dados = aba_origem.get_all_records(head=2, expected_headers=cabecalho_corrigido)

tabela_pedidos = pd.DataFrame(dados)

# 3. GARANTIA DE COLUNAS EXISTENTES
colunas_necessarias = ['Data_Pedido', 'Celular', 'Concluido', 'Observações', 'Quantidade', 'Pago', 'Valor']
for col in colunas_necessarias:
    if col not in tabela_pedidos.columns:
        tabela_pedidos[col] = ''

# 4. Limpeza de linhas fantasmas
tabela_pedidos = tabela_pedidos[tabela_pedidos['Cliente_ID'].astype(str).str.strip() != '']

# Separar Concluídos de Pendentes
print(f"[DIAGNÓSTICO] Colunas lidas da Página1: {list(tabela_pedidos.columns)}")
print(f"[DIAGNÓSTICO] Total de linhas lidas (após limpeza de fantasmas): {len(tabela_pedidos)}")
print(f"[DIAGNÓSTICO] Valores únicos encontrados na coluna 'Concluido': {tabela_pedidos['Concluido'].astype(str).unique().tolist()}")

condicao_concluido = tabela_pedidos['Concluido'].astype(str).str.strip().str.upper().isin(['TRUE', 'SIM', 'VERDADEIRO', 'PRONTO'])
tabela_concluidos = tabela_pedidos[condicao_concluido].copy()
tabela_pendentes = tabela_pedidos[~condicao_concluido].copy()

print(f"[DIAGNÓSTICO] Pedidos identificados como CONCLUÍDOS nesta execução: {len(tabela_concluidos)}")
if len(tabela_concluidos) > 0:
    print(f"[DIAGNÓSTICO] Clientes dos concluídos: {tabela_concluidos['Cliente_ID'].tolist()}")

# ==========================================
# PARTE A: ATUALIZAR ABA HISTÓRICO DE ENTREGAS
# ==========================================
colunas_historico = ['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Concluido', 'Pago', 'Valor']
for col in colunas_historico:
    if col not in tabela_concluidos.columns:
        tabela_concluidos[col] = ''
tabela_final_historico = tabela_concluidos[colunas_historico]

try:
    aba_historico = planilha.worksheet('Historico_Entregas')
except gspread.exceptions.WorksheetNotFound:
    aba_historico = planilha.add_worksheet(title="Historico_Entregas", rows="200", cols="20")

# Conta quantas linhas já existem no histórico (a partir da linha 3), para não
# sobrescrever o que já foi registrado em execuções anteriores.
valores_historico_atual = aba_historico.get_all_values()
linhas_ja_existentes = max(0, len(valores_historico_atual) - 2)  # desconta linhas 1 e 2 (cabeçalho)

dados_historico = tabela_final_historico.values.tolist()
if dados_historico:
    # append_rows adiciona ao final, sem apagar os pedidos já registrados antes.
    aba_historico.append_rows(dados_historico, value_input_option='USER_ENTERED')

total_entregues = linhas_ja_existentes + len(dados_historico)
texto_contador = f"🎉 Pedidos Entregues: {total_entregues}"

# ------------------------------------------
# Total já vendido (soma da coluna Valor de todos os pedidos concluídos,
# incluindo os já registrados em execuções anteriores + os novos de agora)
# ------------------------------------------
def parse_valor(v):
    """Converte texto de valor (ex: 'R$ 25,90', '1.234,56', '25.90') em float."""
    s = str(v).strip().replace('R$', '').replace(' ', '')
    if not s:
        return 0.0
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0

indice_coluna_valor = colunas_historico.index('Valor')  # posição padrão (fallback)

# Em vez de confiar cegamente na posição fixa, procura o cabeçalho "Valor" de
# verdade na linha 2 da planilha — assim, mesmo que uma coluna tenha sido
# inserida/movida manualmente, a soma continua batendo com a coluna certa.
if len(valores_historico_atual) >= 2:
    linha_cabecalho = valores_historico_atual[1]  # linha 2 (índice 1)
    for i, titulo in enumerate(linha_cabecalho):
        if titulo.strip().lower() == 'valor':
            indice_coluna_valor = i
            break

linhas_com_valor_existente = [
    (i + 3, linha[indice_coluna_valor])  # +3 porque a planilha começa a contar na linha 3
    for i, linha in enumerate(valores_historico_atual[2:])
    if len(linha) > indice_coluna_valor and linha[indice_coluna_valor].strip() != ''
]

print(f"[DIAGNÓSTICO] índice da coluna Valor usado: {indice_coluna_valor}")
print(f"[DIAGNÓSTICO] total de linhas já existentes no histórico: {len(valores_historico_atual[2:])}")
print(f"[DIAGNÓSTICO] linhas com Valor preenchido encontradas: {len(linhas_com_valor_existente)}")
for numero_linha, valor_bruto in linhas_com_valor_existente[:20]:
    print(f"[DIAGNÓSTICO]   linha {numero_linha}: '{valor_bruto}' -> {parse_valor(valor_bruto)}")

valor_ja_existente = sum(
    parse_valor(linha[indice_coluna_valor])
    for linha in valores_historico_atual[2:]  # pula as linhas 1 e 2 (cabeçalho)
    if len(linha) > indice_coluna_valor
)

print(f"[DIAGNÓSTICO] valor_ja_existente somado: {valor_ja_existente}")
print(f"[DIAGNÓSTICO] valor_novo (deste lote): {tabela_concluidos['Valor'].apply(parse_valor).sum() if 'Valor' in tabela_concluidos.columns else 0.0}")
valor_novo = tabela_concluidos['Valor'].apply(parse_valor).sum() if 'Valor' in tabela_concluidos.columns else 0.0

total_valor_vendido = valor_ja_existente + valor_novo
texto_valor_total = f"💰 Total já vendido: R$ {total_valor_vendido:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

# O quadro de resumo está mesclado em K1:M2 (não mais J1:L2), então a
# atualização precisa mirar em K1, que é a célula-âncora do bloco mesclado.
texto_quadro = f"{texto_contador}\n{texto_valor_total}"

try:
    aba_historico.update('K1', [[texto_quadro]], value_input_option='USER_ENTERED')
except Exception as e:
    print(f"Erro ao atualizar quadro de resumo (contador + total vendido): {e}")

# ------------------------------------------
# RESUMO MENSAL: pedidos entregues e total vendido por mês de entrega.
# Recalculado do zero a cada execução, lendo todo o Historico_Entregas.
# ------------------------------------------
try:
    from datetime import datetime

    def parse_data_entrega(texto, hoje_dt):
        """Aceita '28/07/2026'. Se vier sem ano ('15/09'), assume o ano atual
        (ou o anterior, se a data ficaria mais de 6 meses no futuro)."""
        s = str(texto).strip()
        if not s:
            return None
        for formato in ('%d/%m/%Y', '%d/%m/%y'):
            try:
                data_lida = datetime.strptime(s, formato)
            except ValueError:
                continue
            # Datas muito antigas (ex.: 01/01/1900) vêm de números soltos que o
            # Sheets converteu em data. Não são entregas reais.
            return data_lida if data_lida.year >= 2020 else None
        partes = s.replace('-', '/').split('/')
        if len(partes) == 2:
            try:
                dia, mes = int(partes[0]), int(partes[1])
                data = datetime(hoje_dt.year, mes, dia)
                if (data - hoje_dt).days > 180:
                    data = datetime(hoje_dt.year - 1, mes, dia)
                return data
            except ValueError:
                return None
        return None

    valores_finais = aba_historico.get_all_values()
    cabecalho_hist = valores_finais[1] if len(valores_finais) >= 2 else []
    idx_data_entrega = colunas_historico.index('Data_Entrega')
    for i, titulo in enumerate(cabecalho_hist):
        if titulo.strip().lower() == 'data_entrega':
            idx_data_entrega = i
            break

    agora = datetime.now()
    por_mes = {}   # (ano, mes) -> [quantidade, valor]
    sem_data = [0, 0.0]

    for linha in valores_finais[2:]:
        if not any(celula.strip() for celula in linha[:len(colunas_historico)]):
            continue  # linha vazia
        texto_data = linha[idx_data_entrega] if len(linha) > idx_data_entrega else ''
        valor_linha = parse_valor(linha[indice_coluna_valor]) if len(linha) > indice_coluna_valor else 0.0
        data = parse_data_entrega(texto_data, agora)
        if data is None:
            sem_data[0] += 1
            sem_data[1] += valor_linha
        else:
            chave = (data.year, data.month)
            por_mes.setdefault(chave, [0, 0.0])
            por_mes[chave][0] += 1
            por_mes[chave][1] += valor_linha

    # ---- SAÍDAS (gastos): aba "Saidas", preenchida manualmente pela cliente ----
    # Colunas: A=Data, B=Descrição, C=Categoria, D=Valor (cabeçalho na linha 2).
    try:
        aba_saidas = planilha.worksheet('Saidas')
    except gspread.exceptions.WorksheetNotFound:
        aba_saidas = planilha.add_worksheet(title='Saidas', rows="500", cols="6")
        aba_saidas.update(values=[['SAÍDAS (GASTOS)'], ['Data', 'Descrição', 'Categoria', 'Valor']],
                          range_name='A1', value_input_option='RAW')
        try:
            aba_saidas.format('A1:D2', {'textFormat': {'bold': True}})
        except Exception:
            pass

    gastos_mes = {}        # (ano, mes) -> total
    gastos_cat = {}        # categoria -> {(ano, mes) ou None: total}
    gasto_sem_data = 0.0
    for linha_s in aba_saidas.get_all_values()[2:]:
        linha_s = linha_s + [''] * (4 - len(linha_s))
        if not any(c.strip() for c in linha_s[:4]):
            continue
        valor_gasto = parse_valor(linha_s[3])
        categoria = linha_s[2].strip().upper() or 'SEM CATEGORIA'
        data_gasto = parse_data_entrega(linha_s[0], agora)
        chave_g = (data_gasto.year, data_gasto.month) if data_gasto else None
        if chave_g:
            gastos_mes[chave_g] = gastos_mes.get(chave_g, 0.0) + valor_gasto
        else:
            gasto_sem_data += valor_gasto
        gastos_cat.setdefault(categoria, {})
        gastos_cat[categoria][chave_g] = gastos_cat[categoria].get(chave_g, 0.0) + valor_gasto

    # ---- Resumo_Mensal: vendas x gastos x sobra ----
    linhas_resumo = [['Mês', 'Pedidos entregues', 'Total vendido (R$)', 'Gastos (R$)', 'Sobra (R$)']]
    for chave in sorted(set(por_mes) | set(gastos_mes)):
        qtd, vendido = por_mes.get(chave, [0, 0.0])
        gasto = gastos_mes.get(chave, 0.0)
        linhas_resumo.append([f"{chave[1]:02d}/{chave[0]}", qtd, round(vendido, 2),
                              round(gasto, 2), round(vendido - gasto, 2)])
    if sem_data[0] > 0 or gasto_sem_data > 0:
        linhas_resumo.append(['Sem data válida', sem_data[0], round(sem_data[1], 2),
                              round(gasto_sem_data, 2), round(sem_data[1] - gasto_sem_data, 2)])
    total_qtd = sum(v[0] for v in por_mes.values()) + sem_data[0]
    total_val = sum(v[1] for v in por_mes.values()) + sem_data[1]
    total_gasto = sum(gastos_mes.values()) + gasto_sem_data
    linhas_resumo.append(['TOTAL', total_qtd, round(total_val, 2), round(total_gasto, 2),
                          round(total_val - total_gasto, 2)])

    try:
        aba_resumo = planilha.worksheet('Resumo_Mensal')
    except gspread.exceptions.WorksheetNotFound:
        aba_resumo = planilha.add_worksheet(title='Resumo_Mensal', rows="100", cols="10")

    aba_resumo.batch_clear(['A1:E100'])
    aba_resumo.update(values=linhas_resumo, range_name='A1', value_input_option='RAW')
    try:
        aba_resumo.format('C2:E100', {'numberFormat': {'type': 'CURRENCY', 'pattern': '"R$" #,##0.00'}})
        aba_resumo.format('A1:E1', {'textFormat': {'bold': True}})
    except Exception as e_fmt:
        print(f"Aviso: não foi possível formatar a aba Resumo_Mensal: {e_fmt}")

    # ---- Gastos_por_Categoria: categorias nas linhas, meses nas colunas ----
    meses_g = sorted({k for d in gastos_cat.values() for k in d if k is not None})
    tem_sem_data = any(None in d for d in gastos_cat.values())
    cabecalho_cat = ['Categoria'] + [f"{m:02d}/{a}" for (a, m) in meses_g]
    if tem_sem_data:
        cabecalho_cat.append('Sem data')
    cabecalho_cat.append('Total')
    linhas_cat = [cabecalho_cat]
    for categoria, d in sorted(gastos_cat.items(), key=lambda x: -sum(x[1].values())):
        linha_c = [categoria] + [round(d.get(k, 0.0), 2) for k in meses_g]
        if tem_sem_data:
            linha_c.append(round(d.get(None, 0.0), 2))
        linha_c.append(round(sum(d.values()), 2))
        linhas_cat.append(linha_c)

    try:
        aba_cat = planilha.worksheet('Gastos_por_Categoria')
    except gspread.exceptions.WorksheetNotFound:
        aba_cat = planilha.add_worksheet(title='Gastos_por_Categoria', rows="200", cols="30")
    aba_cat.batch_clear(['A1:Z200'])
    aba_cat.update(values=linhas_cat, range_name='A1', value_input_option='RAW')
    try:
        aba_cat.format('B2:Z200', {'numberFormat': {'type': 'CURRENCY', 'pattern': '"R$" #,##0.00'}})
        aba_cat.format('A1:Z1', {'textFormat': {'bold': True}})
    except Exception as e_fmt:
        print(f"Aviso: não foi possível formatar a aba Gastos_por_Categoria: {e_fmt}")

    print(f"Resumo_Mensal atualizado: {len(set(por_mes) | set(gastos_mes))} mês(es), {sem_data[0]} pedido(s) sem data válida.")
    print(f"Gastos_por_Categoria atualizado: {len(gastos_cat)} categoria(s).")
except Exception as e:
    print(f"Erro ao atualizar Resumo_Mensal: {e}")


# ==========================================
# PARTE A.1: REMOVER PEDIDOS CONCLUÍDOS DA PÁGINA1
# ==========================================
# Só chega até aqui depois que o Historico_Entregas já foi atualizado acima,
# então mesmo que essa etapa falhe, os dados dos pedidos concluídos já
# estão salvos em Historico_Entregas.
colunas_pagina1 = ['Data_Pedido', 'Celular', 'Cliente_ID', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Concluido', 'Pago', 'Valor']
for col in colunas_pagina1:
    if col not in tabela_pendentes.columns:
        tabela_pendentes[col] = ''

tabela_pagina1_restante = tabela_pendentes[colunas_pagina1]

try:
    aba_origem.batch_clear(['A3:J1000'])
    dados_pagina1_restante = tabela_pagina1_restante.values.tolist()
    if dados_pagina1_restante:
        aba_origem.update('A3', dados_pagina1_restante)
    print(f"Página1 atualizada: {len(dados_pagina1_restante)} pedido(s) pendente(s) mantido(s).")
except Exception as e:
    print(f"Erro ao limpar pedidos concluídos da Página1: {e}")


# ==========================================
# PARTE B: PROCESSAR FILA DE PRIORIDADE (Apenas Pendentes)
# ==========================================
regras_tempos = {
    'CARTÃO DE VISITA':                 {'fixo': 30, 'unitario': 0.5},
    'IMPRESSÃO':                        {'fixo': 5,  'unitario': 0.1},
    'CANECA':                           {'fixo': 0,  'unitario': 40},
    'AGENDA':                           {'fixo': 0,  'unitario': 300},
    'CAMISETA':                         {'fixo': 0,  'unitario': 20},
    'TOPO DE BOLO':                     {'fixo': 45, 'unitario': 15},
    'BALÃO BUBBLE ELABORADO':           {'fixo': 0,  'unitario': 130},
    'PAPEL ADESIVO COM CORTE':          {'fixo': 0, 'unitario': 5},
    'CRACHÁ SIMPLES COM PLASTIFICAÇÃO': {'fixo': 0, 'unitario': 15},
    'CONVITE SIMPLES':                  {'fixo': 0, 'unitario': 20},
    'CONVITE ELABORADO COM CORTE':      {'fixo': 0, 'unitario': 60},
    'CAIXA PADRINHO':                   {'fixo': 0, 'unitario': 60},
    'CARD SIMPLES':                     {'fixo': 0, 'unitario': 5},
    'ETIQUETA ROUPA':                   {'fixo': 0, 'unitario': 60},
    'ETIQUETA SIMPLES SEM LAMINAÇÃO':   {'fixo': 0, 'unitario': 30},
    'APLICAÇÃO NOME CAMISETA':          {'fixo': 0, 'unitario': 30},
    'BLOCO DE PEDIDOS':                 {'fixo': 0, 'unitario': 60},
    'CADERNETA DE VACINA REFORMA':      {'fixo': 0, 'unitario': 240},
    'CARD COM CHOCOLATE':               {'fixo': 0, 'unitario': 10},
    'FLAY SIMPLES':                     {'fixo': 0, 'unitario': 10},
    'PLASTIFICAÇÃO':                    {'fixo': 0, 'unitario': 10},
    'REFORMA AGENDA ESCOLAR':           {'fixo': 0, 'unitario': 30},
    'CONVITE CASAMENTO':                {'fixo': 4320, 'unitario': 0},
    'CORTE LETRAS COLOR PLUSS':         {'fixo': 0, 'unitario': 30},
    'COMANDA':                          {'fixo': 4320, 'unitario': 0},
    'APOSTILA COM IMPRESSÃO':           {'fixo': 0, 'unitario': 240},
    'ENVELOPE COM VALE PRESENTE':       {'fixo': 0, 'unitario': 30},
    'VALE PRESENTE':                    {'fixo': 0, 'unitario': 30},
    'FOTO POLAROID':                    {'fixo': 0, 'unitario': 5},
    'FOTO POLAROID IMÃ DE GELADEIRA':   {'fixo': 0, 'unitario': 5},
    'TAG AGRADECIMENTO':                {'fixo': 0, 'unitario': 5},
    'IMPRESSÃO DE CERTIFICADO':         {'fixo': 0, 'unitario': 5},
    'TAGS ADESIVO PERSONALIZADOS':      {'fixo': 0, 'unitario': 10},
    'FOTO POLAROID DE GELADEIRA':       {'fixo': 0, 'unitario': 5},
    'ESTAMPA DTF':                      {'fixo': 0, 'unitario': 10},
    'ADESIVO DE VINIL':                 {'fixo': 0, 'unitario': 5},
    'BLOQUINHO COLOR PLUS':             {'fixo': 0, 'unitario': 30},
    'CADERNO A5':                       {'fixo': 0,  'unitario': 90},   # 🔶 ESTIMATIVA — validar com a cliente
    'CADERNO A4':                       {'fixo': 0,  'unitario': 120},  # 🔶 ESTIMATIVA — validar com a cliente
    'CALENDARIO GELADEIRA MINI':        {'fixo': 0,  'unitario': 20},   # 🔶 ESTIMATIVA — validar com a cliente
    'CALENDARIO GELADEIRA MEDIO':       {'fixo': 0,  'unitario': 30},   # 🔶 ESTIMATIVA — validar com a cliente
    'CALENDARIO DE GELADEIRA A4':       {'fixo': 0,  'unitario': 45},   # 🔶 ESTIMATIVA — validar com a cliente
    'QUADRO MDF A4':                    {'fixo': 0,  'unitario': 40},   # 🔶 ESTIMATIVA — validar com a cliente
    'TOPO DE BOLO CENARIO':             {'fixo': 60, 'unitario': 20},   # 🔶 ESTIMATIVA — validar com a cliente
    'MARCA PAGINA MAGNETICO':           {'fixo': 0,  'unitario': 10},  # 🔶 ESTIMATIVA — validar com a cliente
    'CAIXA C/ CANECA':                  {'fixo': 0,  'unitario': 60},
    'CORTE LETRA':                      {'fixo': 0,  'unitario': 20},   # 🔶 ESTIMATIVA — validar com a cliente
    'APENAS CORTE':                     {'fixo': 0,  'unitario': 20},   # 🔶 ESTIMATIVA — validar com a cliente
    'APENAS MONTAGEM':                  {'fixo': 0,  'unitario': 20},   # 🔶 ESTIMATIVA — validar com a cliente
    'ITEM DA LOJA':                     {'fixo': 0,  'unitario': 20},   # 🔶 ESTIMATIVA — validar com a cliente
    'APENAS ARTE':                      {'fixo': 0,  'unitario': 20},   # 🔶 ESTIMATIVA — validar com a cliente
    'BRINDES':                          {'fixo': 0,  'unitario': 20},   # 🔶 ESTIMATIVA — validar com a cliente
    'TAG COM CORTE':                    {'fixo': 0,  'unitario': 10},   # 🔶 ESTIMATIVA — validar com a cliente
    'ARTE':                             {'fixo': 0,  'unitario': 20},   # 🔶 ESTIMATIVA — validar com a cliente
    'ARTE + IMPRESSÃO':                 {'fixo': 0,  'unitario': 30},   # 🔶 ESTIMATIVA — validar com a cliente
    'EMBALAGEM PARA PRESENTE':          {'fixo': 0,  'unitario': 15}    # 🔶 ESTIMATIVA — validar com a cliente
}

if not tabela_pendentes.empty:
    tabela_pendentes['Data_Calc'] = pd.to_datetime(tabela_pendentes['Data_Entrega'], format='%d/%m/%Y', errors='coerce')
    hoje = pd.Timestamp.today().normalize()
    tabela_pendentes['Dias_Para_Entrega'] = (tabela_pendentes['Data_Calc'] - hoje).dt.days
    tabela_pendentes['Dias_Para_Entrega'] = tabela_pendentes['Dias_Para_Entrega'].fillna(1)
    tabela_pendentes['Dias_Para_Entrega'] = tabela_pendentes['Dias_Para_Entrega'].apply(lambda x: 1 if x <= 0 else x)

    tabela_pendentes['Quantidade_Limpa'] = tabela_pendentes['Quantidade'].astype(str).str.extract(r'(\d+)')[0]
    tabela_pendentes['Quantidade_Limpa'] = pd.to_numeric(tabela_pendentes['Quantidade_Limpa'], errors='coerce').fillna(1)
    tabela_pendentes['Quantidade_Limpa'] = tabela_pendentes['Quantidade_Limpa'].replace(0, 1)

    def calcular_tempo_real(linha):
        produto = str(linha['Produto']).strip().lower()
        qtd = linha['Quantidade_Limpa']
        for chave_regra in regras_tempos:
            if chave_regra.lower() == produto:
                tempo_fixo = regras_tempos[chave_regra]['fixo']
                tempo_unitario = regras_tempos[chave_regra]['unitario']
                return tempo_fixo + (tempo_unitario * qtd)
        return 15

    tabela_pendentes['Tempo_Total_Minutos'] = tabela_pendentes.apply(calcular_tempo_real, axis=1)

    nota_interna = tabela_pendentes['Tempo_Total_Minutos'] / tabela_pendentes['Dias_Para_Entrega']

    def definir_status(nota):
        if nota >= 30:
            return '🚨 Urgente'
        elif nota >= 15:
            return '⚠️ Atenção'
        else:
            return '✅ Em dia'

    tabela_pendentes['Status_Urgencia'] = nota_interna.apply(definir_status)
    tabela_pendentes['Quantidade'] = tabela_pendentes['Quantidade_Limpa']
    tabela_pendentes['_nota_oculta'] = nota_interna
    tabela_organizada = tabela_pendentes.sort_values(by='_nota_oculta', ascending=False)
else:
    tabela_organizada = pd.DataFrame(columns=['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Tempo_Total_Minutos', 'Status_Urgencia', 'Pago', 'Valor'])

colunas_finais = ['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Tempo_Total_Minutos', 'Status_Urgencia', 'Pago', 'Valor']
tabela_final_sheets = tabela_organizada[colunas_finais] if not tabela_organizada.empty else pd.DataFrame(columns=colunas_finais)

try:
    aba_destino = planilha.worksheet('Fila_Prioridade')
    try:
        aba_destino.batch_clear(['A3:K1000'])
    except Exception:
        pass
except gspread.exceptions.WorksheetNotFound:
    aba_destino = planilha.add_worksheet(title="Fila_Prioridade", rows="100", cols="20")

dados_prioridade = tabela_final_sheets.values.tolist()
if dados_prioridade:
    aba_destino.update('A3', dados_prioridade)


# ==========================================
# PARTE C: DISTRIBUIR PEDIDOS PELAS ABAS DE DIA DA SEMANA
# ==========================================
CAPACIDADE_DIARIA_MINUTOS = 480  # 8 horas — ajuste aqui se necessário

NOMES_DIAS = {
    0: 'SEGUNDA FEIRA',
    1: 'TERÇA FEIRA',
    2: 'QUARTA FEIRA',
    3: 'QUINTA FEIRA',
    4: 'SEXTA FEIRA',
    5: 'SÁBADO',
}

if not tabela_organizada.empty:
    dia_semana_hoje = pd.Timestamp.today().weekday()
    inicio = dia_semana_hoje if dia_semana_hoje < 6 else 0
    ordem_dias = [(inicio + i) % 6 for i in range(6)]
    nomes_em_ordem = [NOMES_DIAS[d] for d in ordem_dias]

    capacidade_restante = {nome: CAPACIDADE_DIARIA_MINUTOS for nome in nomes_em_ordem}
    pedidos_por_dia = {nome: [] for nome in nomes_em_ordem}

    for _, linha in tabela_organizada.iterrows():
        tempo = linha['Tempo_Total_Minutos']
        dia_escolhido = next(
            (nome for nome in nomes_em_ordem if capacidade_restante[nome] >= tempo),
            None
        )
        if dia_escolhido is None:
            dia_escolhido = max(capacidade_restante, key=capacidade_restante.get)

        capacidade_restante[dia_escolhido] -= tempo
        pedidos_por_dia[dia_escolhido].append(linha)

    for nome_aba in nomes_em_ordem:
        try:
            aba_dia = planilha.worksheet(nome_aba)
            try:
                aba_dia.batch_clear(['A3:K1000'])
            except Exception:
                pass
        except gspread.exceptions.WorksheetNotFound:
            aba_dia = planilha.add_worksheet(title=nome_aba, rows="100", cols="20")

        linhas_do_dia = pedidos_por_dia[nome_aba]
        if linhas_do_dia:
            tabela_dia = pd.DataFrame(linhas_do_dia)[colunas_finais]
            aba_dia.update('A3', tabela_dia.values.tolist())

    print("Distribuição por dia da semana concluída!")
else:
    print("Nenhum pedido pendente para distribuir pelos dias da semana.")

print("Processamento concluído com sucesso!")
