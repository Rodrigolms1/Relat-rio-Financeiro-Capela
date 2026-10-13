import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import re
import tkinter as tk
from tkinter import filedialog

output_filename = 'Controle_Financeiro.xlsx'

def selecionar_pastas_ou_arquivos():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    print("Aguardando seleção da pasta...")
    caminho_diretorio = filedialog.askdirectory(
        title="Selecione a Pasta Anual (ex: Pasta de 2026)"
    )
    
    arquivos_encontrados = []
    if caminho_diretorio:
        for raiz, _, arquivos in os.walk(caminho_diretorio):
            for arquivo in arquivos:
                if (arquivo.endswith('.xlsx') or arquivo.endswith('.xls')) and not arquivo.startswith('~$'):
                    arquivos_encontrados.append(os.path.join(raiz, arquivo))
    return arquivos_encontrados

all_records = []

def converter_mes_extenso_para_numero(nome_aba):
    meses_map = {
        'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'MARCO': '03',
        'ABRIL': '04', 'MAIO': '05', 'JUNHO': '06', 'JULHO': '07',
        'AGOSTO': '08', 'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11', 'DEZEMBRO': '12'
    }
    nome_upper = nome_aba.upper()
    for mes_extenso, num_mes in meses_map.items():
        if mes_extenso in nome_upper:
            anos = re.findall(r'20\d{2}', nome_upper)
            ano = anos[0] if anos else '2026'
            return f"{num_mes}/{ano}"
    return None

def eh_aba_de_periodo(nome_aba):
    nome_limpo = nome_aba.strip().upper()
    termos_ignorar = ['EXT', 'EXTRATO', 'CAPELA', 'RESUMO', 'CAPA', 'INICIO', 'GERAL', 'SALDO', 'PLANILHA']
    if any(termo in nome_limpo for termo in termos_ignorar):
        return False
    if converter_mes_extenso_para_numero(nome_limpo):
        return True
    digitos = re.sub(r'\D', '', nome_limpo)
    if len(digitos) in [4, 5, 6]:
        return True
    return False

def obter_mes_ano_da_aba(sheet):
    mes_extenso_conv = converter_mes_extenso_para_numero(sheet)
    if mes_extenso_conv:
        return mes_extenso_conv
    digits = re.sub(r'\D', '', sheet)
    if len(digits) == 6:
        return f"{digits[:2]}/{digits[2:]}"
    elif len(digits) == 5:
        return f"0{digits[:1]}/{digits[1:]}"
    return sheet.strip()

def processar_planilha(caminho_arquivo):
    if not os.path.exists(caminho_arquivo):
        return
        
    nome_arquivo = os.path.basename(caminho_arquivo).upper()
    if 'SICREDI' in nome_arquivo:
        origem_nome = 'Banco Sicredi'
    elif 'INTERNO' in nome_arquivo:
        origem_nome = 'Caixa Interno'
    elif 'CAIXA' in nome_arquivo:
        origem_nome = 'Caixa'
    else:
        origem_nome = 'Outros'

    try:
        xl = pd.ExcelFile(caminho_arquivo)
    except Exception as e:
        print(f"Erro ao abrir o arquivo {os.path.basename(caminho_arquivo)}: {e}")
        return
    
    abas_validas = [s for s in xl.sheet_names if eh_aba_de_periodo(s)]
    
    for idx_aba, sheet in enumerate(abas_validas):
        try:
            df_raw = pd.read_excel(caminho_arquivo, sheet_name=sheet, header=None)
        except Exception:
            continue
            
        if df_raw.empty:
            continue
            
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = [str(cell).upper().strip() for cell in row.values]
            if 'DATA' in row_str:
                header_idx = idx
                break
                
        if header_idx is None:
            continue
            
        df = pd.read_excel(caminho_arquivo, sheet_name=sheet, skiprows=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        date_col = next((c for c in df.columns if 'DATA' in c), None)
        hist_col = next((c for c in df.columns if 'HIST' in c or 'DESC' in c), None)
        entr_col = next((c for c in df.columns if 'ENTRA' in c), None)
        said_col = next((c for c in df.columns if 'SAÍD' in c or 'SAID' in c), None)
        val_col = next((c for c in df.columns if 'VALOR' in c), None)
        
        mes_ano_aba = obter_mes_ano_da_aba(sheet)
        
        for _, row in df.iterrows():
            dt = row[date_col] if date_col else None
            hist = row[hist_col] if hist_col else None
            hist_str = str(hist).strip() if pd.notna(hist) else ''
            
            if hist_str == '' or hist_str.lower() == 'nan':
                continue
                
            hist_upper = hist_str.upper()
            
            is_saldo_inicial = any(w in hist_upper for w in ['SALDO ANTERIOR', 'RESERVA MÊS ANTERIOR', 'RESERVA MES ANTERIOR'])
            
            if is_saldo_inicial:
                if idx_aba > 0:
                    continue
            else:
                palavras_excluir = ['BALANCETE', 'TOTAL', 'VALOR LÍQUIDO', 'VALOR LIQUIDO']
                if any(w in hist_upper for w in palavras_excluir):
                    continue
                
            val_e, val_s = 0.0, 0.0
            
            if is_saldo_inicial:
                for val_cell in row.values:
                    try:
                        v = float(val_cell)
                        if v > 0:
                            val_e = v
                            break
                    except:
                        pass
            else:
                if val_col and pd.notna(row[val_col]):
                    try:
                        v = float(row[val_col])
                        if v >= 0: val_e = v
                        else: val_s = abs(v)
                    except: pass
                else:
                    if entr_col and pd.notna(row[entr_col]):
                        try: val_e = float(row[entr_col])
                        except: pass
                    if said_col and pd.notna(row[said_col]):
                        try: val_s = abs(float(row[said_col]))
                        except: pass
                        
            if val_e == 0.0 and val_s == 0.0 and pd.isna(dt):
                continue
                
            formatted_dt = pd.to_datetime(dt, errors='coerce') if pd.notna(dt) else None
            mes_ano_str = mes_ano_aba
            
            all_records.append({
                'Origem': origem_nome,
                'Mês/Ano': mes_ano_str,
                'Data': formatted_dt,
                'Histórico / Descrição': hist_str,
                'Entrada (R$)': val_e,
                'Saída (R$)': val_s
            })

arquivos_selecionados = selecionar_pastas_ou_arquivos()

if not arquivos_selecionados:
    print("Nenhum arquivo encontrado ou pasta não selecionada. Operação cancelada.")
else:
    print(f"Foram encontrados {len(arquivos_selecionados)} arquivos para processamento.")
    for arquivo in arquivos_selecionados:
        processar_planilha(arquivo)

    all_records = sorted(all_records, key=lambda x: (x['Data'] is None, x['Data']))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado Geral"
    ws.views.sheetView[0].showGridLines = False

    headers = ["Origem", "Mês/Ano", "Data", "Histórico / Descrição", "Entrada (R$)", "Saída (R$)"]
    ws.append(headers)  

    navy_header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    white_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F2F7FA", end_color="F2F7FA", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(top=Side(border_style="thin", color="000000"), bottom=Side(border_style="double", color="000000"))

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = navy_header_fill
        cell.font = white_header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 25

    row_idx = 2
    for rec in all_records:
        dt_val = rec['Data'].strftime('%d/%m/%Y') if pd.notna(rec['Data']) else ''
        
        ws.cell(row=row_idx, column=1, value=rec['Origem'])
        ws.cell(row=row_idx, column=2, value=rec['Mês/Ano'])
        ws.cell(row=row_idx, column=3, value=dt_val)    
        ws.cell(row=row_idx, column=4, value=rec['Histórico / Descrição'])
        ws.cell(row=row_idx, column=5, value=rec['Entrada (R$)'])
        ws.cell(row=row_idx, column=6, value=rec['Saída (R$)'])
        
        fill = zebra_fill if row_idx % 2 == 0 else white_fill
        for c in range(1, 7):
            cell = ws.cell(row=row_idx, column=c)
            cell.fill = fill
            cell.border = thin_border
            cell.font = Font(name="Calibri", size=11)
            
            if c in [1, 2, 3, 4]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c in [5, 6]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00;"-"'
                
        row_idx += 1

    # Linha de Totais com SUBTOTAL
    ws.cell(row=row_idx, column=4, value="TOTAL GERAL").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=row_idx, column=4).alignment = Alignment(horizontal="right", vertical="center")

    ws.cell(row=row_idx, column=5, value=f"=SUBTOTAL(9, E2:E{row_idx-1})")
    ws.cell(row=row_idx, column=6, value=f"=SUBTOTAL(9, F2:F{row_idx-1})")

    for c in range(1, 7):
        cell = ws.cell(row=row_idx, column=c)
        cell.border = total_border
        cell.font = Font(name="Calibri", size=11, bold=True)
        if c in [5, 6]:
            cell.number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00;"-"'

    # Ajuste automático da largura de todas as colunas para evitar texto cortado ou ###
    for col in ws.columns:
        max_len = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        for cell in col:
            # Ignora a fórmula de total longo na contagem para não alargar demais
            val_str = str(cell.value or '')
            if val_str.startswith('='):
                val_str = "R$ 999.999,00" # Tamanho padrão seguro para valores monetários
            if len(val_str) > max_len:
                max_len = len(val_str)
        # Define a largura com uma margem de folga proporcional
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Garante tamanho especial folgado para a coluna de Histórico/Descrição
    ws.column_dimensions['D'].width = 45

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    wb.save(output_filename)
    print(f"Processamento concluído! Arquivo '{output_filename}' gerado com sucesso.")