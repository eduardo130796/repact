import streamlit as st
from openpyxl import load_workbook

def ler_linhas_a_f(arquivo, aba):
    wb = load_workbook(filename=arquivo, data_only=True)
    ws = wb[aba]
    linhas = []
    for row in ws.iter_rows(min_row=1, max_col=6):
        campo_cell = row[0]
        if campo_cell.value is None or str(campo_cell.value).strip() == "":
            continue

        linha = [str(campo_cell.value).strip()]
        for cell in row[1:]:
            valor = cell.value
            formato = cell.number_format

            if formato and "%" in formato:
                tipo = "percentual"
            elif formato and ("R$" in formato or "[$" in formato or "¤" in formato or "0.00" in formato):
                tipo = "reais"
            else:
                tipo = "numero"

            linha.append((valor, tipo))
        linhas.append(linha)
    return linhas

def comparar_linhas(linhas_antigo, linhas_novo):
    max_len = max(len(linhas_antigo), len(linhas_novo))
    comparacao = []
    for i in range(max_len):
        linha_a = linhas_antigo[i] if i < len(linhas_antigo) else [None] + [(None, "numero")] * 5
        linha_b = linhas_novo[i] if i < len(linhas_novo) else [None] + [(None, "numero")] * 5

        campo_a = linha_a[0]
        campo_b = linha_b[0]
        campo = campo_a if campo_a == campo_b else f"{campo_a or ''} / {campo_b or ''}"

        comparacao.append({
            'campo': campo,
            'antigo_B': linha_a[1][0],
            'novo_B': linha_b[1][0],
            'tipo_B': linha_a[1][1],
            'antigo_C': linha_a[2][0],
            'novo_C': linha_b[2][0],
            'tipo_C': linha_a[2][1],
            'antigo_D': linha_a[3][0],
            'novo_D': linha_b[3][0],
            'tipo_D': linha_a[3][1],
            'antigo_E': linha_a[4][0],
            'novo_E': linha_b[4][0],
            'tipo_E': linha_a[4][1],
            'antigo_F': linha_a[5][0],
            'novo_F': linha_b[5][0],
            'tipo_F': linha_a[5][1],
        })
    return comparacao

def formatar_tipo(val, tipo):
    if val in [None, ""]:
        return "-"
    try:
        val_f = float(val)
        if tipo == "percentual":
            return f"{val_f:.2%}"
        elif tipo == "reais":
            return f"R$ {val_f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            return f"{val_f:.2f}".replace(".", ",")
    except:
        return str(val)

st.title("Comparação Visual - PCFP")
col1, col2 = st.columns(2)
with col1:
    arquivo_antigo = st.file_uploader("Planilha Antiga", type="xlsx")
with col2:    
    arquivo_novo = st.file_uploader("Planilha Nova", type="xlsx")

if arquivo_antigo and arquivo_novo:
    wb_antigo = load_workbook(filename=arquivo_antigo, data_only=True)
    wb_novo = load_workbook(filename=arquivo_novo, data_only=True)
    with col1:
        aba_antiga = st.selectbox("Selecione a aba da planilha ANTIGA", wb_antigo.sheetnames)
    with col2:    
        aba_nova = st.selectbox("Selecione a aba da planilha NOVA", wb_novo.sheetnames)

    if aba_antiga and aba_nova:
        linhas_antigo = ler_linhas_a_f(arquivo_antigo, aba_antiga)
        linhas_novo = ler_linhas_a_f(arquivo_novo, aba_nova)

        comparacao = comparar_linhas(linhas_antigo, linhas_novo)


        for item in comparacao:
            mostrar = False
            for col in ['C', 'D', 'E', 'F']:
                val_antigo = item.get(f'antigo_{col}')
                val_novo = item.get(f'novo_{col}')
                if (val_antigo or val_novo) and val_antigo != val_novo:
                    mostrar = True
                    break

            if not mostrar:
                continue

           # Verifica se algum valor antes era zero/vazio e agora tem valor
            # Verifica se algum valor antes era zero/vazio e agora tem valor
            era_zerado_e_virou_valor = any(
                (item.get(f'antigo_{col}') in [None, "", 0, "0", "0.0"]) and (item.get(f'novo_{col}') not in [None, "", 0, "0", "0.0"])
                for col in ['C', 'D', 'E', 'F']
            )

            # Mostra título padrão ou textos lado a lado conforme o caso
            if era_zerado_e_virou_valor:
                st.markdown(f"""
                    <div style='background-color:#fff4e5; padding: 6px 5px; border-left: 4px solid #fbbc05; margin: 5px 0 2px 0; border-radius: 4px;'>
                        <div style="display: grid; grid-template-columns: 29% 48%; gap: 4%; font-size: 12px; color: #333; font-weight: 500;">
                            <div><span style='color:#1a73e8;'>Texto antigo:</span><br>{item['antigo_B'] or '-'}</div>
                            <div><span style='color:#ea4335;'>Texto novo:</span><br>{item['novo_B'] or '-'}</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                    <div style='background-color:#f5f5f5; padding: 4px 8px; border-left: 4px solid #1a73e8; margin: 8px 0 4px 0;'>
                        <div style='font-size: 13px; font-weight: 500; color: #333;'>
                            <strong>{item['campo']} - {item['antigo_B'] or "-"}</strong>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

            st.markdown("""
                <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; font-weight: 500; font-size: 12px; color: #555; border-bottom: 1px solid #ddd; padding-bottom: 2px; margin-bottom: 4px;">
                    <div style='color:#1a73e8;'>Antigo</div>
                    <div style='color:#f9ab00;'>Novo</div>
                    <div style='color:#d93025;'>Diferença</div>
                </div>
            """, unsafe_allow_html=True)
    
            for col in ['C', 'D', 'E', 'F']:
                val_antigo = item.get(f'antigo_{col}')
                val_novo = item.get(f'novo_{col}')
                tipo = item.get(f'tipo_{col}')
                

                if (val_antigo in [None, ""]) and (val_novo in [None, ""]):
                    continue

                try:
                    val1 = float(val_antigo)
                    val2 = float(val_novo)
                    dif = val2 - val1
                    pct = (dif / val1) * 100 if val1 != 0 else None
                    cor = "#34a853" if dif > 0 else "#ea4335" if dif < 0 else "#333"
                    dif_formatado = f"<span style='color:{cor}; font-weight:600;'>R$ {dif:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + "</span>"
                    if pct is not None:
                        dif_formatado += f" <span style='color:{cor}; font-weight:400;'>({pct:+.1f}%)</span>"
                except:
                    dif_formatado = "<span style='color:#888;'>–</span>"

                if val_antigo == val_novo:
                    dif_formatado = "<span style='color:#aaa;'>–</span>"

                destaque_estilo = ""
                if (val_antigo in [None, "", 0, "0", "0.0"]) and (val_novo not in [None, "", 0, "0", "0.0"]):
                    destaque_estilo = "background-color: #e6f4ea; font-weight: 600; border-radius: 4px; padding: 2px 4px;"

                st.markdown(f"""
                    <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; margin-bottom: 6px;">
                        <div>{formatar_tipo(val_antigo, tipo)}</div>
                        <div style="{destaque_estilo}">{formatar_tipo(val_novo, tipo)}{' 🆕' if destaque_estilo else ''}</div>
                        <div>{dif_formatado}</div>
                    </div>
                """, unsafe_allow_html=True)
