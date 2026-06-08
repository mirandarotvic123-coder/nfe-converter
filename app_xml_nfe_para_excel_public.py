# app_xml_nfe_para_excel.py
# Como usar localmente:
#   1) pip install streamlit pandas openpyxl
#   2) streamlit run app_xml_nfe_para_excel.py

import io
import os
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

import pandas as pd
import streamlit as st

COLUNAS = [
    'arquivo_xml', 'tipo_documento', 'chave_acesso', 'modelo', 'serie', 'numero_nf',
    'data_emissao', 'data_saida_entrada', 'tipo_operacao', 'natureza_operacao', 'finalidade',
    'ambiente', 'status_protocolo', 'protocolo_autorizacao', 'emitente_cnpj_cpf',
    'emitente_razao_social', 'emitente_nome_fantasia', 'emitente_ie', 'emitente_crt',
    'origem_uf', 'origem_municipio', 'origem_pais', 'destinatario_cnpj_cpf',
    'destinatario_razao_social', 'destinatario_ie', 'destino_uf', 'destino_municipio',
    'destino_pais', 'n_item', 'codigo_produto', 'descricao_produto', 'ean', 'ncm', 'cest',
    'cfop', 'unidade_comercial', 'quantidade_comercial', 'valor_unitario_comercial',
    'valor_produto', 'unidade_tributavel', 'quantidade_tributavel', 'valor_unitario_tributavel',
    'frete_item', 'seguro_item', 'desconto_item', 'outras_despesas_item',
    'origem_mercadoria_icms', 'icms_cst_csosn', 'bc_icms', 'aliquota_icms', 'valor_icms',
    'bc_icms_st', 'valor_icms_st', 'ipi_cst', 'bc_ipi', 'aliquota_ipi', 'valor_ipi',
    'pis_cst', 'bc_pis', 'aliquota_pis', 'valor_pis', 'cofins_cst', 'bc_cofins',
    'aliquota_cofins', 'valor_cofins', 'valor_total_produtos_nf', 'valor_total_nf',
    'valor_frete_nf', 'valor_seguro_nf', 'valor_desconto_nf', 'valor_outras_despesas_nf',
    'valor_icms_total_nf', 'valor_ipi_total_nf', 'valor_pis_total_nf', 'valor_cofins_total_nf',
    'transportador_cnpj_cpf', 'transportador_nome', 'modalidade_frete', 'placa_veiculo',
    'uf_veiculo', 'peso_bruto', 'peso_liquido', 'informacoes_complementares'
]


def limpar_namespace(root):
    for elem in root.iter():
        if "}" in elem.tag:
            elem.tag = elem.tag.split("}", 1)[1]
    return root


def txt(parent, path, default=""):
    if parent is None:
        return default
    found = parent.find(path)
    if found is None or found.text is None:
        return default
    return found.text.strip()


def attr(parent, name, default=""):
    if parent is None:
        return default
    return parent.attrib.get(name, default)


def num(value):
    if value in (None, ""):
        return None
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return None


def cnpj_ou_cpf(parent):
    return txt(parent, "CNPJ") or txt(parent, "CPF") or txt(parent, "idEstrangeiro")


def primeiro_filho(parent):
    if parent is None:
        return None
    filhos = list(parent)
    return filhos[0] if filhos else None


def detectar_tipo(modelo):
    if modelo == "55":
        return "NF-e"
    if modelo == "65":
        return "NFC-e"
    if modelo in {"57", "67"}:
        return "CT-e/CT-e OS"
    return "XML fiscal"


def parse_xml_nfe(xml_bytes, nome_arquivo="arquivo.xml"):
    try:
        root = ET.fromstring(xml_bytes)
        root = limpar_namespace(root)
    except Exception as e:
        return [], [{"arquivo_xml": nome_arquivo, "erro": f"XML inválido: {e}"}]

    inf_nfe = root.find(".//infNFe")
    if inf_nfe is None:
        return [], [{"arquivo_xml": nome_arquivo, "erro": "Não é NF-e (evento de cancelamento ou outro tipo) — ignorado."}]

    ide = inf_nfe.find("ide")
    emit = inf_nfe.find("emit")
    dest = inf_nfe.find("dest")
    end_emit = emit.find("enderEmit") if emit is not None else None
    end_dest = dest.find("enderDest") if dest is not None else None
    total = inf_nfe.find("total/ICMSTot")
    transp = inf_nfe.find("transp")
    transporta = transp.find("transporta") if transp is not None else None
    veic = transp.find("veicTransp") if transp is not None else None
    vol = transp.find("vol") if transp is not None else None
    inf_prot = root.find(".//protNFe/infProt")

    chave = attr(inf_nfe, "Id").replace("NFe", "") or txt(inf_prot, "chNFe")
    modelo = txt(ide, "mod")
    tp_nf_desc = {"0": "Entrada", "1": "Saída"}.get(txt(ide, "tpNF"), txt(ide, "tpNF"))
    ambiente = {"1": "Produção", "2": "Homologação"}.get(txt(ide, "tpAmb"), txt(ide, "tpAmb"))
    finalidade = {
        "1": "NF-e normal", "2": "NF-e complementar",
        "3": "NF-e de ajuste", "4": "Devolução/retorno",
    }.get(txt(ide, "finNFe"), txt(ide, "finNFe"))
    mod_frete_val = txt(transp, "modFrete") if transp is not None else ""
    mod_frete = {
        "0": "Por conta do emitente",
        "1": "Por conta do destinatário/remetente",
        "2": "Por conta de terceiros",
        "3": "Transporte próprio por conta do remetente",
        "4": "Transporte próprio por conta do destinatário",
        "9": "Sem frete",
    }.get(mod_frete_val, mod_frete_val)

    capa = {
        "arquivo_xml": nome_arquivo,
        "tipo_documento": detectar_tipo(modelo),
        "chave_acesso": chave,
        "modelo": modelo,
        "serie": txt(ide, "serie"),
        "numero_nf": txt(ide, "nNF"),
        "data_emissao": txt(ide, "dhEmi") or txt(ide, "dEmi"),
        "data_saida_entrada": txt(ide, "dhSaiEnt") or txt(ide, "dSaiEnt"),
        "tipo_operacao": tp_nf_desc,
        "natureza_operacao": txt(ide, "natOp"),
        "finalidade": finalidade,
        "ambiente": ambiente,
        "status_protocolo": txt(inf_prot, "cStat"),
        "protocolo_autorizacao": txt(inf_prot, "nProt"),
        "emitente_cnpj_cpf": cnpj_ou_cpf(emit),
        "emitente_razao_social": txt(emit, "xNome"),
        "emitente_nome_fantasia": txt(emit, "xFant"),
        "emitente_ie": txt(emit, "IE"),
        "emitente_crt": txt(emit, "CRT"),
        "origem_uf": txt(end_emit, "UF"),
        "origem_municipio": txt(end_emit, "xMun"),
        "origem_pais": txt(end_emit, "xPais"),
        "destinatario_cnpj_cpf": cnpj_ou_cpf(dest),
        "destinatario_razao_social": txt(dest, "xNome"),
        "destinatario_ie": txt(dest, "IE"),
        "destino_uf": txt(end_dest, "UF"),
        "destino_municipio": txt(end_dest, "xMun"),
        "destino_pais": txt(end_dest, "xPais"),
        "valor_total_produtos_nf": num(txt(total, "vProd")),
        "valor_total_nf": num(txt(total, "vNF")),
        "valor_frete_nf": num(txt(total, "vFrete")),
        "valor_seguro_nf": num(txt(total, "vSeg")),
        "valor_desconto_nf": num(txt(total, "vDesc")),
        "valor_outras_despesas_nf": num(txt(total, "vOutro")),
        "valor_icms_total_nf": num(txt(total, "vICMS")),
        "valor_ipi_total_nf": num(txt(total, "vIPI")),
        "valor_pis_total_nf": num(txt(total, "vPIS")),
        "valor_cofins_total_nf": num(txt(total, "vCOFINS")),
        "transportador_cnpj_cpf": cnpj_ou_cpf(transporta),
        "transportador_nome": txt(transporta, "xNome"),
        "modalidade_frete": mod_frete,
        "placa_veiculo": txt(veic, "placa"),
        "uf_veiculo": txt(veic, "UF"),
        "peso_bruto": num(txt(vol, "pesoB")),
        "peso_liquido": num(txt(vol, "pesoL")),
        "informacoes_complementares": txt(inf_nfe.find("infAdic"), "infCpl"),
    }

    detalhes = inf_nfe.findall("det")
    if not detalhes:
        linha = {col: "" for col in COLUNAS}
        linha.update(capa)
        return [linha], []

    linhas = []
    for det in detalhes:
        prod = det.find("prod")
        imposto = det.find("imposto")
        icms_grupo = primeiro_filho(imposto.find("ICMS") if imposto is not None else None)
        ipi = imposto.find("IPI") if imposto is not None else None
        ipi_grupo = primeiro_filho(ipi) if ipi is not None else None
        pis_grupo = primeiro_filho(imposto.find("PIS") if imposto is not None else None)
        cofins_grupo = primeiro_filho(imposto.find("COFINS") if imposto is not None else None)

        linha = {col: "" for col in COLUNAS}
        linha.update(capa)
        linha.update({
            "n_item": attr(det, "nItem"),
            "codigo_produto": txt(prod, "cProd"),
            "descricao_produto": txt(prod, "xProd"),
            "ean": txt(prod, "cEAN"),
            "ncm": txt(prod, "NCM"),
            "cest": txt(prod, "CEST"),
            "cfop": txt(prod, "CFOP"),
            "unidade_comercial": txt(prod, "uCom"),
            "quantidade_comercial": num(txt(prod, "qCom")),
            "valor_unitario_comercial": num(txt(prod, "vUnCom")),
            "valor_produto": num(txt(prod, "vProd")),
            "unidade_tributavel": txt(prod, "uTrib"),
            "quantidade_tributavel": num(txt(prod, "qTrib")),
            "valor_unitario_tributavel": num(txt(prod, "vUnTrib")),
            "frete_item": num(txt(prod, "vFrete")),
            "seguro_item": num(txt(prod, "vSeg")),
            "desconto_item": num(txt(prod, "vDesc")),
            "outras_despesas_item": num(txt(prod, "vOutro")),
            "origem_mercadoria_icms": txt(icms_grupo, "orig"),
            "icms_cst_csosn": txt(icms_grupo, "CST") or txt(icms_grupo, "CSOSN"),
            "bc_icms": num(txt(icms_grupo, "vBC")),
            "aliquota_icms": num(txt(icms_grupo, "pICMS")),
            "valor_icms": num(txt(icms_grupo, "vICMS")),
            "bc_icms_st": num(txt(icms_grupo, "vBCST")),
            "valor_icms_st": num(txt(icms_grupo, "vICMSST")),
            "ipi_cst": txt(ipi_grupo, "CST"),
            "bc_ipi": num(txt(ipi_grupo, "vBC")),
            "aliquota_ipi": num(txt(ipi_grupo, "pIPI")),
            "valor_ipi": num(txt(ipi_grupo, "vIPI")),
            "pis_cst": txt(pis_grupo, "CST"),
            "bc_pis": num(txt(pis_grupo, "vBC")),
            "aliquota_pis": num(txt(pis_grupo, "pPIS")),
            "valor_pis": num(txt(pis_grupo, "vPIS")),
            "cofins_cst": txt(cofins_grupo, "CST"),
            "bc_cofins": num(txt(cofins_grupo, "vBC")),
            "aliquota_cofins": num(txt(cofins_grupo, "pCOFINS")),
            "valor_cofins": num(txt(cofins_grupo, "vCOFINS")),
        })
        linhas.append(linha)
    return linhas, []


def ler_uploads(uploaded_files):
    arquivos_xml = []
    for up in uploaded_files:
        data = up.read()
        nome = up.name
        if nome.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for item in zf.namelist():
                    if item.lower().endswith(".xml"):
                        nome_xml = os.path.basename(item)
                        arquivos_xml.append((nome_xml, zf.read(item)))
        elif nome.lower().endswith(".xml"):
            arquivos_xml.append((nome, data))
    return arquivos_xml


def gerar_excel(df, df_erros):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Relatorio_XML")
        if not df_erros.empty:
            df_erros.to_excel(writer, index=False, sheet_name="Erros_XML")
        workbook = writer.book
        ws = workbook["Relatorio_XML"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col_cells in ws.columns:
            header = str(col_cells[0].value or "")
            max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col_cells[:200])
            width = min(max(max_len + 2, 12), 42)
            if any(x in header for x in ["descricao", "razao", "informacoes", "natureza"]):
                width = 38
            ws.column_dimensions[col_cells[0].column_letter].width = width
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True, color="FFFFFF")
            cell.fill = cell.fill.copy(fill_type="solid", fgColor="1F4E78")
        if not df_erros.empty:
            we = workbook["Erros_XML"]
            we.freeze_panes = "A2"
            we.auto_filter.ref = we.dimensions
    output.seek(0)
    return output


# ── Interface ──────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Conversor XML NF-e para Excel - Countout Co.", layout="wide")
st.title("📄 Conversor de XML de notas fiscais para Excel - Countout Co.")
st.caption("Envie aqui os arquivos. O relatório sai com uma linha por item da nota.")

if "upload_key" not in st.session_state:
    st.session_state.upload_key = 0

uploaded = st.file_uploader(
    "Selecione os arquivos",
    type=["zip", "xml"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.upload_key}",
)

if uploaded:
    if st.button("🔄 Limpar e recomeçar"):
        st.session_state.upload_key += 1
        st.rerun()

    with st.spinner("Lendo arquivos..."):
        arquivos = ler_uploads(uploaded)

    if not arquivos:
        st.error("Nenhum arquivo .xml encontrado. Verifique o conteúdo e tente novamente.")
    else:
        st.info(f"📦 {len(arquivos)} XML(s) encontrados.")
        progresso = st.progress(0, text="Iniciando...")
        linhas, erros = [], []
        total = len(arquivos)

        for i, (nome, data) in enumerate(arquivos):
            l, e = parse_xml_nfe(data, nome)
            linhas.extend(l)
            erros.extend(e)
            progresso.progress((i + 1) / total, text=f"Processando {i+1}/{total}: {nome[:60]}")

        progresso.empty()

        df = pd.DataFrame(linhas, columns=COLUNAS)
        df_erros = pd.DataFrame(erros)

        st.success(
            f"✅ Processados: {total} XML(s) | "
            f"Linhas de itens: {len(df)} | "
            f"Ignorados (cancelamentos/eventos): {len(df_erros)}"
        )

        st.dataframe(df, use_container_width=True, height=420)

        excel = gerar_excel(df, df_erros)
        st.download_button(
            "⬇️ Baixar relatório Excel",
            data=excel,
            file_name=f"relatorio_xml_nfe_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if not df_erros.empty:
            with st.expander(f"Ver {len(df_erros)} arquivo(s) ignorado(s)"):
                st.dataframe(df_erros, use_container_width=True)
