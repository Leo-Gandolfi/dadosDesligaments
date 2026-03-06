import streamlit as st
import fitz
import pandas as pd
import re
import io
import os

# ─────────────────────────────────────────────
# Extraction logic (same as extract.py, adapted
# to work with bytes instead of a file path)
# ─────────────────────────────────────────────

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "EXEMPLO DO RELATÓRIO 2.xlsx")

def extract_data_from_bytes(pdf_bytes: bytes) -> pd.DataFrame:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = "".join(page.get_text() + "\n" for page in doc)

    records_raw = re.split(r"Solicita..o de Desligamento - (\d+)", full_text)
    data = []

    for i in range(1, len(records_raw), 2):
        sol_id = records_raw[i]
        rec_text = records_raw[i + 1]

        lines = [ln.strip() for ln in rec_text.split("\n") if ln.strip()]

        # ── Header fields ────────────────────────────────
        idx_havera = -1
        for j, line in enumerate(lines):
            if re.search(r"Haver.\s+Reposi..o:", line):
                idx_havera = j
                break

        if idx_havera == -1:
            continue

        try:
            val_colaborador = lines[idx_havera + 2]
            val_tipo_deslig  = lines[idx_havera + 6]
            val_motivo       = lines[idx_havera + 7]
            val_havera_rep   = lines[idx_havera + 9]
        except IndexError:
            continue

        np_match = re.match(r"(\d+)\s*[-–]\s*(.*)", val_colaborador)
        np_val    = np_match.group(1).lstrip("0") if np_match else ""
        colab_val = np_match.group(2)              if np_match else val_colaborador

        # ── Outras Informações ───────────────────────────
        val_outras_info = ""
        try:
            idx_admin = next(
                k for k, ln in enumerate(lines)
                if re.search(r"Administra..o de Pessoal", ln)
            )
            parts = []
            for k in range(idx_admin - 1, -1, -1):
                if lines[k] in ["Sim", "No", "Não"] or re.search(r"Recomendaria-o", lines[k]):
                    break
                parts.insert(0, lines[k])
            val_outras_info = " ".join(parts).strip()
            if val_outras_info in ["Outras Informaes", "Outras Informações"]:
                val_outras_info = ""
        except StopIteration:
            pass

        # ── Recrutamento e Seleção ───────────────────────
        val_recrut_readm = ""
        for j, ln in enumerate(lines):
            if re.search(r"Recrutamento e Sele..o", ln):
                if j + 2 < len(lines) and "Diretoria" not in lines[j + 2]:
                    val_recrut_readm = lines[j + 2]
                break

        # ── Diretoria / Considerações ────────────────────
        val_dir_readm = ""
        val_consid    = ""

        idx_dir = -1
        for j, ln in enumerate(lines):
            if re.search(r"Diretoria\s*/Ger.ncia de Unidade de Neg.cio", ln):
                idx_dir = j
                break

        if idx_dir != -1:
            # Find "Considerações" after the Diretoria header
            idx_consid = -1
            for j in range(idx_dir + 1, len(lines)):
                if re.search(r"Considera..es", lines[j]):
                    idx_consid = j
                    break

            # "Condições de Readmissão" is the label between idx_dir and idx_consid.
            # Its value is the line(s) between that label and "Considerações".
            if idx_consid != -1:
                for j in range(idx_dir + 1, idx_consid):
                    if re.search(r"Condi..es de Readmiss.o", lines[j]):
                        # value is everything between this label and idx_consid
                        readm_parts = []
                        for k in range(j + 1, idx_consid):
                            readm_parts.append(lines[k])
                        val_dir_readm = " ".join(readm_parts).strip()
                        break

            # Collect ALL lines of Considerações until "Fluxo de Aprova" or end
            if idx_consid != -1:
                consid_parts = []
                for k in range(idx_consid + 1, len(lines)):
                    if re.search(r"Fluxo de Aprova", lines[k]):
                        break
                    consid_parts.append(lines[k])
                val_consid = " ".join(consid_parts).strip()

        # ── Aprovador (FIRST row in Fluxo) ───────────────
        aprovador_id   = ""
        aprovador_nome = ""

        idx_fluxo = -1
        for j, ln in enumerate(lines):
            if re.search(r"Fluxo de Aprova", ln):
                idx_fluxo = j
                break

        if idx_fluxo != -1:
            for k in range(idx_fluxo + 1, len(lines)):
                if re.match(r"^\d{2}\.\d{2}\.\d{4}$", lines[k]):
                    aprovador_nome = lines[k - 1]
                    if k >= 2 and re.match(r"^\d{6,}$", lines[k - 2]):
                        aprovador_id = lines[k - 2].lstrip("0")
                    break

        data.append({
            "sol_id":          sol_id,
            "NP":              np_val,
            "Colaborador":     colab_val,
            "havera_rep":      val_havera_rep,
            "tipo_deslig":     val_tipo_deslig,
            "motivo":          val_motivo,
            "outras_info":     val_outras_info,
            "recrut_readm":    val_recrut_readm,
            "dir_readm":       val_dir_readm,
            "consid":          val_consid,
            "aprovador_id":    aprovador_id,
            "aprovador_nome":  aprovador_nome,
        })

    # ── Build DataFrame aligned with template columns ──
    template_df = pd.read_excel(TEMPLATE_PATH)
    cols = template_df.columns

    mapped = [
        {
            cols[0]:  r["sol_id"],
            cols[1]:  r["NP"],
            cols[2]:  r["Colaborador"],
            cols[3]:  r["havera_rep"],
            cols[4]:  r["tipo_deslig"],
            cols[5]:  r["motivo"],
            cols[6]:  r["outras_info"],
            cols[7]:  r["recrut_readm"],
            cols[8]:  r["dir_readm"],
            cols[9]:  r["consid"],
            cols[10]: r["aprovador_id"],
            cols[11]: r["aprovador_nome"],
        }
        for r in data
    ]

    return template_df, pd.DataFrame(mapped)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Relatório")
    return buf.getvalue()


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Extrator de Desligamentos",
    page_icon="",
    layout="centered",
)

# ── Password gate ─────────────────────────────
def check_password():
    """Returns True if the user has entered the correct password."""
    def submit():
        entered = st.session_state.get("pwd_input", "")
        correct  = st.secrets.get("passwords", {}).get("app_password", "")
        if entered == correct:
            st.session_state["authenticated"] = True
        else:
            st.session_state["auth_error"] = True

    if st.session_state.get("authenticated"):
        return True

    st.markdown("""
    <div style='
        max-width:380px; margin:80px auto; padding:2.4rem;
        background:#f7faf8; border:1px solid #d4edda;
        border-radius:16px; box-shadow:0 8px 32px rgba(0,0,0,.08);
        text-align:center;'>
        <div style='font-size:2.4rem;margin-bottom:.6rem;'>🔒</div>
        <h2 style='margin:0 0 .3rem; color:#1a6b3c;'>Acesso Restrito</h2>
        <p style='color:#666; font-size:.9rem; margin-bottom:1.4rem;'>
            Insira a senha para continuar
        </p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.text_input(
            "Senha", type="password",
            key="pwd_input",
            label_visibility="collapsed",
            placeholder="Digite a senha...",
            on_change=submit,
        )
        st.button("Entrar", on_click=submit, use_container_width=True)
        if st.session_state.get("auth_error"):
            st.error("Senha incorreta. Tente novamente.")
            st.session_state["auth_error"] = False

    return False

if not check_password():
    st.stop()


# ── Custom styles ────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .hero {
        background: linear-gradient(135deg, #1a6b3c 0%, #2d9e5f 100%);
        border-radius: 16px;
        padding: 2.4rem 2rem;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 8px 32px rgba(26,107,60,.25);
    }
    .hero h1 { color: #fff; font-size: 2rem; font-weight: 700; margin: 0 0 .4rem; }
    .hero p  { color: rgba(255,255,255,.85); font-size: 1rem; margin: 0; }

    .card {
        background: #f7faf8;
        border: 1px solid #d4edda;
        border-radius: 12px;
        padding: 1.6rem;
        margin-bottom: 1.2rem;
    }

    .stat-row {
        display: flex;
        gap: 1rem;
        margin-top: 1rem;
    }
    .stat {
        flex: 1;
        background: #fff;
        border: 1px solid #c8e6c9;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
    }
    .stat .num { font-size: 2rem; font-weight: 700; color: #1a6b3c; }
    .stat .lbl { font-size: .8rem; color: #555; margin-top: .2rem; }

    div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #1a6b3c, #2d9e5f) !important;
        color: #fff !important;
        border: none !important;
        border-radius: 8px !important;
        padding: .65rem 2rem !important;
        font-weight: 600 !important;
        letter-spacing: .02em !important;
        font-size: 1rem !important;
        width: 100% !important;
        transition: opacity .2s;
    }
    div[data-testid="stDownloadButton"] > button:hover { opacity: .88 !important; }

    div[data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #2d9e5f !important;
        border-radius: 12px !important;
        background: #f0f9f3 !important;
    }
    
    .stAlert { border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

# ── Hero ─────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>📋 Extrator de Desligamentos</h1>
    <p>Faça upload do PDF das Solicitações de Desligamento e baixe o relatório em Excel</p>
</div>
""", unsafe_allow_html=True)

# ── Upload ────────────────────────────────────
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("1. Selecione o PDF")
    uploaded = st.file_uploader(
        label="Arraste o PDF aqui ou clique para selecionar",
        type=["pdf"],
        help="Arquivo PDF das Solicitações de Desligamento gerado pelo sistema",
    )
    st.markdown("</div>", unsafe_allow_html=True)

# ── Process ───────────────────────────────────
if uploaded:
    with st.spinner("Processando o PDF, aguarde..."):
        try:
            pdf_bytes = uploaded.read()
            template_df, extracted_df = extract_data_from_bytes(pdf_bytes)

            n_records = len(extracted_df)
            n_empty_outras = extracted_df[template_df.columns[6]].isna().sum() + (extracted_df[template_df.columns[6]] == "").sum()
            n_with_aprovador = extracted_df[template_df.columns[11]].notna().sum()

            # Keep header row (example) + new data
            final_df = pd.concat([template_df, extracted_df], ignore_index=True)
            excel_bytes = to_excel_bytes(final_df)

            st.success(f"✅ PDF processado com sucesso! **{n_records}** cadastros extraídos.")

            # Stats
            st.markdown(f"""
            <div class="stat-row">
                <div class="stat"><div class="num">{n_records}</div><div class="lbl">Cadastros extraídos</div></div>
                <div class="stat"><div class="num">{n_with_aprovador}</div><div class="lbl">Com Nome do Aprovador</div></div>
                <div class="stat"><div class="num">{n_records - n_empty_outras}</div><div class="lbl">Com Outras Informações</div></div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Preview
            with st.expander("👁️ Pré-visualizar primeiras 10 linhas"):
                st.dataframe(extracted_df.head(10), use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Download
            st.subheader("2. Baixe o Relatório")
            output_name = uploaded.name.replace(".pdf", "_relatorio.xlsx")
            st.download_button(
                label="⬇️ Baixar Excel",
                data=excel_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Erro ao processar o PDF: {e}")
            st.exception(e)

else:
    st.info("⬆️ Faça upload de um PDF para começar.")

# ── Footer ────────────────────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center; color:#999; font-size:.8rem;'>Extrator de Solicitações de Desligamento</p>",
    unsafe_allow_html=True,
)
