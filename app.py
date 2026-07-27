import base64
from contextlib import contextmanager
import io
import subprocess
import tempfile
from datetime import date
from pathlib import Path
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from docxtpl import DocxTemplate

# --------------------------------------------------------------------------
# Konfigurasi path
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
TEMPLATE_DIR = BASE_DIR / "templates"
DATA_DIR = BASE_DIR / "data"
PEGAWAI_CSV = DATA_DIR / "pegawai.csv"

BULAN_ID = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]


def format_tanggal(d: date) -> str:
    return f"{d.day} {BULAN_ID[d.month]} {d.year}"


JENIS_SURAT = {
    "Surat Pernyataan Tidak Menginap di Hotel": {
        "template": TEMPLATE_DIR / "template_hotel.docx",
        "file_prefix": "Surat_Pernyataan_Tidak_Menginap_Hotel",
        "verb": "tidak menginap di hotel",
        "title_line2": "TIDAK MENGINAP DI HOTEL",
        "paragraf1": (
            "Menerangkan bahwa dalam rangka melaksanakan perjalanan dinas dalam kota di "
            "Kabupaten {kabupaten_kota} untuk melaksanakan tugas kedinasan sesuai surat "
            "tugas nomor {nomor_surat_tugas} tanggal {tanggal_surat_tugas}, saya benar-benar "
            "tidak menginap di hotel pada tanggal {tanggal_perjalanan}."
        ),
    },
    "Surat Pernyataan Tidak Menggunakan Kendaraan Dinas": {
        "template": TEMPLATE_DIR / "template_kendaraan.docx",
        "file_prefix": "Surat_Pernyataan_Tidak_Menggunakan_Kendaraan_Dinas",
        "verb": "tidak menggunakan kendaraan dinas",
        "title_line2": "TIDAK MENGGUNAKAN KENDARAAN DINAS",
        "paragraf1": (
            "Menerangkan bahwa dalam rangka melaksanakan perjalanan dinas dalam kota di "
            "kabupaten/kota {kabupaten_kota} untuk melaksanakan tugas kedinasan sesuai surat "
            "tugas nomor {nomor_surat_tugas} tanggal {tanggal_surat_tugas}, saya benar-benar "
            "tidak menggunakan kendaraan dinas pada tanggal {tanggal_perjalanan}."
        ),
    },
}

PARAGRAF2 = (
    "Demikian pernyataan ini saya buat dengan sebenar-benarnya untuk dipergunakan sebagaimana "
    "mestinya. Apabila terdapat kekeliruan dalam pertanggungjawaban SPD dan mengakibatkan kerugian "
    "negara, saya bersedia dituntut sesuai peraturan yang berlaku dan mengembalikan biaya transport "
    "yang sudah terlanjur saya terima ke kas negara."
)


def render_preview_html(info: dict, ctx: dict) -> str:
    """Render an HTML mock-up of the letter that mirrors the .docx layout, for live preview."""
    paragraf1 = info["paragraf1"].format(**ctx)

    def esc(s):
        return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    fields = [
        ("Nama", ctx["nama"]),
        ("NIP", ctx["nip"]),
        ("Pangkat/Golongan", ctx["pangkat_gol"]),
        ("Jabatan", ctx["jabatan"]),
        ("Unit Kerja", ctx["unit_kerja"]),
    ]
    field_rows = "".join(
        f'<tr><td class="label">{esc(l)}</td><td class="colon">:</td><td>{esc(v) or "&nbsp;"}</td></tr>'
        for l, v in fields
    )

    return f"""
    <div class="page">
      <style>
        .page {{
          font-family: 'Arial', Arial, sans-serif;
          background: white;
          color: #111;
          padding: 40px 50px;
          max-width: 720px;
          margin: 0 auto;
          box-shadow: 0 0 12px rgba(0,0,0,0.15);
          font-size: 12px;
          line-height: 1.5;
        }}
        .page h1, .page h2 {{
          text-align: center;
          margin: 0;
          font-size: 14px;
          font-weight: bold;
        }}
        .page .spacer {{ height: 18px; }}
        .page table.fields {{ margin-left: 40px; margin-top: 10px; }}
        .page table.fields td {{ padding: 1px 0; vertical-align: top; }}
        .page table.fields td.label {{ min-width: 150px; }}
        .page table.fields td.colon {{ padding: 0 8px; }}
        .page p {{ text-align: justify; margin: 16px 0; }}
        .page .signblock {{ text-align: right; margin-top: 24px; }}
        .page .signname {{ font-weight: bold; text-decoration: underline; margin-top: 70px; }}
      </style>
      <h1>SURAT PERNYATAAN</h1>
      <h2>{esc(info['title_line2'])}</h2>
      <div class="spacer"></div>
      <p style="margin:0;">Yang bertanda tangan di bawah ini:</p>
      <table class="fields">{field_rows}</table>
      <p>{esc(paragraf1)}</p>
      <p>{esc(PARAGRAF2)}</p>
      <div class="signblock">
        {esc(ctx['kota_ttd'])}, {esc(ctx['tanggal_surat'])}<br/>
        Pelaksana Perjalanan Dinas
        <div class="signname">{esc(ctx['nama'])}</div>
        NIP. {esc(ctx['nip'])}
      </div>
    </div>
    """

# --------------------------------------------------------------------------
# Data pegawai (CSV)
# --------------------------------------------------------------------------
DATA_DIR.mkdir(exist_ok=True)
if not PEGAWAI_CSV.exists():
    pd.DataFrame(columns=["nama", "nip", "pangkat_gol", "jabatan"]).to_csv(PEGAWAI_CSV, index=False)


def load_pegawai() -> pd.DataFrame:
    df = pd.read_csv(PEGAWAI_CSV, dtype=str).fillna("")
    for col in ["nama", "nip", "pangkat_gol", "jabatan"]:
        if col not in df.columns:
            df[col] = ""
    return df


def save_pegawai(df: pd.DataFrame) -> None:
    df.to_csv(PEGAWAI_CSV, index=False)


# =============================
# STREAMLIT UI
# =============================
st.set_page_config("Super by AK", "📄", layout="wide", initial_sidebar_state="collapsed")

def load_css():
    with open("style.css") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css()

pegawai_df = load_pegawai()
nama_list = pegawai_df["nama"].tolist()

@contextmanager
def card(title, icon=""):
    """A real bordered container (native Streamlit nesting) with a title
    rendered as its first child - no separate empty HTML block above it."""
    with st.container(border=True):
        st.markdown(f'<div class="card-title">{icon} {title}</div>', unsafe_allow_html=True)
        yield

# ---- TOP BAR ----
st.markdown(
    """
    <div class="topbar">
        <div class="brand">
            <div class="brand-name">Su<span class="accent">per</span></div>
            <div class="brand-sub">by <b>AK</b></div>
        </div>
        <div class="topbar-badge">Generator Super</div>
    </div>
    """,
    unsafe_allow_html=True,
)

left, right = st.columns([1, 1], gap="medium")
with left:

    # 1. Jenis surat
    with card("1. Jenis Surat", ""):
        jenis = st.selectbox("Pilih jenis surat pernyataan", list(JENIS_SURAT.keys()), label_visibility="collapsed")
        info = JENIS_SURAT[jenis]

    if not info["template"].exists():
        st.error(f"Template belum ditemukan: `{info['template'].name}` di folder `templates/`.")
        st.stop()

    # 2. Data pegawai
    with card("2. Data Pegawai", ""):
        if not nama_list:
            st.warning("Belum ada data pegawai. Silahkan Hubungi Admin.")
            st.stop()

        nama_pilihan = st.selectbox("Pilih Nama", nama_list)
        row = pegawai_df[pegawai_df["nama"] == nama_pilihan].iloc[0]

        c1, c2 = st.columns(2)
        with c1:
            st.text_input("NIP", value=row["nip"], disabled=True)
            st.text_input("Pangkat/Golongan", value=row["pangkat_gol"], disabled=True)
        with c2:
            st.text_input("Jabatan", value=row["jabatan"], disabled=True)
            unit_kerja = st.text_input("Unit Kerja", value="BPS Kabupaten Donggala")


    # 3. Data surat
    with card("3. Data Surat", ""):
        c3, c4 = st.columns(2)
        with c3:
            nomor_surat_tugas = st.text_input("Nomor Surat Tugas (dasar)", placeholder="B-1461B/72050/SS.100/2026")
        with c4:
            tanggal_surat_tugas = st.date_input("Tanggal Surat Tugas", value=date.today())

        kabupaten_kota = st.text_input("Kabupaten/Kota Pelaksanaan Perjalanan Dinas", value="Donggala")
        tanggal_perjalanan = st.date_input(f"Tanggal {info['verb']}", value=date.today())

        c5, c6 = st.columns(2)
        with c5:
            kota_ttd = st.text_input("Kota Tanda Tangan", value="Donggala")
        with c6:
            tanggal_surat = st.date_input("Tanggal Surat Pernyataan", value=date.today())

# --------------------------------------------------------------------------
# 4. Preview (live) + Generate & unduh
# --------------------------------------------------------------------------
with right:
    with card("4. Preview", ""):
        preview_context = {
            "nama": nama_pilihan,
            "nip": row["nip"],
            "pangkat_gol": row["pangkat_gol"],
            "jabatan": row["jabatan"],
            "unit_kerja": unit_kerja,
            "kabupaten_kota": kabupaten_kota or "…",
            "nomor_surat_tugas": nomor_surat_tugas or "…",
            "tanggal_surat_tugas": format_tanggal(tanggal_surat_tugas),
            "tanggal_perjalanan": format_tanggal(tanggal_perjalanan),
            "kota_ttd": kota_ttd,
            "tanggal_surat": format_tanggal(tanggal_surat),
        }

        components.html(render_preview_html(info, preview_context), height=693, scrolling=False)

        generate = st.button("🔄 Generate Surat", use_container_width=True, type="primary")

        if generate:
            if not nomor_surat_tugas.strip():
                st.error("Nomor Surat Tugas wajib diisi.")
                st.stop()

            context = dict(preview_context)
            context["kabupaten_kota"] = kabupaten_kota
            context["nomor_surat_tugas"] = nomor_surat_tugas

            doc = DocxTemplate(str(info["template"]))
            doc.render(context)

            docx_buffer = io.BytesIO()
            doc.save(docx_buffer)
            docx_buffer.seek(0)

            file_name_base = f"{info['file_prefix']}_{nama_pilihan.replace(' ', '_').replace(',', '')}"

            st.success("Surat berhasil dibuat ✅")

            col_a, col_b = st.columns(2)
            with col_a:
                st.download_button(
                    "⬇️ Download Word (.docx)",
                    data=docx_buffer,
                    file_name=f"{file_name_base}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

            pdf_bytes = None
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    tmp_docx = Path(tmp) / "surat.docx"
                    doc.save(tmp_docx)
                    subprocess.run(
                        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(tmp_docx)],
                        check=True,
                        timeout=60,
                        capture_output=True,
                    )
                    tmp_pdf = Path(tmp) / "surat.pdf"
                    if tmp_pdf.exists():
                        pdf_bytes = tmp_pdf.read_bytes()
            except Exception:
                pdf_bytes = None

            with col_b:
                if pdf_bytes:
                    st.download_button(
                        "🖨️ Download / Print PDF",
                        data=pdf_bytes,
                        file_name=f"{file_name_base}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                else:
                    st.caption("Konversi PDF tidak tersedia — unduh Word lalu print dari Word.")

            if pdf_bytes:
                st.markdown("**Pratinjau PDF:**")
                b64 = base64.b64encode(pdf_bytes).decode()
                st.markdown(
                    f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="600" type="application/pdf"></iframe>',
                    unsafe_allow_html=True,
                )
