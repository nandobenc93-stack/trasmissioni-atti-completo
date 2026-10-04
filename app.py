import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from google.genai import types
import datetime
import io
import json
import os
import traceback
from pathlib import Path

st.set_page_config(page_title="Automazione Notifiche e Atti", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
MODELLI_DIR = BASE_DIR / "modelli_riferimento"
MODELLI_DIR.mkdir(parents=True, exist_ok=True)
RUBRICA_FILE = BASE_DIR / "enti_rubrica.json"

# Inizializzazione session state
if "dati_elaborati" not in st.session_state:
    st.session_state["dati_elaborati"] = None
if "file_salvato_bytes" not in st.session_state:
    st.session_state["file_salvato_bytes"] = None
if "file_salvato_nome" not in st.session_state:
    st.session_state["file_salvato_nome"] = ""
if "file_salvato_mime" not in st.session_state:
    st.session_state["file_salvato_mime"] = ""

def get_api_key():
    if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"].strip():
        return st.secrets["GEMINI_API_KEY"].strip()
    if "gemini_key_manuale" in st.session_state and st.session_state["gemini_key_manuale"].strip():
        return st.session_state["gemini_key_manuale"].strip()
    return ""

chiave_attuale = get_api_key()

with st.sidebar:
    st.title("Impostazioni")
    st.subheader("Chiave Gemini API")
    if chiave_attuale:
        st.success("Chiave API collegata")
    else:
        st.error("Nessuna chiave API trovata")

    nuova_chiave = st.text_input(
        "Inserisci/Sostituisci Chiave:", 
        type="password", 
        key="gemini_key_manuale"
    )
    if nuova_chiave:
        chiave_attuale = nuova_chiave.strip()

    st.divider()
    st.subheader("Modelli Word")
    nuovo_modello = st.file_uploader("Carica modello (.docx)", type=["docx"], key="side_modello_uploader")
    if nuovo_modello:
        with open(MODELLI_DIR / nuovo_modello.name, "wb") as f:
            f.write(nuovo_modello.getbuffer())
        st.success(f"Modello {nuovo_modello.name} pronto!")
        st.rerun()

def carica_rubrica():
    if RUBRICA_FILE.exists():
        try:
            with open(RUBRICA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def aggiorna_rubrica(nome_ente, blocco_indirizzo):
    if not nome_ente or not blocco_indirizzo:
        return
    rubrica = carica_rubrica()
    rubrica[nome_ente.strip()] = blocco_indirizzo.strip()
    try:
        with open(RUBRICA_FILE, "w", encoding="utf-8") as f:
            json.dump(rubrica, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

st.title("Generatore Lettere di Trasmissione")

oggi = datetime.date.today()
data_odierna_str = oggi.strftime("%d/%m/%Y")
anno_corrente = oggi.year

# Selezione modello Word
modelli_docx = [f for f in os.listdir(MODELLI_DIR) if f.endswith(".docx") and not f.startswith("~$")]
opzioni_modello = ["-- Modello Istituzionale Standard (Generato al volo) --"] + modelli_docx
modello_scelto = st.selectbox("Formato / Modello Word di base da utilizzare:", opzioni_modello)

col_prot, col_data = st.columns(2)
with col_prot:
    protocollo_input = st.text_input("Protocollo Uscita (in alto a sinistra)", value=f"N. /{anno_corrente}")
with col_data:
    st.info(f"Data documento (in alto a destra): {data_odierna_str}")

rubrica_enti = carica_rubrica()
opzioni_enti = ["-- Rileva automaticamente dall'atto caricato --"] + list(rubrica_enti.keys())
ente_selezionato = st.selectbox("Destinatario da Rubrica:", opzioni_enti)

st.subheader("1. Atto/Richiesta di Notifica ricevuta (Opzionale)")
st.caption("Puoi caricare un PDF, una foto, oppure non caricare nulla e inserire i dettagli solo a mano.")

c_up1, c_up2 = st.columns(2)
with c_up1:
    carica_doc = st.file_uploader("Carica File (PDF, Foto, Word):", key="file_up_atto")
    if carica_doc is not None:
        st.session_state["file_salvato_bytes"] = carica_doc.getvalue()
        st.session_state["file_salvato_nome"] = carica_doc.name
        st.session_state["file_salvato_mime"] = carica_doc.type or ""
with c_up2:
    scatta_doc = st.camera_input("Oppure scatta foto da fotocamera:")
    if scatta_doc is not None:
        st.session_state["file_salvato_bytes"] = scatta_doc.getvalue()
        st.session_state["file_salvato_nome"] = "foto_scattata.jpg"
        st.session_state["file_salvato_mime"] = "image/jpeg"

# Mostra lo stato di aggancio
if st.session_state["file_salvato_bytes"]:
    col_info, col_del = st.columns([4, 1])
    with col_info:
        st.success(f"Documento agganciato per l'analisi: **{st.session_state['file_salvato_nome']}** ({round(len(st.session_state['file_salvato_bytes'])/1024, 1)} KB)")
    with col_del:
        if st.button("Rimuovi Allegato"):
            st.session_state["file_salvato_bytes"] = None
            st.session_state["file_salvato_nome"] = ""
            st.session_state["file_salvato_mime"] = ""
            st.rerun()
else:
    st.info("Nessun allegato presente (l'app genererà la lettera solo dalle note e dal modello Word).")

st.subheader("2. Note operative ed estremi notifica")
note_input = st.text_area(
    "Dettagli operativi / anagrafica / estremi (se non hai caricato un file o per precisazioni):",
    placeholder="Es. Notificato in data odierna a mani proprie del destinatario Rossi Mario nato a..."
)

def estrai_testo_docx(file_path):
    try:
        doc = Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception:
        return ""

def raccogli_esempi_stile():
    esempi = []
    for f in os.listdir(MODELLI_DIR):
        if f.endswith(".docx") and not f.startswith("~$"):
            percorso = MODELLI_DIR / f
            testo = estrai_testo_docx(percorso)
            if testo:
                esempi.append(f"--- ESEMPIO DA '{f}' ---\n{testo[:1500]}\n")
    return "\n".join(esempi)

def genera_docx_standard(data_str, prot_uscita, prot_rif, destinatari, oggetto, corpo):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    p_top = doc.add_paragraph()
    r_prot = p_top.add_run(f"Prot. {prot_uscita}")
    r_prot.bold = True
    r_prot.font.size = Pt(10)
    p_top.add_run("\t\t\t\t\t\t")
    r_data = p_top.add_run(f"Data: {data_str}")
    r_data.font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_before = Pt(18)

    p_dest = doc.add_paragraph()
    p_dest.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for riga in destinatari.split("\n"):
        if riga.strip():
            r = p_dest.add_run(riga.strip() + "\n")
            r.bold = True
            r.font.size = Pt(11)

    doc.add_paragraph().paragraph_format.space_before = Pt(14)

    p_ogg = doc.add_paragraph()
    r_ogg_l = p_ogg.add_run("OGGETTO: ")
    r_ogg_l.bold = True
    r_ogg = p_ogg.add_run(oggetto)
    r_ogg.bold = True

    doc.add_paragraph().paragraph_format.space_before = Pt(14)

    p_corpo = doc.add_paragraph()
    p_corpo.paragraph_format.line_spacing = 1.15
    p_corpo.paragraph_format.space_after = Pt(12)
    p_corpo.add_run(corpo)

    doc.add_paragraph().paragraph_format.space_before = Pt(36)

    p_firma = doc.add_paragraph()
    p_firma.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_firma.add_run("IL COMANDANTE\n(firma e timbro)")

    if prot_rif:
        doc.add_paragraph().paragraph_format.space_before = Pt(30)
        p_rif = doc.add_paragraph()
        r_rif = p_rif.add_run(prot_rif)
        r_rif.font.size = Pt(9)
        r_rif.italic = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def elabora_con_gemini(chiave, note_op, esempi_stile, indirizzo_fisso=None):
    client = genai.Client(api_key=chiave)
    contenuti = []

    prompt = f"""
Sei un addetto esperto delle Forze dell'Ordine / Pubblica Amministrazione italiana incaricato di redigere la lettera di trasmissione per la restituzione di un atto notificato.

REQUISITI DI REDAZIONE:
1. PROTOCOLLO IN BASSO A SINISTRA:
   - Estrai o individua gli estremi del protocollo, R.G.N.R. o riferimento della delega (es. "Rif. nota prot. n. ... del ...").

2. DESTINATARI (INDIRIZZO):
   - Se indicato nel testo o nel documento allegato, individua autorità/cancelleria e indirizzo/PEC.
   - Fornisci un "nome_breve_ente" sintetico per la rubrica.

3. OGGETTO:
   - DEVE iniziare con:
     "Trasmissione atti notificati a carico di [COGNOME Nome, nato a LUOGO (PROV) il GG/MM/AAAA, residente a COMUNE (PROV) in VIA/PIAZZA, domiciliato a COMUNE (PROV) in VIA/PIAZZA]"

4. CORPO DELLA LETTERA:
   - DEVE iniziare tassativamente con:
     "Allegato alla presente si restituisce debitamente notificato al soggetto sopra meglio indicato "
   - Prosegui con il tipo di atto e i riferimenti precisi, terminando con "per il prosieguo di competenza.".

ESEMPI DI STILE DELL'UFFICIO:
{esempi_stile}

NOTE OPERATIVE INSERITE:
{note_op if note_op else "Notifica eseguita regolarmente."}

Rispondi ESCLUSIVAMENTE con un JSON valido con questa struttura esatta:
{{
  "protocollo_riferimento": "Rif. prot. ...",
  "nome_breve_ente": "Nome sintetico ente",
  "destinatari": "Ente e Cancelleria\\nIndirizzo o PEC",
  "oggetto": "Trasmissione atti notificati a carico di ...",
  "corpo_lettera": "Allegato alla presente si restituisce debitamente notificato al soggetto sopra meglio indicato ..."
}}
"""
    contenuti.append(prompt)

    # Se c'è un file allegato salvato in sessione, inseriscilo nel prompt di Gemini
    if st.session_state["file_salvato_bytes"] is not None:
        b = st.session_state["file_salvato_bytes"]
        nome = st.session_state["file_salvato_nome"].lower()
        mime = st.session_state["file_salvato_mime"].lower()

        if nome.endswith(".pdf") or "pdf" in mime:
            contenuti.append(types.Part.from_bytes(data=b, mime_type="application/pdf"))
        elif any(nome.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".webp"]) or "image" in mime:
            tipo_img = mime if "image" in mime else "image/jpeg"
            contenuti.append(types.Part.from_bytes(data=b, mime_type=tipo_img))
        elif nome.endswith(".docx"):
            tdoc = Document(io.BytesIO(b))
            txt = "\n".join([p.text for p in tdoc.paragraphs if p.text.strip()])
            contenuti.append(f"\nTESTO ATTO:\n{txt}")
        else:
            contenuti.append(types.Part.from_bytes(data=b, mime_type="application/pdf"))

    config = types.GenerateContentConfig(
        response_mime_type="application/json"
    )

    try:
        res = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=contenuti,
            config=config
        )
    except Exception:
        res = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=contenuti,
            config=config
        )

    t = res.text.strip()
    if t.startswith("```json"):
        t = t[7:]
    if t.startswith("```"):
        t = t[3:]
    if t.endswith("```"):
        t = t[:-3]

    dati = json.loads(t.strip())
    if indirizzo_fisso:
        dati["destinatari"] = indirizzo_fisso
    return dati

st.divider()

# PULSANTE UNICO DI AVVIO
btn_elabora = st.button("🚀 Elabora e Genera Lettera di Trasmissione", type="primary", use_container_width=True)

if btn_elabora:
    if not chiave_attuale:
        st.error("Inserisci la chiave Gemini API nella barra laterale sinistra.")
    else:
        with st.status("Elaborazione e redazione lettera in corso...", expanded=True) as status:
            try:
                stile = raccogli_esempi_stile()
                ind_fisso = rubrica_enti.get(ente_selezionato) if not ente_selezionato.startswith("--") else None
                
                risultato = elabora_con_gemini(chiave_attuale, note_input, stile, ind_fisso)
                st.session_state["dati_elaborati"] = risultato

                if ente_selezionato.startswith("--") and risultato.get("nome_breve_ente") and risultato.get("destinatari"):
                    aggiorna_rubrica(risultato["nome_breve_ente"], risultato["destinatari"])

                status.update(label="Completato con successo!", state="complete", expanded=False)
                st.success("Lettera generata con successo!")
            except Exception as exc:
                status.update(label="Errore riscontrato", state="error", expanded=True)
                st.error(f"Errore: {exc}")
                st.code(traceback.format_exc())

# AREA DI DOWNLOAD E MODIFICA
if st.session_state.get("dati_elaborati"):
    dati = st.session_state["dati_elaborati"]
    st.markdown("---")
    st.subheader("Dati Generati (Verifica e Modifica)")

    cp1, cp2 = st.columns(2)
    with cp1:
        prot_u = st.text_input("Protocollo Uscita (in alto a sinistra):", value=protocollo_input)
    with cp2:
        prot_r = st.text_input("Protocollo/Riferimento Atto (in basso a sinistra):", value=dati.get("protocollo_riferimento", ""))

    c1, c2 = st.columns(2)
    with c1:
        dest_f = st.text_area("Destinatari / Ente:", value=dati.get("destinatari", ""), height=130)
    with c2:
        ogg_f = st.text_area("Oggetto Generato:", value=dati.get("oggetto", ""), height=130)

    corpo_f = st.text_area("Testo Trasmissione:", value=dati.get("corpo_lettera", ""), height=180)

    # Creazione documento Word finale
    if modello_scelto.startswith("--"):
        word_buf = genera_docx_standard(data_odierna_str, prot_u, prot_r, dest_f, ogg_f, corpo_f)
    else:
        doc = Document(MODELLI_DIR / modello_scelto)
        mappa = {
            "{{DATA}}": data_odierna_str,
            "{{PROTOCOLLO}}": prot_u,
            "{{PROTOCOLLO_USCITA}}": prot_u,
            "{{PROTOCOLLO_RIFERIMENTO}}": prot_r,
            "{{RIFERIMENTO}}": prot_r,
            "{{DESTINATARI}}": dest_f,
            "{{OGGETTO}}": ogg_f,
            "{{CORPO_LETTERA}}": corpo_f
        }
        for p in doc.paragraphs:
            for k, v in mappa.items():
                if k in p.text:
                    for r in p.runs:
                        if k in r.text:
                            r.text = r.text.replace(k, v)
        word_buf = io.BytesIO()
        doc.save(word_buf)
        word_buf.seek(0)

    st.download_button(
        label="📥 SCARICA LETTERA DI TRASMISSIONE (.DOCX)",
        data=word_buf,
        file_name=f"trasmissione_{oggi.strftime('%Y%m%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )
