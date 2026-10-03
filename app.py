import streamlit as st
from docx import Document
from google import genai
from google.genai import types
import datetime
import io
import json
import os
import time
from pathlib import Path

st.set_page_config(page_title="Automazione Notifiche e Atti", layout="wide")

# Directory di lavoro
BASE_DIR = Path(__file__).resolve().parent
MODELLI_DIR = BASE_DIR / "modelli_riferimento"
MODELLI_DIR.mkdir(parents=True, exist_ok=True)
RUBRICA_FILE = BASE_DIR / "enti_rubrica.json"

# Gestione API Key
def get_api_key():
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    if "gemini_key" in st.session_state:
        return st.session_state["gemini_key"]
    return ""

api_key = get_api_key()

if not api_key:
    st.sidebar.title("Configurazione")
    api_key_input = st.sidebar.text_input("Inserisci Gemini API Key", type="password")
    if api_key_input:
        st.session_state["gemini_key"] = api_key_input
        api_key = api_key_input
        st.rerun()

# Gestione Rubrica
def carica_rubrica():
    if "rubrica" not in st.session_state:
        if RUBRICA_FILE.exists():
            try:
                with open(RUBRICA_FILE, "r", encoding="utf-8") as f:
                    st.session_state["rubrica"] = json.load(f)
            except Exception:
                st.session_state["rubrica"] = {}
        else:
            st.session_state["rubrica"] = {}
    return st.session_state["rubrica"]

def aggiorna_rubrica(nome_ente, blocco_indirizzo):
    if not nome_ente or not blocco_indirizzo:
        return
    rubrica = carica_rubrica()
    rubrica[nome_ente.strip()] = blocco_indirizzo.strip()
    st.session_state["rubrica"] = rubrica
    try:
        with open(RUBRICA_FILE, "w", encoding="utf-8") as f:
            json.dump(rubrica, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

st.title("Generatore Lettere di Trasmissione")

oggi = datetime.date.today()
data_odierna_str = oggi.strftime("%d/%m/%Y")
anno_corrente = oggi.year

# Barra laterale per i Modelli Word
with st.sidebar:
    st.subheader("Modelli Word")
    nuovo_modello = st.file_uploader("Carica modello base (.docx)", type=["docx"], key="upload_modello")
    if nuovo_modello:
        salva_path = MODELLI_DIR / nuovo_modello.name
        with open(salva_path, "wb") as f:
            f.write(nuovo_modello.getbuffer())
        st.success(f"Modello {nuovo_modello.name} registrato!")
        st.rerun()

modelli_disponibili = [f for f in os.listdir(MODELLI_DIR) if f.endswith(".docx") and not f.startswith("~$")]

if not modelli_disponibili:
    st.warning("Carica prima il tuo modello Word (.docx) dal menu laterale a sinistra.")
    modello_scelto = None
else:
    modello_scelto = st.selectbox("Modello Word selezionato:", modelli_disponibili)

st.divider()

col_prot, col_data = st.columns(2)
with col_prot:
    protocollo_input = st.text_input("Protocollo Uscita (in alto a sinistra)", value=f"N. /{anno_corrente}")
with col_data:
    st.info(f"📅 **Data documento (in alto a destra):** {data_odierna_str}")

# Rubrica
rubrica_enti = carica_rubrica()
opzioni_enti = ["-- Rileva automaticamente dall'atto caricato --"] + list(rubrica_enti.keys())
ente_selezionato = st.selectbox("Seleziona Ente Destinatario dalla Rubrica (o lascia automatico):", opzioni_enti)

# Caricamento Atto - ACCETTA QUALSIASI TIPO DI FILE SENZA FILTRI
st.subheader("1. Atto/Richiesta di Notifica ricevuta")
file_atto = st.file_uploader(
    "Carica o scatta foto all'atto (PDF, Foto/Immagini o Word - qualsiasi formato):",
    type=None,
    key="uploader_atto_universale"
)

# Note operative
st.subheader("2. Note operative")
note_input = st.text_input(
    "Dettagli operativi (opzionale):",
    placeholder="Es. Notificato a mani proprie / irreperibile..."
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

def trova_modelli_validi(client):
    candidati = []
    try:
        for m in client.models.list():
            nome = getattr(m, "name", "").split("/")[-1]
            metodi = getattr(m, "supported_generation_methods", []) or getattr(m, "supported_actions", [])
            if not metodi or "generateContent" in str(metodi):
                candidati.append(nome)
    except Exception:
        pass
    priorita = ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash", "gemini-flash"]
    ordinati = [p for p in priorita if any(p in c for c in candidati)]
    return ordinati if ordinati else ["gemini-2.5-flash", "gemini-2.0-flash"]

def elabora_con_gemini(chiave, file_caricato, note_op, esempi_stile, indirizzo_fisso=None):
    client = genai.Client(api_key=chiave)
    contenuti = []

    prompt = f"""
Sei un addetto esperto delle Forze dell'Ordine / Pubblica Amministrazione italiana incaricato di redigere la lettera di trasmissione per la restituzione di un atto notificato.

ISTRUZIONI DI ESTRAZIONE E FORMULAZIONE:

1. PROTOCOLLO / RIFERIMENTO ATTO IN ENTRATA (PER LA POSIZIONE IN BASSO A SINISTRA):
   - Estrai il numero di protocollo, numero di registro/procedimento o identificativo dell'atto delegato dall'ente mittente (es. "Rif. nota prot. n. 1234/2026 del 12/03/2026" oppure "Proc. Pen. n. 567/2026 R.G.N.R.").

2. ENTE RICHIEDENTE / DESTINATARI (INDIRIZZI):
   - Se l'ente non è già preselezionato, individua con esattezza l'autorità o ufficio mittente (es. Tribunale, Procura, Questura, Prefettura, Ufficio NEP, ecc.), con cancelleria/sezione di competenza e indirizzo o PEC.
   - Fornisci un "nome_breve_ente" sintetico per la rubrica (es. "Tribunale di Reggio Calabria - Cancelleria Penale") e il blocco completo degli indirizzi.

3. DATI ANAGRAFICI:
   - Estrai: Cognome e Nome, data di nascita, luogo di nascita (con provincia), residenza (comune, prov, via) e domicilio (se indicato). Se residenza o domicilio non sono presenti, ometti solo la voce assente.

4. OGGETTO DELLA LETTERA:
   - Deve essere redatto TASSATIVAMENTE nel seguente formato:
     "Trasmissione atti notificati a carico di [COGNOME Nome, nato a LUOGO (PROV) il GG/MM/AAAA, residente a COMUNE (PROV) in VIA/PIAZZA, domiciliato a COMUNE (PROV) in VIA/PIAZZA]"

5. CORPO DELLA LETTERA:
   - Deve iniziare TASSATIVAMENTE con la formula:
     "Allegato alla presente si restituisce debitamente notificato al soggetto sopra meglio indicato "
   - Prosegui specificando l'atto notificato e i relativi estremi legali/giudiziari dedotti dai file (es. "l'atto di citazione a giudizio relativo al procedimento penale n. ... R.G.N.R. e n. ... R.G. DIB.").
   - Concludi con la formula: "per il prosieguo di competenza.".
   - Ricalca fedelmente lo stile formale presente in questi atti d'esempio:
{esempi_stile}

NOTE AGGIUNTIVE OPERATORE:
{note_op if note_op else "Notifica eseguita regolarmente."}

Rispondi ESCLUSIVAMENTE in formato JSON:
{{
  "protocollo_riferimento": "Rif. prot. n. ... del ... (o estremi atto)",
  "nome_breve_ente": "Nome sintetico per rubrica",
  "destinatari": "Riga 1 Ente e Ufficio\\nRiga 2 Indirizzo o PEC",
  "oggetto": "Oggetto completo conforme alle istruzioni",
  "corpo_lettera": "Testo completo che inizia con 'Allegato alla presente si restituisce debitamente notificato al soggetto sopra meglio indicato...'"
}}
"""
    contenuti.append(prompt)

    if file_caricato is not None:
        b = file_caricato.getvalue()
        nome_file = file_caricato.name.lower()
        mime_rilevato = file_caricato.type or ""

        # Rilevamento estensione e mime flessibile per accettare qualsiasi file
        if nome_file.endswith(".pdf") or "pdf" in mime_rilevato:
            contenuti.append(types.Part.from_bytes(data=b, mime_type="application/pdf"))
        elif any(nome_file.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".webp"]) or "image" in mime_rilevato:
            tipo_img = mime_rilevato if "image" in mime_rilevato else "image/jpeg"
            contenuti.append(types.Part.from_bytes(data=b, mime_type=tipo_img))
        elif nome_file.endswith(".docx"):
            try:
                tdoc = Document(io.BytesIO(b))
                txt = "\n".join([p.text for p in tdoc.paragraphs if p.text.strip()])
                contenuti.append(f"\nTESTO DELL'ATTO CARICATO:\n{txt}")
            except Exception:
                pass
        else:
            # Fallback generico: tenta come testo o PDF binario
            try:
                txt = b.decode("utf-8", errors="ignore")
                contenuti.append(f"\nCONTENUTO DELL'ATTO:\n{txt}")
            except Exception:
                contenuti.append(types.Part.from_bytes(data=b, mime_type="application/pdf"))

    lista_modelli = trova_modelli_validi(client)
    ultimo_err = None
    for mod in lista_modelli:
        for _ in range(2):
            try:
                res = client.models.generate_content(
                    model=mod,
                    contents=contenuti,
                    config={"response_mime_type": "application/json"}
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
            except Exception as e:
                ultimo_err = e
                if "404" in str(e) or "NOT_FOUND" in str(e):
                    break
                time.sleep(2)
                continue
    raise ultimo_err

def sostituisci_placeholder_docx(doc, sostituzioni):
    for p in doc.paragraphs:
        for k, v in sostituzioni.items():
            if k in p.text:
                for r in p.runs:
                    if k in r.text:
                        r.text = r.text.replace(k, v)

    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in sostituzioni.items():
                        if k in p.text:
                            for r in p.runs:
                                if k in r.text:
                                    r.text = r.text.replace(k, v)

st.divider()

if st.button("Elabora e Genera Lettera di Trasmissione"):
    if not api_key:
        st.error("Chiave Gemini API mancante.")
    elif not modello_scelto:
        st.error("Nessun modello Word disponibile. Caricalo dal menu laterale.")
    elif not file_atto:
        st.error("Carica o scatta una foto all'atto per procedere.")
    else:
        with st.spinner("Elaborazione e compilazione in corso..."):
            try:
                stile = raccogli_esempi_stile()
                indirizzo_prefissato = rubrica_enti.get(ente_selezionato) if ente_selezionato in rubrica_enti else None
                
                risultato = elabora_con_gemini(api_key, file_atto, note_input, stile, indirizzo_prefissato)
                
                prot_rif_estratto = risultato.get("protocollo_riferimento", "")
                dest_estratto = risultato.get("destinatari", "")
                ente_breve = risultato.get("nome_breve_ente", "")
                oggetto_estratto = risultato.get("oggetto", "")
                corpo_estratto = risultato.get("corpo_lettera", "")

                if ente_selezionato.startswith("--") and ente_breve and dest_estratto:
                    aggiorna_rubrica(ente_breve, dest_estratto)

                st.success("Dati estratti con successo!")

                c_p1, c_p2 = st.columns(2)
                with c_p1:
                    prot_uscita_val = st.text_input("Protocollo Uscita (in alto a sinistra):", value=protocollo_input)
                with c_p2:
                    prot_rif_val = st.text_input("Protocollo/Riferimento Atto (in basso a sinistra):", value=prot_rif_estratto)

                c1, c2 = st.columns(2)
                with c1:
                    dest_finale = st.text_area("Destinatari / Ente:", value=dest_estratto, height=130)
                with c2:
                    ogg_finale = st.text_area("Oggetto generato:", value=oggetto_estratto, height=130)

                corpo_finale = st.text_area("Testo trasmissione:", value=corpo_estratto, height=180)

                percorso_docx = MODELLI_DIR / modello_scelto
                doc = Document(percorso_docx)

                mappa = {
                    "{{DATA}}": data_odierna_str,
                    "{{PROTOCOLLO}}": prot_uscita_val,
                    "{{PROTOCOLLO_USCITA}}": prot_uscita_val,
                    "{{PROTOCOLLO_RIFERIMENTO}}": prot_rif_val,
                    "{{RIFERIMENTO}}": prot_rif_val,
                    "{{DESTINATARI}}": dest_finale,
                    "{{OGGETTO}}": ogg_finale,
                    "{{CORPO_LETTERA}}": corpo_finale
                }

                sostituisci_placeholder_docx(doc, mappa)

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)

                st.download_button(
                    label="📥 Scarica Lettera di Trasmissione Compilata (.docx)",
                    data=doc_buffer,
                    file_name=f"trasmissione_{datetime.date.today().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"Si è verificato un errore: {e}")
