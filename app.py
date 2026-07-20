import streamlit as st
import streamlit.components.v1 as components
from gtts import gTTS
from langdetect import detect
import tempfile
import os
import json
import torch
import pandas as pd
import fitz
import pytesseract
from PIL import Image
import io
from transformers import AutoTokenizer, AutoModelForSeq2SeqLM

# Optional-feature imports — wrapped so the app still runs (with that one
# feature disabled) if a package isn't installed yet.
try:
    import docx as docxlib
    HAVE_DOCX = True
except ImportError:
    HAVE_DOCX = False

try:
    import speech_recognition as sr
    HAVE_SR = True
except ImportError:
    HAVE_SR = False

try:
    from pydub import AudioSegment
    HAVE_PYDUB = True
except ImportError:
    HAVE_PYDUB = False

st.set_page_config(
    page_title="Ọ̀rọ̀ — NMT Translator",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------------------------
# DESIGN SYSTEM
# Paper background + deep green (Nigerian flag green, muted) as the single
# accent in light mode; the same palette inverted to a dark "ink" background
# with a brighter green in dark mode. Every text/background color in the
# app is driven from these variables, so nothing is hardcoded elsewhere.
# Serif display face for headings (evokes a dictionary / manuscript),
# monospace tags for language codes (evokes glossary entries).
# ---------------------------------------------------------------------------
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False
dark = st.session_state.dark_mode

if dark:
    BG        = "#14171A"
    PANEL     = "#1D2124"
    MUTED     = "#252B2E"
    INK       = "#EDEFEE"
    INK_SOFT  = "#9AA0A6"
    ACCENT    = "#2FBE7C"
    ACCENT_BG = "rgba(47,190,124,0.14)"
    BORDER    = "#2E3437"
    WARN      = "#E3A75B"
    DANGER    = "#E58880"
else:
    BG        = "#F7F6F2"
    PANEL     = "#FFFFFF"
    MUTED     = "#EFEDE6"
    INK       = "#17181A"
    INK_SOFT  = "#5B5D63"
    ACCENT    = "#00693E"
    ACCENT_BG = "rgba(0,105,62,0.08)"
    BORDER    = "#E2DFD5"
    WARN      = "#B5792B"
    DANGER    = "#B23A34"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    background-color: {BG};
    color: {INK};
}}
.stApp {{ background-color: {BG}; }}

section[data-testid="stSidebar"] {{
    background-color: {PANEL};
    border-right: 1px solid {BORDER};
}}
section[data-testid="stSidebar"] * {{ color: {INK} !important; }}

h1, h2, h3 {{ font-family: 'Fraunces', serif; }}

.brand {{
    font-family: 'Fraunces', serif;
    font-size: 26px;
    font-weight: 600;
    letter-spacing: -0.3px;
    color: {INK};
}}
.brand-sub {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    color: {INK_SOFT};
    margin-top: 2px;
}}

.stButton > button {{
    background-color: {ACCENT};
    color: #ffffff;
    border: none;
    border-radius: 3px;
    font-weight: 500;
    font-size: 14px;
    padding: 10px 22px;
    font-family: 'Inter', sans-serif;
    transition: opacity 0.15s ease;
}}
.stButton > button:hover {{ opacity: 0.85; }}

.stTextArea textarea {{
    background-color: {PANEL} !important;
    color: {INK} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 4px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 15px !important;
    padding: 14px !important;
}}
.stTextArea textarea:focus {{
    border-color: {ACCENT} !important;
    box-shadow: 0 0 0 1px {ACCENT} !important;
}}
.stTextArea textarea::placeholder {{
    color: {INK_SOFT} !important;
    opacity: 0.8 !important;
}}
/* Disabled textareas (translation results, OCR/PDF previews) otherwise get
   washed-out browser default styling that ignores our color — override the
   webkit fill color explicitly so text stays legible in both modes. */
.stTextArea textarea:disabled {{
    color: {INK} !important;
    -webkit-text-fill-color: {INK} !important;
    opacity: 1 !important;
    background-color: {MUTED} !important;
}}
.stSelectbox > div > div, .stSelectbox * {{
    color: {INK} !important;
}}
.stSelectbox > div > div {{
    background-color: {PANEL} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 4px !important;
}}
/* The dropdown option list renders in a portal outside .stApp, so it needs
   its own explicit theming or it falls back to browser default (white on
   white in dark mode, invisible). */
div[data-baseweb="popover"], div[data-baseweb="popover"] * {{
    background-color: {PANEL} !important;
    color: {INK} !important;
}}
li[role="option"]:hover, div[aria-selected="true"] {{
    background-color: {MUTED} !important;
}}
.stTextInput input {{
    background-color: {PANEL} !important;
    color: {INK} !important;
    border: 1px solid {BORDER} !important;
}}
.stMarkdown, .stMarkdown p, .stMarkdown li, .stMarkdown span {{
    color: {INK};
}}
.stCaption, [data-testid="stCaptionContainer"] {{
    color: {INK_SOFT} !important;
}}
.stDataFrame, .stDataFrame * {{
    color: {INK} !important;
}}
[data-testid="stMetricValue"], [data-testid="stMetricLabel"] {{
    color: {INK} !important;
}}
[data-testid="stFileUploaderDropzone"] {{
    background-color: {MUTED} !important;
    border: 1px dashed {BORDER} !important;
}}
[data-testid="stFileUploaderDropzone"] * {{
    color: {INK_SOFT} !important;
}}
div[data-testid="metric-container"] {{
    background-color: {PANEL};
    border: 1px solid {BORDER};
    border-radius: 4px;
    padding: 16px;
}}
hr {{ border: none; border-top: 1px solid {BORDER}; margin: 28px 0; }}

/* NOTE: we deliberately do NOT hide the <header> element — in current
   Streamlit versions that element also contains the sidebar collapse /
   expand arrow, so hiding it makes the sidebar unreachable. Only the
   hamburger menu and the "made with Streamlit" footer are hidden. */
footer {{ visibility: hidden; }}
#MainMenu {{ visibility: hidden; }}

.page-eyebrow {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    letter-spacing: 1.4px;
    text-transform: uppercase;
    color: {ACCENT};
    margin-bottom: 6px;
}}
.page-title {{
    font-family: 'Fraunces', serif;
    font-size: 30px;
    font-weight: 600;
    color: {INK};
    margin-bottom: 4px;
    letter-spacing: -0.4px;
}}
.page-sub {{
    font-size: 13px;
    color: {INK_SOFT};
    margin-bottom: 28px;
}}

.tag {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.4px;
    padding: 3px 9px;
    border: 1px solid {BORDER};
    border-radius: 3px;
    color: {INK_SOFT};
    background: {MUTED};
}}
.tag-accent {{
    border-color: {ACCENT};
    color: {ACCENT};
    background: {ACCENT_BG};
}}

.lang-label {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: {INK_SOFT};
    margin-bottom: 6px;
}}
.result-label {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: {ACCENT};
    margin-bottom: 6px;
}}
.char-counter {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: {INK_SOFT};
    text-align: right;
    margin-top: 4px;
}}
.card {{
    background: {PANEL};
    border: 1px solid {BORDER};
    border-radius: 6px;
    padding: 24px;
    margin-bottom: 16px;
}}
.info-row {{
    padding: 10px 14px;
    background: {ACCENT_BG};
    border-left: 2px solid {ACCENT};
    font-size: 13px;
    color: {INK};
    margin: 8px 0;
    border-radius: 0 3px 3px 0;
}}
.back-trans {{
    padding: 10px 14px;
    background: {MUTED};
    border-left: 2px solid {INK_SOFT};
    font-size: 13px;
    color: {INK};
    margin: 8px 0;
    border-radius: 0 3px 3px 0;
}}
.conf-badge {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 6px 14px;
    border-radius: 3px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12px;
    font-weight: 500;
    margin: 8px 0;
}}
.tip-box {{
    background: {MUTED};
    border: 1px solid {BORDER};
    border-radius: 6px;
    padding: 16px;
    font-size: 13px;
    color: {INK_SOFT};
    line-height: 1.8;
}}
table.entry {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
table.entry td {{ padding: 7px 0; border-bottom: 1px solid {BORDER}; }}
table.entry td:first-child {{
    color: {INK_SOFT};
    width: 40%;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}}
.fav-card {{
    background: {PANEL};
    border: 1px solid {BORDER};
    border-radius: 4px;
    padding: 14px 16px;
    margin-bottom: 10px;
}}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# LANGUAGE DATA
# ---------------------------------------------------------------------------
DETECT_MAP = {
    "en": "English",    "yo": "Yoruba",      "ig": "Igbo",
    "ha": "Hausa",      "fr": "French",      "es": "Spanish",
    "de": "German",     "pt": "Portuguese",  "it": "Italian",
    "nl": "Dutch",      "pl": "Polish",      "ru": "Russian",
    "uk": "Ukrainian",  "el": "Greek",       "cs": "Czech",
    "sk": "Slovak",     "hr": "Croatian",    "ro": "Romanian",
    "hu": "Hungarian",  "fi": "Finnish",     "sv": "Swedish",
    "no": "Norwegian",  "da": "Danish",      "bg": "Bulgarian",
    "zh": "Chinese",    "ja": "Japanese",    "ko": "Korean",
    "hi": "Hindi",      "bn": "Bengali",     "ta": "Tamil",
    "ur": "Urdu",       "vi": "Vietnamese",  "th": "Thai",
    "id": "Indonesian", "ms": "Malay",       "tl": "Filipino",
    "ar": "Arabic",     "tr": "Turkish",     "he": "Hebrew",
    "fa": "Persian",    "sw": "Swahili",     "so": "Somali",
    "am": "Amharic",
}

LANGUAGES = {
    "English":        {"nllb": "eng_Latn", "gtts": "en", "code": "eng", "sr": "en-US"},
    "Yoruba":         {"nllb": "yor_Latn", "gtts": "yo", "code": "yor", "sr": "en-US"},
    "Igbo":           {"nllb": "ibo_Latn", "gtts": "ig", "code": "ibo", "sr": "en-US"},
    "Hausa":          {"nllb": "hau_Latn", "gtts": "ha", "code": "hau", "sr": "en-US"},
    "Swahili":        {"nllb": "swh_Latn", "gtts": "sw", "code": "swh", "sr": "sw-KE"},
    "Amharic":        {"nllb": "amh_Ethi", "gtts": "am", "code": "amh", "sr": "am-ET"},
    "Somali":         {"nllb": "som_Latn", "gtts": "so", "code": "som", "sr": "en-US"},
    "French":         {"nllb": "fra_Latn", "gtts": "fr", "code": "fra", "sr": "fr-FR"},
    "Spanish":        {"nllb": "spa_Latn", "gtts": "es", "code": "spa", "sr": "es-ES"},
    "German":         {"nllb": "deu_Latn", "gtts": "de", "code": "deu", "sr": "de-DE"},
    "Portuguese":     {"nllb": "por_Latn", "gtts": "pt", "code": "por", "sr": "pt-PT"},
    "Italian":        {"nllb": "ita_Latn", "gtts": "it", "code": "ita", "sr": "it-IT"},
    "Dutch":          {"nllb": "nld_Latn", "gtts": "nl", "code": "nld", "sr": "nl-NL"},
    "Polish":         {"nllb": "pol_Latn", "gtts": "pl", "code": "pol", "sr": "pl-PL"},
    "Russian":        {"nllb": "rus_Cyrl", "gtts": "ru", "code": "rus", "sr": "ru-RU"},
    "Ukrainian":      {"nllb": "ukr_Cyrl", "gtts": "uk", "code": "ukr", "sr": "uk-UA"},
    "Greek":          {"nllb": "ell_Grek", "gtts": "el", "code": "ell", "sr": "el-GR"},
    "Czech":          {"nllb": "ces_Latn", "gtts": "cs", "code": "ces", "sr": "cs-CZ"},
    "Slovak":         {"nllb": "slk_Latn", "gtts": "sk", "code": "slk", "sr": "sk-SK"},
    "Croatian":       {"nllb": "hrv_Latn", "gtts": "hr", "code": "hrv", "sr": "hr-HR"},
    "Romanian":       {"nllb": "ron_Latn", "gtts": "ro", "code": "ron", "sr": "ro-RO"},
    "Hungarian":      {"nllb": "hun_Latn", "gtts": "hu", "code": "hun", "sr": "hu-HU"},
    "Finnish":        {"nllb": "fin_Latn", "gtts": "fi", "code": "fin", "sr": "fi-FI"},
    "Swedish":        {"nllb": "swe_Latn", "gtts": "sv", "code": "swe", "sr": "sv-SE"},
    "Norwegian":      {"nllb": "nob_Latn", "gtts": "no", "code": "nob", "sr": "no-NO"},
    "Danish":         {"nllb": "dan_Latn", "gtts": "da", "code": "dan", "sr": "da-DK"},
    "Bulgarian":      {"nllb": "bul_Cyrl", "gtts": "bg", "code": "bul", "sr": "bg-BG"},
    "Serbian":        {"nllb": "srp_Cyrl", "gtts": "sr", "code": "srp", "sr": "sr-RS"},
    "Chinese":        {"nllb": "zho_Hans", "gtts": "zh", "code": "zho", "sr": "zh-CN"},
    "Japanese":       {"nllb": "jpn_Jpan", "gtts": "ja", "code": "jpn", "sr": "ja-JP"},
    "Korean":         {"nllb": "kor_Hang", "gtts": "ko", "code": "kor", "sr": "ko-KR"},
    "Hindi":          {"nllb": "hin_Deva", "gtts": "hi", "code": "hin", "sr": "hi-IN"},
    "Bengali":        {"nllb": "ben_Beng", "gtts": "bn", "code": "ben", "sr": "bn-IN"},
    "Tamil":          {"nllb": "tam_Taml", "gtts": "ta", "code": "tam", "sr": "ta-IN"},
    "Telugu":         {"nllb": "tel_Telu", "gtts": "te", "code": "tel", "sr": "te-IN"},
    "Marathi":        {"nllb": "mar_Deva", "gtts": "mr", "code": "mar", "sr": "mr-IN"},
    "Gujarati":       {"nllb": "guj_Gujr", "gtts": "gu", "code": "guj", "sr": "gu-IN"},
    "Kannada":        {"nllb": "kan_Knda", "gtts": "kn", "code": "kan", "sr": "kn-IN"},
    "Malayalam":      {"nllb": "mal_Mlym", "gtts": "ml", "code": "mal", "sr": "ml-IN"},
    "Punjabi":        {"nllb": "pan_Guru", "gtts": "pa", "code": "pan", "sr": "pa-IN"},
    "Urdu":           {"nllb": "urd_Arab", "gtts": "ur", "code": "urd", "sr": "ur-PK"},
    "Vietnamese":     {"nllb": "vie_Latn", "gtts": "vi", "code": "vie", "sr": "vi-VN"},
    "Thai":           {"nllb": "tha_Thai", "gtts": "th", "code": "tha", "sr": "th-TH"},
    "Indonesian":     {"nllb": "ind_Latn", "gtts": "id", "code": "ind", "sr": "id-ID"},
    "Malay":          {"nllb": "zsm_Latn", "gtts": "ms", "code": "zsm", "sr": "ms-MY"},
    "Filipino":       {"nllb": "fil_Latn", "gtts": "tl", "code": "fil", "sr": "fil-PH"},
    "Nepali":         {"nllb": "npi_Deva", "gtts": "ne", "code": "npi", "sr": "ne-NP"},
    "Sinhala":        {"nllb": "sin_Sinh", "gtts": "si", "code": "sin", "sr": "si-LK"},
    "Arabic":         {"nllb": "arb_Arab", "gtts": "ar", "code": "arb", "sr": "ar-SA"},
    "Turkish":        {"nllb": "tur_Latn", "gtts": "tr", "code": "tur", "sr": "tr-TR"},
    "Hebrew":         {"nllb": "heb_Hebr", "gtts": "iw", "code": "heb", "sr": "he-IL"},
    "Persian":        {"nllb": "pes_Arab", "gtts": "fa", "code": "pes", "sr": "fa-IR"},
    "Haitian Creole": {"nllb": "hat_Latn", "gtts": "ht", "code": "hat", "sr": "fr-FR"},
}

LANG_NAMES = list(LANGUAGES.keys())

# ---------------------------------------------------------------------------
# DEVICE — forced to CPU. float32 is used deliberately (see previous notes:
# float16 is a poor fit for CPU inference in PyTorch).
# ---------------------------------------------------------------------------
DEVICE        = torch.device("cpu")
torch.set_num_threads(os.cpu_count() or 4)

# BASE now points to a local writable folder in the Streamlit Cloud
# container instead of Google Drive. Files here are NOT persistent
# across app restarts/redeploys -- that's expected for a free-tier demo.
BASE          = "."
FEEDBACK_FILE = BASE + "/feedback.json"
FAVORITES_FILE = BASE + "/favorites.json"


# Model + tokenizer are now pulled from the Hugging Face Hub repos
# instead of local Drive paths.
@st.cache_resource
def load_models():
    tokenizer = AutoTokenizer.from_pretrained("facebook/nllb-200-distilled-600M")
    yo = AutoModelForSeq2SeqLM.from_pretrained(
        "Oluwanifemianft/nllb-yoruba-nmt", torch_dtype=torch.float32
    ).to(DEVICE)
    ig = AutoModelForSeq2SeqLM.from_pretrained(
        "Oluwanifemianft/nllb-igbo-nmt", torch_dtype=torch.float32
    ).to(DEVICE)
    ha = AutoModelForSeq2SeqLM.from_pretrained(
        "Oluwanifemianft/nllb-hausa-nmt", torch_dtype=torch.float32
    ).to(DEVICE)
    yo.eval(); ig.eval(); ha.eval()
    return tokenizer, yo, ig, ha


def get_model(lang, yo, ig, ha):
    if lang == "Yoruba": return yo
    elif lang == "Igbo":  return ig
    elif lang == "Hausa": return ha
    else:                  return yo


def translate_nllb(text, tokenizer, model, tgt_lang, src_lang="eng_Latn"):
    tokenizer.src_lang = src_lang
    inputs = tokenizer(text, return_tensors="pt",
                        padding=True, truncation=True, max_length=256).to(DEVICE)
    with torch.no_grad():
        out = model.generate(
            **inputs,
            forced_bos_token_id=tokenizer.convert_tokens_to_ids(tgt_lang),
            max_new_tokens=128, num_beams=4, early_stopping=True
        )
    return tokenizer.decode(out[0], skip_special_tokens=True)


def translate_batch(lines, tokenizer, model, tgt_nllb, src_nllb):
    return [translate_nllb(l, tokenizer, model, tgt_nllb, src_nllb)
            for l in lines if l.strip()]


def get_confidence(text, tokenizer, model, src_nllb, tgt_nllb):
    try:
        tokenizer.src_lang = src_nllb
        inputs = tokenizer(text, return_tensors="pt",
                            truncation=True, max_length=256).to(DEVICE)
        with torch.no_grad():
            outputs = model.generate(
                **inputs,
                forced_bos_token_id=tokenizer.convert_tokens_to_ids(tgt_nllb),
                max_new_tokens=128, num_beams=4, early_stopping=True,
                return_dict_in_generate=True, output_scores=True
            )
        probs = [torch.softmax(s, dim=-1).max().item() for s in outputs.scores]
        return round(sum(probs) / len(probs) * 100, 1)
    except Exception:
        return None


def speak(text, gtts_code, slow=False):
    try:
        tts = gTTS(text=text, lang=gtts_code, slow=slow)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".mp3")
        tts.save(tmp.name)
        return tmp.name
    except Exception:
        return None


def auto_detect(text):
    try:
        code = detect(text)
        name = DETECT_MAP.get(code)
        if name and name in LANGUAGES:
            return name
    except Exception:
        pass
    return "English"


def save_feedback(source, translation, src, tgt, rating):
    entry = {"source": source, "translation": translation,
             "src": src, "tgt": tgt, "rating": rating}
    data = []
    if os.path.exists(FEEDBACK_FILE):
        with open(FEEDBACK_FILE) as f:
            data = json.load(f)
    data.append(entry)
    with open(FEEDBACK_FILE, "w") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_favorites():
    if os.path.exists(FAVORITES_FILE):
        try:
            with open(FAVORITES_FILE) as f:
                return json.load(f)
        except Exception:
            return []
    return []


def save_favorites(favs):
    try:
        with open(FAVORITES_FILE, "w") as f:
            json.dump(favs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def extract_pdf(pdf_bytes):
    doc   = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    for i, page in enumerate(doc):
        text = page.get_text().strip()
        if text:
            pages.append({"page": i + 1, "text": text})
    return pages


def extract_docx(file_bytes):
    """Returns a list of {'page': n, 'text': ...} shaped like extract_pdf,
    but chunked into blocks of ~25 paragraphs so the rest of the Documents
    page (which is written around a page list) works unchanged."""
    doc = docxlib.Document(io.BytesIO(file_bytes))
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    if not paras:
        return []
    chunk_size = 25
    chunks = []
    for i in range(0, len(paras), chunk_size):
        block = "\n".join(paras[i:i + chunk_size])
        chunks.append({"page": (i // chunk_size) + 1, "text": block})
    return chunks


def extract_txt(file_bytes):
    text = file_bytes.decode("utf-8", errors="ignore").strip()
    if not text:
        return []
    return [{"page": 1, "text": text}]


def extract_image(image_bytes):
    image = Image.open(io.BytesIO(image_bytes))
    return pytesseract.image_to_string(image).strip()


def transcribe_audio(audio_bytes, filename, lang_sr_code="en-US"):
    """Speech-to-text for an uploaded audio clip. Uses Google's free web
    speech API via SpeechRecognition, which requires internet access
    (available in Colab) but no API key."""
    if not HAVE_SR:
        return None, "SpeechRecognition is not installed."
    try:
        suffix = os.path.splitext(filename)[1].lower() or ".wav"
        tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp_in.write(audio_bytes)
        tmp_in.close()

        wav_path = tmp_in.name
        if suffix != ".wav":
            if not HAVE_PYDUB:
                return None, "Non-WAV audio needs pydub — upload a .wav file instead."
            audio = AudioSegment.from_file(tmp_in.name)
            wav_path = tmp_in.name + ".conv.wav"
            audio.export(wav_path, format="wav")

        r = sr.Recognizer()
        with sr.AudioFile(wav_path) as source:
            audio_data = r.record(source)
        text = r.recognize_google(audio_data, language=lang_sr_code)
        return text, None
    except sr.UnknownValueError:
        return None, "Could not understand the audio — try a clearer recording."
    except Exception as e:
        return None, "Transcription failed: " + str(e)


def wc(text):
    return len(text.split()) if text.strip() else 0


def lang_tag(name):
    info = LANGUAGES.get(name, {})
    return name + "  ·  " + info.get("code", "").upper()


def conf_html(conf):
    if conf is None:
        return ""
    color = ACCENT if conf >= 70 else WARN if conf >= 45 else DANGER
    label = "High" if conf >= 70 else "Medium" if conf >= 45 else "Low"
    return (
        "<div class='conf-badge' style='background:" + MUTED + ";color:" + color + ";"
        "border:1px solid " + color + "'>"
        "CONFIDENCE " + str(conf) + "% — " + label + "</div>"
    )


def copy_button(text, key):
    """Small JS-backed copy-to-clipboard button — the closest offline
    equivalent to Google Translate's copy icon."""
    safe = json.dumps(text)
    components.html(
        "<button id='cp_" + key + "' style=\""
        "background:" + PANEL + ";color:" + ACCENT + ";border:1px solid " + ACCENT + ";"
        "border-radius:3px;font-family:'IBM Plex Mono',monospace;font-size:11px;"
        "padding:6px 12px;cursor:pointer;width:100%;\" "
        "onclick=\"navigator.clipboard.writeText(" + safe + ");"
        "this.innerText='Copied';setTimeout(()=>{this.innerText='Copy';},1200);\">"
        "Copy</button>",
        height=38
    )


# ---------------------------------------------------------------------------
# SESSION STATE
# ---------------------------------------------------------------------------
defaults = {
    "history": [], "lang_counts": {}, "result": "",
    "last_src": "", "last_tgt": "",
    "prev_text": "", "prev_tgt": "",
    "favorites": load_favorites(),
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------
with st.sidebar:
    col_brand, col_theme = st.columns([4, 1])
    with col_brand:
        st.markdown(
            "<div style='padding:6px 0 4px'>"
            "<div class='brand'>Ọ̀rọ̀</div>"
            "<div class='brand-sub'>Neural Machine Translation</div>"
            "</div>",
            unsafe_allow_html=True
        )
    with col_theme:
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        if st.button("☀" if dark else "☾", help="Switch to " + ("light" if dark else "dark") + " mode"):
            st.session_state.dark_mode = not dark
            st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)

    page = st.radio(
        "",
        ["Translator", "Documents", "Image", "Voice", "Favorites", "Analytics", "About"],
        label_visibility="collapsed"
    )

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<div class='lang-label'>Settings</div>", unsafe_allow_html=True)
    use_auto     = st.checkbox("Auto-detect language",   value=True)
    auto_trans   = st.checkbox("Auto-translate on type", value=True)
    slow_speech  = st.checkbox("Slow speech",             value=False)
    show_back    = st.checkbox("Back-translation",        value=True)
    show_conf    = st.checkbox("Confidence score",        value=True)
    compare_mode = st.checkbox("Compare two languages",   value=False)

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown(
        "<div style='font-size:12px;line-height:1.9;color:" + INK_SOFT + "'>"
        "<div style='font-weight:600;color:" + INK + "'>Atuase Oluwanifemi Favour</div>"
        "<div>CSC/2022/1033</div>"
        "<div>Supervisor: Dr Fagbuagun</div>"
        "</div>",
        unsafe_allow_html=True
    )

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<div class='lang-label'>Session</div>", unsafe_allow_html=True)
    st.metric("Translations", len(st.session_state.history))
    if st.session_state.lang_counts:
        top = max(st.session_state.lang_counts, key=st.session_state.lang_counts.get)
        st.metric("Top language", top)
    st.markdown(
        "<span class='tag'>DEVICE · CPU</span> "
        "<span class='tag'>" + str(len(st.session_state.favorites)) + " SAVED</span> "
        "<span class='tag'>" + ("DARK" if dark else "LIGHT") + "</span>",
        unsafe_allow_html=True
    )

# ---------------------------------------------------------------------------
# LOAD MODELS
# ---------------------------------------------------------------------------
with st.spinner("Loading models on CPU — this can take a minute..."):
    nllb_tokenizer, yo_model, ig_model, ha_model = load_models()


def run_translation(text, src_lang, target_lang):
    """Shared translation routine used by Translator / Voice pages."""
    src_nllb = LANGUAGES[src_lang]["nllb"]
    tgt_nllb = LANGUAGES[target_lang]["nllb"]
    model    = get_model(target_lang, yo_model, ig_model, ha_model)
    if "\n" in text:
        lines = [l for l in text.splitlines() if l.strip()]
        return "\n".join(translate_batch(lines, nllb_tokenizer, model, tgt_nllb, src_nllb))
    return translate_nllb(text, nllb_tokenizer, model, tgt_nllb, src_nllb)


def add_favorite(source_text, result_text, src_lang, tgt_lang):
    entry = {"source": source_text, "translation": result_text,
              "src": src_lang, "tgt": tgt_lang}
    if entry not in st.session_state.favorites:
        st.session_state.favorites.append(entry)
        save_favorites(st.session_state.favorites)


# ---------------------------------------------------------------------------
# PAGE: TRANSLATOR
# ---------------------------------------------------------------------------
if page == "Translator":

    st.markdown("<div class='page-eyebrow'>01 · Translate</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-title'>Translation</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-sub'>Translate between 50+ languages, "
                "fine-tuned for Yoruba, Igbo and Hausa.</div>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([5, 1, 5])
    with col1:
        st.markdown("<div class='lang-label'>Source language</div>", unsafe_allow_html=True)
        source_lang = st.selectbox(
            "", LANG_NAMES, index=0, disabled=use_auto,
            key="src_lang", label_visibility="collapsed", format_func=lang_tag
        )
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        swap = st.button("⇄")
    with col3:
        st.markdown("<div class='lang-label'>Target language</div>", unsafe_allow_html=True)
        target_lang = st.selectbox(
            "", LANG_NAMES, index=1,
            key="tgt_lang", label_visibility="collapsed", format_func=lang_tag
        )

    if compare_mode:
        st.markdown("<div class='lang-label'>Second target language</div>", unsafe_allow_html=True)
        target_lang2 = st.selectbox(
            "", LANG_NAMES, index=2,
            key="tgt_lang2", label_visibility="collapsed", format_func=lang_tag
        )

    if swap and not use_auto:
        st.session_state.src_lang = target_lang
        st.session_state.tgt_lang = source_lang
        st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)

    col_src, col_tgt = st.columns(2)
    with col_src:
        st.markdown("<div class='lang-label'>Original text</div>", unsafe_allow_html=True)
        source_text = st.text_area(
            "", height=220, placeholder="Start typing to translate...",
            max_chars=500, key="source_input", label_visibility="collapsed"
        )
        char_count = len(source_text)
        c_color = DANGER if char_count > 450 else WARN if char_count > 350 else INK_SOFT
        st.markdown(
            "<div class='char-counter' style='color:" + c_color + "'>" +
            str(char_count) + " / 500</div>", unsafe_allow_html=True
        )
        if source_text.strip():
            copy_button(source_text, "src")

    with col_tgt:
        st.markdown("<div class='result-label'>Translation</div>", unsafe_allow_html=True)
        result_box = st.empty()
        result_box.text_area(
            "", height=220, value=st.session_state.result,
            disabled=True, key="result_static", label_visibility="collapsed"
        )
        if st.session_state.result:
            copy_button(st.session_state.result, "res")
            dl1, dl2, dl3 = st.columns(3)
            with dl1:
                st.download_button(
                    "Download", data=st.session_state.result,
                    file_name="translation.txt", mime="text/plain",
                    use_container_width=True
                )
            with dl2:
                if st.session_state.last_tgt and st.session_state.last_tgt in LANGUAGES:
                    tgt_gtts_dl = LANGUAGES[st.session_state.last_tgt]["gtts"]
                    audio_dl = speak(st.session_state.result.splitlines()[0],
                                      tgt_gtts_dl, slow=slow_speech)
                    if audio_dl:
                        st.audio(audio_dl)
            with dl3:
                if st.button("★ Save", use_container_width=True, key="fav_btn"):
                    add_favorite(source_text, st.session_state.result,
                                 st.session_state.last_src, st.session_state.last_tgt)
                    st.success("Saved to favorites")

    st.markdown("<br>", unsafe_allow_html=True)
    translate_btn = st.button("Translate", type="primary", use_container_width=True)

    should_translate = translate_btn
    if auto_trans and source_text.strip():
        if (source_text != st.session_state.prev_text or
                target_lang != st.session_state.prev_tgt):
            should_translate = True

    if should_translate and source_text.strip():
        src_lang = source_lang
        if use_auto:
            src_lang = auto_detect(source_text)
            st.markdown("<div class='info-row'>Detected: <strong>" +
                        lang_tag(src_lang) + "</strong></div>", unsafe_allow_html=True)

        if src_lang != target_lang:
            src_nllb = LANGUAGES[src_lang]["nllb"]
            tgt_nllb = LANGUAGES[target_lang]["nllb"]
            model    = get_model(target_lang, yo_model, ig_model, ha_model)

            with st.spinner("Translating on CPU..."):
                result = run_translation(source_text, src_lang, target_lang)

            st.session_state.result    = result
            st.session_state.last_src  = src_lang
            st.session_state.last_tgt  = target_lang
            st.session_state.prev_text = source_text
            st.session_state.prev_tgt  = target_lang

            result_box.text_area(
                "", height=220, value=result,
                disabled=True, key="result_filled", label_visibility="collapsed"
            )

            if compare_mode:
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown("<div class='lang-label'>Side-by-side comparison</div>",
                            unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("<span class='tag tag-accent'>" + lang_tag(target_lang) +
                                "</span>", unsafe_allow_html=True)
                    st.text_area("", value=result, height=160, disabled=True,
                                 key="cmp1", label_visibility="collapsed")
                with c2:
                    tgt_nllb2 = LANGUAGES[target_lang2]["nllb"]
                    model2    = get_model(target_lang2, yo_model, ig_model, ha_model)
                    result2   = translate_nllb(
                        source_text, nllb_tokenizer, model2, tgt_nllb2, src_nllb
                    )
                    st.markdown("<span class='tag tag-accent'>" + lang_tag(target_lang2) +
                                "</span>", unsafe_allow_html=True)
                    st.text_area("", value=result2, height=160, disabled=True,
                                 key="cmp2", label_visibility="collapsed")

            if show_conf:
                conf = get_confidence(source_text.splitlines()[0],
                                       nllb_tokenizer, model, src_nllb, tgt_nllb)
                st.markdown(conf_html(conf), unsafe_allow_html=True)

            if show_back:
                back_model = get_model(src_lang, yo_model, ig_model, ha_model)
                back = translate_nllb(result.splitlines()[0], nllb_tokenizer,
                                       back_model, src_nllb, tgt_nllb)
                st.markdown(
                    "<div class='back-trans'><span style='color:" + INK_SOFT + ";"
                    "font-weight:600;font-size:11px;text-transform:uppercase'>"
                    "Back-translation</span><br>"
                    "<span style='font-size:13px'>" + back + "</span></div>",
                    unsafe_allow_html=True
                )

            st.session_state.history.append({
                "From": src_lang, "To": target_lang,
                "Source": source_text[:60], "Translation": result[:60]
            })
            tgt_key = target_lang
            st.session_state.lang_counts[tgt_key] = (
                st.session_state.lang_counts.get(tgt_key, 0) + 1
            )

    if st.session_state.result:
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("<div class='lang-label'>Rate this translation</div>", unsafe_allow_html=True)
        fb1, fb2, fb3 = st.columns([1, 1, 6])
        with fb1:
            if st.button("Good"):
                save_feedback(source_text, st.session_state.result,
                              st.session_state.last_src, st.session_state.last_tgt, "good")
                st.success("Saved")
        with fb2:
            if st.button("Poor"):
                save_feedback(source_text, st.session_state.result,
                              st.session_state.last_src, st.session_state.last_tgt, "poor")
                st.info("Saved")

    if st.session_state.history:
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("<div class='lang-label'>History</div>", unsafe_allow_html=True)
        df = pd.DataFrame(list(reversed(st.session_state.history[-10:])))
        st.dataframe(df, use_container_width=True, hide_index=True)
        if st.button("Clear history"):
            st.session_state.history = []
            st.rerun()

# ---------------------------------------------------------------------------
# PAGE: DOCUMENTS  (PDF + DOCX + TXT — Google Translate calls this "Documents")
# ---------------------------------------------------------------------------
elif page == "Documents":

    st.markdown("<div class='page-eyebrow'>02 · Documents</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-title'>Document translation</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-sub'>Upload a PDF, Word (.docx) or text file — "
                "translated section by section.</div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        doc_auto = st.checkbox("Auto-detect source language", value=True, key="doc_auto")
        doc_src  = st.selectbox("Source", LANG_NAMES, index=0,
                                 key="doc_src", disabled=doc_auto, format_func=lang_tag)
    with col2:
        doc_tgt = st.selectbox("Target", LANG_NAMES, index=1,
                                key="doc_tgt", format_func=lang_tag)

    accepted_types = ["pdf", "txt"] + (["docx"] if HAVE_DOCX else [])
    if not HAVE_DOCX:
        st.markdown(
            "<div class='info-row'>.docx support needs the <code>python-docx</code> "
            "package — run <code>!pip install python-docx</code> to enable it.</div>",
            unsafe_allow_html=True
        )

    uploaded_doc = st.file_uploader("Upload document", type=accepted_types)

    if uploaded_doc:
        doc_bytes = uploaded_doc.read()
        ext = os.path.splitext(uploaded_doc.name)[1].lower()

        if ext == ".pdf":
            sections = extract_pdf(doc_bytes)
        elif ext == ".docx" and HAVE_DOCX:
            sections = extract_docx(doc_bytes)
        elif ext == ".txt":
            sections = extract_txt(doc_bytes)
        else:
            sections = []

        if not sections:
            st.error("No text found in this file. If it's a scanned PDF, try the Image tab instead.")
        else:
            st.markdown(
                "<div class='info-row'><strong>" + uploaded_doc.name + "</strong> — " +
                str(len(sections)) + " section(s) extracted</div>", unsafe_allow_html=True
            )

            with st.expander("Preview extracted text"):
                for p in sections[:3]:
                    st.markdown("**Section " + str(p["page"]) + "**")
                    st.text(p["text"][:400] + ("..." if len(p["text"]) > 400 else ""))

            max_sections = 1 if len(sections) == 1 else st.slider(
                "Sections to translate", 1, len(sections), min(5, len(sections))
            )

            if st.button("Translate document", type="primary", use_container_width=True):
                src_lang_doc = doc_src
                if doc_auto:
                    src_lang_doc = auto_detect(sections[0]["text"][:200])
                    st.markdown("<div class='info-row'>Detected: <strong>" +
                                lang_tag(src_lang_doc) + "</strong></div>", unsafe_allow_html=True)

                src_nllb = LANGUAGES[src_lang_doc]["nllb"]
                tgt_nllb = LANGUAGES[doc_tgt]["nllb"]
                tgt_gtts = LANGUAGES[doc_tgt]["gtts"]
                model    = get_model(doc_tgt, yo_model, ig_model, ha_model)

                all_translations = []
                prog = st.progress(0)

                for i, p in enumerate(sections[:max_sections]):
                    pct = int((i + 1) / max_sections * 100)
                    prog.progress(pct, text="Translating section " + str(p["page"]) + "...")
                    lines = [l for l in p["text"].splitlines() if l.strip()]
                    trans = translate_batch(lines[:30], nllb_tokenizer, model, tgt_nllb, src_nllb)
                    all_translations.append({
                        "page": p["page"], "original": p["text"],
                        "translation": "\n".join(trans)
                    })

                st.success("Done — " + str(len(all_translations)) + " section(s) translated")

                for item in all_translations:
                    with st.expander("Section " + str(item["page"])):
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown("<div class='lang-label'>Original</div>", unsafe_allow_html=True)
                            st.text_area("", value=item["original"], height=180, disabled=True,
                                         key="orig_" + str(item["page"]), label_visibility="collapsed")
                        with c2:
                            st.markdown("<div class='result-label'>" + lang_tag(doc_tgt) + "</div>",
                                        unsafe_allow_html=True)
                            st.text_area("", value=item["translation"], height=180, disabled=True,
                                         key="tr_" + str(item["page"]), label_visibility="collapsed")

                full = "\n\n".join([
                    "=== SECTION " + str(x["page"]) + " ===\n" + x["translation"]
                    for x in all_translations
                ])
                st.download_button(
                    "Download full translation", data=full,
                    file_name="translated_" + os.path.splitext(uploaded_doc.name)[0] + ".txt",
                    mime="text/plain", use_container_width=True
                )

                out_path = BASE + "/document_translation.txt"
                with open(out_path, "w", encoding="utf-8") as f:
                    f.write(full)
                st.markdown("<div class='back-trans'>Saved to Drive</div>", unsafe_allow_html=True)

                if all_translations:
                    first = all_translations[0]["translation"].splitlines()[0]
                    audio = speak(first, tgt_gtts)
                    if audio:
                        st.audio(audio)

                st.session_state.history.append({
                    "From": src_lang_doc, "To": doc_tgt,
                    "Source": "Doc: " + uploaded_doc.name,
                    "Translation": str(len(all_translations)) + " sections"
                })

# ---------------------------------------------------------------------------
# PAGE: IMAGE
# ---------------------------------------------------------------------------
elif page == "Image":

    st.markdown("<div class='page-eyebrow'>03 · Documents</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-title'>Image translation</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-sub'>OCR extracts text from an image, then translates it.</div>",
                unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        img_auto = st.checkbox("Auto-detect source language", value=True, key="img_auto")
        img_src  = st.selectbox("Source", LANG_NAMES, index=0,
                                 key="img_src", disabled=img_auto, format_func=lang_tag)
    with col2:
        img_tgt = st.selectbox("Target", LANG_NAMES, index=1,
                                key="img_tgt", format_func=lang_tag)

    uploaded_img = st.file_uploader("Upload image", type=["png", "jpg", "jpeg", "bmp", "tiff", "webp"])

    extracted = ""
    edited = ""

    if uploaded_img:
        image_bytes = uploaded_img.read()
        image       = Image.open(io.BytesIO(image_bytes))

        col_img, col_ocr = st.columns(2)
        with col_img:
            st.markdown("<div class='lang-label'>Uploaded image</div>", unsafe_allow_html=True)
            st.image(image, use_column_width=True)

        with col_ocr:
            st.markdown("<div class='lang-label'>Extracted text</div>", unsafe_allow_html=True)
            with st.spinner("Running OCR..."):
                extracted = extract_image(image_bytes)

            if not extracted.strip():
                st.warning("No text detected — try a clearer image.")
            else:
                st.markdown("<div class='info-row'>" + str(wc(extracted)) +
                            " words extracted</div>", unsafe_allow_html=True)
                edited = st.text_area("Edit before translating (optional)",
                                       value=extracted, height=180, key="ocr_edit")

        if extracted.strip():
            if st.button("Translate Image Text", type="primary", use_container_width=True):
                text_to_use  = edited.strip() or extracted.strip()
                src_lang_img = img_src
                if img_auto:
                    src_lang_img = auto_detect(text_to_use[:200])
                    st.markdown("<div class='info-row'>Detected: <strong>" +
                                lang_tag(src_lang_img) + "</strong></div>", unsafe_allow_html=True)

                src_nllb = LANGUAGES[src_lang_img]["nllb"]
                tgt_nllb = LANGUAGES[img_tgt]["nllb"]
                tgt_gtts = LANGUAGES[img_tgt]["gtts"]
                model    = get_model(img_tgt, yo_model, ig_model, ha_model)

                with st.spinner("Translating on CPU..."):
                    lines  = [l for l in text_to_use.splitlines() if l.strip()]
                    result = "\n".join(translate_batch(lines, nllb_tokenizer, model, tgt_nllb, src_nllb))

                st.markdown("<div class='result-label'>" + lang_tag(img_tgt) + "</div>",
                            unsafe_allow_html=True)
                st.text_area("", value=result, height=200, disabled=True,
                             key="img_result", label_visibility="collapsed")
                copy_button(result, "img_res")

                if show_conf and lines:
                    conf = get_confidence(lines[0], nllb_tokenizer, model, src_nllb, tgt_nllb)
                    st.markdown(conf_html(conf), unsafe_allow_html=True)

                if show_back and result:
                    back_model = get_model(src_lang_img, yo_model, ig_model, ha_model)
                    back = translate_nllb(result.splitlines()[0], nllb_tokenizer,
                                           back_model, src_nllb, tgt_nllb)
                    st.markdown(
                        "<div class='back-trans'><span style='color:" + INK_SOFT + ";"
                        "font-weight:600;font-size:11px;text-transform:uppercase'>"
                        "Back-translation</span><br>"
                        "<span style='font-size:13px'>" + back + "</span></div>",
                        unsafe_allow_html=True
                    )

                audio = speak(result.splitlines()[0], tgt_gtts, slow=slow_speech)
                if audio:
                    st.audio(audio)

                if st.button("★ Save to favorites", key="img_fav"):
                    add_favorite(text_to_use, result, src_lang_img, img_tgt)
                    st.success("Saved to favorites")

                combined = ("=== ORIGINAL (OCR) ===\n" + text_to_use +
                            "\n\n=== TRANSLATION (" + img_tgt + ") ===\n" + result)
                st.download_button("Download result", data=combined,
                                    file_name="image_translation.txt", mime="text/plain",
                                    use_container_width=True)

                out_path = BASE + "/image_translation.txt"
                with open(out_path, "w", encoding="utf-8") as f:
                    f.write(combined)

                st.session_state.history.append({
                    "From": src_lang_img, "To": img_tgt,
                    "Source": "Image: " + uploaded_img.name, "Translation": result[:60]
                })

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown(
        "<div class='tip-box'><strong>Tips for best OCR results</strong><br>"
        "Use high-resolution images with clear printed text. Good contrast and "
        "straight alignment improve accuracy. Handwriting is not supported. "
        "Formats: PNG, JPG, JPEG, BMP, TIFF, WEBP.</div>",
        unsafe_allow_html=True
    )

# ---------------------------------------------------------------------------
# PAGE: VOICE
# ---------------------------------------------------------------------------
elif page == "Voice":

    st.markdown("<div class='page-eyebrow'>04 · Speech</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-title'>Voice translation</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-sub'>Upload a short recording — it's transcribed, "
                "then translated. Closest offline-friendly match to Google Translate's "
                "microphone input, since live browser mic capture needs extra plugins "
                "this Colab/Streamlit setup doesn't have.</div>", unsafe_allow_html=True)

    if not HAVE_SR:
        st.markdown(
            "<div class='info-row'>This page needs the <code>SpeechRecognition</code> "
            "package. Run <code>!pip install SpeechRecognition pydub</code> in a cell above, "
            "then restart the app.</div>", unsafe_allow_html=True
        )
    else:
        col1, col2 = st.columns(2)
        with col1:
            voice_src = st.selectbox("Spoken language", LANG_NAMES, index=0,
                                      key="voice_src", format_func=lang_tag)
        with col2:
            voice_tgt = st.selectbox("Translate to", LANG_NAMES, index=1,
                                      key="voice_tgt", format_func=lang_tag)

        st.markdown(
            "<div class='tip-box'>Accepted formats: WAV works everywhere. "
            "MP3/M4A/OGG need <code>pydub</code> + <code>ffmpeg</code> installed "
            "in the Colab environment to convert automatically." +
            (" pydub is available." if HAVE_PYDUB else " pydub is NOT currently installed — "
             "upload WAV files or run <code>!pip install pydub</code>.") +
            "</div>", unsafe_allow_html=True
        )

        uploaded_audio = st.file_uploader(
            "Upload a voice recording", type=["wav", "mp3", "m4a", "ogg", "flac"]
        )

        if uploaded_audio:
            st.audio(uploaded_audio)
            if st.button("Transcribe & translate", type="primary", use_container_width=True):
                audio_bytes = uploaded_audio.read()
                sr_code = LANGUAGES[voice_src]["sr"]
                with st.spinner("Transcribing..."):
                    text, err = transcribe_audio(audio_bytes, uploaded_audio.name, sr_code)

                if err:
                    st.error(err)
                elif text:
                    st.markdown("<div class='lang-label'>Transcript</div>", unsafe_allow_html=True)
                    st.text_area("", value=text, height=100, disabled=True,
                                 key="voice_transcript", label_visibility="collapsed")

                    with st.spinner("Translating on CPU..."):
                        result = run_translation(text, voice_src, voice_tgt)

                    st.markdown("<div class='result-label'>" + lang_tag(voice_tgt) + "</div>",
                                unsafe_allow_html=True)
                    st.text_area("", value=result, height=100, disabled=True,
                                 key="voice_result", label_visibility="collapsed")
                    copy_button(result, "voice_res")

                    tgt_gtts = LANGUAGES[voice_tgt]["gtts"]
                    audio_out = speak(result.splitlines()[0], tgt_gtts, slow=slow_speech)
                    if audio_out:
                        st.audio(audio_out)

                    if st.button("★ Save to favorites", key="voice_fav"):
                        add_favorite(text, result, voice_src, voice_tgt)
                        st.success("Saved to favorites")

                    st.session_state.history.append({
                        "From": voice_src, "To": voice_tgt,
                        "Source": "Voice: " + uploaded_audio.name, "Translation": result[:60]
                    })

# ---------------------------------------------------------------------------
# PAGE: FAVORITES
# ---------------------------------------------------------------------------
elif page == "Favorites":

    st.markdown("<div class='page-eyebrow'>05 · Saved</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-title'>Favorites</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-sub'>Translations you've starred, kept across sessions.</div>",
                unsafe_allow_html=True)

    if not st.session_state.favorites:
        st.info("No favorites yet — star a translation from any page to save it here.")
    else:
        for i, fav in enumerate(reversed(st.session_state.favorites)):
            real_idx = len(st.session_state.favorites) - 1 - i
            st.markdown(
                "<div class='fav-card'>"
                "<span class='tag tag-accent'>" + lang_tag(fav.get("src", "")) + " → " +
                lang_tag(fav.get("tgt", "")) + "</span>"
                "<div style='margin-top:10px;font-size:14px'>" + fav.get("source", "") + "</div>"
                "<div style='margin-top:6px;font-size:14px;color:" + ACCENT + "'>" +
                fav.get("translation", "") + "</div>"
                "</div>", unsafe_allow_html=True
            )
            if st.button("Remove", key="rm_fav_" + str(real_idx)):
                st.session_state.favorites.pop(real_idx)
                save_favorites(st.session_state.favorites)
                st.rerun()

# ---------------------------------------------------------------------------
# PAGE: ANALYTICS
# ---------------------------------------------------------------------------
elif page == "Analytics":

    st.markdown("<div class='page-eyebrow'>06 · Results</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-title'>Analytics</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-sub'>Model performance and session statistics.</div>",
                unsafe_allow_html=True)

    st.markdown("<div class='lang-label'>Fine-tuned model performance</div>", unsafe_allow_html=True)
    m1, m2, m3 = st.columns(3)
    m1.metric("Yoruba BLEU", "7.95",  delta="chrF 30.41")
    m2.metric("Igbo BLEU",   "18.90", delta="chrF 46.06")
    m3.metric("Hausa BLEU",  "27.10", delta="chrF 54.99")

    st.markdown("<hr>", unsafe_allow_html=True)
    chart_data = pd.DataFrame({
        "Language": ["Yoruba", "Igbo", "Hausa"],
        "BLEU":     [7.95, 18.90, 27.10],
        "chrF":     [30.41, 46.06, 54.99]
    }).set_index("Language")
    st.bar_chart(chart_data, color=[ACCENT, INK_SOFT])

    st.markdown("<hr>", unsafe_allow_html=True)
    loss_data = pd.DataFrame({
        "Language": ["Yoruba", "Igbo", "Hausa"],
        "Epoch 1":  [2.698, 2.191, 1.815],
        "Epoch 2":  [2.578, 2.074, 1.736],
        "Epoch 3":  [2.551, 2.054, 1.735],
    }).set_index("Language")
    st.dataframe(loss_data, use_container_width=True)
    st.line_chart(loss_data.T)

    st.markdown("<hr>", unsafe_allow_html=True)
    if st.session_state.lang_counts:
        usage = pd.DataFrame(
            list(st.session_state.lang_counts.items()), columns=["Language", "Count"]
        ).set_index("Language").sort_values("Count", ascending=False)
        st.bar_chart(usage)
    else:
        st.info("No translations yet this session.")

    st.markdown("<hr>", unsafe_allow_html=True)
    if os.path.exists(FEEDBACK_FILE):
        with open(FEEDBACK_FILE) as f:
            fb = json.load(f)
        good = sum(1 for x in fb if x["rating"] == "good")
        poor = sum(1 for x in fb if x["rating"] == "poor")
        f1, f2, f3 = st.columns(3)
        f1.metric("Total feedback", len(fb))
        f2.metric("Good ratings",   good)
        f3.metric("Poor ratings",   poor)
    else:
        st.info("No feedback yet.")

# ---------------------------------------------------------------------------
# PAGE: ABOUT
# ---------------------------------------------------------------------------
elif page == "About":

    st.markdown("<div class='page-eyebrow'>07 · Project</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-title'>About</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-sub'>Project details, architecture and results.</div>",
                unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(
            "<div class='card'><div class='lang-label'>Project details</div>"
            "<table class='entry'>"
            "<tr><td>Title</td><td>Neural Machine Translation for Nigerian Languages</td></tr>"
            "<tr><td>Author</td><td>Atuase Oluwanifemi Favour</td></tr>"
            "<tr><td>Matric No</td><td>CSC/2022/1033</td></tr>"
            "<tr><td>Supervisor</td><td>Dr Fagbuagun</td></tr>"
            "<tr><td>Department</td><td>Computer Science</td></tr>"
            "</table></div>", unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            "<div class='card'><div class='lang-label'>Tech stack</div>"
            "<table class='entry'>"
            "<tr><td>Framework</td><td>PyTorch + HuggingFace</td></tr>"
            "<tr><td>Base model</td><td>NLLB-200 (600M params)</td></tr>"
            "<tr><td>TTS</td><td>gTTS</td></tr>"
            "<tr><td>OCR</td><td>Tesseract + PyMuPDF</td></tr>"
            "<tr><td>Speech-to-text</td><td>SpeechRecognition (Google Web Speech API)</td></tr>"
            "<tr><td>Documents</td><td>PyMuPDF + python-docx</td></tr>"
            "<tr><td>Inference device</td><td>CPU</td></tr>"
            "</table></div>", unsafe_allow_html=True
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    a1, a2, a3, a4, a5, a6 = st.columns(6)
    for col, num, title, desc in [
        (a1, "01", "Seq2Seq",   "PyTorch GRU encoder-decoder with Bahdanau attention"),
        (a2, "02", "NLLB-200",  "Meta fine-tuned for Yoruba, Igbo and Hausa"),
        (a3, "03", "Documents", "PDF, DOCX and TXT extracted section by section"),
        (a4, "04", "Image OCR", "Tesseract reads printed text from images"),
        (a5, "05", "Voice",     "Speech-to-text, then translated"),
        (a6, "06", "Favorites", "Starred translations saved to Drive"),
    ]:
        col.markdown(
            "<div class='card' style='text-align:center;padding:16px 10px'>"
            "<div style='font-family:IBM Plex Mono,monospace;font-size:18px;"
            "font-weight:500;color:" + ACCENT + ";opacity:0.6;margin-bottom:6px'>" + num + "</div>"
            "<div style='font-size:12px;font-weight:600;margin-bottom:4px'>" + title + "</div>"
            "<div style='font-size:10.5px;color:" + INK_SOFT + ";line-height:1.5'>" + desc + "</div>"
            "</div>", unsafe_allow_html=True
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    r1, r2, r3 = st.columns(3)
    for col, lang, bleu, chrf in [
        (r1, "Yoruba", "7.95",  "30.41"),
        (r2, "Igbo",   "18.90", "46.06"),
        (r3, "Hausa",  "27.10", "54.99"),
    ]:
        col.markdown(
            "<div class='card' style='text-align:center'>"
            "<span class='tag tag-accent'>" + lang.upper() + "</span>"
            "<div style='font-family:Fraunces,serif;font-size:30px;font-weight:600;"
            "color:" + ACCENT + ";margin:10px 0 2px'>" + bleu + "</div>"
            "<div style='font-size:11px;color:" + INK_SOFT + "'>BLEU</div>"
            "<div style='margin-top:10px;font-size:17px;font-weight:600'>" + chrf + "</div>"
            "<div style='font-size:11px;color:" + INK_SOFT + "'>chrF</div>"
            "</div>", unsafe_allow_html=True
        )

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown(
        "<div class='tip-box'>This project contributes to African language NLP by "
        "providing fine-tuned translation models for Yoruba, Igbo and Hausa — three "
        "of Nigeria's most widely spoken languages. Developed by Atuase Oluwanifemi "
        "Favour (CSC/2022/1033) under the supervision of Dr Fagbuagun.</div>",
        unsafe_allow_html=True
    )
