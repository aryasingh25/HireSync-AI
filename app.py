"""
AI Resume Screening & Candidate Shortlisting  (Gemini + Streamlit)
Run:  streamlit run app.py
"""
import io
import hashlib
from datetime import datetime
import json
import os
from dotenv import load_dotenv
import re
import shutil
import sqlite3
import subprocess
import tempfile
import time
from html import escape

import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document
from google import genai
from google.genai import types

# =====================================================================
# >>> PUT YOUR GEMINI API KEY ON THE NEXT LINE (between the quotes) <<<
load_dotenv()

try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    st.error("Gemini API key is missing!")
    st.stop()
# =====================================================================

# The app tries these models one by one until one works.
MODELS = ["gemini-flash-latest", "gemini-3-flash-preview", "gemini-2.5-flash",
          "gemini-3.1-flash-lite-preview"]
DELAY_BETWEEN_RESUMES = 2  # seconds (avoids free-tier rate limits)
IMAGE_TYPES = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "webp": "image/webp"}
RADAR_COLORS = ["#0E8F82", "#E0893A", "#4C7DF0", "#C0453A"]

st.set_page_config(page_title="HireSync AI", page_icon="✦", layout="wide",
                   initial_sidebar_state="collapsed")

# ------------------------------------------------------------------
# THEMES (light + dark)
# ------------------------------------------------------------------
LIGHT = {
    "bg": "#F5F7FA", "surface": "#FFFFFF", "input": "#FFFFFF", "text": "#14263A", "muted": "#5B6F84",
    "body": "#33475C", "border": "#E1E8F0", "input_border": "#C9D5E2", "primary": "#0E8F82",
    "on_primary": "#FFFFFF", "hero": "#0F2A43", "hero_text": "#B9C8D8", "link": "#0B6E9E",
    "ok_bg": "#DDF4F0", "ok_fg": "#076B61", "no_bg": "#FBE1DE", "no_fg": "#9B2C22",
    "mid_bg": "#FBEFD3", "mid_fg": "#8A5D07", "chip_bg": "#EAEFF5", "chip_fg": "#33475C",
}
DARK = {
    "bg": "#0D1520", "surface": "#152131", "input": "#1B2A3D", "text": "#E7EEF6", "muted": "#93A6BB",
    "body": "#C3D0DF", "border": "#26374B", "input_border": "#33475E", "primary": "#2BC4B0",
    "on_primary": "#062A25", "hero": "#13304D", "hero_text": "#BCD0E3", "link": "#6CB8F0",
    "ok_bg": "#12392F", "ok_fg": "#6EE7C8", "no_bg": "#43201F", "no_fg": "#FCA5A5",
    "mid_bg": "#453612", "mid_fg": "#F5C76B", "chip_bg": "#223246", "chip_fg": "#C3D0DF",
}

top_left, top_right = st.columns([6, 1])
with top_right:
    dark_mode = st.toggle("Dark mode", key="dark_mode")
T = DARK if dark_mode else LIGHT

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
html, body, .stApp, .stMarkdown, textarea, input, button { font-family: 'Manrope', sans-serif; }

.stApp, [data-testid="stAppViewContainer"] { background: var(--bg) !important; color: var(--text); }
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stSidebar"],
[data-testid="collapsedControl"], #MainMenu, footer { display: none !important; }
[data-testid="stElementContainer"]:has(style), .element-container:has(style) { display: none !important; }
.block-container { max-width: 1180px; padding-top: 1.2rem; padding-bottom: 4rem; }

/* ---- normal text ---- */
.stApp [data-testid="stMarkdownContainer"] p, .stApp [data-testid="stMarkdownContainer"] li,
.stApp [data-testid="stMarkdownContainer"] strong, .stApp label, .stApp label p,
.stApp [data-testid="stWidgetLabel"] p { color: var(--text) !important; }
.stApp [data-testid="stCaptionContainer"], .stApp [data-testid="stCaptionContainer"] p { color: var(--muted) !important; }
.stApp [data-testid="stSliderThumbValue"], .stApp [data-testid="stTickBarMin"],
.stApp [data-testid="stTickBarMax"] { color: var(--text) !important; }

/* ---- inputs ---- */
.stApp textarea, .stApp input:not([type="checkbox"]) { color: var(--text) !important;
  -webkit-text-fill-color: var(--text) !important; caret-color: var(--text) !important; background: var(--input) !important; }
.stApp textarea::placeholder, .stApp input::placeholder { color: var(--muted) !important;
  -webkit-text-fill-color: var(--muted) !important; opacity: 0.8; }
.stApp div[data-baseweb="textarea"], .stApp div[data-baseweb="input"], .stApp div[data-baseweb="base-input"]
  { background: var(--input) !important; }
.stApp div[data-baseweb="textarea"], .stApp div[data-baseweb="input"]
  { border: 1px solid var(--input_border) !important; border-radius: 10px !important; }
.stApp .stNumberInput button { background: var(--input) !important; color: var(--text) !important; }
.stApp div[data-baseweb="select"] > div { background: var(--input) !important;
  border: 1px solid var(--input_border) !important; border-radius: 10px !important; }
.stApp div[data-baseweb="select"] *, .stApp div[data-baseweb="select"] input
  { color: var(--text) !important; -webkit-text-fill-color: var(--text) !important; }
div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li, div[data-baseweb="menu"]
  { background: var(--surface) !important; color: var(--text) !important; }
div[data-baseweb="popover"] li * { color: var(--text) !important; }
.stApp span[data-baseweb="tag"] { background: var(--chip_bg) !important; }
.stApp span[data-baseweb="tag"] * { color: var(--chip_fg) !important; }

/* ---- file uploader ---- */
.stApp [data-testid="stFileUploaderDropzone"] { background: var(--input) !important;
  border: 1.5px dashed var(--input_border) !important; border-radius: 12px; }
.stApp [data-testid="stFileUploaderDropzone"] *, .stApp [data-testid="stFileUploaderFile"] *,
.stApp [data-testid="stFileUploader"] small { color: var(--body) !important; }
.stApp [data-testid="stFileUploaderDropzone"] button { background: var(--surface) !important;
  border: 1px solid var(--input_border) !important; }

/* ---- buttons, tabs, expanders, alerts ---- */
.stApp button[data-testid="stBaseButton-primary"] { background: var(--primary) !important; border: 0 !important;
  border-radius: 12px; padding: 0.7rem 1.5rem; }
.stApp button[data-testid="stBaseButton-primary"] * { color: var(--on_primary) !important; font-weight: 700; }
.stApp button[data-testid="stBaseButton-secondary"] { background: var(--surface) !important;
  border: 1px solid var(--input_border) !important; border-radius: 12px; }
.stApp button[data-testid="stBaseButton-secondary"] *,
.stApp button[data-testid="stBaseButton-secondary"]:hover *,
.stApp button[data-testid="stBaseButton-secondary"]:focus * { color: var(--text) !important; -webkit-text-fill-color: var(--text) !important; }
.stApp button[data-testid="stBaseButton-secondary"]:hover,
.stApp button[data-testid="stBaseButton-secondary"]:focus { background: var(--surface) !important; color: var(--text) !important; border-color: var(--primary) !important; }
.stApp button:not([data-testid="stBaseButton-primary"]):not([data-testid="stBaseButton-secondary"]):hover * { color: var(--text) !important; -webkit-text-fill-color: var(--text) !important; }
.stApp button[role="tab"] * { color: var(--muted) !important; }
.stApp button[role="tab"][aria-selected="true"] * { color: var(--primary) !important; font-weight: 700; }
.stApp [data-baseweb="tab-highlight"] { background: var(--primary) !important; }
.stApp [data-baseweb="tab-border"] { background: var(--border) !important; }
.stApp [data-testid="stExpander"] { background: var(--surface); border: 1px solid var(--border) !important; border-radius: 12px; }
.stApp [data-testid="stExpander"] summary, .stApp [data-testid="stExpander"] summary * { color: var(--text) !important; }
.stApp [data-testid="stAlert"] * { color: var(--text) !important; }
.stApp hr { border-color: var(--border) !important; }

/* ---- header ---- */
.hero-kicker { color: var(--primary); font-size: .74rem; font-weight: 800; letter-spacing: .14em; margin-bottom: 8px; }\n.brand { font-weight: 800; font-size: 1.05rem; color: var(--text); padding-top: 6px; }
.hero { background: linear-gradient(135deg, var(--hero), #173D5A); border-radius: 16px; padding: 34px 38px 30px; margin: 6px 0 26px; }
.stApp div.hero h1 { color: #FFFFFF !important; font-size: 2.1rem; font-weight: 800;
  letter-spacing: -0.02em; margin: 0 0 8px; padding: 0; }
.stApp div.hero p { color: var(--hero_text) !important; font-size: 1.02rem; max-width: 640px; margin: 0 0 22px; line-height: 1.55; }
.flow { display: flex; gap: 14px; flex-wrap: wrap; }
.flow div { display: flex; align-items: center; gap: 10px; background: rgba(255,255,255,0.10);
  padding: 9px 16px 9px 10px; border-radius: 999px; font-size: 0.9rem; font-weight: 600; color: #FFFFFF; }
.flow b { background: #2BC4B0; color: #0F2A43; width: 24px; height: 24px; border-radius: 50%;
  display: inline-flex; align-items: center; justify-content: center; font-size: 0.8rem; }
.sec { display: flex; align-items: center; gap: 10px; font-size: 1.1rem; font-weight: 700; color: var(--text); margin: 14px 0 8px; }
.sec b { background: var(--primary); color: var(--on_primary); width: 26px; height: 26px; border-radius: 50%;
  display: inline-flex; align-items: center; justify-content: center; font-size: 0.85rem; }

/* ---- stat tiles ---- */
.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin: 8px 0 22px; }
.stat { background: var(--surface); box-shadow: 0 8px 24px rgba(20,38,58,.05); border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px; }
.stat b { display: block; font-size: 1.9rem; font-weight: 800; color: var(--text); line-height: 1.1; }
.stat span { color: var(--muted); font-size: 0.88rem; font-weight: 500; }
.stat.hl { background: var(--primary); border-color: var(--primary); }
.stat.hl b, .stat.hl span { color: var(--on_primary); }

/* ---- candidate card ---- */
.cand { display: flex; background: var(--surface); border: 1px solid var(--border); border-radius: 14px; overflow: hidden; margin-top: 14px; }
.score { flex: 0 0 104px; display: flex; flex-direction: column; align-items: center; justify-content: center; color: #FFFFFF; }
.score b { font-size: 2.3rem; font-weight: 800; line-height: 1; }
.score span { font-size: 0.78rem; font-weight: 600; opacity: 0.92; margin-top: 4px; }
.score.good { background: #0E8F82; } .score.mid { background: #C98A12; } .score.low { background: #C0453A; }
.cbody { flex: 1; padding: 18px 22px 16px; min-width: 0; }
.chead { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; }
.cname { font-size: 1.2rem; font-weight: 800; color: var(--text); }
.crole { color: var(--muted); font-size: 0.93rem; margin-top: 2px; }
.badge { padding: 5px 13px; border-radius: 999px; font-size: 0.82rem; font-weight: 700; white-space: nowrap; }
.badge.good { background: var(--ok_bg); color: var(--ok_fg); } .badge.mid { background: var(--mid_bg); color: var(--mid_fg); }
.badge.low { background: var(--no_bg); color: var(--no_fg); }
.contact { display: flex; flex-wrap: wrap; gap: 6px 18px; margin: 12px 0 4px; font-size: 0.88rem; color: var(--body); }
.contact a { color: var(--link); text-decoration: none; font-weight: 600; }
.mini { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; }
.mini span { background: var(--chip_bg); color: var(--chip_fg); border-radius: 8px; padding: 3px 10px; font-size: 0.8rem; font-weight: 600; }
.mini span b { color: var(--text); }
.stApp div.cbody p.sum { color: var(--body) !important; font-size: 0.95rem; line-height: 1.55; margin: 10px 0 6px; }
.lbl { font-size: 0.83rem; font-weight: 700; color: var(--muted); margin: 12px 0 6px; }
.chips { display: flex; flex-wrap: wrap; gap: 6px; }
.chip { padding: 3px 11px; border-radius: 8px; font-size: 0.82rem; font-weight: 600; }
.chip.ok { background: var(--ok_bg); color: var(--ok_fg); } .chip.no { background: var(--no_bg); color: var(--no_fg); }
.chip.n { background: var(--chip_bg); color: var(--chip_fg); }
.empty { background: var(--surface); border: 1px dashed var(--input_border); border-radius: 14px; padding: 34px; text-align: center; color: var(--muted); }
.empty b { color: var(--text); font-size: 1.05rem; }

/* ---- table and bars ---- */
.tbl-wrap { overflow-x: auto; background: var(--surface); border: 1px solid var(--border); border-radius: 12px; }
table.tbl { width: 100%; border-collapse: collapse; font-size: 0.86rem; margin: 0; }
table.tbl th { text-align: left; padding: 11px 12px; color: var(--muted); font-weight: 700; white-space: nowrap;
  border-bottom: 1px solid var(--border); background: var(--surface); }
table.tbl td { padding: 10px 12px; color: var(--text); border-bottom: 1px solid var(--border); vertical-align: middle; background: var(--surface); }
.sbar { position: relative; min-width: 110px; height: 22px; background: var(--chip_bg); border-radius: 6px; overflow: hidden; }
.sbar i { position: absolute; left: 0; top: 0; bottom: 0; background: var(--primary); opacity: 0.85; }
.sbar b { position: relative; padding-left: 8px; font-size: 0.8rem; color: var(--text); line-height: 22px; }
.brow { display: flex; align-items: center; gap: 12px; margin: 10px 0; }
.brow .bl { flex: 0 0 230px; color: var(--text); font-size: 0.9rem; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.brow .bt { position: relative; flex: 1; height: 16px; background: var(--chip_bg); border-radius: 8px; }
.brow .bt i { position: absolute; left: 0; top: 0; bottom: 0; border-radius: 8px; }
.brow .bt u { position: absolute; top: -4px; bottom: -4px; border-left: 2px dashed var(--muted); text-decoration: none; }
.brow b { flex: 0 0 34px; color: var(--text); text-align: right; }

/* ---- stat tiles ---- */
.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin: 8px 0 22px; }
.stat { background: var(--surface); box-shadow: 0 8px 24px rgba(20,38,58,.05); border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px; }
.stat b { display: block; font-size: 1.9rem; font-weight: 800; color: var(--text); line-height: 1.1; }
.stat span { color: var(--muted); font-size: 0.88rem; font-weight: 500; }
.stat.hl { background: var(--primary); border-color: var(--primary); }
.stat.hl b, .stat.hl span { color: var(--on_primary); }

/* ---- candidate card ---- */
.cand { display: flex; background: var(--surface); border: 1px solid var(--border); border-radius: 14px; overflow: hidden; margin-top: 14px; }
.score { flex: 0 0 104px; display: flex; flex-direction: column; align-items: center; justify-content: center; color: #FFFFFF; }
.score b { font-size: 2.3rem; font-weight: 800; line-height: 1; }
.score span { font-size: 0.78rem; font-weight: 600; opacity: 0.92; margin-top: 4px; }
.score.good { background: #0E8F82; } .score.mid { background: #C98A12; } .score.low { background: #C0453A; }
.cbody { flex: 1; padding: 18px 22px 16px; min-width: 0; }
.chead { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; }
.cname { font-size: 1.2rem; font-weight: 800; color: var(--text); }
.crole { color: var(--muted); font-size: 0.93rem; margin-top: 2px; }
.badge { padding: 5px 13px; border-radius: 999px; font-size: 0.82rem; font-weight: 700; white-space: nowrap; }
.badge.good { background: var(--ok_bg); color: var(--ok_fg); } .badge.mid { background: var(--mid_bg); color: var(--mid_fg); }
.badge.low { background: var(--no_bg); color: var(--no_fg); }
.contact { display: flex; flex-wrap: wrap; gap: 6px 18px; margin: 12px 0 4px; font-size: 0.88rem; color: var(--body); }
.contact a { color: var(--link); text-decoration: none; font-weight: 600; }
.mini { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; }
.mini span { background: var(--chip_bg); color: var(--chip_fg); border-radius: 8px; padding: 3px 10px; font-size: 0.8rem; font-weight: 600; }
.mini span b { color: var(--text); }
.stApp div.cbody p.sum { color: var(--body) !important; font-size: 0.95rem; line-height: 1.55; margin: 10px 0 6px; }
.lbl { font-size: 0.83rem; font-weight: 700; color: var(--muted); margin: 12px 0 6px; }
.chips { display: flex; flex-wrap: wrap; gap: 6px; }
.chip { padding: 3px 11px; border-radius: 8px; font-size: 0.82rem; font-weight: 600; }
.chip.ok { background: var(--ok_bg); color: var(--ok_fg); } .chip.no { background: var(--no_bg); color: var(--no_fg); }
.chip.n { background: var(--chip_bg); color: var(--chip_fg); }
.empty { background: var(--surface); border: 1px dashed var(--input_border); border-radius: 14px; padding: 34px; text-align: center; color: var(--muted); }
.empty b { color: var(--text); font-size: 1.05rem; }

/* ---- table and bars ---- */
.tbl-wrap { overflow-x: auto; background: var(--surface); border: 1px solid var(--border); border-radius: 12px; }
table.tbl { width: 100%; border-collapse: collapse; font-size: 0.86rem; margin: 0; }
table.tbl th { text-align: left; padding: 11px 12px; color: var(--muted); font-weight: 700; white-space: nowrap;
  border-bottom: 1px solid var(--border); background: var(--surface); }
table.tbl td { padding: 10px 12px; color: var(--text); border-bottom: 1px solid var(--border); vertical-align: middle; background: var(--surface); }
.sbar { position: relative; min-width: 110px; height: 22px; background: var(--chip_bg); border-radius: 6px; overflow: hidden; }
.sbar i { position: absolute; left: 0; top: 0; bottom: 0; background: var(--primary); opacity: 0.85; }
.sbar b { position: relative; padding-left: 8px; font-size: 0.8rem; color: var(--text); line-height: 22px; }
.brow { display: flex; align-items: center; gap: 12px; margin: 10px 0; }
.brow .bl { flex: 0 0 230px; color: var(--text); font-size: 0.9rem; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.brow .bt { position: relative; flex: 1; height: 16px; background: var(--chip_bg); border-radius: 8px; }
.brow .bt i { position: absolute; left: 0; top: 0; bottom: 0; border-radius: 8px; }
.brow .bt u { position: absolute; top: -4px; bottom: -4px; border-left: 2px dashed var(--muted); text-decoration: none; }
.brow b { flex: 0 0 34px; color: var(--text); text-align: right; }

/* ---- HireSync feature cards ---- */
.feature-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:18px 0 24px; }
.feature-card { background:var(--surface); border:1px solid var(--border); border-radius:14px; padding:16px; min-height:108px; box-shadow:0 7px 22px rgba(20,38,58,.045); transition:transform .18s ease, box-shadow .18s ease; }
.feature-card:hover { transform:translateY(-2px); box-shadow:0 12px 28px rgba(20,38,58,.09); }
.feature-card .ico { font-size:1.25rem; margin-bottom:8px; }
.feature-card b { display:block; color:var(--text); font-size:.95rem; margin-bottom:4px; }
.feature-card span { color:var(--muted); font-size:.78rem; line-height:1.35; }
.insight-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin:12px 0; }
.insight { background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:14px; }
.insight h4 { margin:0 0 7px; color:var(--text); font-size:.9rem; }
.insight p { margin:0; color:var(--body); font-size:.82rem; line-height:1.45; }
.health-row { display:flex; justify-content:space-between; align-items:center; padding:9px 0; border-bottom:1px solid var(--border); font-size:.86rem; }
.health-row:last-child { border-bottom:0; }
.health-ok { color:var(--ok_fg); font-weight:700; }
.health-miss { color:var(--no_fg); font-weight:700; }
.pipeline { display:grid; grid-template-columns:repeat(6,1fr); gap:8px; margin:12px 0 20px; }
.pipeline-step { background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:12px 8px; text-align:center; }
.pipeline-step b { display:block; font-size:1.25rem; color:var(--text); }
.pipeline-step span { font-size:.72rem; color:var(--muted); }
@media (max-width: 900px) { .feature-grid {grid-template-columns:repeat(2,1fr)} .insight-grid {grid-template-columns:1fr} .pipeline {grid-template-columns:repeat(3,1fr)} }
@media (max-width: 760px) {
  .stats { grid-template-columns: repeat(2, 1fr); }
  .cand { flex-direction: column; } .score { flex-basis: auto; padding: 12px; flex-direction: row; gap: 8px; }
  .hero { padding: 24px 20px; } .brow .bl { flex-basis: 110px; }
}
"""
root_vars = ":root{" + "".join(f"--{k}:{v};" for k, v in T.items()) + "}"
st.markdown(f"<style>{root_vars}{CSS}</style>", unsafe_allow_html=True)
with top_left:
    st.markdown('<div class="brand">✦ HireSync AI</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------
# API KEY CHECK
# ------------------------------------------------------------------
if not API_KEY or API_KEY.startswith("PASTE_YOUR"):
    st.error("Gemini API key missing. Add GEMINI_API_KEY as an environment variable or in Streamlit secrets, then refresh the app.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# ------------------------------------------------------------------
# PERSISTENT LOCAL STORAGE
# ------------------------------------------------------------------
# Candidate metadata/status is stored in SQLite so it survives Streamlit
# reruns and app restarts. Resume binaries remain in data/resumes/.
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
RESUME_DIR = os.path.join(DATA_DIR, "resumes")
DB_FILE = os.path.join(DATA_DIR, "hiresync.db")
os.makedirs(RESUME_DIR, exist_ok=True)

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("""CREATE TABLE IF NOT EXISTS candidates (
        id TEXT PRIMARY KEY, file_name TEXT NOT NULL, stored_file TEXT,
        saved_at TEXT NOT NULL, job_description TEXT, result_json TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'AI screened'
    )""")
    conn.commit()
    return conn

def migrate_json_store():
    legacy = os.path.join(DATA_DIR, "analyses.json")
    if not os.path.exists(legacy):
        return
    conn = get_db()
    try:
        count = conn.execute("SELECT COUNT(*) FROM candidates").fetchone()[0]
        if count == 0:
            with open(legacy, "r", encoding="utf-8") as fh:
                old = json.load(fh)
            for r in old if isinstance(old, list) else []:
                result = r.get("result", {})
                conn.execute(
                    "INSERT OR IGNORE INTO candidates(id,file_name,stored_file,saved_at,job_description,result_json,status) VALUES(?,?,?,?,?,?,?)",
                    (r.get("id"), r.get("file_name", "resume"), r.get("stored_file"),
                     r.get("saved_at", datetime.now().isoformat(timespec="seconds")),
                     r.get("job_description", ""), json.dumps(result, ensure_ascii=False),
                     r.get("status", result.get("status", "AI screened")))
                )
            conn.commit()
    except Exception:
        pass
    finally:
        conn.close()

migrate_json_store()

def file_key(name, data):
    return hashlib.sha256(name.encode("utf-8") + data).hexdigest()[:20]

def safe_filename(name):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name)

def save_resume_and_result(file_name, file_bytes, result, job_description):
    key = file_key(file_name, file_bytes)
    stored_name = f"{key}_{safe_filename(file_name)}"
    resume_path = os.path.join(RESUME_DIR, stored_name)
    if not os.path.exists(resume_path):
        with open(resume_path, "wb") as fh:
            fh.write(file_bytes)
    conn = get_db()
    conn.execute("""INSERT OR REPLACE INTO candidates
        (id,file_name,stored_file,saved_at,job_description,result_json,status)
        VALUES(?,?,?,?,?,?,?)""",
        (key, file_name, stored_name, datetime.now().isoformat(timespec="seconds"),
         job_description, json.dumps(result, ensure_ascii=False), result.get("status", "AI screened")))
    conn.commit(); conn.close()
    return key

def load_saved_records():
    conn = get_db()
    rows = conn.execute("SELECT * FROM candidates ORDER BY saved_at ASC").fetchall()
    conn.close()
    return [{"id": r["id"], "file_name": r["file_name"], "stored_file": r["stored_file"],
             "saved_at": r["saved_at"], "job_description": r["job_description"],
             "result": json.loads(r["result_json"]), "status": r["status"]} for r in rows]

def load_saved_results():
    return [r.get("result", {}) for r in load_saved_records() if r.get("result")]

def delete_saved_record(record_id):
    conn = get_db()
    row = conn.execute("SELECT stored_file FROM candidates WHERE id=?", (record_id,)).fetchone()
    conn.execute("DELETE FROM candidates WHERE id=?", (record_id,))
    conn.commit(); conn.close()
    if row and row["stored_file"]:
        path = os.path.join(RESUME_DIR, row["stored_file"])
        if os.path.exists(path):
            os.remove(path)

def update_saved_record(record_id, **updates):
    if not updates:
        return
    conn = get_db()
    if "status" in updates:
        conn.execute("UPDATE candidates SET status=? WHERE id=?", (updates["status"], record_id))
    conn.commit(); conn.close()

def tokenize_skills(items):
    return {re.sub(r"[^a-z0-9+#.]", "", str(x).lower()) for x in items if str(x).strip()}

def resume_health(result):
    score = 0
    details = {}
    checks = [
        ("Contact details", bool(not is_missing(result.get("email")) or not is_missing(result.get("phone"))), 15),
        ("Skills section", len(result.get("skills", [])) >= 4, 20),
        ("Projects", len(result.get("projects", [])) >= 1, 20),
        ("Education", len(result.get("education", [])) >= 1, 15),
        ("Experience", len(result.get("companies", [])) >= 1 or result.get("experience_years", 0) == 0, 10),
        ("Professional links", bool(url_of(result.get("linkedin")) or url_of(result.get("github"))), 10),
        ("Certifications / extras", bool(result.get("certifications")), 10),
    ]
    for name, ok, weight in checks:
        details[name] = weight if ok else 0
        score += details[name]
    return score, details

def skill_gap_summary(row):
    missing = row.get("missing_skills", [])
    if not missing:
        return "No obvious skill gaps were detected against the current role."
    return "Priority areas: " + ", ".join(missing[:5]) + ". Consider these for interview validation or onboarding."

def saved_candidate_match(result, jd_skills):
    candidate = tokenize_skills(result.get("skills", []) + result.get("matched_skills", []))
    required = tokenize_skills(jd_skills)
    if not required:
        return int(to_num(result.get("ai_score", 0)))
    overlap = len(candidate & required) / max(1, len(required))
    return int(round(overlap * 100))


# ------------------------------------------------------------------
# READ RESUME FILES (pdf, docx, doc, txt, jpg, jpeg, png, webp)
# ------------------------------------------------------------------
def extract_pdf_text(file_bytes):
    text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text.strip()


def extract_docx_text(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text)
    return "\n".join(parts).strip()


def extract_doc_text(file_bytes):
    """Old Word .doc files. Tries 3 methods, best one first."""
    # 1) Some .doc files are really .docx files with the wrong extension
    try:
        return extract_docx_text(file_bytes)
    except Exception:  # noqa: BLE001
        pass
    # 2) LibreOffice (if installed on this computer) converts .doc to .docx
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if exe:
        try:
            with tempfile.TemporaryDirectory() as tmp:
                src = os.path.join(tmp, "resume.doc")
                with open(src, "wb") as fh:
                    fh.write(file_bytes)
                subprocess.run([exe, "--headless", "--convert-to", "docx", "--outdir", tmp, src],
                               capture_output=True, timeout=90)
                out = os.path.join(tmp, "resume.docx")
                if os.path.exists(out):
                    with open(out, "rb") as fh:
                        return extract_docx_text(fh.read())
        except Exception:  # noqa: BLE001
            pass
    # 3) Fallback: pull the readable text out of the raw file
    best = []
    for enc in ("cp1252", "utf-16le"):
        runs = re.findall(r"[A-Za-z0-9@.,:;/&()'\-+#%| ]{6,}", file_bytes.decode(enc, errors="ignore"))
        if sum(map(len, runs)) > sum(map(len, best)):
            best = runs
    return "\n".join(best).strip()


def read_resume(uploaded_file):
    """Returns (text, file_bytes, mime). mime is set for PDFs and images."""
    data = uploaded_file.getvalue()
    ext = uploaded_file.name.lower().rsplit(".", 1)[-1]
    if ext == "pdf":
        return extract_pdf_text(data), data, "application/pdf"
    if ext == "docx":
        return extract_docx_text(data), data, None
    if ext == "doc":
        return extract_doc_text(data), data, None
    if ext == "txt":
        return data.decode("utf-8", errors="ignore").strip(), data, None
    if ext in IMAGE_TYPES:
        return "", data, IMAGE_TYPES[ext]
    return "", data, None


# ------------------------------------------------------------------
# GEMINI CALLS
# ------------------------------------------------------------------
def parse_json(raw):
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?", "", raw).strip()
    raw = re.sub(r"```$", "", raw).strip()
    data = json.loads(raw)
    if isinstance(data, list) and data:
        data = data[0]
    return data


def call_gemini(contents):
    """Sends a request. Tries several models and retries on temporary errors."""
    order = MODELS[:]
    good = st.session_state.get("model")
    if good in order:
        order.remove(good)
        order.insert(0, good)

    last_error = None
    for model in order:
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=model, contents=contents,
                    config=types.GenerateContentConfig(response_mime_type="application/json"))
                data = parse_json(response.text)
                st.session_state["model"] = model
                return data
            except Exception as err:  # noqa: BLE001
                last_error = err
                msg = str(err)
                if "API key not valid" in msg or "API_KEY_INVALID" in msg:
                    raise RuntimeError("Your API key is not valid. Check the API_KEY line in app.py.")
                if "404" in msg or "NOT_FOUND" in msg:
                    break  # this model does not exist -> try the next one
                time.sleep(4 * (attempt + 1))
    raise RuntimeError(f"Gemini did not respond. Last error: {last_error}")


def get_jd_skills(job_description):
    """Key skills of the job used for candidate matching."""
    cache = st.session_state.setdefault("jd_cache", {})
    if job_description in cache:
        return cache[job_description]
    prompt = ("List the 6 to 8 most important skills or requirements in this job description. "
              "Each one is a short label of at most 3 words. "
              'Return ONLY JSON like {"skills": ["Python", "SQL"]}.\n\nJOB DESCRIPTION:\n' + job_description)
    try:
        skills = as_list(call_gemini(prompt).get("skills", []))[:8]
    except Exception:  # noqa: BLE001
        return []
    cache[job_description] = skills
    return skills


def analyze_job_description(job_description):
    prompt = f"""Analyze this job description for a recruiter. Return ONLY valid JSON.
{{
  "role_title": "...",
  "seniority": "...",
  "required_skills": ["..."],
  "preferred_skills": ["..."],
  "experience": "...",
  "education": "...",
  "responsibilities": ["..."],
  "quality_score": 0,
  "quality_notes": ["..."]
}}
JOB DESCRIPTION:
{job_description}
"""
    return call_gemini(prompt)

def ask_copilot(job_description, results, question):
    compact = []
    for r in results[:30]:
        compact.append({
            "name": r.get("name"), "score": r.get("ai_score", r.get("match_score", 0)),
            "skills": r.get("skills", [])[:15], "missing": r.get("missing_skills", [])[:10],
            "experience": r.get("experience_years", 0), "role": r.get("current_role", "")
        })
    prompt = f"""You are HireSync AI Recruiter Copilot. Answer the recruiter's question using ONLY the supplied job and candidate data. Be concise, practical and transparent. Never invent candidate facts.
JOB:
{job_description}
CANDIDATES:
{json.dumps(compact, ensure_ascii=False)}
QUESTION:
{question}
"""
    return call_gemini(prompt)

def build_prompt(job_description, jd_skills, resume_text=""):
    resume_part = (f"RESUME:\n{resume_text}" if resume_text
                   else "RESUME: The resume is attached as a file. Read it carefully.")
    if jd_skills:
        rating_rule = ("- skill_ratings: rate the candidate from 0 (no evidence) to 10 (expert) on EACH of these "
                       f"skills, using exactly these names as keys: {json.dumps(jd_skills)}")
        rating_json = '  "skill_ratings": {"skill name": 0},\n'
    else:
        rating_rule, rating_json = "", ""
    return f"""
You are an expert HR recruiter. Do two jobs for the resume below:
1) Extract the candidate's details exactly as written in the resume.
2) Score how well the candidate fits the job description.

RULES:
- Never invent information. If something is missing use "Not found" for text and [] for lists.
- experience_years = total professional work experience in years (0 for freshers, decimals allowed).
- Score ONLY on skills, experience, projects, education and achievements.
  Ignore name, gender, age, religion, nationality, photo and marital status when scoring.
- skills_score, experience_score, education_score, projects_score: each an integer 0-100 showing how well
  that area of the candidate fits the job. match_score is your overall integer score 0-100. Be fair but strict.
{rating_rule}

JOB DESCRIPTION:
{job_description}

{resume_part}

Return ONLY valid JSON (no markdown, no extra text) in exactly this format:
{{
  "name": "full name",
  "email": "email address",
  "phone": "phone number",
  "location": "city / country",
  "linkedin": "linkedin url",
  "github": "github or portfolio url",
  "current_role": "latest job title or 'Fresher / Student'",
  "experience_years": 0,
  "education": ["degree, institute, year"],
  "companies": ["job title at company (duration)"],
  "skills": ["every technical and soft skill found in the resume"],
  "certifications": ["certification"],
  "projects": ["project name: one line description"],
  "match_score": 0,
  "skills_score": 0,
  "experience_score": 0,
  "education_score": 0,
  "projects_score": 0,
{rating_json}  "matched_skills": ["skills from the job description that the candidate has"],
  "missing_skills": ["skills from the job description that the candidate lacks"],
  "strengths": "1-2 lines",
  "weaknesses": "1-2 lines",
  "summary": "2 line summary of the candidate for a recruiter",
  "match_explanation": "2-4 concise sentences explaining WHY this candidate matches the role, using only evidence from the resume",
  "recommendation": "Review"
}}
"""


def analyze_resume(job_description, jd_skills, text, file_bytes, mime):
    # Images and scanned PDFs are sent to Gemini as files (it reads them like a human).
    send_file = bool(mime) and (mime != "application/pdf" or len(text) < 100)
    if send_file:
        contents = [types.Part.from_bytes(data=file_bytes, mime_type=mime),
                    build_prompt(job_description, jd_skills)]
    else:
        contents = build_prompt(job_description, jd_skills, text)
    return call_gemini(contents)


def generate_questions(job_description, cand, difficulty, count):
    prompt = f"""
You are a senior interviewer. Write interview questions for this candidate applying for the job below.
Difficulty level: {difficulty}. Base the questions on the candidate's real projects, skills and skill gaps.

JOB DESCRIPTION:
{job_description}

CANDIDATE:
Name: {cand['name']}
Current role: {cand['current_role']}
Experience: {cand['experience_years']} years
Skills: {', '.join(cand['skills'])}
Projects: {' | '.join(cand['projects'])}
Skills matching the job: {', '.join(cand['matched_skills'])}
Skills missing for the job: {', '.join(cand['missing_skills'])}
Weaknesses: {cand['weaknesses']}

Return ONLY valid JSON in exactly this format:
{{
  "technical": [{{"question": "...", "why": "what it tests", "ideal_answer": "key points of a good answer"}}],
  "project": [{{"question": "...", "why": "...", "ideal_answer": "..."}}],
  "skill_gap": [{{"question": "...", "why": "...", "ideal_answer": "..."}}],
  "hr": [{{"question": "...", "why": "..."}}]
}}
Give exactly {count} technical questions, 2 project questions, 2 skill gap questions and 3 HR questions.
"""
    return call_gemini(prompt)


# ------------------------------------------------------------------
# CLEAN-UP HELPERS
# ------------------------------------------------------------------
def e(value):
    return escape(str(value))


def is_missing(value):
    return str(value).strip().lower() in ("", "not found", "unknown", "none", "n/a", "null")


def as_list(value):
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    return [] if is_missing(value) else [str(value).strip()]


def to_num(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def url_of(value):
    if is_missing(value):
        return ""
    url = str(value).strip()
    return url if url.lower().startswith("http") else "https://" + url


def clamp(value, low=0, high=100):
    return int(min(max(value, low), high))


def normalize(data, file_name):
    ai = clamp(to_num(data.get("match_score")))
    ratings = data.get("skill_ratings", {})
    ratings = ({str(k).lower().strip(): to_num(v) for k, v in ratings.items()}
               if isinstance(ratings, dict) else {})
    out = {
        "file_name": file_name,
        "name": data.get("name") if not is_missing(data.get("name", "")) else file_name,
        "email": data.get("email", "Not found"),
        "phone": data.get("phone", "Not found"),
        "location": data.get("location", "Not found"),
        "linkedin": data.get("linkedin", "Not found"),
        "github": data.get("github", "Not found"),
        "current_role": data.get("current_role", "Not found"),
        "experience_years": round(to_num(data.get("experience_years")), 1),
        "ai_score": ai,
        "skills_score": clamp(to_num(data.get("skills_score"), ai)),
        "experience_score": clamp(to_num(data.get("experience_score"), ai)),
        "education_score": clamp(to_num(data.get("education_score"), ai)),
        "projects_score": clamp(to_num(data.get("projects_score"), ai)),
        "skill_ratings": ratings,
        "strengths": data.get("strengths", ""),
        "weaknesses": data.get("weaknesses", ""),
        "summary": data.get("summary", ""),
        "match_explanation": data.get("match_explanation", ""),
        "ai_recommendation": data.get("recommendation", "Review"),
    }
    for key in ("education", "companies", "skills", "certifications", "projects",
                "matched_skills", "missing_skills"):
        out[key] = as_list(data.get(key, []))
    health, health_parts = resume_health(out)
    out["resume_health"] = health
    out["resume_health_parts"] = health_parts
    out["status"] = "AI screened"
    return out


# ------------------------------------------------------------------
# UI PIECES
# ------------------------------------------------------------------
def chips(items, cls, limit=None):
    shown = items[:limit] if limit else items
    html = "".join(f'<span class="chip {cls}">{e(i)}</span>' for i in shown)
    if limit and len(items) > limit:
        html += f'<span class="chip n">+{len(items) - limit} more</span>'
    return html


def decision(score):
    # No hard shortlisting/rejection criteria. The AI score is shown as evidence only.
    return "AI match", "good"


def candidate_card(row):
    score = row["match_score"]
    label, tone = decision(score)

    contact = []
    if not is_missing(row["email"]):
        contact.append(f'<span>✉ {e(row["email"])}</span>')
    if not is_missing(row["phone"]):
        contact.append(f'<span>☎ {e(row["phone"])}</span>')
    if not is_missing(row["location"]):
        contact.append(f'<span>⌖ {e(row["location"])}</span>')
    if url_of(row["linkedin"]):
        contact.append(f'<a href="{e(url_of(row["linkedin"]))}" target="_blank">LinkedIn</a>')
    if url_of(row["github"]):
        contact.append(f'<a href="{e(url_of(row["github"]))}" target="_blank">GitHub / Portfolio</a>')
    exp = row["experience_years"]
    contact.append("<span>" + ("Fresher" if exp == 0 else f"{exp:g} yrs experience") + "</span>")

    role = "" if is_missing(row["current_role"]) else e(row["current_role"])
    html = (
        '<div class="cand">'
        f'<div class="score {tone}"><b>{score}</b><span>final score</span></div>'
        '<div class="cbody">'
        f'<div class="chead"><div><div class="cname">{row["rank"]}. {e(row["name"])}</div>'
        f'<div class="crole">{role}</div></div>'
        f'<span class="badge {tone}">{label}</span></div>'
        f'<div class="contact">{"".join(contact)}</div>'
        '<div class="mini">'
        f'<span>Skills <b>{row["skills_score"]}</b></span><span>Experience <b>{row["experience_score"]}</b></span>'
        f'<span>Education <b>{row["education_score"]}</b></span><span>Projects <b>{row["projects_score"]}</b></span>'
        '</div>'
        f'<p class="sum">{e(row["summary"])}</p>'
    )
    if row["matched_skills"]:
        html += f'<div class="lbl">Matches the job</div><div class="chips">{chips(row["matched_skills"], "ok")}</div>'
    if row["missing_skills"]:
        html += f'<div class="lbl">Missing</div><div class="chips">{chips(row["missing_skills"], "no")}</div>'
    return html + "</div></div>"


def bullet_block(title, items):
    if items:
        st.markdown(f"**{title}**")
        st.markdown("\n".join(f"- {i}" for i in items))


def profile_details(row, axes=None, use_skills=False):
    c1, c2, c3 = st.columns([1.15, 1.15, 0.9])
    with c1:
        bullet_block("Education", row["education"])
        bullet_block("Certifications", row["certifications"])
    with c2:
        bullet_block("Work experience", row["companies"])
        bullet_block("Projects", row["projects"])
    with c3:
        st.metric("AI match", f"{int(row.get('match_score', row.get('ai_score', 0)))}%")
        st.metric("Skills", f"{int(row.get('skills_score', 0))}%")
        st.metric("Experience", f"{int(row.get('experience_score', 0))}%")

    if row["skills"]:
        st.markdown("**Skills found**")
        st.markdown(f'<div class="chips">{chips(row["skills"], "n", 30)}</div>', unsafe_allow_html=True)
    if row["strengths"]:
        st.markdown(f"**Strengths:** {row['strengths']}")
    if row["weaknesses"]:
        st.markdown(f"**Weaknesses:** {row['weaknesses']}")

    st.markdown("### Why this match?")
    st.info(row.get("match_explanation") or "The AI score is based on job-relevant skills, experience, projects, education and achievements found in the resume.")

    st.markdown("### Resume health")
    st.progress(row.get("resume_health", 0) / 100)
    st.caption(f"{row.get('resume_health', 0)}/100 · completeness and professional-readiness signal")
    parts = row.get("resume_health_parts", {})
    if parts:
        cols = st.columns(4)
        for i, (label, points) in enumerate(parts.items()):
            max_points = {"Contact details":15,"Skills section":20,"Projects":20,"Education":15,"Experience":10,"Professional links":10,"Certifications / extras":10}.get(label, 0)
            cols[i % 4].caption(f"{label}: {points}/{max_points}")

    st.markdown("### Skill gap")
    if row.get("matched_skills"):
        st.markdown(f"**Demonstrated:** {chips(row.get('matched_skills', []), 'ok')}", unsafe_allow_html=True)
    if row.get("missing_skills"):
        st.markdown(f"**Not evidenced:** {chips(row.get('missing_skills', []), 'no')}", unsafe_allow_html=True)
    st.caption(skill_gap_summary(row))


def table_html(d):
    heads = ["#", "Name", "Match", "Fit", "Exp (yrs)", "Role", "Email", "Phone", "Location", "Top skills"]
    out = ['<div class="tbl-wrap"><table class="tbl"><thead><tr>',
           "".join(f"<th>{h}</th>" for h in heads), "</tr></thead><tbody>"]
    for _, r in d.iterrows():
        label, tone = decision(r["match_score"])
        out.append(
            "<tr>"
            f'<td>{r["rank"]}</td><td><b>{e(r["name"])}</b></td>'
            f'<td><div class="sbar"><i style="width:{r["match_score"]}%"></i><b>{r["match_score"]}</b></div></td>'
            f'<td><span class="badge {tone}">{label}</span></td>'
            f'<td>{r["experience_years"]:g}</td><td>{e(r["current_role"])}</td><td>{e(r["email"])}</td>'
            f'<td>{e(r["phone"])}</td><td>{e(r["location"])}</td>'
            f'<td><div class="chips">{chips(r["skills"], "n", 5)}</div></td>'
            "</tr>")
    out.append("</tbody></table></div>")
    return "".join(out)


def questions_to_text(name, data):
    titles = {"technical": "TECHNICAL QUESTIONS", "project": "PROJECT QUESTIONS",
              "skill_gap": "SKILL GAP QUESTIONS", "hr": "HR QUESTIONS"}
    lines = [f"Interview questions for {name}", ""]
    for key, title in titles.items():
        lines.append(title)
        for i, q in enumerate(data.get(key, []), 1):
            q = q if isinstance(q, dict) else {"question": str(q)}
            lines.append(f"{i}. {q.get('question', '')}")
            if q.get("why"):
                lines.append(f"   Tests: {q['why']}")
            if q.get("ideal_answer"):
                lines.append(f"   Good answer: {q['ideal_answer']}")
        lines.append("")
    return "\n".join(lines)


# ------------------------------------------------------------------
# PAGE: HEADER
# ------------------------------------------------------------------
st.markdown("""
<div class="hero">
  <div class="hero-kicker">AI-POWERED RECRUITING WORKSPACE</div>
  <h1>Meet HireSync AI</h1>
  <p>Turn a folder of resumes into a clear, recruiter-friendly candidate view. HireSync AI extracts details, highlights skills and ranks profiles against your role.</p>
  <div class="flow">
    <div><b>1</b>Add the role</div>
    <div><b>2</b>Upload resumes</div>
    <div><b>3</b>Explore candidates</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""

""", unsafe_allow_html=True)

# ------------------------------------------------------------------
# PAGE: INPUTS
# ------------------------------------------------------------------
# Keep JD intelligence available across the whole page, including Overview.
jd_intel = st.session_state.get("jd_intel")

left, right = st.columns(2, gap="large")
with left:
    st.markdown('<div class="sec"><b>1</b>Job description</div>', unsafe_allow_html=True)
    job_description = st.text_area(
        "Job description", height=250, label_visibility="collapsed",
        placeholder="Example: Python Developer with 1-3 years of experience in Django, SQL and REST APIs. "
                    "Knowledge of Git and AWS is a plus.")
with right:
    st.markdown('<div class="sec"><b>2</b>Resumes</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Resumes", type=["pdf", "docx", "doc", "txt", "jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True, label_visibility="collapsed")
    st.caption("PDF, Word (doc, docx), text files and photos (jpg, png). Scanned resumes work too.")

st.markdown('<div class="sec"><b>3</b>Role intelligence</div>', unsafe_allow_html=True)
if job_description.strip():
    if st.button("Analyze job description", key="analyze_jd"):
        with st.spinner("Reading the role..."):
            try:
                st.session_state["jd_intel"] = analyze_job_description(job_description)
            except Exception as err:
                st.error(str(err))
    jd_intel = st.session_state.get("jd_intel")
    if jd_intel:
        a,b,c = st.columns(3)
        a.metric("JD quality", f"{int(to_num(jd_intel.get('quality_score', 0)))}%")
        b.metric("Required skills", len(jd_intel.get("required_skills", [])))
        c.metric("Preferred skills", len(jd_intel.get("preferred_skills", [])))
        st.markdown(f"**{e(jd_intel.get('role_title', 'Role'))}** · {e(jd_intel.get('seniority', 'Seniority not specified'))}")
        if jd_intel.get("required_skills"):
            st.markdown(f"**Required:** {chips(as_list(jd_intel['required_skills']), 'ok')}", unsafe_allow_html=True)
        if jd_intel.get("preferred_skills"):
            st.markdown(f"**Preferred:** {chips(as_list(jd_intel['preferred_skills']), 'n')}", unsafe_allow_html=True)
        if jd_intel.get("quality_notes"):
            st.caption(" · ".join(map(str, jd_intel["quality_notes"][:3])))
else:
    st.caption("Add a job description to unlock role intelligence, candidate matching and rediscovery.")

analyze_clicked = st.button("✦ Analyze resumes", type="primary", use_container_width=True)

# ------------------------------------------------------------------
# RUN ANALYSIS
# ------------------------------------------------------------------
if analyze_clicked:
    if not job_description.strip():
        st.warning("Add a job description first.")
    elif not uploaded_files:
        st.warning("Upload at least one resume.")
    else:
        results, errors = [], []
        bar = st.progress(0.0)
        note = st.empty()
        note.info("Reading the job description...")
        jd_skills = get_jd_skills(job_description)
        st.session_state["jd_skills"] = jd_skills
        st.session_state["job_description"] = job_description
        st.session_state["questions"] = {}
        for i, f in enumerate(uploaded_files):
            note.info(f"Reading {f.name} ({i + 1} of {len(uploaded_files)})")
            try:
                text, file_bytes, mime = read_resume(f)
                if not text and not mime:
                    raise ValueError("No readable text found in this file.")
                data = analyze_resume(job_description, jd_skills, text, file_bytes, mime)
                normalized = normalize(data, f.name)
                save_resume_and_result(f.name, file_bytes, normalized, job_description)
                results.append(normalized)
            except Exception as err:  # noqa: BLE001
                errors.append(f"{f.name}: {err}")
            bar.progress((i + 1) / len(uploaded_files))
            if i < len(uploaded_files) - 1:
                time.sleep(DELAY_BETWEEN_RESUMES)
        note.empty()
        bar.empty()
        st.session_state["results"] = results
        st.session_state["errors"] = errors

saved_records = load_saved_records()

# ------------------------------------------------------------------
# RESULTS
for msg in st.session_state.get("errors", []):
    st.error(msg)

results = st.session_state.get("results")
if not results:
    results = load_saved_results()
    if results:
        st.session_state["results"] = results

if not results:
    st.markdown('<div class="empty"><b>Start screening</b><br>Add a job description, upload resumes and click <b>Analyze resumes</b>. HireSync AI will extract the candidate details, match them to the role and save the profiles for later.</div>', unsafe_allow_html=True)
    st.stop()

jd_skills = st.session_state.get("jd_skills", [])
jd_text = st.session_state.get("job_description", job_description)
df = pd.DataFrame(results)
df["match_score"] = df["ai_score"].astype(int)
df = df.sort_values(["match_score", "experience_years"], ascending=False).reset_index(drop=True)
df.insert(0, "rank", df.index + 1)
saved_records = load_saved_records()

# Simple, non-cluttered summary: useful orientation, not a separate analytics feature.
st.markdown(f"""
<div class="stats">
  <div class="stat"><b>{len(df)}</b><span>Resumes screened</span></div>
  <div class="stat hl"><b>{df['match_score'].mean():.0f}%</b><span>Average role match</span></div>
  <div class="stat"><b>{int(df['match_score'].max())}%</b><span>Top AI match</span></div>
  <div class="stat"><b>{len(saved_records)}</b><span>Saved profiles</span></div>
</div>
""", unsafe_allow_html=True)

export_df = df.copy()
for col in ["education", "companies", "skills", "certifications", "projects", "matched_skills", "missing_skills"]:
    export_df[col] = export_df[col].apply(lambda items: "; ".join(items))
export_df["skill_ratings"] = export_df["skill_ratings"].apply(lambda d: "; ".join(f"{k}: {v:g}" for k, v in d.items()))

# Keep the product understandable: five destinations, with secondary features nested where they belong.
tab_overview, tab_candidates, tab_hiring, tab_interview, tab_library = st.tabs(
    ["Overview", "Candidates", "Hiring", "Interview Kit", "Resume Library"])

# ---- Overview ----
with tab_overview:
    st.markdown("### Your hiring workspace")
    st.caption("The simple flow: define the role → screen resumes → review AI evidence → move candidates through hiring.")

    o1, o2 = st.columns([1.15, 1])
    with o1:
        st.markdown("#### Current role")
        if jd_intel:
            st.markdown(f"**{e(jd_intel.get('role_title', 'Role'))}** · {e(jd_intel.get('seniority', 'Seniority not specified'))}")
            st.caption(f"JD quality: {int(to_num(jd_intel.get('quality_score', 0)))}%")
            if jd_intel.get("required_skills"):
                st.markdown(f"**Required:** {chips(as_list(jd_intel['required_skills']), 'ok')}", unsafe_allow_html=True)
            if jd_intel.get("preferred_skills"):
                st.markdown(f"**Preferred:** {chips(as_list(jd_intel['preferred_skills']), 'n')}", unsafe_allow_html=True)
        else:
            st.info("Analyze the job description above to see the role summary.")

    with o2:
        st.markdown("#### Top candidates")
        for _, rr in df.head(3).iterrows():
            st.markdown(f"**{e(rr['name'])}** · {int(rr['match_score'])}% match")
            st.progress(clamp(to_num(rr['match_score'])) / 100)

    st.markdown("#### Ask HireSync")
    st.caption("Ask a quick question about the candidates instead of searching through every profile.")
    question = st.text_input("Recruiter question", placeholder="Who has the strongest Python and SQL background?", key="overview_copilot_question")
    if st.button("Ask HireSync", type="primary", key="overview_ask_copilot") and question.strip():
        with st.spinner("Thinking..."):
            try:
                answer = ask_copilot(jd_text, results, question)
                if isinstance(answer, dict):
                    answer = answer.get("answer", json.dumps(answer, indent=2))
                st.markdown(f"<div class='copilot'>{e(answer)}</div>", unsafe_allow_html=True)
            except Exception as err:
                st.error(str(err))

# ---- Candidates + comparison ----
with tab_candidates:
    st.markdown("### Candidates")
    st.caption("Open a profile to see the AI evidence, resume health and skill gaps. The AI score ranks candidates; it does not make the final hiring decision.")
    search = st.text_input("Search candidates", placeholder="Name, skill, role or location", key="candidate_search")
    shown = df
    if search.strip():
        mask = export_df.apply(lambda r: search.lower() in " ".join(map(str, r.values)).lower(), axis=1)
        shown = df[mask]
    for _, row in shown.iterrows():
        st.markdown(candidate_card(row), unsafe_allow_html=True)
        with st.expander(f"View {row['name']} — match evidence & resume health"):
            profile_details(row)

    with st.expander("Compare candidates"):
        labels = {f"{r['rank']}. {r['name']}": i for i, r in df.iterrows()}
        picked = st.multiselect("Select 2–3 candidates", list(labels), max_selections=3, key="compare_candidates")
        if len(picked) >= 2:
            rows = [df.loc[labels[x]] for x in picked]
            cols = st.columns(len(rows))
            for col, r in zip(cols, rows):
                with col:
                    st.markdown(f"### {e(r['name'])}")
                    st.metric("AI match", f"{r['match_score']}%")
                    st.progress(r["skills_score"] / 100); st.caption(f"Skills · {r['skills_score']}%")
                    st.progress(r["projects_score"] / 100); st.caption(f"Projects · {r['projects_score']}%")
                    st.progress(r["experience_score"] / 100); st.caption(f"Experience · {r['experience_score']}%")
                    st.progress(r["resume_health"] / 100); st.caption(f"Resume health · {r['resume_health']}%")
        else:
            st.caption("Select at least two candidates to compare them side by side.")

    st.download_button("Export candidate results", data=export_df.to_csv(index=False).encode("utf-8"),
                       file_name="hiresync_candidates.csv", mime="text/csv")

# ---- Hiring pipeline ----
with tab_hiring:
    st.markdown("### Hiring pipeline")
    st.caption("Move candidates through your process. This status is controlled by you, not by the AI.")
    statuses = ["AI screened", "Shortlisted", "Interview", "Selected", "On hold", "Rejected"]
    counts = {s: 0 for s in statuses}
    for rec in saved_records:
        counts[rec.get("status", "AI screened")] = counts.get(rec.get("status", "AI screened"), 0) + 1
    st.markdown(" · ".join(f"**{s}: {counts[s]}**" for s in statuses), unsafe_allow_html=True)
    for _, row in df.iterrows():
        rec = next((x for x in saved_records if x.get("result", {}).get("file_name") == row.get("file_name")), None)
        current = rec.get("status", "AI screened") if rec else "AI screened"
        c1, c2, c3 = st.columns([4, 2, 1])
        with c1:
            st.markdown(f"**{e(row['name'])}** · {row['match_score']}% AI match")
        with c2:
            new_status = st.selectbox("Status", statuses, index=statuses.index(current) if current in statuses else 0, key=f"status_{row['file_name']}")
        with c3:
            if st.button("Save", key=f"save_status_{row['file_name']}") and rec:
                update_saved_record(rec.get("id"), status=new_status)
                st.success("Saved")

# ---- Interview Kit ----
with tab_interview:
    st.markdown("### Interview Kit")
    st.caption("Generate questions from the candidate's resume, role requirements and detected skill gaps.")
    candidate_options = [f"{r['rank']}. {r['name']}" for _, r in df.iterrows()]
    selected_candidate = st.selectbox("Choose candidate", candidate_options, key="interview_candidate")
    selected_idx = candidate_options.index(selected_candidate)
    selected_row = df.iloc[selected_idx]
    q1, q2, q3 = st.columns([2, 1, 1])
    with q1:
        difficulty = st.selectbox("Difficulty", ["Beginner", "Intermediate", "Advanced"], index=1, key="interview_difficulty")
    with q2:
        technical_count = st.number_input("Technical questions", min_value=3, max_value=8, value=5, step=1, key="interview_count")
    with q3:
        st.metric("Candidate match", f"{int(to_num(selected_row['match_score']))}%")
    st.markdown(f"**{e(selected_row['name'])}** · {e(selected_row['current_role'])}")
    if selected_row.get("matched_skills"):
        st.markdown(f"**Matched skills:** {chips(selected_row['matched_skills'], 'ok')}", unsafe_allow_html=True)
    if selected_row.get("missing_skills"):
        st.markdown(f"**Skill gaps to validate:** {chips(selected_row['missing_skills'], 'no')}", unsafe_allow_html=True)
    if st.button("✦ Generate interview questions", type="primary", use_container_width=True, key="generate_interview"):
        with st.spinner("Building a candidate-specific interview kit..."):
            try:
                generated = generate_questions(job_description, selected_row.to_dict(), difficulty, int(technical_count))
                st.session_state["interview_questions"] = generated
                st.session_state["interview_candidate_name"] = selected_row["name"]
            except Exception as err:
                st.error(str(err))
    generated = st.session_state.get("interview_questions")
    generated_for = st.session_state.get("interview_candidate_name")
    if generated and generated_for:
        st.markdown(f"#### Interview kit for {e(generated_for)}")
        section_titles = {"technical": "Technical questions", "project": "Project deep-dive", "skill_gap": "Skill-gap questions", "hr": "HR / behavioural questions"}
        for key, title in section_titles.items():
            items = generated.get(key, []) if isinstance(generated, dict) else []
            if not items:
                continue
            st.markdown(f"**{title}**")
            for i, item in enumerate(items, 1):
                item = item if isinstance(item, dict) else {"question": str(item)}
                with st.expander(f"{i}. {item.get('question', '')}", expanded=(i == 1)):
                    if item.get("why"): st.markdown(f"**What it tests:** {item['why']}")
                    if item.get("ideal_answer"): st.markdown(f"**Good answer should cover:** {item['ideal_answer']}")
        st.download_button("Download interview kit", data=questions_to_text(generated_for, generated).encode("utf-8"),
                           file_name=f"{re.sub(r'[^A-Za-z0-9_-]+', '_', str(generated_for)).strip('_') or 'candidate'}_interview_kit.txt",
                           mime="text/plain", key="download_interview_kit")

# ---- Resume Library ----
with tab_library:
    st.markdown("### Resume Library")
    st.caption("Analyzed resumes are saved in a local SQLite database and the original files are kept in data/resumes/. They remain available after closing and reopening the app on this computer.")
    library_query = st.text_input("Search saved resumes", placeholder="Name, skill, role or location", key="library_tab_search")
    library_records = load_saved_records()
    if not library_records:
        st.info("No saved resumes yet. Analyze a resume above and it will appear here automatically.")
    else:
        visible = []
        for rec in reversed(library_records):
            rr = rec.get("result", {})
            hay = " ".join(map(str, [rr.get("name", ""), rr.get("current_role", ""), rr.get("location", ""), *rr.get("skills", [])])).lower()
            if library_query.strip() and library_query.lower() not in hay:
                continue
            visible.append((rec, rr))
        st.markdown(f"**{len(visible)} saved profile(s)**")
        for rec, rr in visible:
            with st.container(border=True):
                a, b, c, d = st.columns([4, 1.2, 1.4, 1.1])
                with a:
                    st.markdown(f"**{e(rr.get('name', rec.get('file_name', 'Resume')))}**")
                    st.caption(f"{rec.get('file_name','')} · saved {rec.get('saved_at','')}")
                    st.caption(e(rr.get('current_role','Not found')))
                with b: st.metric("AI match", f"{int(to_num(rr.get('ai_score',0)))}%")
                with c: st.metric("Resume health", f"{int(to_num(rr.get('resume_health',0)))}%")
                with d: st.caption("Status"); st.write(rec.get("status", "AI screened"))
                st.markdown(f"**Skills:** {', '.join(rr.get('skills', [])[:10]) or 'Not found'}")
                if st.button("Delete saved resume", key=f"lib_delete_{rec.get('id')}"):
                    delete_saved_record(rec.get("id")); st.rerun()

        st.markdown("---")
        st.markdown("#### Candidate rediscovery")
        st.caption("Older saved candidates who show useful skill overlap with the current role.")
        if jd_skills:
            current_ids = {r.get("file_name") for r in results}
            rediscovered = []
            for rec in library_records:
                rr = rec.get("result", {})
                if rr.get("file_name") in current_ids: continue
                overlap = saved_candidate_match(rr, jd_skills)
                if overlap > 0: rediscovered.append((overlap, rec, rr))
            rediscovered.sort(key=lambda x: x[0], reverse=True)
            if rediscovered:
                for overlap, rec, rr in rediscovered[:6]:
                    a, b = st.columns([4, 1])
                    with a:
                        st.markdown(f"**{e(rr.get('name', 'Candidate'))}** · {e(rr.get('current_role', ''))}")
                        st.caption(f"Saved {rec.get('saved_at', '')} · useful overlap with this role")
                    with b: st.metric("Overlap", f"{overlap}%")
            else:
                st.caption("No older saved profile has a detectable skill overlap with this role yet.")
        else:
            st.caption("Analyze the current role to enable rediscovery.")

