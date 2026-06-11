import re
import os
import math
from pathlib import Path
from io import BytesIO

import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from mplsoccer import Pitch
import pandas as pd
import numpy as np
from PIL import Image
from matplotlib.lines import Line2D
from matplotlib.patches import FancyArrowPatch, Rectangle, FancyBboxPatch
from matplotlib.colors import Normalize, LinearSegmentedColormap
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.gridspec import GridSpec


# ── PAGE CONFIG ──
st.set_page_config(layout="wide", page_title="Hudson Cicala — Dashboard")


# ── OPTIONAL DOCX IMPORT ──
DOCX_AVAILABLE = True
try:
    from docx import Document
except Exception:
    DOCX_AVAILABLE = False


# ── STYLE ──
st.markdown("""<style>.stApp {background: #0d0d1f; color: #ffffff;} .stSelectbox label, .stRadio label {color: #ffffff !important;} div[data-testid=\"stMarkdownContainer\"] h1, h2, h3, h4 {color: #ffffff;} .stTabs [role=\"tab\"] {color: #d0d0e8;} .stTabs [role=\"tab\"][aria-selected=\"true\"] {color: #ffffff !important; border-bottom-color: #2F80ED !important;} div[data-testid=\"column\"] {background: #1a1a2e; border-radius: 8px; padding: 10px;}</style>""", unsafe_allow_html=True)


# ── CONSTANTS ──
FIELD_X, FIELD_Y = 120.0, 80.0
HALF_LINE_X = FIELD_X / 2
FINAL_THIRD_LINE_X = 80.0
LANE_LEFT_MIN = 53.33
LANE_RIGHT_MAX = 26.67
GOAL_X = 120.0
GOAL_Y = 40.0
FIG_W, FIG_H = 7.0, 4.7
FIG_DPI = 180
COLOR_SUCCESS = "#c8c8c8"
COLOR_PROGRESSIVE = "#2F80ED"
COLOR_FAIL = "#E07070"
ALPHA_SUCCESS = 0.07
C_BLUE = "#2F80ED"
C_BLUE_DARK = "#1a56db"
C_GREEN = "#10b981"
C_AMBER = "#f59e0b"
C_PURPLE_LIGHT = "#a78bfa"
C_BLUE_PASTEL = "#5b9bd5"
C_GREEN_PASTEL = "#70ad47"
C_AMBER_PASTEL = "#d4a843"
CMAP_TOP10 = LinearSegmentedColormap.from_list("top10", ["#fef08a", "#f97316", "#b91c1c"])
NORM_TOP10 = Normalize(vmin=0.05, vmax=0.40)
NX_XT, NY_XT = 16, 12
D_REF, D_SCALE, BONUS_CAP = 10.0, 20.0, 0.60
LATERAL_MIN_DIST = 12.0
PENALTY_AREA_X = 18.0
FUNNEL_X_EXTEND = 33.0
PENALTY_AREA_Y_MIN = 18.0
PENALTY_AREA_Y_MAX = 62.0


# ── PDF STYLE CONSTANTS ──
PDF_BG = "#0d0d1f"
PDF_TEXT_WHITE = "#ffffff"
PDF_TEXT_LIGHT = "#d0d0e8"
PDF_TEXT_DIM = "#5a5a7a"


# ── BASE PASSES DATA ──
BASE_MATCHES_DATA = {
    "Connecticut United (03-27)": [
        ("PASS WON", 26.75, 68.34, 8.97, 51.05, None),
        ("PASS WON", 15.46, 64.12, 20.97, 58.76, None),
        ("PASS WON", 15.46, 64.12, 20.97, 58.76, None),
        ("PASS LOST", 18.23, 55.80, 25.10, 48.90, None),
        ("PASS WON", 22.50, 45.30, 35.80, 42.10, None),
        ("PASS WON", 30.12, 38.45, 42.67, 35.20, None),
        ("PASS WON", 28.90, 50.30, 38.45, 55.60, None),
        ("PASS WON", 12.34, 60.20, 18.90, 52.80, None),
        ("PASS WON", 35.67, 44.50, 48.23, 40.10, None),
        ("PASS WON", 20.00, 70.10, 15.50, 65.30, None),
        ("PASS LOST", 40.12, 32.80, 52.45, 28.40, None),
        ("PASS WON", 8.50, 55.00, 14.20, 48.60, None),
        ("PASS WON", 25.30, 62.40, 32.10, 58.20, None),
        ("PASS WON", 42.00, 36.70, 55.30, 32.50, None),
        ("PASS WON", 16.80, 48.90, 22.40, 42.30, None),
        ("PASS LOST", 32.50, 52.10, 28.00, 58.70, None),
        ("PASS WON", 10.20, 72.30, 5.80, 66.40, None),
        ("PASS WON", 38.90, 40.50, 50.60, 38.20, None),
        ("PASS WON", 22.80, 58.90, 30.40, 52.60, None),
        ("PASS WON", 45.10, 30.20, 58.30, 26.80, None),
        ("PASS WON", 14.50, 66.70, 20.80, 60.10, None),
        ("PASS WON", 28.00, 44.30, 35.60, 48.70, None),
        ("PASS LOST", 36.40, 50.80, 42.00, 56.40, None),
        ("PASS WON", 8.00, 58.40, 12.60, 52.00, None),
        ("PASS WON", 32.20, 46.80, 44.50, 42.90, None),
        ("PASS WON", 18.60, 62.00, 25.10, 56.80, None),
        ("PASS WON", 40.80, 34.60, 52.90, 30.40, None),
        ("PASS WON", 12.00, 70.80, 8.40, 64.20, None),
        ("PASS LOST", 34.70, 42.30, 46.20, 38.50, None),
        ("PASS WON", 26.00, 54.60, 34.80, 50.20, None),
        ("PASS WON", 20.40, 60.50, 28.60, 54.90, None),
        ("PASS WON", 44.00, 38.20, 56.70, 34.60, None),
        ("PASS WON", 16.20, 52.40, 22.90, 46.10, None),
        ("PASS WON", 30.80, 48.20, 40.10, 44.50, None),
        ("PASS LOST", 38.20, 56.40, 30.50, 60.80, None),
        ("PASS WON", 10.80, 62.60, 16.40, 56.00, None),
        ("PASS WON", 24.60, 46.00, 32.20, 50.40, None),
        ("PASS WON", 36.80, 42.80, 48.40, 38.60, None),
        ("PASS WON", 14.00, 68.00, 20.60, 62.40, None),
        ("PASS WON", 42.20, 32.40, 54.60, 28.20, None),
        ("PASS LOST", 28.40, 58.20, 34.00, 62.80, None),
        ("PASS WON", 18.80, 56.80, 26.40, 50.60, None),
        ("PASS WON", 34.20, 44.20, 46.80, 40.80, None),
        ("PASS WON", 22.20, 64.40, 30.80, 58.60, None),
        ("PASS LOST", 40.40, 38.80, 50.20, 34.40, None),
        ("PASS WON", 12.40, 74.20, 6.80, 68.40, None),
        ("PASS WON", 26.80, 50.80, 36.40, 46.20, None),
        ("PASS WON", 44.40, 36.20, 58.00, 32.80, None),
        ("PASS WON", 16.60, 60.80, 24.20, 54.40, None),
        ("PASS WON", 32.60, 42.60, 44.20, 38.40, None),
    ],
    "Nashville (03-29)": [
        ("PASS WON", 32.10, 58.40, 38.50, 52.20, None),
        ("PASS WON", 24.30, 62.70, 30.80, 56.40, None),
        ("PASS WON", 18.50, 66.20, 24.90, 60.10, None),
        ("PASS LOST", 36.70, 44.30, 42.10, 38.60, None),
        ("PASS WON", 20.20, 70.50, 14.80, 64.30, None),
        ("PASS WON", 28.40, 52.60, 36.20, 48.20, None),
        ("PASS WON", 40.10, 38.40, 52.30, 34.80, None),
        ("PASS WON", 14.60, 58.20, 20.40, 52.60, None),
        ("PASS WON", 34.80, 48.20, 44.60, 44.00, None),
        ("PASS WON", 22.40, 64.80, 30.60, 58.40, None),
        ("PASS LOST", 44.20, 34.60, 54.40, 30.20, None),
        ("PASS WON", 12.80, 60.40, 18.20, 54.80, None),
        ("PASS WON", 30.20, 56.20, 38.40, 50.80, None),
        ("PASS WON", 42.40, 40.20, 56.20, 36.40, None),
        ("PASS WON", 16.40, 54.80, 22.60, 48.40, None),
        ("PASS LOST", 38.60, 48.40, 44.20, 54.60, None),
        ("PASS WON", 26.20, 68.40, 34.80, 62.20, None),
        ("PASS WON", 46.10, 36.80, 58.40, 32.60, None),
        ("PASS WON", 20.60, 60.20, 28.20, 54.00, None),
        ("PASS WON", 36.20, 42.40, 48.60, 38.20, None),
        ("PASS WON", 10.40, 64.80, 16.20, 58.20, None),
        ("PASS WON", 32.40, 54.40, 42.20, 50.40, None),
        ("PASS LOST", 40.60, 46.20, 48.20, 52.80, None),
        ("PASS WON", 24.80, 56.40, 32.40, 50.20, None),
        ("PASS WON", 44.60, 38.80, 56.80, 34.20, None),
        ("PASS WON", 18.80, 62.40, 26.20, 56.80, None),
        ("PASS LOST", 34.40, 50.60, 40.20, 44.40, None),
        ("PASS WON", 28.60, 66.20, 36.40, 60.40, None),
        ("PASS WON", 38.20, 44.60, 50.40, 40.20, None),
        ("PASS WON", 14.20, 56.80, 20.80, 50.40, None),
        ("PASS LOST", 42.80, 40.40, 52.60, 36.80, None),
        ("PASS WON", 22.80, 58.60, 30.20, 52.40, None),
        ("PASS WON", 46.40, 36.40, 58.80, 32.40, None),
        ("PASS WON", 16.80, 64.20, 24.40, 58.60, None),
        ("PASS WON", 34.60, 46.80, 44.40, 42.60, None),
        ("PASS LOST", 30.40, 60.40, 36.80, 66.20, None),
        ("PASS WON", 12.60, 72.40, 6.80, 66.80, None),
        ("PASS WON", 40.80, 42.20, 54.20, 38.40, None),
        ("PASS WON", 26.40, 54.80, 34.20, 48.60, None),
        ("PASS WON", 44.80, 34.20, 56.40, 30.60, None),
        ("PASS LOST", 36.60, 52.20, 42.40, 58.40, None),
        ("PASS WON", 20.40, 68.20, 28.80, 62.60, None),
        ("PASS WON", 32.80, 48.40, 42.40, 44.80, None),
        ("PASS WON", 14.40, 60.60, 20.20, 54.20, None),
    ],
    "Seongnam (04-05)": [
        ("PASS WON", 28.30, 54.60, 34.70, 48.40, None),
        ("PASS WON", 22.10, 60.80, 28.50, 54.20, None),
        ("PASS WON", 16.40, 66.40, 22.80, 60.60, None),
        ("PASS LOST", 34.50, 42.80, 40.90, 36.40, None),
        ("PASS WON", 18.20, 62.20, 24.60, 56.80, None),
        ("PASS WON", 30.70, 50.40, 38.20, 46.20, None),
        ("PASS WON", 42.40, 36.60, 54.80, 32.40, None),
        ("PASS WON", 12.20, 58.80, 18.60, 52.40, None),
        ("PASS WON", 36.60, 46.20, 46.40, 42.00, None),
        ("PASS WON", 24.80, 64.20, 32.40, 58.80, None),
        ("PASS LOST", 44.60, 34.20, 54.80, 30.40, None),
        ("PASS WON", 14.20, 56.40, 20.80, 50.20, None),
        ("PASS WON", 32.40, 52.60, 40.20, 48.40, None),
        ("PASS WON", 40.80, 38.40, 54.20, 34.60, None),
        ("PASS WON", 20.20, 58.60, 26.80, 52.40, None),
        ("PASS LOST", 38.20, 50.20, 44.60, 56.40, None),
        ("PASS WON", 26.40, 66.80, 34.60, 60.40, None),
        ("PASS WON", 44.20, 34.80, 56.80, 30.60, None),
        ("PASS WON", 18.60, 60.40, 26.20, 54.60, None),
        ("PASS WON", 34.80, 44.20, 46.40, 40.20, None),
        ("PASS WON", 10.80, 62.80, 16.40, 56.40, None),
        ("PASS WON", 30.20, 56.60, 38.80, 52.20, None),
        ("PASS LOST", 42.40, 48.20, 50.20, 44.40, None),
        ("PASS WON", 22.40, 54.80, 30.20, 48.60, None),
        ("PASS WON", 46.20, 36.20, 58.80, 32.80, None),
        ("PASS WON", 16.60, 64.40, 24.20, 58.20, None),
        ("PASS LOST", 36.40, 52.80, 42.60, 58.40, None),
        ("PASS WON", 28.80, 60.20, 36.40, 54.60, None),
        ("PASS WON", 40.40, 40.60, 52.80, 36.40, None),
        ("PASS WON", 12.40, 60.80, 18.80, 54.40, None),
    ],
    "NY Red Bulls (03-31)": [
        ("PASS WON", 30.50, 56.80, 36.90, 50.40, None),
        ("PASS WON", 24.30, 62.40, 30.70, 56.20, None),
        ("PASS WON", 18.60, 68.20, 25.00, 62.40, None),
        ("PASS LOST", 36.80, 44.60, 42.40, 38.20, None),
        ("PASS WON", 20.40, 64.80, 26.80, 58.40, None),
        ("PASS WON", 32.60, 52.40, 40.20, 48.20, None),
        ("PASS WON", 44.20, 38.20, 56.80, 34.40, None),
        ("PASS WON", 14.80, 60.20, 20.40, 54.60, None),
        ("PASS WON", 38.40, 48.60, 48.20, 44.40, None),
        ("PASS WON", 26.60, 66.20, 34.80, 60.80, None),
        ("PASS LOST", 46.40, 36.40, 56.60, 32.20, None),
        ("PASS WON", 12.60, 58.80, 18.20, 52.40, None),
        ("PASS WON", 34.20, 54.20, 42.40, 50.20, None),
        ("PASS WON", 42.60, 40.60, 56.40, 36.20, None),
        ("PASS WON", 22.40, 60.40, 28.80, 54.20, None),
        ("PASS LOST", 40.40, 50.80, 46.20, 56.60, None),
        ("PASS WON", 28.20, 68.40, 36.60, 62.80, None),
        ("PASS WON", 48.20, 36.80, 60.40, 32.40, None),
        ("PASS WON", 16.80, 62.80, 24.40, 56.60, None),
        ("PASS WON", 36.80, 44.80, 48.60, 40.60, None),
        ("PASS WON", 10.60, 66.40, 16.20, 60.20, None),
        ("PASS WON", 30.80, 56.20, 40.40, 52.40, None),
        ("PASS LOST", 44.80, 46.60, 52.40, 42.60, None),
        ("PASS WON", 24.60, 58.40, 32.40, 52.20, None),
        ("PASS WON", 46.80, 38.40, 58.80, 34.60, None),
        ("PASS WON", 18.40, 66.40, 26.20, 60.20, None),
        ("PASS LOST", 38.60, 52.40, 44.80, 58.80, None),
        ("PASS WON", 32.40, 62.40, 40.20, 56.60, None),
        ("PASS WON", 42.40, 42.40, 54.80, 38.80, None),
        ("PASS WON", 14.40, 62.60, 20.80, 56.20, None),
        ("PASS LOST", 46.20, 42.40, 54.60, 38.20, None),
        ("PASS WON", 26.80, 56.80, 34.60, 50.40, None),
        ("PASS WON", 48.60, 36.20, 60.80, 32.20, None),
        ("PASS WON", 20.80, 60.60, 28.40, 54.40, None),
        ("PASS WON", 34.60, 48.40, 44.80, 44.20, None),
        ("PASS LOST", 32.60, 64.40, 38.80, 68.20, None),
        ("PASS WON", 12.80, 70.80, 6.60, 64.40, None),
        ("PASS WON", 44.60, 40.20, 58.40, 36.80, None),
        ("PASS WON", 28.60, 52.60, 36.20, 48.40, None),
        ("PASS WON", 40.60, 36.80, 54.20, 32.40, None),
        ("PASS LOST", 36.40, 54.80, 42.20, 60.60, None),
        ("PASS WON", 22.60, 66.40, 30.80, 60.80, None),
        ("PASS WON", 34.40, 50.20, 44.60, 46.40, None),
        ("PASS WON", 16.20, 62.20, 22.80, 56.60, None),
    ],
    "Vardar (04-13)": [
        ("PASS WON", 32.80, 56.20, 38.40, 50.60, None),
        ("PASS WON", 24.60, 64.80, 30.20, 58.40, None),
        ("PASS WON", 18.40, 68.40, 24.80, 62.60, None),
        ("PASS LOST", 36.40, 46.80, 42.60, 40.20, None),
        ("PASS WON", 20.20, 66.40, 26.80, 60.20, None),
        ("PASS WON", 34.20, 54.40, 42.40, 50.20, None),
        ("PASS WON", 46.40, 38.60, 58.60, 34.80, None),
        ("PASS WON", 14.20, 60.80, 20.60, 54.40, None),
        ("PASS WON", 38.80, 48.20, 48.40, 44.60, None),
        ("PASS WON", 28.40, 64.60, 36.80, 58.80, None),
        ("PASS LOST", 44.40, 36.20, 56.60, 32.40, None),
        ("PASS WON", 12.40, 58.60, 18.80, 52.20, None),
        ("PASS WON", 32.60, 56.80, 40.40, 52.40, None),
        ("PASS WON", 42.80, 42.20, 56.40, 38.60, None),
        ("PASS WON", 22.60, 60.20, 28.80, 54.40, None),
        ("PASS LOST", 38.40, 52.60, 44.80, 58.40, None),
        ("PASS WON", 26.20, 68.80, 34.40, 62.40, None),
        ("PASS WON", 48.40, 38.20, 60.80, 34.40, None),
        ("PASS WON", 18.60, 62.40, 26.20, 56.80, None),
        ("PASS WON", 36.60, 46.40, 48.80, 42.40, None),
        ("PASS WON", 10.40, 64.40, 16.80, 58.20, None),
        ("PASS WON", 30.40, 58.40, 38.60, 54.20, None),
        ("PASS LOST", 42.40, 48.80, 50.40, 44.80, None),
        ("PASS WON", 24.80, 56.40, 32.60, 50.80, None),
        ("PASS WON", 46.60, 36.40, 58.80, 32.80, None),
        ("PASS WON", 16.60, 66.40, 24.40, 60.20, None),
        ("PASS LOST", 36.20, 54.80, 42.80, 60.40, None),
        ("PASS WON", 30.80, 62.40, 38.40, 56.80, None),
        ("PASS WON", 44.60, 40.80, 56.80, 36.40, None),
        ("PASS WON", 14.60, 62.80, 20.80, 56.40, None),
        ("PASS LOST", 46.80, 44.20, 54.40, 40.60, None),
        ("PASS WON", 28.60, 54.80, 36.40, 48.80, None),
    ],
    "Real Salt Lake (04-26)": [
        ("PASS WON", 34.20, 52.80, 40.60, 46.40, None),
        ("PASS WON", 26.80, 60.40, 32.40, 54.20, None),
        ("PASS WON", 20.40, 66.80, 26.80, 60.40, None),
        ("PASS LOST", 38.60, 44.20, 44.80, 38.60, None),
        ("PASS WON", 22.40, 64.20, 28.60, 58.40, None),
        ("PASS WON", 36.20, 52.80, 44.40, 48.60, None),
        ("PASS WON", 48.60, 36.80, 60.40, 32.40, None),
        ("PASS WON", 16.80, 58.40, 22.40, 52.60, None),
        ("PASS WON", 40.20, 46.80, 50.40, 42.20, None),
        ("PASS WON", 30.40, 62.60, 38.80, 56.80, None),
        ("PASS LOST", 46.60, 34.40, 58.60, 30.60, None),
        ("PASS WON", 14.60, 56.80, 20.60, 50.40, None),
        ("PASS WON", 34.80, 54.60, 42.40, 50.20, None),
        ("PASS WON", 44.80, 40.20, 58.20, 36.80, None),
        ("PASS WON", 24.40, 58.60, 30.80, 52.40, None),
        ("PASS LOST", 40.60, 50.20, 46.80, 56.60, None),
        ("PASS WON", 28.20, 66.80, 36.40, 60.40, None),
        ("PASS WON", 50.40, 36.60, 62.80, 32.20, None),
        ("PASS WON", 18.60, 60.80, 26.20, 54.40, None),
        ("PASS WON", 38.40, 44.40, 50.40, 40.80, None),
        ("PASS WON", 12.60, 64.20, 18.80, 58.60, None),
        ("PASS WON", 32.60, 56.60, 42.20, 52.40, None),
        ("PASS LOST", 44.80, 46.20, 52.40, 42.80, None),
        ("PASS WON", 26.40, 54.80, 34.20, 48.80, None),
        ("PASS WON", 48.80, 36.20, 60.80, 32.80, None),
        ("PASS WON", 18.80, 64.20, 26.40, 58.60, None),
        ("PASS LOST", 38.20, 52.40, 44.60, 58.80, None),
        ("PASS WON", 32.80, 60.40, 40.40, 54.60, None),
        ("PASS WON", 46.40, 42.60, 58.80, 38.40, None),
        ("PASS WON", 16.40, 60.80, 22.80, 54.40, None),
        ("PASS LOST", 48.60, 42.40, 56.80, 38.40, None),
        ("PASS WON", 30.20, 52.80, 38.40, 46.80, None),
        ("PASS WON", 52.40, 34.60, 64.80, 30.40, None),
        ("PASS WON", 22.60, 62.80, 30.20, 56.60, None),
        ("PASS WON", 36.60, 46.80, 46.80, 42.60, None),
        ("PASS LOST", 34.60, 62.60, 40.80, 68.40, None),
        ("PASS WON", 14.80, 72.60, 8.40, 66.40, None),
    ],
    "Real Futbol (05-23)": [
        ("PASS WON", 30.40, 54.60, 36.80, 48.40, None),
        ("PASS WON", 22.60, 62.80, 28.40, 56.60, None),
        ("PASS WON", 18.20, 66.40, 24.80, 60.40, None),
        ("PASS LOST", 34.60, 46.40, 40.80, 40.20, None),
        ("PASS WON", 20.80, 64.20, 26.40, 58.80, None),
        ("PASS WON", 32.40, 54.80, 40.20, 50.40, None),
        ("PASS WON", 44.80, 38.20, 56.60, 34.40, None),
        ("PASS WON", 16.20, 60.40, 22.60, 54.60, None),
        ("PASS WON", 36.80, 48.20, 46.40, 44.60, None),
        ("PASS WON", 26.80, 64.60, 34.40, 58.40, None),
        ("PASS LOST", 42.40, 36.80, 54.60, 32.40, None),
        ("PASS WON", 14.20, 58.40, 20.80, 52.20, None),
        ("PASS WON", 34.40, 56.80, 42.60, 52.20, None),
        ("PASS WON", 46.60, 42.60, 58.80, 38.80, None),
        ("PASS WON", 24.40, 60.20, 30.60, 54.60, None),
        ("PASS LOST", 40.80, 50.40, 46.60, 56.80, None),
        ("PASS WON", 28.40, 66.60, 36.80, 60.20, None),
        ("PASS WON", 48.40, 38.60, 60.80, 34.80, None),
        ("PASS WON", 20.60, 62.60, 28.40, 56.80, None),
        ("PASS WON", 38.60, 46.20, 48.80, 42.40, None),
        ("PASS WON", 12.80, 66.20, 18.40, 60.40, None),
        ("PASS WON", 32.80, 58.20, 42.40, 54.40, None),
        ("PASS LOST", 46.80, 48.40, 54.40, 44.60, None),
        ("PASS WON", 26.20, 56.80, 34.40, 50.40, None),
        ("PASS WON", 48.60, 38.20, 60.60, 34.80, None),
        ("PASS WON", 18.60, 66.20, 26.20, 60.40, None),
        ("PASS LOST", 38.40, 54.60, 44.80, 60.40, None),
        ("PASS WON", 32.60, 62.20, 40.20, 56.60, None),
        ("PASS WON", 46.20, 40.80, 58.40, 36.60, None),
        ("PASS WON", 16.60, 64.60, 22.80, 58.40, None),
    ],
    "San Jose (05-24)": [
        ("PASS WON", 34.80, 52.40, 40.60, 46.80, None),
        ("PASS WON", 26.20, 60.80, 32.60, 54.60, None),
        ("PASS WON", 20.60, 66.40, 26.80, 60.60, None),
        ("PASS LOST", 38.40, 44.80, 44.40, 38.40, None),
        ("PASS WON", 22.60, 64.80, 28.40, 58.60, None),
        ("PASS WON", 36.40, 52.80, 44.60, 48.80, None),
        ("PASS WON", 48.80, 36.40, 60.40, 32.80, None),
        ("PASS WON", 16.40, 58.80, 22.80, 52.40, None),
        ("PASS WON", 40.40, 46.80, 50.60, 42.40, None),
        ("PASS WON", 30.80, 62.40, 38.60, 56.80, None),
        ("PASS LOST", 46.80, 34.80, 58.40, 30.60, None),
        ("PASS WON", 14.80, 56.40, 20.40, 50.80, None),
        ("PASS WON", 36.60, 54.40, 44.20, 50.40, None),
        ("PASS WON", 46.40, 42.20, 58.80, 38.40, None),
        ("PASS WON", 24.60, 58.20, 30.80, 52.40, None),
        ("PASS LOST", 42.40, 50.60, 48.80, 56.80, None),
        ("PASS WON", 28.60, 66.20, 36.40, 60.80, None),
        ("PASS WON", 50.80, 36.40, 62.40, 32.60, None),
        ("PASS WON", 18.80, 60.40, 26.40, 54.60, None),
        ("PASS WON", 38.60, 44.80, 50.60, 40.60, None),
        ("PASS WON", 12.80, 64.80, 18.60, 58.20, None),
        ("PASS WON", 34.60, 56.40, 44.40, 52.80, None),
        ("PASS LOST", 46.80, 46.80, 54.60, 42.40, None),
        ("PASS WON", 26.80, 54.60, 34.20, 48.80, None),
        ("PASS WON", 48.40, 36.60, 60.80, 32.40, None),
        ("PASS WON", 18.40, 64.60, 26.80, 58.60, None),
        ("PASS LOST", 38.80, 52.80, 44.80, 58.80, None),
        ("PASS WON", 34.20, 60.80, 42.40, 54.80, None),
        ("PASS WON", 48.60, 42.40, 60.60, 38.20, None),
        ("PASS WON", 16.80, 62.80, 22.60, 56.40, None),
    ],
    "Houston Dynamo (05-26)": [
        ("PASS WON", 32.60, 56.40, 38.80, 50.20, None),
        ("PASS WON", 24.80, 64.40, 30.80, 58.20, None),
        ("PASS WON", 18.60, 68.80, 24.80, 62.40, None),
        ("PASS LOST", 36.80, 46.40, 42.40, 40.60, None),
        ("PASS WON", 20.80, 66.80, 26.60, 60.40, None),
        ("PASS WON", 34.40, 54.60, 42.40, 50.80, None),
        ("PASS WON", 46.80, 38.40, 58.60, 34.80, None),
        ("PASS WON", 14.60, 60.80, 20.80, 54.40, None),
        ("PASS WON", 40.80, 48.40, 50.60, 44.20, None),
        ("PASS WON", 28.80, 64.40, 36.80, 58.80, None),
        ("PASS LOST", 44.80, 36.40, 56.60, 32.80, None),
        ("PASS WON", 12.80, 58.80, 18.60, 52.40, None),
        ("PASS WON", 34.80, 56.40, 42.60, 52.40, None),
        ("PASS WON", 44.80, 42.80, 58.40, 38.60, None),
        ("PASS WON", 22.60, 60.80, 28.80, 54.80, None),
        ("PASS LOST", 40.60, 52.40, 46.80, 58.60, None),
        ("PASS WON", 28.40, 68.80, 36.60, 62.80, None),
        ("PASS WON", 48.60, 38.80, 60.80, 34.60, None),
        ("PASS WON", 18.80, 62.60, 26.40, 56.80, None),
        ("PASS WON", 38.60, 46.80, 50.80, 42.60, None),
        ("PASS WON", 10.80, 66.60, 16.60, 60.40, None),
        ("PASS WON", 32.80, 58.60, 42.60, 54.40, None),
        ("PASS LOST", 46.60, 48.60, 54.80, 44.80, None),
        ("PASS WON", 26.80, 56.80, 34.40, 50.40, None),
        ("PASS WON", 48.80, 38.60, 60.80, 34.40, None),
        ("PASS WON", 18.40, 66.60, 26.60, 60.80, None),
        ("PASS LOST", 38.60, 54.40, 44.80, 60.60, None),
        ("PASS WON", 34.60, 62.40, 42.20, 56.80, None),
        ("PASS WON", 48.60, 42.60, 60.60, 38.80, None),
        ("PASS WON", 16.80, 64.80, 22.80, 58.60, None),
    ],
}


# ── DEFENSIVE ACTIONS DATA ──
DEFENSIVE_MATCHES_DATA = {
    "Michigan Wolves (02-20)": [
        ("DUEL_WON", 53.85, 25.21),
        ("INTERCEPTION", 68.40, 48.20),
        ("DUEL_LOST", 42.30, 58.60),
        ("DUEL_WON", 55.20, 30.80),
        ("DUEL_WON", 70.60, 42.40),
        ("INTERCEPTION", 62.80, 36.50),
        ("DUEL_WON", 48.40, 52.20),
        ("DUEL_LOST", 58.60, 28.40),
        ("DUEL_WON", 72.40, 44.60),
        ("INTERCEPTION", 66.20, 38.80),
        ("DUEL_WON", 50.60, 54.80),
        ("DUEL_WON", 60.40, 32.40),
    ],
    "Connecticut United (03-27)": [
        ("DUEL_WON", 56.40, 28.60),
        ("INTERCEPTION", 64.80, 46.20),
        ("DUEL_WON", 52.20, 54.40),
        ("DUEL_LOST", 44.60, 60.80),
        ("DUEL_WON", 68.40, 40.20),
        ("INTERCEPTION", 60.40, 34.60),
        ("DUEL_WON", 46.80, 56.40),
        ("DUEL_WON", 70.80, 42.80),
    ],
    "Nashville (03-29)": [
        ("DUEL_WON", 54.80, 30.40),
        ("INTERCEPTION", 66.40, 44.80),
        ("DUEL_WON", 50.40, 58.20),
        ("DUEL_LOST", 46.20, 62.40),
        ("DUEL_WON", 72.40, 38.60),
        ("INTERCEPTION", 62.80, 36.20),
        ("DUEL_WON", 48.60, 52.80),
        ("DUEL_WON", 68.80, 46.40),
    ],
    "Seongnam (04-05)": [
        ("DUEL_WON", 58.40, 32.60),
        ("INTERCEPTION", 70.20, 48.40),
        ("DUEL_WON", 54.60, 56.80),
        ("DUEL_LOST", 48.80, 64.20),
        ("DUEL_WON", 74.60, 42.20),
        ("INTERCEPTION", 66.40, 38.40),
        ("DUEL_WON", 52.80, 54.60),
        ("DUEL_WON", 72.60, 44.80),
    ],
    "NY Red Bulls (03-31)": [
        ("DUEL_WON", 56.80, 26.40),
        ("INTERCEPTION", 68.80, 42.60),
        ("DUEL_WON", 52.40, 58.40),
        ("DUEL_LOST", 46.60, 66.20),
        ("DUEL_WON", 74.20, 40.40),
        ("INTERCEPTION", 64.40, 36.80),
        ("DUEL_WON", 50.80, 56.40),
        ("DUEL_WON", 70.20, 48.60),
    ],
    "Vardar (04-13)": [
        ("DUEL_WON", 60.20, 34.40),
        ("INTERCEPTION", 72.40, 50.20),
        ("DUEL_WON", 56.60, 60.80),
        ("DUEL_LOST", 50.40, 68.40),
        ("DUEL_WON", 76.80, 44.40),
        ("INTERCEPTION", 68.60, 40.20),
        ("DUEL_WON", 54.80, 58.60),
        ("DUEL_WON", 74.80, 46.80),
    ],
    "Real Salt Lake (04-26)": [
        ("DUEL_WON", 58.80, 30.80),
        ("INTERCEPTION", 66.80, 46.80),
        ("DUEL_WON", 54.40, 58.80),
        ("DUEL_LOST", 48.60, 64.60),
        ("DUEL_WON", 72.60, 42.60),
        ("INTERCEPTION", 64.80, 38.60),
        ("DUEL_WON", 52.40, 56.40),
        ("DUEL_WON", 70.40, 48.20),
    ],
    "Real Futbol (05-23)": [
        ("DUEL_WON", 56.40, 28.80),
        ("INTERCEPTION", 70.40, 44.40),
        ("DUEL_WON", 54.80, 60.20),
        ("DUEL_LOST", 48.40, 66.80),
        ("DUEL_WON", 74.40, 40.80),
        ("INTERCEPTION", 66.60, 38.40),
        ("DUEL_WON", 52.80, 56.80),
        ("DUEL_WON", 72.80, 48.40),
    ],
    "San Jose (05-24)": [
        ("DUEL_WON", 60.60, 32.80),
        ("INTERCEPTION", 68.40, 48.60),
        ("DUEL_WON", 56.80, 58.60),
        ("DUEL_LOST", 50.60, 66.40),
        ("DUEL_WON", 76.40, 42.80),
        ("INTERCEPTION", 68.80, 40.40),
        ("DUEL_WON", 54.40, 58.40),
        ("DUEL_WON", 74.60, 46.40),
    ],
    "Houston Dynamo (05-26)": [
        ("DUEL_WON", 58.40, 30.60),
        ("INTERCEPTION", 70.60, 46.40),
        ("DUEL_WON", 56.40, 60.40),
        ("DUEL_LOST", 50.20, 68.60),
        ("DUEL_WON", 74.60, 42.40),
        ("INTERCEPTION", 66.80, 38.80),
        ("DUEL_WON", 54.60, 58.80),
        ("DUEL_WON", 72.40, 48.80),
    ],
}

mapping = {
    "Connecticut United": "Connecticut United (03-27)",
    "Nashville": "Nashville (03-29)",
    "Seongnam": "Seongnam (04-05)",
    "NY Red Bulls": "NY Red Bulls (03-31)",
    "Vardar": "Vardar (04-13)",
    "Real Salt Lake": "Real Salt Lake (04-26)",
    "Real Futbol": "Real Futbol (05-23)",
    "San Jose": "San Jose (05-24)",
    "Houston Dynamo": "Houston Dynamo (05-26)",
    "Houston": "Houston Dynamo (05-26)",
    "Michigan Wolves": "Michigan Wolves (02-20)",
}


# ── HELPERS ──
def _hex_to_rgba(hex_color, alpha=1.0):
    if hex_color.startswith('#'):
        h = hex_color.lstrip('#')
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return f'rgba({r},{g},{b},{alpha})'
    return hex_color


def get_lane(y):
    if y >= LANE_LEFT_MIN:
        return "left"
    elif y <= LANE_RIGHT_MAX:
        return "right"
    else:
        return "center"


def is_in_funnel_zone(x, y):
    return x <= FUNNEL_X_EXTEND and PENALTY_AREA_Y_MIN <= y <= PENALTY_AREA_Y_MAX


def apply_date_mapping(name):
    for k, v in mapping.items():
        if k.lower() == name.lower().strip():
            return v
    return name


def get_match_minutes(match_name):
    if match_name == "All Matches":
        total = 0.0
        for k in dfs_by_match:
            total += get_match_minutes(k)
        return total
    name_lower = match_name.lower()
    if "connecticut" in name_lower:
        return 60.0
    if "nashville" in name_lower:
        return 60.0
    if "seongnam" in name_lower:
        return 32.0
    if "red bulls" in name_lower:
        return 60.0
    if "houston" in name_lower:
        return 63.0
    if "vardar" in name_lower:
        return 65.0
    return 90.0


def read_docx_text(docx_path):
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx is not installed.")
    doc = Document(str(docx_path))
    return "\n".join(p.text for p in doc.paragraphs if p.text and p.text.strip())


def parse_docx_events(raw_text):
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    matches = {}
    current_match = None
    current_state = None
    re_match = re.compile(r"^Vs\s+(.+)$", re.IGNORECASE)
    re_success = re.compile(r"^Sucesso$", re.IGNORECASE)
    re_fail = re.compile(r"^Errado[s]?$", re.IGNORECASE)
    re_arrow = re.compile(
        r"^Seta\s+\d+:\s(([-+]?\d+\.?\d*),\s*([-+]?\d+\.?\d*))\s->\s(([-+]?\d+\.?\d*),\s*([-+]?\d+\.?\d*))$",
        re.IGNORECASE,
    )
    for ln in lines:
        m_match = re_match.match(ln)
        if m_match:
            current_match = m_match.group(1).strip()
            matches.setdefault(current_match, [])
            current_state = None
            continue
        if re_success.match(ln):
            current_state = "PASS WON"
            continue
        if re_fail.match(ln):
            current_state = "PASS LOST"
            continue
        m_arrow = re_arrow.match(ln)
        if m_arrow and current_match and current_state:
            x1, y1, x2, y2 = map(float, m_arrow.groups()[2:6])
            matches[current_match].append(("PASS WON" if current_state == "PASS WON" else "PASS LOST", x1, y1, x2, y2, None))
    return {k: v for k, v in matches.items() if len(v) > 0}


def load_docx_matches(docx_filename="Passes - Hudson Cicala.docx"):
    p = Path(docx_filename)
    if not p.exists():
        return {}
    txt = read_docx_text(p)
    return parse_docx_events(txt)


# ── DATA LOADING ──
docx_matches_data = {}
try:
    docx_matches_data = load_docx_matches()
except Exception:
    pass

combined_matches_data = {}
for k, v in docx_matches_data.items():
    mapped_k = apply_date_mapping(k)
    name = mapped_k if mapped_k not in combined_matches_data else f"DOCX - {mapped_k}"
    combined_matches_data[name] = v
for k, v in BASE_MATCHES_DATA.items():
    combined_matches_data[k] = v

if len(combined_matches_data) == 0:
    st.error("Could not load data.")
    st.stop()


# ── BUILD DATAFRAMES & REORDER MATCHES ──
dfs_by_match = {}
for match_name, events in combined_matches_data.items():
    dfm = pd.DataFrame(events, columns=["type", "x_start", "y_start", "x_end", "y_end", "video"])
    dfm["match"] = match_name
    dfm["number"] = np.arange(1, len(dfm) + 1)
    dfm["is_won"] = dfm["type"].str.contains("WON", case=False)
    dfm["progressive"] = (
        (dfm["x_start"] < HALF_LINE_X) &
        ((dfm["x_end"] - dfm["x_start"]) / (FIELD_X - dfm["x_start"] + 1e-6) >= 0.25)
    )
    dfm["delta_x"] = dfm["x_end"] - dfm["x_start"]
    dfm["delta_y"] = dfm["y_end"] - dfm["y_start"]
    dfm["dist"] = np.sqrt(dfm["delta_x"]**2 + dfm["delta_y"]**2)
    dfm["end_lane"] = dfm["y_end"].apply(get_lane)
    x_grid = np.linspace(0, FIELD_X, NX_XT + 1)
    y_grid = np.linspace(0, FIELD_Y, NY_XT + 1)
    dfm["xt_bin_x"] = np.digitize(dfm["x_start"], x_grid) - 1
    dfm["xt_bin_y"] = np.digitize(dfm["y_start"], y_grid) - 1
    dfm["xt_bin_x_end"] = np.digitize(dfm["x_end"], x_grid) - 1
    dfm["xt_bin_y_end"] = np.digitize(dfm["y_end"], y_grid) - 1
    xt_grid = np.zeros((NX_XT, NY_XT))
    xt_grid[2, 4] = 0.02
    xt_grid[3, 5] = 0.04
    xt_grid[4, 6] = 0.08
    xt_grid[5, 5] = 0.12
    xt_grid[6, 4] = 0.15
    xt_grid[7, 5] = 0.20
    xt_grid[8, 6] = 0.25
    xt_grid[9, 5] = 0.30
    xt_grid[10, 4] = 0.35
    xt_grid[11, 6] = 0.40
    xt_grid[12, 5] = 0.45
    xt_grid[13, 4] = 0.50
    xt_grid[14, 5] = 0.55
    xt_grid[15, 6] = 0.60
    dfm["xt_start"] = dfm.apply(lambda r: xt_grid[min(r["xt_bin_x"], NX_XT - 1), min(r["xt_bin_y"], NY_XT - 1)], axis=1)
    dfm["xt_end"] = dfm.apply(lambda r: xt_grid[min(r["xt_bin_x_end"], NX_XT - 1), min(r["xt_bin_y_end"], NY_XT - 1)], axis=1)
    dfm["delta_xt"] = dfm["xt_end"] - dfm["xt_start"]
    max_xt_in_grid = np.max(xt_grid)
    penalty = np.clip((dfm["x_end"] - PENALTY_AREA_X) / (FIELD_X - PENALTY_AREA_X + 1e-6), 0, 1) * 0.3
    dfm["delta_xt_adj"] = np.clip(dfm["delta_xt"] + penalty, -BONUS_CAP, BONUS_CAP)
    dfs_by_match[match_name] = dfm

items = list(dfs_by_match.items())
if len(items) >= 18:
    part1 = items[:6]
    part2 = items[14:18]
    part3 = items[6:14]
    part4 = items[18:]
    dfs_by_match = dict(part1 + part2 + part3 + part4)

df_all = pd.concat(dfs_by_match.values(), ignore_index=True)


# ── DEFENSIVE DATA LOADING ──
defensive_dfs_by_match = {}
for match_name, events in DEFENSIVE_MATCHES_DATA.items():
    df_def = pd.DataFrame(events, columns=["type", "x", "y"])
    df_def["match"] = match_name
    df_def["is_attacking_half"] = df_def["x"] >= FIELD_X / 2
    df_def["is_duel_won"] = df_def["type"] == "DUEL_WON"
    df_def["is_duel_lost"] = df_def["type"] == "DUEL_LOST"
    df_def["is_duel"] = df_def["is_duel_won"] | df_def["is_duel_lost"]
    df_def["is_interception"] = df_def["type"] == "INTERCEPTION"
    df_def["in_funnel"] = df_def.apply(lambda r: is_in_funnel_zone(r["x"], r["y"]), axis=1)
    defensive_dfs_by_match[match_name] = df_def


# ── STATS ──
def compute_stats(df, match_name):
    total = len(df)
    mins = get_match_minutes(match_name)
    p90_factor = 90.0 / mins if mins > 0 else 1.0
    progressive = df["progressive"].sum() if "progressive" in df.columns else 0
    won = df["is_won"].sum() if "is_won" in df.columns else 0
    total_impact = float(df.loc[df["is_won"], "delta_xt_adj"].sum()) if "delta_xt_adj" in df.columns and won > 0 else 0.0
    return {
        "total": total,
        "total_p90": total * p90_factor,
        "progressive": progressive,
        "progressive_p90": progressive * p90_factor,
        "accuracy": (won / total * 100) if total > 0 else 0,
        "won": won,
        "total_impact": total_impact,
        "total_impact_p90": total_impact * p90_factor,
        "mins": mins,
    }


def compute_defensive_stats(df, match_name):
    total = len(df)
    mins = get_match_minutes(match_name)
    p90_factor = 90.0 / mins if mins > 0 else 1.0
    duels = df["is_duel"].sum() if "is_duel" in df.columns else 0
    duel_won = df["is_duel_won"].sum() if "is_duel_won" in df.columns else 0
    interceptions = df["is_interception"].sum() if "is_interception" in df.columns else 0
    funnel_actions = df["in_funnel"].sum() if "in_funnel" in df.columns else 0
    return {
        "total": total,
        "total_p90": total * p90_factor,
        "duels": duels,
        "duel_won": duel_won,
        "duel_accuracy": (duel_won / duels * 100) if duels > 0 else 0,
        "interceptions": interceptions,
        "interceptions_p90": interceptions * p90_factor,
        "funnel_protections": funnel_actions,
        "funnel_protections_p90": funnel_actions * p90_factor,
        "mins": mins,
    }


# ── UI HELPERS ──
def section_card_html(title, metrics_lines):
    lines_html = "".join(f"<p style='margin:2px 0;font-size:13px;color:#d0d0e8;'>{line}</p>" for line in metrics_lines)
    return f"""
    <div style="background:#1a1a2e;border-radius:8px;padding:12px;border:1px solid #2a2a4e;">
        <h4 style="color:#ffffff;margin:0 0 8px 0;font-size:14px;">{title}</h4>
        {lines_html}
    </div>
    """


def save_fig(fig):
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=FIG_DPI, bbox_inches="tight", facecolor=fig.get_facecolor())
    buf.seek(0)
    return Image.open(buf)


def attack_arrow(fig):
    fig.text(0.5, 0.02, "→ Attack →", ha="center", va="center",
             fontsize=8, color="#5a5a7a", fontstyle="italic")


# ── PDF EXPORT FUNCTION ──
def export_dashboard_pdf(passes_images, def_images):
    total_pages = 2
    buf = BytesIO()
    with PdfPages(buf) as pdf:
        fig = plt.figure(figsize=(11, 8.5), facecolor=PDF_BG)
        fig.suptitle("Passes Analysis", fontsize=20, fontweight=700, color=PDF_TEXT_WHITE, y=0.97, x=0.06, ha="left")
        labels_p = ["Pass Map", "Zone Heatmap (Destination)", "Top 10 Pass Impact"]
        for i, img in enumerate(passes_images):
            left = 0.03 + i * 0.33
            width = 0.31
            ax_img = fig.add_axes([left, 0.12, width, 0.78])
            ax_img.imshow(img)
            ax_img.axis("off")
            ax_img.text(0.5, 1.01, labels_p[i], ha="center", va="bottom",
                        fontsize=9, fontweight=600, color=PDF_TEXT_LIGHT, transform=ax_img.transAxes)
        pdf.savefig(fig, facecolor=fig.get_facecolor())
        plt.close(fig)

        fig = plt.figure(figsize=(11, 8.5), facecolor=PDF_BG)
        fig.suptitle("Defensive Analysis", fontsize=20, fontweight=700, color=PDF_TEXT_WHITE, y=0.97, x=0.06, ha="left")
        labels_d = ["Defensive Actions Map", "Defensive Heatmap", "Funnel Protection"]
        for i, img in enumerate(def_images):
            left = 0.03 + i * 0.33
            width = 0.31
            ax_img = fig.add_axes([left, 0.12, width, 0.78])
            ax_img.imshow(img)
            ax_img.axis("off")
            ax_img.text(0.5, 1.01, labels_d[i], ha="center", va="bottom",
                        fontsize=9, fontweight=600, color=PDF_TEXT_LIGHT, transform=ax_img.transAxes)
        pdf.savefig(fig, facecolor=fig.get_facecolor())
        plt.close(fig)

    buf.seek(0)
    return buf.read()


# ── DRAW HELPERS (PITCH) ──
def draw_pass_map(df):
    pitch = Pitch(pitch_type="custom", pitch_length=FIELD_X, pitch_width=FIELD_Y,
                  line_color="#444466", pitch_color="#1a1a2e", goal_type="box")
    fig, ax = pitch.draw(figsize=(FIG_W, FIG_H), dpi=FIG_DPI)
    fig.patch.set_facecolor("#1a1a2e")
    ax.set_facecolor("#1a1a2e")

    for _, row in df.iterrows():
        is_prog = row.get("progressive", False)
        is_success = row.get("is_won", True)
        if is_success and not is_prog:
            color, alpha = COLOR_SUCCESS, ALPHA_SUCCESS
        elif not is_success:
            color, alpha = COLOR_FAIL, 0.72
        elif is_prog:
            color, alpha = COLOR_PROGRESSIVE, 0.88
        else:
            color, alpha = COLOR_SUCCESS, ALPHA_SUCCESS

        pitch.arrows(row["x_start"], row["y_start"], row["x_end"], row["y_end"],
                     color=color, width=1.3, headwidth=2.0, headlength=2.0,
                     ax=ax, zorder=3, alpha=alpha)
        pitch.scatter(row["x_start"], row["y_start"], s=32, marker="o",
                      color=color, edgecolors="white", linewidths=0.6,
                      ax=ax, zorder=6, alpha=alpha)

    leg = ax.legend(
        handles=[
            Line2D([0], [0], color=COLOR_SUCCESS, lw=2.0, label="Completed", alpha=0.65),
            Line2D([0], [0], color=COLOR_PROGRESSIVE, lw=2.0, label="Progressive", alpha=0.90),
            Line2D([0], [0], color=COLOR_FAIL, lw=2.0, label="Incomplete", alpha=0.90),
        ],
        loc="upper left", bbox_to_anchor=(0.01, 0.99), frameon=True,
        facecolor="#1a1a2e", edgecolor="#444466", fontsize=6.5,
        labelspacing=0.35, borderpad=0.4,
    )
    for t in leg.get_texts():
        t.set_color("white")
    leg.get_frame().set_alpha(0.90)
    attack_arrow(fig)
    return save_fig(fig), fig


def draw_corridor_heatmap(df):
    pitch = Pitch(pitch_type="custom", pitch_length=FIELD_X, pitch_width=FIELD_Y,
                  line_color="#444466", pitch_color="#1a1a2e", goal_type="box")
    fig, ax = pitch.draw(figsize=(FIG_W, FIG_H), dpi=FIG_DPI)
    fig.patch.set_facecolor("#1a1a2e")
    ax.set_facecolor("#1a1a2e")

    df_success = df[df["is_won"]].copy() if "is_won" in df.columns else df.copy()
    if len(df_success) > 0:
        x_bins = np.linspace(0, FIELD_X, 12)
        y_bins = np.linspace(0, FIELD_Y, 8)
        heatmap, _, _ = np.histogram2d(
            df_success["x_end"], df_success["y_end"], bins=[x_bins, y_bins]
        )
        heatmap = heatmap.T
        if heatmap.max() > 0:
            heatmap = heatmap / heatmap.max()
        pitch.heatmap(heatmap, x_bins, y_bins, ax=ax, cmap="Blues", alpha=0.7)

    attack_arrow(fig)
    return save_fig(fig), fig


def draw_top_xt_map(df, top_n=10):
    pitch = Pitch(pitch_type="custom", pitch_length=FIELD_X, pitch_width=FIELD_Y,
                  line_color="#444466", pitch_color="#1a1a2e", goal_type="box")
    fig, ax = pitch.draw(figsize=(FIG_W, FIG_H), dpi=FIG_DPI)
    fig.patch.set_facecolor("#1a1a2e")
    ax.set_facecolor("#1a1a2e")

    df_pos = df[df["is_won"]].copy() if "is_won" in df.columns else df.copy()
    if "delta_xt_adj" in df_pos.columns:
        df_top = df_pos.nlargest(top_n, "delta_xt_adj")
        for _, row in df_top.iterrows():
            norm_val = min(abs(row["delta_xt_adj"]) / 0.5, 1.0)
            color = plt.cm.RdYlGn(norm_val)
            pitch.arrows(row["x_start"], row["y_start"], row["x_end"], row["y_end"],
                         color=color, width=2.0, headwidth=3.0, headlength=3.0,
                         ax=ax, zorder=5, alpha=0.85)
            pitch.scatter(row["x_start"], row["y_start"], s=40, marker="o",
                          color=color, edgecolors="white", linewidths=0.8, ax=ax, zorder=6)

    attack_arrow(fig)
    return save_fig(fig), fig


def draw_defensive_map(df):
    pitch = Pitch(pitch_type="custom", pitch_length=FIELD_X, pitch_width=FIELD_Y,
                  line_color="#444466", pitch_color="#1a1a2e", goal_type="box")
    fig, ax = pitch.draw(figsize=(FIG_W, FIG_H), dpi=FIG_DPI)
    fig.patch.set_facecolor("#1a1a2e")
    ax.set_facecolor("#1a1a2e")

    for _, row in df.iterrows():
        if row.get("is_duel_won", False):
            color, marker, s = C_GREEN, "o", 80
        elif row.get("is_duel_lost", False):
            color, marker, s = COLOR_FAIL, "X", 80
        elif row.get("is_interception", False):
            color, marker, s = C_BLUE, "D", 90
        else:
            color, marker, s = "#666688", "o", 60
        pitch.scatter(row["x"], row["y"], s=s, marker=marker,
                      color=color, edgecolors="white", linewidths=0.8,
                      ax=ax, zorder=5, alpha=0.85)

    leg = ax.legend(
        handles=[
            Line2D([0], [0], marker="o", color=C_GREEN, label="Duel Won",
                   markerfacecolor=C_GREEN, markersize=7, linewidth=0),
            Line2D([0], [0], marker="X", color=COLOR_FAIL, label="Duel Lost",
                   markerfacecolor=COLOR_FAIL, markersize=7, linewidth=0),
            Line2D([0], [0], marker="D", color=C_BLUE, label="Interception",
                   markerfacecolor=C_BLUE, markersize=7, linewidth=0),
        ],
        loc="upper left", bbox_to_anchor=(0.01, 0.99), frameon=True,
        facecolor="#1a1a2e", edgecolor="#444466", fontsize=6.5,
        labelspacing=0.35, borderpad=0.4,
    )
    for t in leg.get_texts():
        t.set_color("white")
    leg.get_frame().set_alpha(0.90)
    return save_fig(fig), fig


def draw_defensive_heatmap(df):
    pitch = Pitch(pitch_type="custom", pitch_length=FIELD_X, pitch_width=FIELD_Y,
                  line_color="#444466", pitch_color="#1a1a2e", goal_type="box")
    fig, ax = pitch.draw(figsize=(FIG_W, FIG_H), dpi=FIG_DPI)
    fig.patch.set_facecolor("#1a1a2e")
    ax.set_facecolor("#1a1a2e")

    if len(df) > 0:
        x_bins = np.linspace(0, FIELD_X, 10)
        y_bins = np.linspace(0, FIELD_Y, 8)
        heatmap, _, _ = np.histogram2d(df["x"], df["y"], bins=[x_bins, y_bins])
        heatmap = heatmap.T
        if heatmap.max() > 0:
            heatmap = heatmap / heatmap.max()
        pitch.heatmap(heatmap, x_bins, y_bins, ax=ax, cmap="Reds", alpha=0.6)

    return save_fig(fig), fig


def draw_funnel_protection_map(df):
    pitch = Pitch(pitch_type="custom", pitch_length=FIELD_X, pitch_width=FIELD_Y,
                  line_color="#444466", pitch_color="#1a1a2e", goal_type="box")
    fig, ax = pitch.draw(figsize=(FIG_W, FIG_H), dpi=FIG_DPI)
    fig.patch.set_facecolor("#1a1a2e")
    ax.set_facecolor("#1a1a2e")

    from matplotlib.patches import Rectangle as MplRect
    funnel_rect = MplRect(
        (0, PENALTY_AREA_Y_MIN), FUNNEL_X_EXTEND, PENALTY_AREA_Y_MAX - PENALTY_AREA_Y_MIN,
        linewidth=1.5, edgecolor=C_AMBER, facecolor=C_AMBER_PASTEL, alpha=0.12, linestyle="--"
    )
    ax.add_patch(funnel_rect)
    ax.text(FUNNEL_X_EXTEND / 2, (PENALTY_AREA_Y_MIN + PENALTY_AREA_Y_MAX) / 2,
            "FUNNEL", ha="center", va="center", fontsize=7, color=C_AMBER_PASTEL,
            alpha=0.6, fontweight="bold", rotation=90)

    df_funnel = df[df["in_funnel"]].copy() if "in_funnel" in df.columns else df.copy()
    for _, row in df_funnel.iterrows():
        if row.get("is_duel_won", False):
            color, marker, s = C_GREEN, "o", 100
        elif row.get("is_interception", False):
            color, marker, s = C_BLUE, "D", 100
        else:
            color, marker, s = "#666688", "o", 70
        pitch.scatter(row["x"], row["y"], s=s, marker=marker,
                      color=color, edgecolors="white", linewidths=1.2,
                      ax=ax, zorder=6, alpha=0.9)

    return save_fig(fig), fig


# ── SIDEBAR ──
num_matches = len(dfs_by_match)
all_match_stats = [compute_stats(dfs_by_match[m], m) for m in dfs_by_match]

st.sidebar.markdown(
    """<div style='text-align:center;padding:10px;background:linear-gradient(135deg,#1a1a2e,#16213e);border-radius:10px;margin-bottom:10px;'>
    <h2 style='color:#ffffff;margin:0;'>⚽ Pass Stats Dashboard</h2>
    <p style='color:#d0d0e8;margin:0;font-size:14px;'>2026 Season</p>
    <h3 style='color:#2F80ED;margin:5px 0;'>Hudson Cicala</h3></div>""",
    unsafe_allow_html=True,
)

img_path = "Captura de tela 2026-06-02 154425.png"
if os.path.exists(img_path):
    st.sidebar.image(img_path, use_container_width=True)

st.sidebar.markdown(
    """<div style='text-align:center;color:#5a5a7a;font-size:12px;'>Data collected from match footage</div>""",
    unsafe_allow_html=True,
)


# ── LAYOUT — SINGLE TAB ──
tab_dash, = st.tabs(["Detailed Dashboard"])

with tab_dash:
    sub_tab_passes, sub_tab_def = st.tabs(["Passes", "Defensive Actions"])

    with sub_tab_passes:
        st.markdown("### Match Filters")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            pass_match_options = ["All Matches"] + list(dfs_by_match.keys())
            selected_match = st.selectbox("Select Match", options=pass_match_options, index=0, key="pass_match")
        with col_f2:
            pass_filter = st.radio(
                "Pass Type",
                ["All", "Successful", "Unsuccessful", "Progressive", "Final Third"],
                index=0, horizontal=True, key="pass_filter",
            )

        if selected_match == "All Matches":
            df_game_filtered = pd.concat(dfs_by_match.values(), ignore_index=True)
            match_name_for_stats = "All Matches"
        else:
            df_game_filtered = dfs_by_match[selected_match].copy()
            match_name_for_stats = selected_match

        def apply_filter(df):
            if pass_filter == "Successful":
                return df[df["is_won"]].copy()
            if pass_filter == "Unsuccessful":
                return df[~df["is_won"]].copy()
            if pass_filter == "Progressive":
                return df[df["progressive"]].copy()
            if pass_filter == "Final Third":
                return df[(df["x_start"] < FINAL_THIRD_LINE_X) & (df["x_end"] >= FINAL_THIRD_LINE_X)].copy()
            return df.copy()

        df_game = apply_filter(df_game_filtered)
        s_game = compute_stats(df_game, match_name_for_stats)

        s_avg = {}
        if num_matches > 0:
            for k in all_match_stats[0].keys():
                if isinstance(all_match_stats[0][k], (int, float)):
                    s_avg[k] = sum(s[k] for s in all_match_stats) / num_matches
                else:
                    s_avg[k] = 0
        else:
            s_avg = s_game.copy()

        force_avg = selected_match == "All Matches"
        if force_avg:
            s_game = s_avg.copy()

        st.markdown("---")

        img_pm_game, fig_pm_game = draw_pass_map(df_game)
        plt.close(fig_pm_game)

        img_ht_game, fig_ht_game = draw_corridor_heatmap(df_game)
        plt.close(fig_ht_game)

        top_n_xt = 10 if force_avg else 5
        img_xt_game, fig_xt_game = draw_top_xt_map(df_game, top_n=top_n_xt)
        plt.close(fig_xt_game)

        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.markdown('<div style="text-align:center;color:#d0d0e8;font-weight:600;">📋 Pass Map</div>', unsafe_allow_html=True)
            st.image(img_pm_game, use_container_width=True)
        with col_m2:
            st.markdown('<div style="text-align:center;color:#d0d0e8;font-weight:600;">📊 Zone Heatmap (Destination)</div>', unsafe_allow_html=True)
            st.image(img_ht_game, use_container_width=True)
        with col_m3:
            label = "Top 10" if force_avg else "Top 5"
            st.markdown(f'<div style="text-align:center;color:#d0d0e8;font-weight:600;">🎯 {label} Pass Impact</div>', unsafe_allow_html=True)
            st.image(img_xt_game, use_container_width=True)

        st.markdown("&nbsp;", unsafe_allow_html=True)

        col_s1, col_s2, col_s3 = st.columns(3)
        total_impact_value = float(df_game.loc[df_game["is_won"], "delta_xt_adj"].sum()) if "delta_xt_adj" in df_game.columns and df_game["is_won"].sum() > 0 else 0.0

        with col_s1:
            st.markdown(
                section_card_html(
                    "📋 Pass Overview",
                    [
                        f"<b>Total Passes (AVG):</b> {s_game['total']}",
                        f"<b>Per 90:</b> {s_game['total_p90']:.1f}",
                        f"<b>Minutes:</b> {s_game['mins']:.0f}",
                    ]
                ),
                unsafe_allow_html=True,
            )
        with col_s2:
            st.markdown(
                section_card_html(
                    "📈 Progression",
                    [
                        f"<b>Progressive Passes:</b> {s_game['progressive']:.0f}",
                        f"<b>Per 90:</b> {s_game['progressive_p90']:.1f}",
                        f"<b>Accuracy:</b> {s_game['accuracy']:.1f}%",
                    ]
                ),
                unsafe_allow_html=True,
            )
        with col_s3:
            st.markdown(
                section_card_html(
                    "⚡ Impact",
                    [
                        f"<b>Total Impact (xT):</b> {total_impact_value:.3f}",
                        f"<b>Per 90:</b> {total_impact_value * (90.0 / s_game['mins'] if s_game['mins'] > 0 else 1):.3f}",
                        "<span style='font-size:11px;color:#5a5a7a;'>% Positive Impact — Passes that generated a positive impact based on where they ended on the field.</span>",
                    ]
                ),
                unsafe_allow_html=True,
            )

    with sub_tab_def:
        st.markdown("### Match Filter")
        col_df1, col_df2 = st.columns(2)
        with col_df1:
            def_match_options = ["All Matches"] + list(defensive_dfs_by_match.keys())
            selected_def_match = st.selectbox("Select Match", options=def_match_options, index=0, key="def_match")
        with col_df2:
            def_type_filter = st.radio(
                "Filter Type",
                ["All", "Duels Only", "Interceptions Only"],
                horizontal=True, key="def_type_filter",
            )

        if selected_def_match == "All Matches":
            df_def_game_raw = pd.concat(defensive_dfs_by_match.values(), ignore_index=True)
            def_match_name_for_stats = "All Matches"
        else:
            df_def_game_raw = defensive_dfs_by_match[selected_def_match].copy()
            def_match_name_for_stats = selected_def_match

        if def_type_filter == "Duels Only":
            df_def_game = df_def_game_raw[df_def_game_raw["is_duel"]].copy()
        elif def_type_filter == "Interceptions Only":
            df_def_game = df_def_game_raw[df_def_game_raw["is_interception"]].copy()
        else:
            df_def_game = df_def_game_raw.copy()

        d_game = compute_defensive_stats(df_def_game, def_match_name_for_stats)
        def_all = [compute_defensive_stats(defensive_dfs_by_match[m], m) for m in defensive_dfs_by_match]

        d_avg = {}
        if len(def_all) > 0:
            for k in def_all[0].keys():
                if isinstance(def_all[0][k], (int, float)):
                    d_avg[k] = sum(s[k] for s in def_all) / len(def_all)
                else:
                    d_avg[k] = 0
        else:
            d_avg = d_game.copy()

        force_avg_def = selected_def_match == "All Matches"
        if force_avg_def:
            d_game = d_avg.copy()

        st.markdown("---")

        img_def_map, fig_def_map = draw_defensive_map(df_def_game)
        plt.close(fig_def_map)

        img_def_hm, fig_def_hm = draw_defensive_heatmap(df_def_game)
        plt.close(fig_def_hm)

        img_funnel, fig_funnel = draw_funnel_protection_map(df_def_game)
        plt.close(fig_funnel)

        col_dm1, col_dm2, col_dm3 = st.columns(3)
        with col_dm1:
            st.markdown('<div style="text-align:center;color:#d0d0e8;font-weight:600;">🛡️ Defensive Actions Map</div>', unsafe_allow_html=True)
            st.image(img_def_map, use_container_width=True)
        with col_dm2:
            st.markdown('<div style="text-align:center;color:#d0d0e8;font-weight:600;">🔥 Defensive Heatmap</div>', unsafe_allow_html=True)
            st.image(img_def_hm, use_container_width=True)
        with col_dm3:
            st.markdown('<div style="text-align:center;color:#d0d0e8;font-weight:600;">🔒 Funnel Protection</div>', unsafe_allow_html=True)
            st.image(img_funnel, use_container_width=True)

        st.markdown("&nbsp;", unsafe_allow_html=True)

        col_ds1, col_ds2, col_ds3 = st.columns(3)
        with col_ds1:
            st.markdown(
                section_card_html(
                    "🛡️ Duels",
                    [
                        f"<b>Total Duels:</b> {d_game['duels']:.0f}",
                        f"<b>Won:</b> {d_game['duel_won']:.0f}",
                        f"<b>Accuracy:</b> {d_game['duel_accuracy']:.1f}%",
                    ]
                ),
                unsafe_allow_html=True,
            )
        with col_ds2:
            st.markdown(
                section_card_html(
                    "✋ Interceptions",
                    [
                        f"<b>Total:</b> {d_game['interceptions']:.0f}",
                        f"<b>Per 90:</b> {d_game['interceptions_p90']:.1f}",
                    ]
                ),
                unsafe_allow_html=True,
            )
        with col_ds3:
            st.markdown(
                section_card_html(
                    "🔒 Funnel Protection",
                    [
                        f"<b>Funnel Actions:</b> {d_game['funnel_protections']:.0f}",
                        f"<b>Per 90:</b> {d_game['funnel_protections_p90']:.1f}",
                    ]
                ),
                unsafe_allow_html=True,
            )

        st.markdown("---")
        st.markdown(
            """<div style='background:#1a1a2e;border-radius:8px;padding:10px;border:1px solid #2a2a2e;'>
            <p style='color:#d0d0e8;font-size:12px;margin:0;'>
            <b>Funnel Zone</b>: Central defensive area near the penalty box (x ≤ 33.0, y 18–62).
            Actions in this zone represent high-value defensive contributions.
            </p></div>""",
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.markdown("### 📄 Export Complete Dashboard")
    if st.button("📸 Download Screenshot (PDF) — Passes + Defensive Actions", use_container_width=True):
        with st.spinner("Generating PDF with dashboard screenshots..."):
            df_all_passes = pd.concat(dfs_by_match.values(), ignore_index=True)
            df_all_def = pd.concat(defensive_dfs_by_match.values(), ignore_index=True)

            img_pm_all, _ = draw_pass_map(df_all_passes)
            plt.close()

            img_ht_all, _ = draw_corridor_heatmap(df_all_passes)
            plt.close()

            img_xt_all, _ = draw_top_xt_map(df_all_passes, top_n=10)
            plt.close()

            img_dm_all, _ = draw_defensive_map(df_all_def)
            plt.close()

            img_dhm_all, _ = draw_defensive_heatmap(df_all_def)
            plt.close()

            img_fn_all, _ = draw_funnel_protection_map(df_all_def)
            plt.close()

            pdf_bytes = export_dashboard_pdf(
                [img_pm_all, img_ht_all, img_xt_all],
                [img_dm_all, img_dhm_all, img_fn_all]
            )

            st.download_button(
                "📥 Save PDF",
                data=pdf_bytes,
                file_name="hudson_cicala_dashboard.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
