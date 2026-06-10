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

# PAGE CONFIG
st.set_page_config(layout="wide", page_title="Hudson Cicala — Dashboard")

# OPTIONAL DOCX IMPORT
DOCX_AVAILABLE = True
try:
    from docx import Document
except Exception:
    DOCX_AVAILABLE = False

# STYLE
st.markdown("""
<style>
    /* GLOBAL */
    .stApp { background: #0d0d1f; color: #e0e0f0; }
    .stSelectbox label, .stRadio label { color: #c0c0d8 !important; font-weight: 500; }
    .st-bw { background: #1a1a2e; }
    div[data-testid="stSelectbox"] > div { background: #1a1a2e; border: 1px solid #3a3a5c; }
    div[data-testid="stRadio"] > div { color: #c0c0d8; }
    .stTabs [data-baseweb="tab-list"] { background: #1a1a2e; border-radius: 8px; padding: 4px; gap: 4px; }
    .stTabs [data-baseweb="tab"] { color: #a0a0c0; font-weight: 500; border-radius: 6px; padding: 6px 18px; }
    .stTabs [aria-selected="true"] { background: #2a2a4e !important; color: #ffffff !important; }
    h1, h2, h3, h4, h5, h6 { color: #f0f0ff; }
    .stMetric label { color: #a0a0c0 !important; }
    .stMetric [data-testid="stMetricValue"] { color: #ffffff !important; }
    div[data-testid="stImage"] { border: 1px solid #2a2a4e; border-radius: 8px; overflow: hidden; }
    button[data-testid="baseButton-secondary"] {
        background: #1a1a2e !important; color: #c0c0d8 !important; border: 1px solid #3a3a5c !important;
        border-radius: 8px !important; font-weight: 500 !important;
    }
    button[data-testid="baseButton-secondary"]:hover { background: #2a2a4e !important; border-color: #5a5a7a !important; }
    .stDownloadButton button {
        background: linear-gradient(135deg, #1a56db, #2F80ED) !important;
        color: #ffffff !important; border: none !important; border-radius: 8px !important;
        font-weight: 600 !important; padding: 10px 24px !important;
    }
    .stDownloadButton button:hover { background: linear-gradient(135deg, #2563eb, #3b82f6) !important; }
    hr { border-color: #2a2a4e; margin: 20px 0; }
    /* SIDEBAR */
    section[data-testid="stSidebar"] { background: #0d0d1f; border-right: 1px solid #1a1a2e; }
    section[data-testid="stSidebar"] h1 { color: #f0f0ff; font-size: 1.3rem; }
    section[data-testid="stSidebar"] .stMarkdown p { color: #a0a0c0; font-size: 0.85rem; }
    /* TOOLTIP */
    .custom-tooltip { position: relative; cursor: help; border-bottom: 1px dotted #5a5a7a; }
    .custom-tooltip .tooltip-text {
        visibility: hidden; width: 220px; background: #1a1a2e; color: #d0d0e8;
        border: 1px solid #3a3a5c; border-radius: 6px; padding: 8px 10px;
        position: absolute; z-index: 100; bottom: 120%; left: 50%; margin-left: -110px;
        font-size: 0.75rem; line-height: 1.3; font-weight: 400;
        opacity: 0; transition: opacity 0.2s;
    }
    .custom-tooltip:hover .tooltip-text { visibility: visible; opacity: 1; }
    /* METRIC CARDS */
    .metric-card {
        background: linear-gradient(135deg, #12122a 0%, #1a1a2e 100%);
        border-radius: 12px; padding: 16px 18px;
        border-left: 3px solid var(--accent-color, #5b9bd5);
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
        margin-bottom: 12px;
    }
    .metric-card .label { color: #a0a0c0; font-size: 0.78rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-card .value { color: #ffffff; font-size: 1.35rem; font-weight: 700; margin-top: 2px; }
    .metric-card .sub { color: #7a7a9a; font-size: 0.72rem; margin-top: 1px; }
    .metric-card .arrow { font-size: 0.70rem; font-weight: 600; }
    .metric-card .arrow.up { color: #10b981; }
    .metric-card .arrow.down { color: #E07070; }
    .metric-card .sep { border-top: 1px solid rgba(255,255,255,0.06); margin: 6px 0; }
    .metric-card .avg-label { color: #5a5a7a; font-size: 0.70rem; }
    .metric-card .avg-value { color: #a0a0c0; font-size: 0.85rem; }
    /* EXPANDED EXPLANATION */
    .explanation-card {
        background: linear-gradient(135deg, #12122a 0%, #1a1a2e 100%);
        border-radius: 12px; padding: 14px 16px;
        border-left: 3px solid var(--accent-color);
        margin-bottom: 12px;
    }
    .explanation-card .explabel { color: #a0a0c0; font-size: 0.75rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
    .explanation-card .exptext { color: #b0b0d0; font-size: 0.80rem; line-height: 1.4; }
    /* IMAGE LABELS */
    .img-label {
        text-align: center; font-size: 0.85rem; font-weight: 600; color: #d0d0e8;
        padding: 6px 0 2px 0; border-bottom: 1px solid #2a2a4e; margin-bottom: 8px;
    }
    /* SECTION CARD HEADER */
    .section-header {
        font-size: 0.90rem; font-weight: 700; color: #f0f0ff;
        padding: 6px 0; margin-bottom: 6px;
    }
    /* SPACER */
    .stat-spacer { height: 6px; }
</style>
""", unsafe_allow_html=True)

# CONSTANTS
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

# PDF STYLE CONSTANTS
PDF_BG = "#0d0d1f"
PDF_TEXT_WHITE = "#ffffff"
PDF_TEXT_LIGHT = "#d0d0e8"
PDF_TEXT_DIM = "#5a5a7a"

def _hex_to_rgba(hex_color, alpha=1.0):
    if hex_color.startswith('#'):
        h = hex_color.lstrip('#')
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return f'rgba({r},{g},{b},{alpha})'
    return hex_color

def get_lane(y):
    if y >= LANE_LEFT_MIN:
        return "left"
    elif y < LANE_RIGHT_MAX:
        return "right"
    return "center"

def distance_to_goal(x, y):
    return np.sqrt((GOAL_X - x) ** 2 + (GOAL_Y - y) ** 2)

def is_progressive_pass(x_start, y_start, x_end, y_end):
    if x_start < 35:
        return False
    start_dist = distance_to_goal(x_start, y_start)
    end_dist = distance_to_goal(x_end, y_end)
    if start_dist == 0:
        return False
    return ((start_dist - end_dist) / start_dist) >= 0.25

def classify_pass_direction(x_start, y_start, x_end, y_end):
    dx = x_end - x_start
    dy = y_end - y_start
    dist = np.sqrt(dx ** 2 + dy ** 2)
    angle_deg = np.degrees(np.arctan2(abs(dy), dx))
    if angle_deg <= 45.0:
        return "forward"
    if angle_deg >= 135.0:
        return "backward"
    if dist > LATERAL_MIN_DIST:
        return "lateral_right" if dy > 0 else "lateral_left"
    return "forward" if dx >= 0 else "backward"

def distance_bonus(distance):
    excess = np.maximum(0.0, np.asarray(distance, dtype=float) - D_REF)
    return np.minimum(BONUS_CAP, np.log1p(excess / D_SCALE))

@st.cache_data(show_spinner=False)
def compute_xt_grid(NX=16, NY=12, sub=24):
    ncols_hr = NX * sub
    nrows_hr = NY * sub
    xe = np.linspace(0, FIELD_X, ncols_hr + 1)
    ye = np.linspace(0, FIELD_Y, nrows_hr + 1)
    xc = (xe[:-1] + xe[1:]) / 2
    yc_arr = (ye[:-1] + ye[1:]) / 2
    Xc, Yc = np.meshgrid(xc, yc_arr)

    xp = 0.01 + (Xc / FIELD_X) * 0.99
    yc = 1.0 - np.abs((Yc / FIELD_Y) - 0.5) * 2.0
    base = xp * (0.8 + 0.2 * yc)
    base = (base - base.min()) / (base.max() - base.min() + 1e-12)

    XT = base.copy()
    XT = (XT - XT.min()) / (XT.max() - XT.min() + 1e-12)

    XTc = np.zeros((NY, NX))
    for iy in range(NY):
        for ix in range(NX):
            XTc[iy, ix] = XT[iy * sub:(iy + 1) * sub, ix * sub:(ix + 1) * sub].mean()
    XTc = (XTc - XTc.min()) / (XTc.max() - XTc.min() + 1e-12)
    return XTc

XT_GRID = compute_xt_grid()

def xt_value(x, y):
    ix = int(np.clip((x / FIELD_X) * NX_XT, 0, NX_XT - 1))
    iy = int(np.clip((y / FIELD_Y) * NY_XT, 0, NY_XT - 1))
    return float(XT_GRID[iy, ix])

def is_in_funnel_zone(x, y):
    return x <= FUNNEL_X_EXTEND and PENALTY_AREA_Y_MIN <= y <= PENALTY_AREA_Y_MAX

# BASE PASSES
BASE_MATCHES_DATA = {
    "Connecticut United (03-27)": [
        ("PASS WON", 26.75, 68.34, 8.97, 51.05, None),
        ("PASS WON", 31.24, 51.22, 34.57, 72.50, None),
        ("PASS WON", 36.06, 46.90, 44.37, 57.04, None),
        ("PASS WON", 48.36, 64.02, 58.17, 51.72, None),
        ("PASS WON", 58.17, 64.02, 62.49, 55.21, None),
        ("PASS WON", 54.51, 49.72, 64.82, 61.69, None),
        ("PASS WON", 42.21, 70.84, 34.90, 76.49, None),
        ("PASS WON", 43.54, 75.32, 36.73, 67.84, None),
        ("PASS WON", 32.24, 53.96, 6.81, 38.50, None),
        ("PASS WON", 33.57, 65.77, 36.56, 75.57, None),
        ("PASS WON", 37.39, 61.11, 43.04, 75.41, None),
        ("PASS WON", 65.49, 53.63, 56.18, 70.42, None),
        ("PASS WON", 55.68, 48.15, 46.87, 30.86, None),
        ("PASS WON", 52.02, 22.05, 46.70, 41.99, None),
        ("PASS WON", 62.16, 35.51, 71.80, 35.18, None),
        ("PASS WON", 54.02, 33.35, 63.99, 22.55, None),
        ("PASS WON", 60.00, 22.21, 76.62, 32.85, None),
        ("PASS WON", 87.10, 9.41, 77.45, 16.23, None),
        ("PASS WON", 62.66, 20.05, 117.18, 8.25, None),
        ("PASS WON", 98.90, 43.49, 103.22, 47.15, None),
        ("PASS WON", 70.31, 45.98, 82.28, 60.11, None),
        ("PASS WON", 85.10, 75.24, 101.39, 74.08, None),
        ("PASS WON", 53.18, 67.59, 39.05, 59.62, None),
        ("PASS WON", 55.18, 49.64, 54.85, 13.07, None),
        ("PASS WON", 68.64, 19.22, 49.03, 24.37, None),
        ("PASS WON", 53.35, 22.71, 59.34, 30.19, None),
        ("PASS WON", 44.37, 24.71, 40.05, 46.82, None),
        ("PASS WON", 43.88, 39.34, 41.38, 73.08, None),
        ("PASS WON", 56.84, 53.46, 70.81, 76.24, None),
        ("PASS WON", 82.77, 12.24, 91.42, 4.59, None),
        ("PASS WON", 108.04, 11.74, 115.69, 58.29, None),
        ("PASS WON", 93.08, 3.93, 111.03, 13.74, None),
        ("PASS WON", 84.60, 17.89, 96.74, 22.05, None),
        ("PASS WON", 58.34, 16.06, 65.65, 2.43, None),
        ("PASS WON", 52.02, 8.58, 44.37, 15.73, None),
        ("PASS WON", 61.00, 23.21, 49.36, 15.23, None),
        ("PASS WON", 32.74, 30.69, 50.03, 33.02, None),
        ("PASS WON", 51.85, 33.68, 60.66, 40.00, None),
        ("PASS WON", 79.95, 60.45, 98.23, 60.28, None),
        ("PASS WON", 31.24, 52.14, 39.05, 72.08, None),
        ("PASS WON", 39.72, 48.98, 33.40, 57.62, None),
        ("PASS WON", 70.64, 51.47, 61.00, 51.64, None),
        ("PASS LOST", 53.35, 19.55, 73.96, 11.24, None),
        ("PASS LOST", 63.82, 20.55, 88.76, 22.55, None),
        ("PASS LOST", 85.60, 27.86, 94.41, 37.17, None),
        ("PASS LOST", 77.79, 27.53, 96.41, 25.37, None),
        ("PASS LOST", 91.09, 27.86, 109.54, 50.47, None),
        ("PASS LOST", 58.17, 26.04, 95.41, 40.33, None),
        ("PASS LOST", 53.35, 28.53, 73.80, 27.86, None),
        ("PASS LOST", 53.35, 34.02, 84.60, 58.62, None),
        ("PASS LOST", 56.18, 49.48, 97.07, 62.11, None),
        ("PASS LOST", 34.23, 74.91, 65.65, 78.57, None),
    ],
    "Nashville SC (03-28)": [
        ("PASS WON", 21.27, 14.23, 29.25, 31.02, None),
        ("PASS WON", 29.41, 23.38, 34.40, 64.60, None),
        ("PASS WON", 41.55, 39.67, 41.88, 6.92, None),
        ("PASS WON", 44.54, 32.52, 43.54, 14.23, None),
        ("PASS WON", 23.59, 56.46, 34.57, 47.48, None),
        ("PASS WON", 30.58, 64.44, 21.10, 49.48, None),
        ("PASS WON", 33.07, 56.79, 49.53, 69.59, None),
        ("PASS WON", 33.24, 59.78, 44.04, 71.75, None),
        ("PASS WON", 61.50, 71.58, 54.68, 75.57, None),
        ("PASS WON", 63.16, 50.81, 78.45, 67.26, None),
        ("PASS WON", 63.49, 76.90, 84.44, 62.77, None),
        ("PASS WON", 76.96, 56.96, 86.93, 57.79, None),
        ("PASS WON", 82.61, 59.12, 96.41, 68.43, None),
        ("PASS WON", 79.78, 35.35, 106.21, 11.74, None),
        ("PASS WON", 45.37, 49.64, 40.72, 32.02, None),
        ("PASS LOST", 78.62, 64.94, 96.57, 67.10, None),
        ("PASS LOST", 85.43, 68.76, 106.05, 77.74, None),
    ],
    "Seongnam FC (03-29)": [
        ("PASS WON", 28.08, 28.53, 29.75, 8.25, None),
        ("PASS WON", 33.74, 26.54, 29.41, 43.82, None),
        ("PASS WON", 28.08, 47.15, 31.57, 64.60, None),
        ("PASS WON", 39.39, 43.82, 51.69, 53.46, None),
        ("PASS WON", 43.88, 46.15, 55.84, 40.66, None),
        ("PASS WON", 47.03, 49.97, 44.04, 28.03, None),
        ("PASS WON", 47.53, 50.81, 71.97, 33.18, None),
        ("PASS WON", 67.65, 52.63, 64.32, 33.85, None),
        ("PASS WON", 73.63, 65.10, 69.31, 73.25, None),
        ("PASS WON", 77.29, 63.27, 79.12, 72.91, None),
        ("PASS WON", 81.61, 56.62, 93.91, 73.75, None),
        ("PASS WON", 86.43, 66.43, 81.78, 54.96, None),
        ("PASS WON", 111.03, 71.42, 99.56, 67.59, None),
        ("PASS WON", 89.76, 59.62, 97.74, 48.98, None),
        ("PASS WON", 88.43, 52.47, 96.41, 74.24, None),
        ("PASS WON", 87.93, 50.97, 77.12, 27.70, None),
        ("PASS WON", 81.61, 53.63, 74.30, 27.03, None),
        ("PASS WON", 79.28, 51.14, 94.91, 70.42, None),
        ("PASS WON", 52.85, 32.85, 65.49, 25.37, None),
        ("PASS WON", 82.77, 33.18, 69.31, 47.65, None),
        ("PASS LOST", 72.14, 16.56, 78.45, 1.60, None),
        ("PASS LOST", 79.62, 27.53, 97.07, 47.98, None),
        ("PASS LOST", 91.75, 50.14, 109.70, 65.77, None),
        ("PASS LOST", 96.41, 56.79, 107.04, 67.26, None),
    ],
    "NY Red Bulls (03-31)": [
        ("PASS WON", 39.39, 19.39, 52.35, 4.76, None),
        ("PASS WON", 63.82, 7.92, 72.63, 1.43, None),
        ("PASS WON", 70.47, 11.91, 80.95, 13.74, None),
        ("PASS WON", 64.49, 22.55, 97.24, 10.24, None),
        ("PASS WON", 32.07, 35.51, 43.04, 28.20, None),
        ("PASS WON", 53.52, 46.32, 54.02, 33.68, None),
        ("PASS WON", 77.12, 48.64, 84.94, 50.14, None),
        ("PASS WON", 78.12, 52.47, 117.52, 69.42, None),
        ("PASS WON", 88.76, 65.93, 97.40, 76.74, None),
        ("PASS WON", 82.61, 69.26, 86.60, 77.40, None),
        ("PASS WON", 78.62, 66.26, 79.62, 78.40, None),
        ("PASS WON", 83.61, 75.91, 62.49, 57.12, None),
        ("PASS WON", 34.40, 50.14, 88.76, 75.41, None),
        ("PASS WON", 56.68, 64.27, 78.29, 64.27, None),
        ("PASS WON", 51.85, 73.25, 54.18, 78.07, None),
        ("PASS WON", 41.05, 57.45, 46.04, 74.91, None),
        ("PASS WON", 37.39, 60.61, 41.71, 73.91, None),
        ("PASS WON", 30.41, 63.44, 36.89, 77.40, None),
        ("PASS WON", 26.09, 63.94, 28.42, 76.74, None),
        ("PASS WON", 22.43, 56.62, 22.10, 76.41, None),
        ("PASS WON", 33.90, 64.77, 25.42, 73.58, None),
        ("PASS LOST", 41.88, 42.49, 56.18, 52.97, None),
        ("PASS LOST", 37.56, 41.16, 46.37, 53.96, None),
        ("PASS LOST", 54.68, 56.96, 54.85, 64.44, None),
        ("PASS LOST", 51.69, 68.43, 66.15, 76.57, None),
    ],
}

# DEFENSIVE ACTIONS
DEFENSIVE_MATCHES_DATA = {
    "Michigan Wolves (02-20)": [
        ("DUEL_WON", 53.85, 25.21),
        ("DUEL_WON", 23.59, 29.69),
        ("DUEL_WON", 43.88, 50.31),
        ("DUEL_WON", 16.28, 50.47),
        ("DUEL_WON", 15.62, 72.08),
        ("DUEL_LOST", 73.63, 27.70),
        ("DUEL_LOST", 17.78, 75.41),
        ("INTERCEPTION", 65.82, 19.05),
        ("INTERCEPTION", 72.80, 57.62),
    ],
    "Philadelphia Union (02-27)": [
        ("DUEL_WON", 67.98, 34.68),
        ("DUEL_WON", 41.05, 23.54),
        ("DUEL_WON", 21.27, 31.36),
        ("DUEL_WON", 39.55, 60.95),
        ("DUEL_LOST", 29.08, 40.50),
        ("INTERCEPTION", 53.52, 19.05),
        ("INTERCEPTION", 28.08, 27.03),
        ("INTERCEPTION", 29.75, 53.63),
        ("INTERCEPTION", 30.58, 69.42),
        ("INTERCEPTION", 52.19, 58.12),
        ("INTERCEPTION", 59.17, 63.11),
        ("INTERCEPTION", 80.78, 68.92),
    ],
    "Columbus Crew (03-06)": [
        ("DUEL_WON", 22.76, 29.69),
        ("DUEL_WON", 48.36, 20.72),
        ("DUEL_LOST", 63.32, 56.96),
        ("DUEL_LOST", 25.42, 53.63),
        ("DUEL_LOST", 27.42, 34.35),
        ("DUEL_LOST", 35.06, 36.84),
        ("INTERCEPTION", 29.75, 35.84),
        ("INTERCEPTION", 29.41, 40.00),
        ("INTERCEPTION", 37.39, 60.95),
    ],
    "Minnesota United (03-13)": [
        ("DUEL_WON", 44.04, 58.95),
        ("DUEL_WON", 14.78, 18.56),
        ("DUEL_WON", 17.61, 12.24),
        ("DUEL_LOST", 77.29, 27.20),
        ("DUEL_LOST", 39.89, 3.43),
        ("DUEL_LOST", 33.24, 10.91),
        ("DUEL_LOST", 35.90, 57.12),
        ("DUEL_LOST", 0.99, 69.26),
        ("INTERCEPTION", 31.74, 38.50),
        ("INTERCEPTION", 35.06, 36.34),
        ("INTERCEPTION", 38.39, 41.00),
        ("INTERCEPTION", 46.54, 26.37),
        ("INTERCEPTION", 40.38, 19.22),
    ],
    "Vardar Soccer (03-14)": [
        ("INTERCEPTION", 72.63, 35.18),
        ("INTERCEPTION", 12.29, 44.99),
    ],
    "Colorado Rapids (03-20)": [
        ("DUEL_WON", 36.39, 73.75),
        ("DUEL_WON", 39.39, 68.76),
        ("DUEL_WON", 52.02, 66.10),
        ("DUEL_WON", 21.60, 53.63),
        ("DUEL_WON", 35.06, 43.32),
        ("DUEL_WON", 36.39, 31.36),
        ("DUEL_WON", 45.54, 25.04),
        ("DUEL_WON", 34.40, 21.71),
        ("DUEL_WON", 53.68, 17.23),
        ("DUEL_WON", 57.67, 22.55),
        ("DUEL_LOST", 78.95, 4.59),
        ("DUEL_LOST", 75.46, 65.43),
        ("DUEL_LOST", 33.07, 54.46),
        ("INTERCEPTION", 67.31, 9.58),
        ("INTERCEPTION", 39.89, 24.54),
        ("INTERCEPTION", 43.38, 28.86),
        ("INTERCEPTION", 27.92, 35.01),
        ("INTERCEPTION", 64.49, 53.80),
        ("INTERCEPTION", 36.56, 55.96),
        ("INTERCEPTION", 30.58, 62.11),
    ],
    "Connecticut United (03-27)": [
        ("DUEL_WON", 82.94, 3.43),
        ("DUEL_WON", 70.47, 21.05),
        ("DUEL_WON", 67.31, 27.53),
        ("DUEL_WON", 27.58, 32.52),
        ("DUEL_LOST", 65.49, 22.71),
        ("DUEL_LOST", 3.48, 72.42),
        ("INTERCEPTION", 82.28, 31.02),
        ("INTERCEPTION", 66.15, 26.04),
        ("INTERCEPTION", 83.94, 56.29),
        ("INTERCEPTION", 59.00, 61.44),
    ],
    "Nashville SC (03-28)": [
        ("DUEL_WON", 84.77, 54.79),
        ("DUEL_WON", 62.33, 55.46),
        ("DUEL_WON", 35.90, 62.61),
        ("DUEL_WON", 40.38, 70.09),
        ("DUEL_WON", 40.38, 40.33),
        ("DUEL_WON", 26.92, 23.71),
        ("DUEL_LOST", 92.91, 24.54),
        ("DUEL_LOST", 90.59, 53.63),
        ("DUEL_LOST", 64.82, 59.78),
        ("DUEL_LOST", 51.02, 71.58),
        ("INTERCEPTION", 85.60, 23.38),
        ("INTERCEPTION", 65.65, 57.12),
        ("INTERCEPTION", 77.45, 61.78),
    ],
    "Seongnam FC (03-29)": [
        ("DUEL_LOST", 73.80, 21.71),
        ("INTERCEPTION", 38.06, 30.36),
    ],
    "NY Red Bulls (03-31)": [
        ("DUEL_WON", 33.87, 59.39),
        ("DUEL_WON", 37.58, 67.14),
        ("DUEL_LOST", 66.32, 60.28),
        ("INTERCEPTION", 34.90, 34.02),
        ("INTERCEPTION", 56.34, 42.66),
        ("INTERCEPTION", 68.15, 54.30),
    ],
    "Minnesota United (04-10)": [
        ("DUEL_WON", 15.62, 54.30),
        ("DUEL_LOST", 36.39, 39.34),
        ("DUEL_LOST", 10.79, 64.27),
        ("INTERCEPTION", 66.15, 68.59),
        ("INTERCEPTION", 25.42, 54.79),
        ("INTERCEPTION", 35.06, 48.15),
        ("INTERCEPTION", 22.76, 21.88),
        ("INTERCEPTION", 56.84, 25.87),
        ("INTERCEPTION", 82.11, 20.72),
    ],
    "Sporting Kansas City (04-17)": [
        ("DUEL_WON", 85.43, 17.06),
        ("DUEL_WON", 76.12, 20.72),
        ("DUEL_WON", 54.68, 12.07),
        ("DUEL_WON", 53.18, 24.87),
        ("DUEL_WON", 24.92, 34.35),
        ("DUEL_WON", 31.24, 49.64),
        ("DUEL_WON", 39.05, 52.14),
        ("DUEL_WON", 43.71, 62.61),
        ("DUEL_WON", 49.69, 73.25),
        ("DUEL_WON", 75.79, 62.77),
        ("DUEL_LOST", 30.24, 69.09),
        ("INTERCEPTION", 60.83, 15.40),
        ("INTERCEPTION", 10.79, 25.87),
        ("INTERCEPTION", 52.35, 52.97),
        ("INTERCEPTION", 70.14, 61.28),
        ("INTERCEPTION", 54.85, 62.11),
        ("INTERCEPTION", 39.89, 66.60),
    ],
    "Cedar Stars (04-22)": [
        ("DUEL_WON", 9.30, 22.88),
        ("DUEL_WON", 59.00, 15.06),
        ("DUEL_WON", 60.83, 44.65),
        ("INTERCEPTION", 75.46, 28.20),
        ("INTERCEPTION", 79.95, 57.29),
        ("INTERCEPTION", 27.09, 66.43),
    ],
    "South Florida (04-23)": [
        ("DUEL_WON", 36.23, 32.85),
        ("DUEL_WON", 42.05, 54.79),
        ("DUEL_WON", 35.56, 57.62),
        ("DUEL_WON", 70.97, 18.72),
        ("INTERCEPTION", 55.18, 63.77),
        ("INTERCEPTION", 22.26, 62.94),
    ],
    "Real Salt Lake (04-26)": [
        ("DUEL_WON", 47.70, 56.96),
        ("DUEL_WON", 26.75, 55.29),
        ("DUEL_WON", 21.93, 26.37),
        ("DUEL_WON", 68.15, 2.93),
        ("DUEL_LOST", 76.29, 32.02),
        ("INTERCEPTION", 15.78, 53.30),
        ("INTERCEPTION", 35.23, 24.54),
        ("INTERCEPTION", 76.79, 21.55),
    ],
    "Real Futbol (05-23)": [
        ("DUEL_WON", 72.63, 10.24),
        ("DUEL_WON", 73.80, 13.90),
        ("DUEL_WON", 54.68, 40.50),
        ("DUEL_LOST", 69.97, 22.55),
        ("DUEL_LOST", 30.24, 5.26),
        ("DUEL_LOST", 39.22, 71.75),
        ("INTERCEPTION", 75.46, 56.12),
    ],
    "San Jose (05-24)": [
        ("DUEL_WON", 8.97, 23.21),
        ("DUEL_WON", 23.76, 23.71),
        ("DUEL_WON", 24.09, 41.50),
        ("DUEL_WON", 30.91, 61.61),
        ("DUEL_WON", 65.15, 39.17),
        ("DUEL_WON", 69.31, 29.36),
        ("DUEL_LOST", 27.42, 52.97),
        ("DUEL_LOST", 30.74, 49.48),
        ("DUEL_LOST", 34.73, 52.80),
        ("DUEL_LOST", 43.38, 59.62),
        ("DUEL_LOST", 34.90, 63.77),
        ("DUEL_LOST", 31.08, 62.61),
        ("DUEL_LOST", 21.27, 66.93),
        ("DUEL_LOST", 70.47, 57.79),
        ("INTERCEPTION", 76.62, 21.38),
        ("INTERCEPTION", 80.78, 60.61),
        ("INTERCEPTION", 21.93, 57.45),
        ("INTERCEPTION", 25.59, 70.59),
        ("INTERCEPTION", 34.90, 31.52),
        ("INTERCEPTION", 38.39, 33.68),
        ("INTERCEPTION", 29.91, 23.38),
    ],
    "Houston Dynamo (05-26)": [
        ("DUEL_WON", 68.31, 37.84),
        ("DUEL_WON", 68.15, 42.33),
        ("DUEL_WON", 83.27, 73.75),
        ("DUEL_WON", 55.51, 62.77),
        ("DUEL_WON", 49.53, 75.91),
        ("DUEL_WON", 31.24, 70.92),
        ("DUEL_WON", 24.59, 55.29),
        ("DUEL_LOST", 21.60, 21.88),
        ("DUEL_LOST", 26.59, 60.45),
    ],
}

# HELPERS
def apply_date_mapping(name: str) -> str:
    mapping = {
        "Connecticut United": "Connecticut United (03-27)",
        "Nashville SC": "Nashville SC (03-28)",
        "Seongnam FC": "Seongnam FC (03-29)",
        "NY Red Bulls": "NY Red Bulls (03-31)",
        "Real Salt Lake": "Real Salt Lake (04-26)",
        "Real Futbol": "Real Futbol (05-23)",
        "San Jose": "San Jose (05-24)",
        "Houston Dynamo": "Houston Dynamo (05-26)"
    }
    for k, v in mapping.items():
        if k.lower() == name.lower().strip():
            return v
    return name

def get_match_minutes(match_name: str) -> float:
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

def read_docx_text(docx_path: Path) -> str:
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx is not installed.")
    doc = Document(str(docx_path))
    return "\n".join(p.text for p in doc.paragraphs if p.text and p.text.strip())

def parse_docx_events(raw_text: str) -> dict:
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    matches = {}
    current_match = None
    current_state = None

    re_match = re.compile(r"^Vs\s+(.+)$", re.IGNORECASE)
    re_success = re.compile(r"^Sucesso$", re.IGNORECASE)
    re_fail = re.compile(r"^Errado[s]?$", re.IGNORECASE)
    re_arrow = re.compile(
        r"^Seta\s+\d+:\s(([-+]?\d\.?\d+),\s*([-+]?\d\.?\d+))\s->\s(([-+]?\d\.?\d+),\s*([-+]?\d\.?\d+))$",
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
            x1, y1, x2, y2 = map(float, m_arrow.groups())
            matches[current_match].append((current_state, x1, y1, x2, y2, None))

    return {k: v for k, v in matches.items() if len(v) > 0}

def load_docx_matches(docx_filename="Passes - Hudson Cicala.docx") -> dict:
    p = Path(docx_filename)
    if not p.exists():
        return {}
    txt = read_docx_text(p)
    return parse_docx_events(txt)

# DATA LOADING
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

# BUILD DATAFRAMES & REORDER MATCHES
dfs_by_match = {}
for match_name, events in combined_matches_data.items():
    dfm = pd.DataFrame(events, columns=["type", "x_start", "y_start", "x_end", "y_end", "video"])
    dfm["match"] = match_name
    dfm["number"] = np.arange(1, len(dfm) + 1)
    dfm["is_won"] = dfm["type"].str.contains("WON", case=False)
    dfm["progressive"] = dfm.apply(
        lambda r: r["is_won"] and is_progressive_pass(r["x_start"], r["y_start"], r["x_end"], r["y_end"]),
        axis=1
    )
    dfm["direction"] = dfm.apply(
        lambda r: classify_pass_direction(r["x_start"], r["y_start"], r["x_end"], r["y_end"]),
        axis=1
    )
    dfm["is_forward"] = dfm["direction"] == "forward"
    dfm["is_backward"] = dfm["direction"] == "backward"
    dfm["is_lateral"] = dfm["direction"].isin(["lateral_left", "lateral_right"])
    dfm["pass_distance"] = np.sqrt(
        (dfm["x_end"] - dfm["x_start"]) ** 2 + (dfm["y_end"] - dfm["y_start"]) ** 2
    )
    dfm["xt_start"] = dfm.apply(lambda r: xt_value(r["x_start"], r["y_start"]), axis=1)
    dfm["xt_end"] = dfm.apply(lambda r: xt_value(r["x_end"], r["y_end"]), axis=1)
    dfm["delta_xt"] = np.where(dfm["is_won"], dfm["xt_end"] - dfm["xt_start"], 0.0)
    dfm["dist_bonus"] = distance_bonus(dfm["pass_distance"].values)
    dfm["delta_xt_adj"] = np.where(dfm["is_won"], dfm["delta_xt"] * (1.0 + dfm["dist_bonus"]), 0.0)
    dfs_by_match[match_name] = dfm

# REORDER LOGIC
items = list(dfs_by_match.items())
if len(items) >= 18:
    part1 = items[:6]
    part2 = items[14:18]
    part3 = items[6:14]
    part4 = items[18:]
    dfs_by_match = dict(part1 + part2 + part3 + part4)

df_all = pd.concat(dfs_by_match.values(), ignore_index=True)

# DEFENSIVE DATA LOADING
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

# STATS
def compute_stats(df: pd.DataFrame, match_name: str) -> dict:
    total = len(df)
    mins = get_match_minutes(match_name)
    p90_factor = 90.0 / mins if mins > 0 else 1.0

    if total == 0:
        return {
            "total_passes": 0, "successful_passes": 0, "unsuccessful_passes": 0, "accuracy_pct": 0.0,
            "progressive_attempted": 0, "progressive_successful": 0, "progressive_accuracy_pct": 0.0,
            "to_final_third_total": 0, "to_final_third_success": 0, "to_final_third_accuracy_pct": 0.0,
            "fwd": 0, "fwd_pct": 0.0, "bwd": 0, "bwd_pct": 0.0, "lat": 0, "lat_pct": 0.0,
            "pos_count": 0, "pos_pct": 0.0, "high_xt_pct": 0.0, "sum_dxt": 0.0,
            "total_p90": 0.0, "prog_p90": 0.0, "f3_p90": 0.0, "xt_p90": 0.0, "neg_xt_p90": 0.0,
            "minutes": mins, "long_acc_pct": 0.0, "high_xt_p90": 0.0, "dz_p90": 0.0,
            "advanced_passes_p90": 0.0, "advanced_accuracy_pct": 0.0,
        }

    successful = int(df["is_won"].sum())
    unsuccessful = total - successful
    accuracy = successful / total * 100.0

    progressive_total = int(df["progressive"].sum())
    progressive_unsuccessful = int(
        (df["is_won"] & df.apply(
            lambda r: is_progressive_pass(r["x_start"], r["y_start"], r["x_end"], r["y_end"]),
            axis=1
        )).sum()
    )
    progressive_attempted = progressive_total + progressive_unsuccessful
    progressive_accuracy = (progressive_total / progressive_attempted * 100.0) if progressive_attempted else 0.0

    to_final_third = (df["x_start"] < FINAL_THIRD_LINE_X) & (df["x_end"] >= FINAL_THIRD_LINE_X)
    to_final_third_total = int(to_final_third.sum())
    to_final_third_success = int((to_final_third & df["is_won"]).sum())
    to_final_third_accuracy = (to_final_third_success / to_final_third_total * 100.0) if to_final_third_total else 0.0

    long_passes = df[df["pass_distance"] > 25.0]
    long_total = len(long_passes)
    long_success = int(long_passes["is_won"].sum())
    long_acc_pct = (long_success / long_total * 100.0) if long_total > 0 else 0.0

    dz_mask = df["is_won"] & (
        (df["x_end"] >= 100.0) |
        ((df["x_end"] >= 80.0) & (df["x_end"] < 100.0) & (df["y_end"] >= LANE_RIGHT_MAX) & (df["y_end"] < LANE_LEFT_MIN))
    )
    dz_passes = int(dz_mask.sum())

    fwd = int(df["is_forward"].sum())
    bwd = int(df["is_backward"].sum())
    lat = int(df["is_lateral"].sum())

    pos_count = int((df["is_won"] & (df["delta_xt_adj"] > 0)).sum())
    pos_pct = (pos_count / total * 100.0) if total > 0 else 0.0

    high_xt = int((df["delta_xt_adj"] > 0.1).sum())
    sum_dxt = float(df.loc[df["is_won"], "delta_xt_adj"].sum())
    neg_xt = float(df.loc[df["is_won"] & (df["delta_xt_adj"] < 0), "delta_xt_adj"].sum())

    advanced_successful = progressive_total + to_final_third_success
    advanced_attempted = progressive_attempted + to_final_third_total
    advanced_accuracy_pct = (advanced_successful / advanced_attempted * 100.0) if advanced_attempted else 0.0
    advanced_passes_p90 = round((progressive_total + to_final_third_success) * p90_factor, 2)

    return {
        "total_passes": total,
        "successful_passes": successful,
        "unsuccessful_passes": unsuccessful,
        "accuracy_pct": round(accuracy, 2),
        "progressive_attempted": progressive_attempted,
        "progressive_successful": progressive_total,
        "progressive_accuracy_pct": round(progressive_accuracy, 2),
        "to_final_third_total": to_final_third_total,
        "to_final_third_success": to_final_third_success,
        "to_final_third_accuracy_pct": round(to_final_third_accuracy, 2),
        "fwd": fwd,
        "fwd_pct": round(fwd / total * 100.0, 1),
        "bwd": bwd,
        "bwd_pct": round(bwd / total * 100.0, 1),
        "lat": lat,
        "lat_pct": round(lat / total * 100.0, 1),
        "pos_count": pos_count,
        "pos_pct": round(pos_pct, 1),
        "high_xt_pct": round(high_xt / total * 100.0, 1),
        "sum_dxt": round(sum_dxt, 3),
        "total_p90": round(total * p90_factor, 1),
        "prog_p90": round(progressive_total * p90_factor, 2),
        "f3_p90": round(to_final_third_success * p90_factor, 2),
        "xt_p90": round(sum_dxt * p90_factor, 3),
        "neg_xt_p90": round(neg_xt * p90_factor, 3),
        "minutes": mins,
        "long_acc_pct": round(long_acc_pct, 1),
        "high_xt_p90": round(high_xt * p90_factor, 2),
        "dz_p90": round(dz_passes * p90_factor, 2),
        "advanced_passes_p90": round(advanced_passes_p90, 1),
        "advanced_accuracy_pct": round(advanced_accuracy_pct, 2),
    }

def compute_defensive_stats(df: pd.DataFrame, match_name: str) -> dict:
    total_actions = len(df)
    if match_name == "All Matches":
        mins = sum(get_match_minutes(k) for k in defensive_dfs_by_match)
    else:
        mins = get_match_minutes(match_name)
    p90_factor = 90.0 / mins if mins > 0 else 1.0

    duels_won = int(df["is_duel_won"].sum())
    duels_lost = int(df["is_duel_lost"].sum())
    total_duels = duels_won + duels_lost
    duels_won_pct = (duels_won / total_duels * 100.0) if total_duels > 0 else 0.0
    interceptions = int(df["is_interception"].sum())

    attacking_half = df[df["is_attacking_half"]]
    actions_attacking = len(attacking_half)
    interceptions_attacking = int(attacking_half["is_interception"].sum())

    # own half
    own_half = df[~df["is_attacking_half"]]
    actions_own = len(own_half)

    # funnel
    funnel_total = int(df["in_funnel"].sum())
    funnel_df = df[df["in_funnel"]]
    funnel_successful = int(funnel_df["is_duel_won"].sum() + funnel_df["is_interception"].sum())
    funnel_success_pct = (funnel_successful / funnel_total * 100.0) if funnel_total > 0 else 0.0

    return {
        "total_actions": total_actions,
        "total_actions_p90": round(total_actions * p90_factor, 1),
        "actions_own": actions_own,
        "actions_own_p90": round(actions_own * p90_factor, 1),
        "actions_attacking": actions_attacking,
        "actions_attacking_p90": round(actions_attacking * p90_factor, 1),
        "total_duels": total_duels,
        "duels_p90": round(total_duels * p90_factor, 1),
        "duels_won_pct": round(duels_won_pct, 1),
        "duels_won": duels_won,
        "interceptions": interceptions,
        "interceptions_p90": round(interceptions * p90_factor, 1),
        "interceptions_attacking": interceptions_attacking,
        "interceptions_attacking_p90": round(interceptions_attacking * p90_factor, 1),
        "funnel_actions": funnel_total,
        "funnel_actions_p90": round(funnel_total * p90_factor, 1),
        "funnel_success_pct": round(funnel_success_pct, 1),
    }

# UI HELPERS
def _safe_pct_diff(a: float, b: float) -> float:
    base = max(abs(b), 1.0)
    pct = (abs(a - b) / base) * 100.0
    return min(pct, 999.0)

def _arrow_html(val_game: float, val_avg: float) -> str:
    if np.isclose(val_game, val_avg, atol=1e-9):
        return ""
    if abs(val_game) < 1 and abs(val_avg) < 1:
        return ""
    if val_game > val_avg:
        pct = _safe_pct_diff(val_game, val_avg)
        return f'<span class="arrow up"> ▲ +{pct:.0f}%</span>'
    else:
        pct = _safe_pct_diff(val_avg, val_game)
        return f'<span class="arrow down"> ▼ -{pct:.0f}%</span>'

def section_card(title, border_color, items):
    bg = _hex_to_rgba(border_color, 0.55)
    bd = _hex_to_rgba(border_color, 0.30)
    html = f'<div class="metric-card" style="--accent-color:{border_color};">'
    html += f'<div class="label">{title}</div>'
    html += f'<div class="stat-spacer"></div>'
    for idx, item in enumerate(items):
        label = item[0]; value = item[1]; sub = item[2] if len(item) > 2 else ""; tooltip = item[3] if len(item) > 3 else ""
        is_last = idx == len(items) - 1
        sep = "" if is_last else 'style="border-bottom:1px solid rgba(255,255,255,0.06);padding-bottom:6px;margin-bottom:6px"'
        html += f'<div {sep}>'
        if tooltip:
            label_html = f'<span class="custom-tooltip">{label}<span class="tooltip-text">{tooltip}</span></span>'
            html += f'<div class="label">{label_html}</div><div class="value">{value}</div>'
        else:
            html += f'<div class="label">{label}</div><div class="value">{value}</div>'
        if sub:
            html += f'<div class="sub">{sub}</div>'
        html += '</div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

def cmp_section_card(title, border_color, items):
    bg = _hex_to_rgba(border_color, 0.55); bd = _hex_to_rgba(border_color, 0.30)
    html = f'<div class="metric-card" style="--accent-color:{border_color};">'
    html += f'<div class="label">{title}</div>'
    for idx, item in enumerate(items):
        label=item[0]; val_game=item[1]; val_avg=item[2]
        disp_game=item[3] if len(item)>3 else str(val_game)
        disp_avg=item[4] if len(item)>4 else str(val_avg)
        tooltip=item[5] if len(item)>5 else ""
        sub=item[6] if len(item)>6 else ""
        is_last=idx==len(items)-1
        sep="" if is_last else 'style="border-bottom:1px solid rgba(255,255,255,0.06);padding-bottom:6px;margin-bottom:6px"'
        html+=f'<div {sep}>'
        if tooltip:
            label_html=f'<span class="custom-tooltip">{label}<span class="tooltip-text">{tooltip}</span></span>'
            html+=f'<div class="label">{label_html}</div><div class="value">{disp_game}{_arrow_html(float(val_game),float(val_avg))}</div>'
        else:
            html+=f'<div class="label">{label}</div><div class="value">{disp_game}{_arrow_html(float(val_game),float(val_avg))}</div>'
        html+=f'<div class="avg-label">AVG: {disp_avg}</div>'
        if sub:
            html+=f'<div class="sub">{sub}</div>'
        html+='</div>'
    html+='</div>'
    st.markdown(html, unsafe_allow_html=True)

# PDF EXPORT FUNCTION
def _pdf_add_footer(fig, page_num, total_pages):
    ax = fig.add_axes([0.06, 0.01, 0.88, 0.025], zorder=999)
    ax.axis("off")
    ax.text(0, 0.5, "Hudson Cicala — 2026 Season", ha="left", va="center",
            fontsize=7, color=PDF_TEXT_DIM, transform=ax.transAxes)
    ax.text(1, 0.5, f"{page_num}/{total_pages}", ha="right", va="center",
            fontsize=7, color=PDF_TEXT_DIM, transform=ax.transAxes)
    ax.plot([0, 1], [1, 1], color="#3a3a5c", linewidth=0.4, transform=ax.transAxes)

def export_dashboard_pdf(passes_images, def_images):
    """Generate a 2-page PDF with passes and defensive screenshots"""
    total_pages = 2
    buf = BytesIO()
    with PdfPages(buf) as pdf:
        # ── PAGE 1: PASSES ──
        fig = plt.figure(figsize=(11, 8.5), facecolor=PDF_BG)
        fig.suptitle("Passes Analysis", fontsize=20, fontweight=700,
                     color=PDF_TEXT_WHITE, y=0.97, x=0.06, ha="left")
        labels_p = ["Pass Map", "Zone Heatmap (Destination)", "Top 10 Pass Impact"]
        for i, img in enumerate(passes_images):
            left = 0.03 + i * 0.33
            width = 0.31
            ax_img = fig.add_axes([left, 0.12, width, 0.78])
            ax_img.imshow(img)
            ax_img.axis("off")
            ax_img.text(0.5, 1.01, labels_p[i], ha="center", va="bottom",
                        fontsize=9, fontweight=600, color=PDF_TEXT_LIGHT,
                        transform=ax_img.transAxes)
        _pdf_add_footer(fig, 1, total_pages)
        pdf.savefig(fig, facecolor=PDF_BG, bbox_inches="tight")
        plt.close(fig)

        # ── PAGE 2: DEFENSIVE ──
        fig = plt.figure(figsize=(11, 8.5), facecolor=PDF_BG)
        fig.suptitle("Defensive Actions", fontsize=20, fontweight=700,
                     color=PDF_TEXT_WHITE, y=0.97, x=0.06, ha="left")
        labels_d = ["Defensive Actions Map", "Defensive Heatmap", "Funnel Protection Actions"]
        for i, img in enumerate(def_images):
            left = 0.03 + i * 0.33
            width = 0.31
            ax_img = fig.add_axes([left, 0.12, width, 0.78])
            ax_img.imshow(img)
            ax_img.axis("off")
            ax_img.text(0.5, 1.01, labels_d[i], ha="center", va="bottom",
                        fontsize=9, fontweight=600, color=PDF_TEXT_LIGHT,
                        transform=ax_img.transAxes)
        _pdf_add_footer(fig, 2, total_pages)
        pdf.savefig(fig, facecolor=PDF_BG, bbox_inches="tight")
        plt.close(fig)

    buf.seek(0)
    return buf

# DRAW HELPERS (PITCH)
def _base_pitch(bg="#1a1a2e"):
    pitch = Pitch(pitch_type="statsbomb", pitch_color=bg, line_color="#ffffff", line_alpha=0.95)
    fig, ax = pitch.draw(figsize=(FIG_W, FIG_H))
    fig.set_facecolor(bg); fig.set_dpi(FIG_DPI)
    ax.axvline(x=FINAL_THIRD_LINE_X, color="#ffffff", lw=1.2, alpha=0.40, linestyle="--")
    ax.axvline(x=HALF_LINE_X, color="#ffffff", lw=0.7, alpha=0.12, linestyle="--")
    return fig, ax, pitch

def _attack_arrow(fig, has_cbar=False):
    ox = -0.04 if has_cbar else 0.0
    fig.patches.append(FancyArrowPatch(
        (0.44 + ox, 0.045), (0.56 + ox, 0.045),
        transform=fig.transFigure, arrowstyle="-|>",
        mutation_scale=11, linewidth=1.6, color="#aaaaaa"))
    fig.text(0.50 + ox, 0.012, "Attacking Direction", ha="center", va="bottom",
             transform=fig.transFigure, fontsize=7.5, color="#aaaaaa")

def _save_fig(fig):
    fig.canvas.draw(); buf = BytesIO()
    fig.savefig(buf, format="png", dpi=FIG_DPI, facecolor=fig.get_facecolor(), bbox_inches="tight")
    buf.seek(0); return Image.open(buf)

def draw_pass_map(df):
    fig, ax, pitch = _base_pitch()
    for _, row in df.iterrows():
        is_lost = not row["is_won"]; is_prog = bool(row["progressive"])
        if is_lost:
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

    leg = ax.legend(handles=[
        Line2D([0],[0],color=COLOR_SUCCESS,lw=2.0,label="Completed",alpha=0.65),
        Line2D([0],[0],color=COLOR_PROGRESSIVE,lw=2.0,label="Progressive",alpha=0.90),
        Line2D([0],[0],color=COLOR_FAIL,lw=2.0,label="Incomplete",alpha=0.90)
    ],loc="upper left",bbox_to_anchor=(0.01,0.99),frameon=True,
        facecolor="#1a1a2e",edgecolor="#444466",fontsize=6.5,
        labelspacing=0.35,borderpad=0.4)
    for t in leg.get_texts(): t.set_color("white")
    leg.get_frame().set_alpha(0.90)
    _attack_arrow(fig)
    return _save_fig(fig), fig

def draw_corridor_heatmap(df):
    df_s = df[df["is_won"]].copy()
    x_bins = np.linspace(0.0, FIELD_X, 7)
    corridors = {
        "left": (LANE_LEFT_MIN, FIELD_Y),
        "center": (LANE_RIGHT_MAX, LANE_LEFT_MIN),
        "right": (0.0, LANE_RIGHT_MAX)
    }
    counts = {}
    for cname, (y0, y1) in corridors.items():
        arr = np.zeros(6, dtype=int)
        for i in range(6):
            x0, x1_ = x_bins[i], x_bins[i + 1]
            arr[i] = int(((df_s["x_end"] >= x0) & (df_s["x_end"] < x1_) &
                          (df_s["y_end"] >= y0) & (df_s["y_end"] < y1)).sum())
        counts[cname] = arr

    all_vals = np.concatenate([counts[c] for c in counts])
    vmax = max(1, int(all_vals.max()))
    cmap = LinearSegmentedColormap.from_list(
        "wr", ["#ffffff", "#ffecec", "#ffbfbf", "#ff8080", "#ff3b3b", "#ff0000"]
    )
    norm = Normalize(vmin=0, vmax=vmax)
    threshold = max(1, vmax * 0.35)

    fig, ax, pitch = _base_pitch()
    for cname, (y0, y1) in corridors.items():
        for i in range(6):
            x0_, x1_ = x_bins[i], x_bins[i + 1]
            value = counts[cname][i]
            ax.add_patch(Rectangle(
                (x0_, y0), x1_ - x0_, y1 - y0,
                facecolor=cmap(norm(value)),
                edgecolor=(1,1,1,0.12), lw=0.5, alpha=0.95, zorder=2))
            ax.text((x0_+x1_)/2, (y0+y1)/2, str(value), ha="center", va="center",
                    color="#000000" if value <= threshold else "#ffffff",
                    fontsize=9, fontweight="700" if value>=vmax*0.5 else "600", zorder=4)

    ax.axhline(y=LANE_LEFT_MIN, color="#ffffff", lw=0.5, alpha=0.15, linestyle="--", zorder=3)
    ax.axhline(y=LANE_RIGHT_MAX, color="#ffffff", lw=0.5, alpha=0.15, linestyle="--", zorder=3)
    _attack_arrow(fig)
    return _save_fig(fig), fig

def _draw_comet_arrow(ax, x0, y0, x1, y1, color):
    segs = 12; ts = np.linspace(0.0, 1.0, segs + 1)
    for i in range(segs):
        t0, t1 = ts[i], ts[i+1]
        xa = x0 + (x1 - x0) * t0
        ya = y0 + (y1 - y0) * t0
        xb = x0 + (x1 - x0) * t1
        yb = y0 + (y1 - y0) * t1
        alpha = 0.85 * (0.15 + 0.85 * t1)
        lw = 2.5 * (0.80 + 0.20 * t1)
        ax.plot([xa, xb], [ya, yb], color=color, linewidth=lw, alpha=alpha,
                zorder=4, solid_capstyle="round")
    ax.scatter(x0, y0, s=20, marker="o", facecolors="none", edgecolors=color,
               linewidths=1.5, zorder=5, alpha=0.85)
    ax.scatter(x1, y1, s=32, marker="o", facecolors=color, edgecolors="white",
               linewidths=0.9, zorder=6, alpha=0.85)

def draw_top_xt_map(df, top_n=5):
    fig, ax, pitch = _base_pitch()
    top_passes = (df[(df["is_won"]) & (df["delta_xt_adj"] > 0)]
                  .sort_values("delta_xt_adj", ascending=False)
                  .head(top_n).copy().reset_index(drop=True))
    if not top_passes.empty:
        for _, row in top_passes.iterrows():
            val = float(row["delta_xt_adj"])
            color = CMAP_TOP10(NORM_TOP10(np.clip(val, 0.05, 0.40)))
            _draw_comet_arrow(ax, float(row["x_start"]), float(row["y_start"]),
                              float(row["x_end"]), float(row["y_end"]), color)
    sm = plt.cm.ScalarMappable(cmap=CMAP_TOP10, norm=NORM_TOP10)
    cbar = fig.colorbar(sm, ax=ax, fraction=0.020, pad=0.02, shrink=0.60)
    cbar.set_label("Pass Impact", color="#ffffff", fontsize=8)
    cbar.ax.yaxis.set_tick_params(color="#ffffff", labelsize=7)
    plt.setp(plt.getp(cbar.ax.axes, "yticklabels"), color="#ffffff")
    _attack_arrow(fig, has_cbar=True)
    return _save_fig(fig), fig

# DEFENSIVE PITCH DRAW HELPERS
COLOR_DUEL_WON = "#10b981"
COLOR_DUEL_LOST = "#E07070"
COLOR_INTERCEPTION = "#2F80ED"

def draw_defensive_map(df):
    fig, ax, pitch = _base_pitch()
    for _, row in df.iterrows():
        if row["is_duel_won"]:
            color, marker, s, alpha = COLOR_DUEL_WON, "o", 90, 0.85
        elif row["is_duel_lost"]:
            color, marker, s, alpha = COLOR_DUEL_LOST, "X", 100, 0.85
        else:
            color, marker, s, alpha = COLOR_INTERCEPTION, "^", 80, 0.85
        pitch.scatter(row["x"], row["y"], s=s, marker=marker, color=color,
                      edgecolors="white", linewidths=0.8, ax=ax, zorder=6, alpha=alpha)

    leg = ax.legend(handles=[
        Line2D([0],[0],marker="o",color="w",markerfacecolor=COLOR_DUEL_WON,markersize=7,label="Duel Won",alpha=0.90),
        Line2D([0],[0],marker="X",color="w",markerfacecolor=COLOR_DUEL_LOST,markersize=8,label="Duel Lost",alpha=0.90),
        Line2D([0],[0],marker="^",color="w",markerfacecolor=COLOR_INTERCEPTION,markersize=7,label="Interception",alpha=0.90)
    ],loc="upper left",bbox_to_anchor=(0.01,0.99),frameon=True,
        facecolor="#1a1a2e",edgecolor="#444466",fontsize=6.5,
        labelspacing=0.35,borderpad=0.4)
    for t in leg.get_texts(): t.set_color("white")
    leg.get_frame().set_alpha(0.90)
    _attack_arrow(fig)
    return _save_fig(fig), fig

def draw_funnel_protection_map(df):
    fig, ax, pitch = _base_pitch()
    funnel_rect = Rectangle(
        (0, PENALTY_AREA_Y_MIN), FUNNEL_X_EXTEND,
        PENALTY_AREA_Y_MAX - PENALTY_AREA_Y_MIN,
        facecolor="#ffd700", edgecolor="#ffd700", lw=1.5,
        linestyle="--", alpha=0.12, zorder=2)
    ax.add_patch(funnel_rect)

    for _, row in df.iterrows():
        x, y = float(row["x"]), float(row["y"])
        in_funnel = bool(row.get("in_funnel", is_in_funnel_zone(x, y)))
        if in_funnel:
            marker, s, color, edge = "*", 120, "#ffd700", "#b8860b"
        else:
            marker, s, color, edge = "o", 60, "#888888", "#555555"
        pitch.scatter(x, y, s=s, marker=marker, color=color,
                      edgecolors=edge, linewidths=0.5, ax=ax, zorder=6, alpha=0.85)

    leg = ax.legend(handles=[
        Line2D([0],[0],marker="*",color="w",markerfacecolor="#ffd700",markersize=9,label="Funnel Action",alpha=0.95),
        Line2D([0],[0],marker="o",color="w",markerfacecolor="#888888",markersize=6,label="Other Action",alpha=0.50)
    ],loc="upper left",bbox_to_anchor=(0.01,0.99),frameon=True,
        facecolor="#1a1a2e",edgecolor="#444466",fontsize=6.5,
        labelspacing=0.35,borderpad=0.4)
    for t in leg.get_texts(): t.set_color("white")
    leg.get_frame().set_alpha(0.90)
    _attack_arrow(fig)
    return _save_fig(fig), fig

def draw_defensive_heatmap(df):
    corridors = {
        "Right": (LANE_LEFT_MIN, FIELD_Y),
        "Center": (LANE_RIGHT_MAX, LANE_LEFT_MIN),
        "Left": (0.0, LANE_RIGHT_MAX)
    }
    corridor_data = {}
    for cname, (y0, y1) in corridors.items():
        mask = (df["y"] >= y0) & (df["y"] < y1)
        corridor_data[cname] = {
            "total": int(mask.sum()),
            "duels_won": int(df.loc[mask, "is_duel_won"].sum()),
            "duels_lost": int(df.loc[mask, "is_duel_lost"].sum())
        }

    all_vals = np.array([corridor_data[c]["total"] for c in corridors])
    vmax = max(1, int(all_vals.max()))
    cmap = LinearSegmentedColormap.from_list(
        "wb", ["#ffffff", "#dbeafe", "#93c5fd", "#60a5fa", "#3b82f6", "#1d4ed8"]
    )
    norm = Normalize(vmin=0, vmax=vmax)
    threshold = max(1, vmax * 0.35)

    fig, ax, pitch = _base_pitch()
    for cname, (y0, y1) in corridors.items():
        value = corridor_data[cname]["total"]
        ax.add_patch(Rectangle(
            (0, y0), FIELD_X, y1 - y0,
            facecolor=cmap(norm(value)),
            edgecolor=(1, 1, 1, 0.12), lw=0.5, alpha=0.95, zorder=2))

        duel_total = corridor_data[cname]["duels_won"] + corridor_data[cname]["duels_lost"]
        duel_pct = (corridor_data[cname]["duels_won"] / duel_total * 100.0) if duel_total > 0 else None
        label = f"{cname}\nTotal: {value}\nWon: {corridor_data[cname]['duels_won']}/{corridor_data[cname]['duels_won']+corridor_data[cname]['duels_lost']} ({duel_pct:.0f}%)" if duel_pct is not None else f"{cname}\nTotal: {value}"
        ax.text(FIELD_X / 2, (y0 + y1) / 2, label, ha="center", va="center",
                color="#000000" if value <= threshold else "#ffffff",
                fontsize=9, fontweight="600", zorder=4)

    ax.axhline(y=LANE_LEFT_MIN, color="#ffffff", lw=0.5, alpha=0.20, linestyle="--", zorder=3)
    ax.axhline(y=LANE_RIGHT_MAX, color="#ffffff", lw=0.5, alpha=0.20, linestyle="--", zorder=3)
    _attack_arrow(fig)
    return _save_fig(fig), fig

# SIDEBAR
st.sidebar.markdown("""
<div style="text-align:center;padding:10px 0;">
    <h1 style="font-size:1.8rem;font-weight:800;background:linear-gradient(135deg,#2F80ED,#10b981);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">
        Pass Stats Dashboard
    </h1>
    <p style="color:#7a7a9a;font-size:0.90rem;margin-top:-4px;">2026 Season &bull; Hudson Cicala</p>
</div>
""", unsafe_allow_html=True)

img_path = "Captura de tela 2026-06-02 154425.png"
if os.path.exists(img_path):
    st.sidebar.image(img_path, use_container_width=True)

st.sidebar.markdown("""
<div style="padding:8px 0;text-align:center;">
    <p style="color:#5a5a7a;font-size:0.75rem;">
        Data collected from match footage
    </p>
</div>
""", unsafe_allow_html=True)

num_matches = len(dfs_by_match)
all_match_stats = [compute_stats(dfs_by_match[m], m) for m in dfs_by_match]

# LAYOUT — SINGLE TAB
tab_dash, = st.tabs(["Detailed Dashboard"])

with tab_dash:
    sub_tab_passes, sub_tab_def = st.tabs(["Passes", "Defensive Actions"])

    with sub_tab_passes:
        st.markdown("### Match Filters")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            pass_match_options = ["All Matches"] + list(dfs_by_match.keys())
            selected_match = st.selectbox(
                "Select Match", options=pass_match_options, index=0, key="pass_match")
        with col_f2:
            pass_filter = st.radio(
                "Pass Type",
                ["All", "Successful", "Unsuccessful", "Progressive", "Final Third"],
                index=0, horizontal=True, key="pass_filter")

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
        img_pm_game, fig_pm_game = draw_pass_map(df_game); plt.close(fig_pm_game)
        img_ht_game, fig_ht_game = draw_corridor_heatmap(df_game); plt.close(fig_ht_game)
        top_n_xt = 10 if force_avg else 5
        img_xt_game, fig_xt_game = draw_top_xt_map(df_game, top_n=top_n_xt); plt.close(fig_xt_game)

        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.markdown('<div class="img-label">⏹️ Pass Map</div>', unsafe_allow_html=True)
            st.image(img_pm_game, use_container_width=True)
        with col_m2:
            st.markdown('<div class="img-label">🗺️ Zone Heatmap (Destination)</div>', unsafe_allow_html=True)
            st.image(img_ht_game, use_container_width=True)
        with col_m3:
            label = "Top 10" if force_avg else "Top 5"
            st.markdown(f'<div class="img-label">⚡ {label} Pass Impact</div>', unsafe_allow_html=True)
            st.image(img_xt_game, use_container_width=True)

        st.markdown(" ", unsafe_allow_html=True)
        col_s1, col_s2, col_s3 = st.columns(3)

        total_impact_value = float(df_game.loc[df_game["is_won"], "delta_xt_adj"].sum())

        if force_avg:
            with col_s1:
                section_card("📋 Pass Overview", C_BLUE_PASTEL, [
                    ("Total Passes (AVG)", f"{s_game['total_p90']:.1f}"),
                    ("% Accuracy", f"{s_game['accuracy_pct']:.1f}%")
                ])
            with col_s2:
                section_card("📊 Advanced", C_GREEN_PASTEL, [
                    ("Advanced Passes (AVG)", f"{s_game['advanced_passes_p90']:.1f}"),
                    ("% Advanced Accuracy", f"{s_game['advanced_accuracy_pct']:.1f}%")
                ])
            with col_s3:
                section_card("⚡ Impact", C_AMBER_PASTEL, [
                    ("Pass Impact Value (AVG)", f"{s_game['xt_p90']:.1f}", f"Total: {total_impact_value:.3f}"),
                    ("% Positive Impact", f"{s_game['pos_pct']:.1f}%")
                ])
            col_s3_exp = st.columns(3)[2]
            with col_s3_exp:
                expl_bg=_hex_to_rgba(C_AMBER_PASTEL,0.35); expl_bd=_hex_to_rgba(C_AMBER_PASTEL,0.20)
                st.markdown(
                    f'<div class="explanation-card" style="--accent-color:{C_AMBER_PASTEL};">'
                    '<div class="explabel">❓ Explanation</div>'
                    '<div class="exptext">'
                    '<b>Pass Impact Value</b> — Calculation used to evaluate the offensive value added by a pass.<br>'
                    '<b>% Positive Impact</b> — Passes that generated a positive impact based on where they ended on the field.'
                    '</div></div>',
                    unsafe_allow_html=True)
        else:
            with col_s1:
                cmp_section_card("📋 Pass Overview", C_BLUE_PASTEL, [
                    ("Total Passes (AVG)", s_game["total_p90"], f"{s_avg['total_p90']:.1f}"),
                    ("% Accuracy", s_game["accuracy_pct"], s_avg["accuracy_pct"],
                     f"{s_game['accuracy_pct']:.1f}%", f"{s_avg['accuracy_pct']:.1f}%")
                ])
            with col_s2:
                cmp_section_card("📊 Advanced", C_GREEN_PASTEL, [
                    ("Advanced Passes (AVG)", s_game["advanced_passes_p90"], f"{s_avg['advanced_passes_p90']:.1f}"),
                    ("% Advanced Accuracy", s_game["advanced_accuracy_pct"], s_avg["advanced_accuracy_pct"],
                     f"{s_game['advanced_accuracy_pct']:.1f}%", f"{s_avg['advanced_accuracy_pct']:.1f}%")
                ])
            with col_s3:
                cmp_section_card("⚡ Impact", C_AMBER_PASTEL, [
                    ("Pass Impact Value (AVG)", s_game["xt_p90"], s_avg["xt_p90"],
                     f"{s_game['xt_p90']:.1f}", f"{s_avg['xt_p90']:.1f}", "",
                     f"Total: {total_impact_value:.3f}"),
                    ("% Positive Impact", s_game["pos_pct"], s_avg["pos_pct"],
                     f"{s_game['pos_pct']:.1f}%", f"{s_avg['pos_pct']:.1f}%")
                ])
            col_s3_exp = st.columns(3)[2]
            with col_s3_exp:
                expl_bg=_hex_to_rgba(C_AMBER_PASTEL,0.35); expl_bd=_hex_to_rgba(C_AMBER_PASTEL,0.20)
                st.markdown(
                    f'<div class="explanation-card" style="--accent-color:{C_AMBER_PASTEL};">'
                    '<div class="explabel">❓ Explanation</div>'
                    '<div class="exptext">'
                    '<b>Pass Impact Value</b> — Calculation used to evaluate the offensive value added by a pass.<br>'
                    '<b>% Positive Impact</b> — Passes that generated a positive impact based on where they ended on the field.'
                    '</div></div>',
                    unsafe_allow_html=True)

    with sub_tab_def:
        st.markdown("### Match Filter")
        col_df1, col_df2 = st.columns(2)
        with col_df1:
            def_match_options = ["All Matches"] + list(defensive_dfs_by_match.keys())
            selected_def_match = st.selectbox(
                "Select Match", options=def_match_options, index=0, key="def_match")
        with col_df2:
            def_type_filter = st.radio(
                "Filter Type", ["All", "Duels Only", "Interceptions Only"],
                horizontal=True, key="def_type_filter")

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
        img_def_map, fig_def_map = draw_defensive_map(df_def_game); plt.close(fig_def_map)
        img_def_hm, fig_def_hm = draw_defensive_heatmap(df_def_game); plt.close(fig_def_hm)
        img_funnel, fig_funnel = draw_funnel_protection_map(df_def_game); plt.close(fig_funnel)

        col_dm1, col_dm2, col_dm3 = st.columns(3)
        with col_dm1:
            st.markdown('<div class="img-label">🛡️ Defensive Actions Map</div>', unsafe_allow_html=True)
            st.image(img_def_map, use_container_width=True)
        with col_dm2:
            st.markdown('<div class="img-label">📊 Defensive Heatmap</div>', unsafe_allow_html=True)
            st.image(img_def_hm, use_container_width=True)
        with col_dm3:
            st.markdown('<div class="img-label">🔒 Funnel Protection Actions</div>', unsafe_allow_html=True)
            st.image(img_funnel, use_container_width=True)

        st.markdown(" ", unsafe_allow_html=True)
        col_ds1, col_ds2, col_ds3 = st.columns(3)

        if force_avg_def:
            with col_ds1:
                section_card("🛡️ General", C_BLUE_PASTEL, [
                    ("Defensive Actions (AVG)", f"{d_game['total_actions_p90']:.1f}"),
                    ("Actions in Own Half (AVG)", f"{d_game['actions_own_p90']:.1f}")
                ])
            with col_ds2:
                section_card("⚔️ Duels and Interceptions", C_GREEN_PASTEL, [
                    ("Defensive Duels (AVG)", f"{d_game['duels_p90']:.1f}"),
                    ("% Duels Won", f"{d_game['duels_won_pct']:.1f}%"),
                    ("Interceptions (AVG)", f"{d_game['interceptions_p90']:.1f}")
                ])
            with col_ds3:
                section_card("🛡️ Funnel Protection", C_AMBER_PASTEL, [
                    ("Funnel Protection Actions (AVG)", f"{d_game['funnel_actions_p90']:.1f}"),
                    ("%FPA Successful", f"{d_game['funnel_success_pct']:.1f}%")
                ])
        else:
            with col_ds1:
                cmp_section_card("🛡️ General", C_BLUE_PASTEL, [
                    ("Defensive Actions (AVG)", d_game["total_actions_p90"], f"{d_avg['total_actions_p90']:.1f}"),
                    ("Actions in Own Half (AVG)", d_game["actions_own_p90"], f"{d_avg['actions_own_p90']:.1f}")
                ])
            with col_ds2:
                cmp_section_card("⚔️ Duels and Interceptions", C_GREEN_PASTEL, [
                    ("Defensive Duels (AVG)", d_game["duels_p90"], f"{d_avg['duels_p90']:.1f}"),
                    ("% Duels Won", d_game["duels_won_pct"], d_avg["duels_won_pct"],
                     f"{d_game['duels_won_pct']:.1f}%", f"{d_avg['duels_won_pct']:.1f}%"),
                    ("Interceptions (AVG)", d_game["interceptions_p90"], d_avg["interceptions_p90"],
                     f"{d_game['interceptions_p90']:.1f}", f"{d_avg['interceptions_p90']:.1f}")
                ])
            with col_ds3:
                cmp_section_card("🛡️ Funnel Protection", C_AMBER_PASTEL, [
                    ("Funnel Protection Actions (AVG)", d_game["funnel_actions_p90"], f"{d_avg['funnel_actions_p90']:.1f}"),
                    ("%FPA Successful", d_game["funnel_success_pct"], d_avg["funnel_success_pct"],
                     f"{d_game['funnel_success_pct']:.1f}%", f"{d_avg['funnel_success_pct']:.1f}%")
                ])

        # ── PDF EXPORT SECTION ──
        st.markdown("---")
        st.markdown("### 📄 Export Complete Dashboard")
        if st.button("📸 Download Screenshot (PDF) — Passes + Defensive Actions", use_container_width=True):
            with st.spinner("Generating PDF with dashboard screenshots..."):
                df_all_passes = pd.concat(dfs_by_match.values(), ignore_index=True)
                df_all_def = pd.concat(defensive_dfs_by_match.values(), ignore_index=True)

                img_pm_all, _ = draw_pass_map(df_all_passes); plt.close()
                img_ht_all, _ = draw_corridor_heatmap(df_all_passes); plt.close()
                img_xt_all, _ = draw_top_xt_map(df_all_passes, top_n=10); plt.close()
                img_dm_all, _ = draw_defensive_map(df_all_def); plt.close()
                img_dhm_all, _ = draw_defensive_heatmap(df_all_def); plt.close()
                img_fn_all, _ = draw_funnel_protection_map(df_all_def); plt.close()

                pdf_bytes = export_dashboard_pdf(
                    [img_pm_all, img_ht_all, img_xt_all],
                    [img_dm_all, img_dhm_all, img_fn_all]
                )
                st.download_button(
                    "📥 Save PDF",
                    data=pdf_bytes,
                    file_name="hudson_cicala_dashboard.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
