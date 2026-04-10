"""
Energy Management System (EMS) Dashboard
Run: streamlit run ems_dashboard.py
Dependencies: pip install streamlit plotly pandas openpyxl
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io

st.set_page_config(page_title="EMS Dashboard", layout="wide", page_icon="⚡")

# ─────────────────────────────────────────────
# EMS RULES ENGINE
# ─────────────────────────────────────────────
# Rules template (from Rules template sheet):
#   Rule No | RE | Batt_L | Batt_M | Batt_H | Grid | Grid_T_H | Grid_T_L | DG | TOD1 | TOD2 | TOD3 | TOD4
#   Outputs: Load source, Batt mode, Grid mode
#
# Columns mapping:
#   RE: 1 if RE >= Load (RE > Load), else 0
#   Batt_L: SOC 0–30%,  Batt_M: SOC 30–70%,  Batt_H: SOC 70–100%
#   Grid: 1=available, Grid_T_H: tariff high=1, Grid_T_L: tariff low=1
#   TOD1: 00:00-06:00, TOD2: 06:00-12:00, TOD3: 12:00-18:00, TOD4: 18:00-24:00

RULES = [
    # RE=0, Batt_L=1 (SOC 0-30%), Grid=0, DG=0
    {"rule": 1,  "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": 1, "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "No Export", "note": "Use RE with reduced load"},
    {"rule": 2,  "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "No Export", "note": "Use RE with reduced load"},
    {"rule": 3,  "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "No Export", "note": "Use RE with reduced load"},
    {"rule": 4,  "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1, "Load": "RE",      "Batt": "Idle",      "Grid_out": "No Export", "note": "Use RE with reduced load"},
    # RE=0, Batt_L=1, Grid=1, High Tariff
    {"rule": 5,  "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "No Export", "note": "Use RE with reduced load"},
    {"rule": 6,  "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "No Export", "note": "Use RE with reduced load"},
    {"rule": 7,  "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "No Export", "note": "Use RE with reduced load"},
    {"rule": 8,  "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Idle",      "Grid_out": "No Export", "note": "Use RE with reduced load"},
    # RE=0, Batt_L=1, Grid=1, Low Tariff
    {"rule": 9,  "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "Grid",    "Batt": "Charge",    "Grid_out": "No Export", "note": "Grid primary, charge batt"},
    {"rule": 10, "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "Grid",    "Batt": "Charge",    "Grid_out": "Export",    "note": "RE Export"},
    {"rule": 11, "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "Grid",    "Batt": "Charge",    "Grid_out": "Export",    "note": "RE Export"},
    {"rule": 12, "RE": 0, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "Grid",    "Batt": "Charge",    "Grid_out": "No Export", "note": "Grid primary"},
    # RE=0, Batt_M=1 (SOC 30-70%), Grid=0
    {"rule": 13, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Batt till 30%"},
    {"rule": 14, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Batt till 30%"},
    {"rule": 15, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Batt till 30%"},
    {"rule": 16, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Batt till 30%"},
    # RE=0, Batt_M=1, Grid=1, High Tariff
    {"rule": 17, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Batt till 30%, peak tariff avoid grid"},
    {"rule": 18, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Batt till 30%"},
    {"rule": 19, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Batt till 30%"},
    {"rule": 20, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Batt till 30%, peak"},
    # RE=0, Batt_M=1, Grid=1, Low Tariff
    {"rule": 21, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "Grid",    "Batt": "Charge",    "Grid_out": "No Export", "note": "Grid primary"},
    {"rule": 22, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "Grid",    "Batt": "Charge",    "Grid_out": "Export",    "note": "Grid + RE Export"},
    {"rule": 23, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "Grid",    "Batt": "Charge",    "Grid_out": "Export",    "note": "Grid + RE Export"},
    {"rule": 24, "RE": 0, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "Grid",    "Batt": "Charge",    "Grid_out": "No Export", "note": "Grid primary"},
    # RE=0, Batt_H=1 (SOC 70-100%), Grid=0
    {"rule": 25, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Battery primary"},
    {"rule": 26, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "Export",    "note": "Battery + RE Export"},
    {"rule": 27, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "Export",    "note": "Battery + RE Export"},
    {"rule": 28, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Battery primary"},
    # RE=0, Batt_H=1, Grid=1, High Tariff
    {"rule": 29, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Battery primary, peak tariff"},
    {"rule": 30, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "Export",    "note": "Battery + RE Export"},
    {"rule": 31, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "Battery", "Batt": "Discharge", "Grid_out": "Export",    "note": "Battery + RE Export"},
    {"rule": 32, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "Battery", "Batt": "Discharge", "Grid_out": "No Export", "note": "Battery primary"},
    # RE=0, Batt_H=1, Grid=1, Low Tariff
    {"rule": 33, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "Grid",    "Batt": "Charge",    "Grid_out": "No Export", "note": "Grid primary, charge from RE"},
    {"rule": 34, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "Grid",    "Batt": "Charge",    "Grid_out": "No Export", "note": "Grid primary"},
    {"rule": 35, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "Grid",    "Batt": "Charge",    "Grid_out": "No Export", "note": "Grid primary"},
    {"rule": 36, "RE": 0, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "Grid",    "Batt": "Charge",    "Grid_out": "No Export", "note": "Grid primary"},
    # RE=1, Batt_L=1, Grid=0
    {"rule": 37, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "No Export", "note": "RE primary, charge batt"},
    {"rule": 38, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "No Export", "note": "RE primary, charge batt"},
    {"rule": 39, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "No Export", "note": "RE primary, charge batt"},
    {"rule": 40, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Charge",    "Grid_out": "No Export", "note": "RE primary, charge batt"},
    # RE=1, Batt_L=1, Grid=1, High Tariff
    {"rule": 41, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE primary + export"},
    {"rule": 42, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE primary + export"},
    {"rule": 43, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE primary + export"},
    {"rule": 44, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE primary + export"},
    # RE=1, Batt_L=1, Grid=1, Low Tariff
    {"rule": 45, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "Battery Charge - Grid"},
    {"rule": 46, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "Battery Charge - Grid"},
    {"rule": 47, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "Battery Charge - Grid"},
    {"rule": 48, "RE": 1, "BL": 1, "BM": 0, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "Battery Charge - Grid"},
    # RE=1, Batt_M=1, Grid=0
    {"rule": 49, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "No Export", "note": "RE primary"},
    {"rule": 50, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "No Export", "note": "RE primary"},
    {"rule": 51, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "No Export", "note": "RE primary"},
    {"rule": 52, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Charge",    "Grid_out": "No Export", "note": "RE primary"},
    # RE=1, Batt_M=1, Grid=1, High Tariff
    {"rule": 53, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE + Export"},
    {"rule": 54, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE + Export"},
    {"rule": 55, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE + Export"},
    {"rule": 56, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE + Export"},
    # RE=1, Batt_M=1, Grid=1, Low Tariff
    {"rule": 57, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE + Export, Batt Grid"},
    {"rule": 58, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE + Export"},
    {"rule": 59, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE + Export"},
    {"rule": 60, "RE": 1, "BL": 0, "BM": 1, "BH": 0, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Charge",    "Grid_out": "Export",    "note": "RE + Export"},
    # RE=1, Batt_H=1, Grid=0
    {"rule": 61, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE primary, Batt full"},
    {"rule": 62, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE primary + export"},
    {"rule": 63, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE primary + export"},
    {"rule": 64, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 0, "GTH": None, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE primary + export"},
    # RE=1, Batt_H=1, Grid=1, High Tariff
    {"rule": 65, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE + export"},
    {"rule": 66, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE + export"},
    {"rule": 67, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE + export"},
    {"rule": 68, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": 1, "GTL": None, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE + export"},
    # RE=1, Batt_H=1, Grid=1, Low Tariff
    {"rule": 69, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": 1,    "TOD2": None, "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE + export"},
    {"rule": 70, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": 1,    "TOD3": None, "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE + export"},
    {"rule": 71, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": 1,    "TOD4": None, "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE + export"},
    {"rule": 72, "RE": 1, "BL": 0, "BM": 0, "BH": 1, "Grid": 1, "GTH": None, "GTL": 1, "DG": 0, "TOD1": None, "TOD2": None, "TOD3": None, "TOD4": 1,    "Load": "RE",      "Batt": "Idle",      "Grid_out": "Export",    "note": "RE + export"},
]


def get_tod(time_str):
    """Return TOD1-TOD4 flags from HH:MM:SS string."""
    h = int(str(time_str).split(":")[0])
    tod1 = 1 if (0 <= h < 6) else 0
    tod2 = 1 if (6 <= h < 12) else 0
    tod3 = 1 if (12 <= h < 18) else 0
    tod4 = 1 if (18 <= h < 24) else 0
    return tod1, tod2, tod3, tod4


def get_batt_range(soc):
    bl = 1 if soc < 30 else 0
    bm = 1 if (30 <= soc < 70) else 0
    bh = 1 if soc >= 70 else 0
    return bl, bm, bh


def match_rule(row_re, row_bl, row_bm, row_bh, row_grid, row_gth, row_gtl, row_dg, row_tod1, row_tod2, row_tod3, row_tod4):
    """Find the first matching rule and return Load/Batt/Grid outputs."""
    for r in RULES:
        if r["RE"] != row_re:
            continue
        if r["BL"] != row_bl:
            continue
        if r["BM"] != row_bm:
            continue
        if r["BH"] != row_bh:
            continue
        if r["Grid"] != row_grid:
            continue
        # Grid Tariff
        if r["GTH"] is not None and r["GTH"] != row_gth:
            continue
        if r["GTL"] is not None and r["GTL"] != row_gtl:
            continue
        # DG
        if r["DG"] != row_dg:
            continue
        # TOD: at least one TOD must match
        tod_match = False
        for key, val in [("TOD1", row_tod1), ("TOD2", row_tod2), ("TOD3", row_tod3), ("TOD4", row_tod4)]:
            if r[key] == 1 and val == 1:
                tod_match = True
                break
        if not tod_match:
            continue
        return r["Load"], r["Batt"], r["Grid_out"], r["rule"], r["note"]
    return "Grid", "Idle", "No Export", 0, "Fallback: Grid"


def compute_power(row_load_kw, row_re_kw, load_src, batt_mode, grid_out):
    """
    Compute numeric Load (kW), Batt (kW +charge/-discharge), Grid (kW) values.
    Sign convention:
      Load_kW  = actual kW drawn by load (always positive)
      Batt_kW  = positive → charging, negative → discharging
      Grid_kW  = positive → import, negative → export
    """
    load_kw = row_load_kw
    re_kw = row_re_kw

    if load_src == "RE":
        # RE covers the load; excess charges battery or exports
        excess = re_kw - load_kw
        if batt_mode == "Charge":
            batt_kw = max(0, excess)       # charge with excess
            grid_kw = -max(0, excess - batt_kw) if grid_out == "Export" else 0
        elif batt_mode == "Idle":
            batt_kw = 0
            grid_kw = -excess if (grid_out == "Export" and excess > 0) else 0
        else:
            batt_kw = 0
            grid_kw = 0
        # if RE < load, draw shortfall from grid
        shortfall = load_kw - re_kw
        if shortfall > 0:
            grid_kw = shortfall

    elif load_src == "Battery":
        batt_kw = -load_kw  # discharging
        re_excess = re_kw
        grid_kw = -re_excess if (grid_out == "Export" and re_excess > 0) else 0

    elif load_src == "Grid":
        # Grid primary; RE goes to battery or export
        grid_kw = load_kw
        if batt_mode == "Charge":
            batt_kw = re_kw  # RE charges battery
        else:
            batt_kw = 0
        if grid_out == "Export" and re_kw > 0:
            grid_kw -= re_kw   # offset import with RE export

    else:
        batt_kw = 0
        grid_kw = load_kw

    return round(load_kw, 3), round(batt_kw, 3), round(grid_kw, 3)


def process_dataframe(df):
    """Apply EMS logic to the input dataframe and return results."""
    results = []
    for _, row in df.iterrows():
        time_str = str(row["Time"])
        soc = float(row["Battery SOC (%)"])
        re_kw = float(row["RE (kW)"])
        load_kw = float(row["Load (kW)"])
        grid_avail = int(row["Grid Available"])
        tariff = int(row["Tariff"])

        tod1, tod2, tod3, tod4 = get_tod(time_str)
        bl, bm, bh = get_batt_range(soc)

        re_flag = 1 if re_kw >= load_kw else 0
        gth = 1 if tariff == 1 else 0
        gtl = 1 if tariff == 0 else 0
        dg = 0  # DG not in sample data

        load_src, batt_mode, grid_out, rule_no, note = match_rule(
            re_flag, bl, bm, bh, grid_avail, gth, gtl, dg,
            tod1, tod2, tod3, tod4
        )
        load_val, batt_val, grid_val = compute_power(load_kw, re_kw, load_src, batt_mode, grid_out)

        results.append({
            "Time": time_str,
            "Battery SOC (%)": soc,
            "RE (kW)": re_kw,
            "Load (kW)": load_kw,
            "Grid Available": grid_avail,
            "Tariff": tariff,
            "Rule No": rule_no,
            "Load Source": load_src,
            "Batt Mode": batt_mode,
            "Grid Mode": grid_out,
            "Load_out (kW)": load_val,
            "Batt_out (kW)": batt_val,
            "Grid_out (kW)": grid_val,
            "Note": note,
        })
    return pd.DataFrame(results)


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────
st.title("⚡ Energy Management System Dashboard")
st.markdown("Upload time-series data or edit the table below. The EMS logic automatically computes **Load**, **Battery**, and **Grid** dispatch values.")

# Sidebar
st.sidebar.header("📋 Input Options")
input_mode = st.sidebar.radio("Input Mode", ["Upload Excel / CSV", "Manual Entry", "Use Sample Data"])

REQUIRED_COLS = ["Time", "Battery SOC (%)", "RE (kW)", "Load (kW)", "Grid Available", "Tariff"]

sample_data = {
    "Time":            ["00:00:00","06:10:00","08:30:00","10:00:00","12:40:00","17:00:00","18:10:00","20:10:00","23:00:00"],
    "Battery SOC (%)": [65, 65, 71, 81.8, 100, 98.5, 88, 71.5, 71.5],
    "RE (kW)":         [0, 0.22, 3.04, 4.33, 4.92, 1.29, 0, 0, 0],
    "Load (kW)":       [0.95, 1.83, 2.28, 1.94, 3.21, 4.29, 4.89, 4.11, 2.24],
    "Grid Available":  [1, 1, 1, 1, 1, 0, 0, 1, 1],
    "Tariff":          [0, 0, 0, 1, 1, 0, 1, 0, 0],
}

df_input = None

if input_mode == "Upload Excel / CSV":
    uploaded = st.sidebar.file_uploader("Upload file", type=["xlsx", "csv"])
    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df_input = pd.read_csv(uploaded)
            else:
                df_input = pd.read_excel(uploaded)
            # Try to auto-detect columns
            df_input.columns = [c.strip() for c in df_input.columns]
            missing = [c for c in REQUIRED_COLS if c not in df_input.columns]
            if missing:
                st.error(f"Missing columns: {missing}. Required: {REQUIRED_COLS}")
                df_input = None
        except Exception as e:
            st.error(f"Error reading file: {e}")

elif input_mode == "Use Sample Data":
    df_input = pd.DataFrame(sample_data)

else:  # Manual Entry
    st.subheader("✏️ Manual Data Entry")
    n_rows = st.sidebar.slider("Number of rows", 1, 50, 5)
    default_df = pd.DataFrame({c: [""] * n_rows for c in REQUIRED_COLS})
    df_input = st.data_editor(
        default_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Time": st.column_config.TextColumn("Time (HH:MM:SS)"),
            "Battery SOC (%)": st.column_config.NumberColumn("Battery SOC (%)", min_value=0, max_value=100),
            "RE (kW)": st.column_config.NumberColumn("RE (kW)", min_value=0.0),
            "Load (kW)": st.column_config.NumberColumn("Load (kW)", min_value=0.0),
            "Grid Available": st.column_config.SelectboxColumn("Grid Available", options=[0, 1]),
            "Tariff": st.column_config.SelectboxColumn("Tariff (1=High)", options=[0, 1]),
        }
    )
    try:
        df_input = df_input.dropna(subset=["Time"])
        for col in ["Battery SOC (%)", "RE (kW)", "Load (kW)", "Grid Available", "Tariff"]:
            df_input[col] = pd.to_numeric(df_input[col], errors="coerce")
        df_input = df_input.dropna()
    except Exception:
        df_input = None

# ─────────────────────────────────────────────
# PROCESS & DISPLAY
# ─────────────────────────────────────────────
if df_input is not None and not df_input.empty:
    with st.spinner("Running EMS logic..."):
        df_result = process_dataframe(df_input)

    st.success(f"✅ Processed {len(df_result)} rows")

    # KPI summary
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Load (kWh)", f"{df_result['Load_out (kW)'].sum():.2f}")
    col2.metric("RE Utilization (kWh)", f"{df_result['RE (kW)'].sum():.2f}")
    col3.metric("Grid Import (kWh)", f"{df_result[df_result['Grid_out (kW)'] > 0]['Grid_out (kW)'].sum():.2f}")
    col4.metric("Grid Export (kWh)", f"{abs(df_result[df_result['Grid_out (kW)'] < 0]['Grid_out (kW)'].sum()):.2f}")

    st.markdown("---")

    # ─── PLOT SECTION ───
    st.subheader("📊 Interactive Energy Chart")

    ALL_SIGNALS = {
        "Load (kW)": "Load (kW)",
        "RE (kW)": "RE (kW)",
        "Battery SOC (%)": "Battery SOC (%)",
        "Load_out (kW)": "Load_out (kW)",
        "Batt_out (kW)": "Batt_out (kW)",
        "Grid_out (kW)": "Grid_out (kW)",
    }

    pc1, pc2 = st.columns(2)
    with pc1:
        primary_signals = st.multiselect(
            "Primary Y-axis (kW)",
            options=list(ALL_SIGNALS.keys()),
            default=["Load (kW)", "RE (kW)", "Load_out (kW)", "Grid_out (kW)"],
        )
    with pc2:
        secondary_signals = st.multiselect(
            "Secondary Y-axis",
            options=list(ALL_SIGNALS.keys()),
            default=["Battery SOC (%)"],
        )

    COLORS = ["#2196F3", "#4CAF50", "#FF9800", "#E91E63", "#9C27B0", "#00BCD4", "#F44336", "#8BC34A"]

    if primary_signals or secondary_signals:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        color_idx = 0
        x_vals = df_result["Time"].tolist()

        for sig in primary_signals:
            col_name = ALL_SIGNALS[sig]
            if col_name in df_result.columns:
                fig.add_trace(
                    go.Scatter(
                        x=x_vals,
                        y=df_result[col_name].tolist(),
                        name=sig,
                        line=dict(color=COLORS[color_idx % len(COLORS)], width=2),
                        mode="lines",
                    ),
                    secondary_y=False,
                )
                color_idx += 1

        for sig in secondary_signals:
            col_name = ALL_SIGNALS[sig]
            if col_name in df_result.columns:
                fig.add_trace(
                    go.Scatter(
                        x=x_vals,
                        y=df_result[col_name].tolist(),
                        name=sig + " (R)",
                        line=dict(color=COLORS[color_idx % len(COLORS)], width=2, dash="dot"),
                        mode="lines",
                    ),
                    secondary_y=True,
                )
                color_idx += 1

        fig.update_layout(
            height=480,
            template="plotly_dark",
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=40, r=40, t=40, b=60),
            xaxis=dict(title="Time", tickangle=-45, tickmode="auto", nticks=24),
        )
        fig.update_yaxes(title_text="Power (kW)", secondary_y=False)
        fig.update_yaxes(title_text="SOC (%)", secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Select at least one signal to plot.")

    st.markdown("---")

    # ─── RESULTS TABLE ───
    st.subheader("📋 EMS Output Table")
    display_cols = ["Time", "Battery SOC (%)", "RE (kW)", "Load (kW)", "Grid Available", "Tariff",
                    "Rule No", "Load Source", "Batt Mode", "Grid Mode",
                    "Load_out (kW)", "Batt_out (kW)", "Grid_out (kW)", "Note"]

    def color_load_src(val):
        colors = {"RE": "#1b5e20", "Battery": "#0d47a1", "Grid": "#b71c1c"}
        bg = colors.get(val, "")
        return f"background-color: {bg}; color: white;" if bg else ""

    def color_batt(val):
        if val == "Charge": return "background-color: #1565c0; color: white;"
        if val == "Discharge": return "background-color: #e65100; color: white;"
        return ""

    styled = df_result[display_cols].style \
        .applymap(color_load_src, subset=["Load Source"]) \
        .applymap(color_batt, subset=["Batt Mode"]) \
        .format({"Batt_out (kW)": "{:+.3f}", "Grid_out (kW)": "{:+.3f}", "Load_out (kW)": "{:.3f}"})

    st.dataframe(styled, use_container_width=True, height=400)

    # ─── DOWNLOAD ───
    st.markdown("---")
    dl_col1, dl_col2 = st.columns(2)

    csv_buf = io.StringIO()
    df_result[display_cols].to_csv(csv_buf, index=False)
    dl_col1.download_button("⬇️ Download Results as CSV", csv_buf.getvalue(), "ems_output.csv", "text/csv")

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df_result[display_cols].to_excel(writer, index=False, sheet_name="EMS Output")
    dl_col2.download_button("⬇️ Download Results as Excel", excel_buf.getvalue(), "ems_output.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ─── RULE DISTRIBUTION ───
    st.markdown("---")
    st.subheader("🔢 Rule Usage Distribution")
    rule_counts = df_result["Rule No"].value_counts().reset_index()
    rule_counts.columns = ["Rule No", "Count"]
    fig2 = go.Figure(go.Bar(
        x=rule_counts["Rule No"].astype(str),
        y=rule_counts["Count"],
        marker_color="#42A5F5",
        text=rule_counts["Count"],
        textposition="outside",
    ))
    fig2.update_layout(
        height=300, template="plotly_dark",
        xaxis_title="Rule Number", yaxis_title="Count",
        margin=dict(l=40, r=20, t=20, b=40),
    )
    st.plotly_chart(fig2, use_container_width=True)

else:
    st.info("👈 Select an input mode from the sidebar to get started.")
    st.markdown("""
    ### How it works
    
    This EMS Dashboard applies the **72-rule decision tree** from the `Rules template` sheet:
    
    | Input Signal | Description |
    |---|---|
    | `RE (kW)` | Renewable energy (Solar + Wind) generation |
    | `Load (kW)` | Site load demand |
    | `Battery SOC (%)` | State of charge: Low (0–30%), Medium (30–70%), High (70–100%) |
    | `Grid Available` | 1 = available, 0 = not available |
    | `Tariff` | 1 = High tariff (peak), 0 = Low tariff (off-peak) |
    | `Time` | HH:MM:SS → determines Time of Day band (TOD1–TOD4) |
    
    **Outputs per row:**
    - **Load Source**: where power comes from (RE / Battery / Grid)
    - **Batt Mode**: Charge / Discharge / Idle  
    - **Grid Mode**: Import / Export / No Export
    - **Load_out, Batt_out, Grid_out (kW)**: numeric power values
    """)
