# Web App Valuation Dashboard

An interactive Dash dashboard for building, projecting, and valuing web-app financial models. This project integrates dynamic user assumptions, Excel-based projection and valuation models, Monte Carlo simulation, and sensitivity analysis into a seamless web interface.

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Features](#features)
3. [Repository Structure](#repository-structure)
4. [Getting Started](#getting-started)
   - [Prerequisites](#prerequisites)
   - [Installation](#installation)
   - [Configuration](#configuration)
   - [Running the App](#running-the-app)
5. [Usage Guide](#usage-guide)
   - [1. Assumptions Input](#1-assumptions-input)
   - [2. Projection Model Display](#2-projection-model-display)
   - [3. Valuation Model & Simulation](#3-valuation-model--simulation)
   - [4. Sensitivity Analysis](#4-sensitivity-analysis)
6. [Customizing Ranges](#customizing-ranges)
7. [Performance & Caching](#performance--caching)
8. [Dependencies](#dependencies)
9. [Contributing](#contributing)
10. [License](#license)

---

## Project Overview

This dashboard guides users through:

1. Entering **assumptions** via multi‐tab forms
2. **Writing** those inputs to an Excel-based assumptions sheet
3. **Projecting** financials by reading the Excel "Projection Model" sheet
4. Displaying the **valuation** outputs (FCF to Equity and FCF to the Firm)
5. Running a **Monte Carlo simulation** on the final equity value
6. Performing an **editable sensitivity analysis** against the Monte Carlo mean

All Excel formulas remain in the workbook; Dash simply reads, writes, and triggers recalculation via `win32com`.

---

## Features

- **Dynamic Assumption Tabs**: Generate input fields from a JSON structure, preserving values across tabs.
- **Excel Integration**: Push user inputs to `Assumption Table` sheet, trigger full workbook recalculation.
- **Projection & Valuation UI**: Render any sheet as a clean DataTable, insert custom header rows (e.g., Year 1–5).
- **Valuation Selector**: Switch between "FCF to Equity" and "FCF to the Firm" views.
- **Monte Carlo Simulation**: Run 1,000‑trial random simulations invisibly in Excel and display histograms with formatted axes.
- **Sensitivity Analysis**: Editable DataTable for "Input Change" column, writes back to Excel, recalculates, and updates results, overriding final values with MC mean.
- **Formatting**: Comma‑separated thousands, customizable decimal precision in tables and charts.
- **Client‑side Caching**: Use `dcc.Store` to cache the Monte Carlo mean on page load, avoiding expensive re‑runs on every edit.

---

## Repository Structure

```
├── app.py                   # Main Dash application
├── data/                    # Folder for Excel and JSON configuration
│   ├── Approximation Valuation Model.xlsx
│   ├── Combined_Assumptions.json
│   └── input_mapping.json
├── assets/                  # Static assets (e.g., logo)
│   └── FCAT-image.png
├── requirements.txt         # Python dependencies
└── README.md                # This file
```

---

## Getting Started

### Prerequisites

- **Python 3.8+** (Windows recommended for Excel COM)
- **Microsoft Excel** (for COM automation)

### Installation

1. **Clone** the repository:
   ```bash
   git clone https://github.com/yourusername/webapp-valuation-dashboard.git
   cd webapp-valuation-dashboard
   ```
2. **Create** and **activate** a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate   # on Windows: venv\\Scripts\\activate
   ```
3. **Install** dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Configuration

- Place your **Excel workbook** (`Approximation Valuation Model.xlsx`) and JSON files in the `data/` folder.
- Ensure `Combined_Assumptions.json` and `input_mapping.json` match your model’s structure and cell mappings.

### Running the App

```bash
python app.py
```
Navigate to `http://127.0.0.1:8050` in your browser.

---

## Usage Guide

### 1. Assumptions Input

- Select an **Industry** and **Assumption Category**.
- Fill in dynamic inputs generated from your JSON assumptions schema.
- Click **Submit All** to write to the Excel `Assumption Table` sheet and recalculate.

### 2. Projection Model Display

- After submission, the app automatically navigates to the **Valuation Selector** page.
- Select **FCF to Equity** or **FCF to the Firm**.
- View the projection results as a clean HTML table with Year 1–5 headers.

### 3. Valuation Model & Simulation

- On the **Valuation Results** page, a Monte Carlo histogram and summary stats are computed **once** on page load.
- The Monte Carlo mean is cached for later use in sensitivity.

### 4. Sensitivity Analysis

- Below the simulation, find an **editable** table for sensitivity inputs.
- Change values in the **Input Change** column; Excel formulas recalc all other columns.
- The **Final Estimated Value** column is overridden by `MC_mean × (1 + Total Percent Change)`.

---

## Customizing Ranges

- Projection, valuation, and sensitivity slices are defined in `app.py` or helper functions:
  - **Projection rows**: adjust `data[7:21]` or `data[23:33]` slices.
  - **Sensitivity header**: update `header_row`, `start_col`, `end_col` in `read_sensitivity_table()`.
  - **Monte Carlo cell**: change `sht.Range("E18")` to your NPV cell.

---

## Performance & Caching

- Monte Carlo runs once on page load and stores the mean in a hidden `dcc.Store`.
- Sensitivity edits do **not** re‑run the full simulation, only recalc Excel formulas.
- You can increase or decrease `n_trials` in `run_monte_carlo()` for speed vs. accuracy.

---

## Dependencies

Key Python packages in `requirements.txt`:

- dash
- dash-table
- pandas
- numpy
- openpyxl
- pywin32 (win32com)
- plotly

---
