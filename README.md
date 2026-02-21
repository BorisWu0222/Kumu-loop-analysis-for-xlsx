# Kumu Loop Report Generator

## 1. Background & Problem Statement

### Background

When using **Kumu** to build Causal Loop Diagrams (CLDs), the raw map data is stored as a list of connections — each row simply says *"A points to B."* On its own, this flat list tells you nothing about the **closed feedback loops** hiding inside the system.

To find those loops, someone would need to manually trace every possible path through the map and check whether it circles back to its own starting point. For a map with dozens of nodes and connections — such as *"How to solve the low participation rate in career institution activities?"* — this is not just tedious, it is practically impossible to do completely by hand.

### The Problem

Even after building a Kumu map, teams are typically left with two unresolved questions:

- **How many feedback loops actually exist in this system?** Without an exhaustive count, you can't know whether your named loops represent 30% or 90% of the system's real dynamics.
- **What is the exact node-by-node path of each loop?** Knowing *that* a loop exists isn't enough — you need the full path (`A → B → C → A`) to understand *why* it exists and *where* to intervene.

### Importance

A feedback loop that is never fully traced is a leverage point that is never acted on. This script gives your team **a complete, structured inventory** of every cycle in the system — so that strategic decisions are based on the full picture, not just the loops that were easy to spot visually.

---

## 2. Input, Output & Expected Impact

### Input & Output

| | Details |
| :--- | :--- |
| **Input** | An `.xlsx` file exported from Kumu — specifically the **Connections sheet**, which must contain a `From` column and a `To` column. |
| **Output** | An `.xlsx` report file saved to your Desktop (`kumu_loops_report_full.xlsx`) with two sheets (see below). |

### Output File Structure

| Sheet | Contents |
| :--- | :--- |
| **Loop_Report** | One row per detected loop — includes a Loop ID, the number of nodes (Length), and the full path string (e.g., `參與率低 → 活動吸引力不足 → 宣傳不足 → 參與率低`). |
| **For_Kumu_Import** | Every connection annotated with the Loop IDs it belongs to, formatted for re-importing tags back into Kumu. |

### Expected Impact

- **Completeness:** Uses NetworkX's `simple_cycles` algorithm to detect every mathematically possible closed loop — no manual tracing, no gaps.
- **Clarity:** Converts abstract graph data into human-readable path strings that any team member can follow without a coding background.
- **Actionability:** The Kumu Import sheet lets you push loop labels back into your map in one step, so your visual diagram reflects the full analysis.

---

## 3. How to Use

### Prerequisites

Make sure Python is installed, then run:

```bash
pip install pandas networkx openpyxl
```

### Step-by-Step Instructions

**Step 1 — Export from Kumu:**

1. Open your Kumu project.
2. Click the menu (bottom-left) → **Export** → **XLSX (Excel)**.
3. Save the file to your **Desktop**.

**Step 2 — Configure the Script:**

Open the `.py` script and find the `路徑與檔名設定` section at the top. Update the input filename to match your exported file:

```python
# ================= 路徑與檔名設定 =================
input_filename = 'your-kumu-export-filename.xlsx'   # ← update this
output_filename = 'kumu_loops_report_full.xlsx'      # ← rename if needed
sheet_name = 'Connections'                           # ← must match your sheet tab name
```

> **Note:** The script automatically looks for files on your Desktop (`~/Desktop`). If your file is elsewhere, update `desktop_path` accordingly.

**Step 3 — Run the Script:**

```bash
python kumu_loop_report.py
```

The terminal will show a live progress log, including a preview of the first 5 detected loops.

**Step 4 — Review the Output:**

Open `kumu_loops_report_full.xlsx` on your Desktop.

- Use **Loop_Report** to audit all loops and identify which ones your team should name and prioritize.
- Use **For_Kumu_Import** to copy loop tag assignments back into Kumu's connection data, so your visual map stays in sync.

---

## 4. Troubleshooting

| Issue | Solution |
| :--- | :--- |
| `❌ 找不到檔案` | Check that the filename in the script exactly matches the file on your Desktop (including spaces and special characters). |
| `❌ 找不到 'From' 或 'To' 欄位` | Open your Excel file and confirm the Connections sheet has columns named exactly `From` and `To`. |
| `⚠️ 未發現任何閉環` | Your map may have no cycles — check whether all connections are one-directional with no return paths. |
| `ModuleNotFoundError` | Re-run `pip install pandas networkx openpyxl`. |
