# Economic Data Summarizer (AI Workflow) 

A native-style Windows desktop app that helps economists and macro researchers **automate standardized economic data paragraphs** for reports/articles.

I built this because writing “standardized” data blurbs is repetitive: every release cycle you’re updating the same paragraph structure with new readings. This app lets you **upload an Excel sheet** plus an **older paragraph as an example**, then uses a **two-step LLM pipeline** to extract facts and produce a fresh paragraph in the same style.

---

## What the app does

1. **Loads an Excel file** (`.xlsx` / `.xls`)
   - Reads headers from the first row
   - Uses the **last ~24 rows** as the “recent history” window
   - Assumes the **first column is the time/period axis**

2. **Runs a two-step OpenAI Responses API pipeline**
   - **Step 1 (Structured facts extraction):**
     - Extracts facts for **every column**
     - Infers likely **units** (especially using your Example Text when provided)
     - Captures latest/previous readings (using last non-empty cells)
     - Summarizes trends + data quality notes
     - Optionally identifies the **economy** (from user input, example text, or inference)
   - **Step 2 (Final paragraph generation):**
     - Produces a professional paragraph
     - Optional style control to mimic the Example Text:
       - **Strict**: mimic structure and phrasing very closely
       - **Loose**: similar tone/length, more freedom
       - **None**: standard summary style (covers all variables)

3. **Optional editing pass**
   - A separate LLM run that edits the generated paragraph based on user instructions
   - Preserves numbers/units exactly (unless fixing an obvious typo)

4. **Non-blocking UI**
   - Uses a background thread + queue polling so the app remains responsive during API calls

---

## Screenshot of the app (example)

<img width="1268" height="735" alt="example" src="https://github.com/user-attachments/assets/3bf8691d-381f-42a8-a2a5-35783d9243b1" />

---

## Example workflow (how to use)

1. **Prepare your Excel**
   - Put variable names in the **first row** (headers)
   - Put time/period in the **first column**
   - Keep the sheet reasonably “rectangular” (avoid big blank blocks)
   - The app will use the **last 24 rows** to summarize recent movement

2. **Open the app**
   - Click **Browse…** and select your `.xlsx`/`.xls`
   - Enter the **Sheet Name** (default is `Sheet1`)

3. **Provide context (optional but recommended)**
   - **Example Text (optional):** paste a previous standardized paragraph you want to mimic  
     This strongly improves unit inference (e.g., “% YoY”, “SAAR”, “index”, “bps”) and style consistency.
   - **Economy (optional):** e.g., `US`, `China`, `Euro Area`  
     If blank, the app may infer it from your example text or variable names.

4. **Choose a style strictness**
   - **Strict**: closest match to your example’s structure/formatting
   - **Loose**: similar tone and length, but more natural rewriting
   - **None**: generic professional summary across all variables

5. Click **Generate Summary**
   - Output appears on the right
   - Use **Copy to Clipboard** to paste into your report

6. (Optional) Click **Edit Summary**
   - Add instructions like “make it shorter”, “more neutral tone”, “avoid repetition”, etc.

---

## How style filtering works

- If **Example Text is provided** and strictness is **Strict** or **Loose**:
  - The app will **only write about variables that are mentioned in the Example Text** (and exist in the sheet).
- If strictness is **None**:
  - The app writes about **all variables** found in the sheet slice.
- If **Example Text is empty**:
  - The app ignores Strict/Loose filtering and writes a standard summary covering all variables.

This is designed for real-world macro workflows where your “standard paragraph template” usually mentions a subset of indicators from a wider dataset.

---

## Installation & Usage (For Non-Coders)

If you are on Windows and just want to run the app without installing Python, you can download the standalone executable.

1. Go to the **[Releases]([https://github.com/TonyLoWaiYau/Economic-Data-Summarizer/releases/tag/v1.0.0)** page on the right side of this repository.
2. Download the latest `EconomicDataSummarizer.exe` file.
3. Double-click the file to run it. 

> **⚠️ Note on Windows Defender (SmartScreen):** > Because this is a free, open-source tool built with PyInstaller, Windows may show a blue "Windows protected your PC" warning. This is normal for unsigned executable files. Simply click **"More info"** and then **"Run anyway"**.
