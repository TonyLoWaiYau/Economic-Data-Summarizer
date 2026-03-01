# 📊 Economic Data Summarizer (AI Workflow)

An AI workflow that helps economists and macro researchers **automate standardized economic data paragraphs** for reports and articles.

I built this because writing "standardized" data blurbs is incredibly repetitive: every release cycle, you are updating the exact same paragraph structure with new readings. This app lets you **upload an Excel sheet** along with an **older paragraph as an example**, then uses a strict, **two-step LLM pipeline** to extract facts and produce a fresh paragraph in your exact style.

---

## ✨ Core Features

* **Excel Pre-processing:** Reads headers from the first row and targets the **last ~24 rows** to capture recent history. Assumes the first column is the time/period axis.
* **Two-Step LLM Pipeline (OpenAI Responses API):**
    * **Step 1 (Extraction):** Extracts facts for *every* column, infers likely units (cross-referencing your Example Text), captures latest/previous readings, summarizes trends, and identifies the target economy.
    * **Step 2 (Generation):** Produces a professional paragraph with strict, loose, or no stylistic adherence to your example.
* **Dedicated Revision Loop:** An optional, separate LLM pass that edits the generated paragraph based on custom user instructions (e.g., "make it more neutral") while perfectly preserving the extracted numbers and units.
* **Non-Blocking UI:** Built with CustomTkinter using a background thread and queue polling, ensuring the app never freezes during API calls.

---

## 📸 App Interface

<img width="1268" height="735" alt="example" src="https://github.com/user-attachments/assets/3bf8691d-381f-42a8-a2a5-35783d9243b1" />

---

## 🛠️ Example Workflow

1.  **Prepare your Excel File**
    * Put variable names in the **first row** (headers).
    * Put the time/period in the **first column**.
    * Keep the sheet reasonably "rectangular" (avoid large blank blocks or merged cells).
2.  **Open the App & Select Data**
    * Click **Browse…** to upload your `.xlsx` or `.xls` file.
    * Enter the **Sheet Name** (defaults to `Sheet1`).
3.  **Provide Context (Highly Recommended)**
    * **Example Text:** Paste a previous standardized paragraph you want to mimic. *Note: This drastically improves unit inference (e.g., "% YoY", "SAAR", "bps") and style consistency.*
    * **Economy:** Enter the target region (e.g., `US`, `China`, `Euro Area`). If left blank, the app will attempt to infer it from your example text or headers.
4.  **Choose Style Strictness**
    * **Strict:** Closest match to your example's exact structure and formatting.
    * **Loose:** Matches the tone and length, but allows for more natural sentence restructuring.
    * **None:** A generic, professional summary across all data points.
5.  **Generate & Edit**
    * Click **Generate Summary**.
    * *(Optional)* Use the **Editing Requirements** box and click **Edit Summary** to apply specific tweaks (e.g., "make it shorter", "avoid repetition").
    * Click **Copy to Clipboard** to drop it directly into your report.

---

## ⚙️ How Style Filtering Works

In real-world macro workflows, your standard template usually only mentions a subset of indicators from a much wider dataset. The app handles this automatically:

* **If Example Text is provided AND Strictness is Strict/Loose:** The app will *only* write about variables that are explicitly mentioned in the Example Text (and exist in the sheet).
* **If Strictness is None:** The app writes a comprehensive summary covering *all* variables found in the sheet slice.
* **If Example Text is empty:** The app ignores filtering and writes a standard summary covering all variables.

---

## 📥 Installation & Usage (For Non-Coders)

If you are on Windows and just want to run the app without installing Python, you can download the standalone executable:

1.  Go to the **[Releases](https://github.com/TonyLoWaiYau/Economic-Data-Summarizer/releases/tag/v1.0.0)** page.
2.  Download the latest `EconomicDataSummarizer.exe` file.
3.  Double-click the file to run it. 

> **⚠️ Note on Windows Defender (SmartScreen):** > Because this is a free, open-source tool built with PyInstaller, Windows will likely show a blue "Windows protected your PC" warning. This is completely normal for unsigned executable files. Simply click **"More info"** and then **"Run anyway"**.


