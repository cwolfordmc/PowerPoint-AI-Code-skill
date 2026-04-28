# PowerPoint Creator Skill — Setup & Usage Guide

## What Is This Skill?

The `powerpoint-creator` skill teaches GitHub Copilot agent mode to generate professional PowerPoint presentations (.pptx) on demand. It lives at:

```
C:\Users\<you>\.agents\skills\powerpoint-creator\
```

This is a **user-level skill**, meaning it is available in every workspace on your machine automatically — no per-project install needed.

---

## Requirements

### Option A — Python + python-pptx (preferred if PyPI is accessible)

| Requirement | Version |
|-------------|---------|
| Python      | 3.8+    |
| python-pptx | 1.0.0+  |

Install:
```powershell
pip install python-pptx
```

If your environment has SSL/proxy issues (corporate network), use:
```powershell
pip install python-pptx --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org
```

If PyPI is fully blocked, download `python_pptx-1.0.2-py3-none-any.whl` from a machine with internet access and install locally:
```powershell
pip install python_pptx-1.0.2-py3-none-any.whl
```

---

### Option B — PowerShell COM (no pip required)

| Requirement | Notes |
|-------------|-------|
| Windows     | Windows 10/11 |
| Microsoft Office | PowerPoint must be installed |
| PowerShell  | 5.1+ (built into Windows) |

No installation steps needed — just ensure PowerPoint is installed on the machine.

> This is the recommended method for corporate environments where PyPI is blocked or packages fail hash verification due to proxy SSL inspection.

---

## How to Use the Skill

### Step 1 — Open Copilot in Agent Mode

In VS Code, open the Copilot Chat panel and switch to **agent mode** (the dropdown next to the chat input).

### Step 2 — Ask for a Presentation

Describe the presentation you want. Examples:

```
Create a PowerPoint presentation about our Q2 sales results with 8 slides
```
```
Build a technical architecture deck for our Kubernetes migration project
```
```
Generate a 10-slide executive summary presentation based on the README in this project
```

The agent will:
1. Check whether `python-pptx` is available
2. Plan the slide structure and show it to you before generating
3. Generate either a Python script (`create_presentation.py`) or a PowerShell script (`create_presentation.ps1`)

### Step 3 — Run the Generated Script

**If Python was used:**
```powershell
python create_presentation.py
```

**If PowerShell COM was used:**
```powershell
powershell -ExecutionPolicy Bypass -File ".\create_presentation.ps1"
```

The output file (e.g., `MyPresentation.pptx`) will be saved in your workspace folder and will open automatically in PowerPoint.

---

## Troubleshooting

### PyPI SSL / Hash Errors (Corporate Proxy)

Symptom:
```
SSLError: CERTIFICATE_VERIFY_FAILED
```
or
```
ERROR: THESE PACKAGES DO NOT MATCH THE HASHES FROM THE REQUIREMENTS FILE
```

**Solution:** Use the PowerShell COM method (Option B). Ask Copilot to generate a `.ps1` script instead:
```
Create a PowerPoint presentation using PowerShell — python-pptx is not available
```

---

### pip Upgrade Fails

Use trusted-host flags:
```powershell
python -m pip install --upgrade pip --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org
```

---

### PowerShell Script Encoding Error

Symptom:
```
Unexpected token 'SomeWord' in expression or statement
```

This is caused by special characters (em dashes, smart quotes, etc.) in string literals. Open the `.ps1` file, find the offending line, and replace the special character with a plain ASCII equivalent (e.g., `-` instead of `—`).

---

### PowerPoint Doesn't Open / COM Error

Ensure Microsoft Office is installed and that PowerPoint can be launched normally. If PowerPoint is open with an existing file, close it before running the script.

---

## Skill File Location

```
C:\Users\<you>\.agents\skills\powerpoint-creator\
    SKILL.md                  # Skill instructions (loaded by Copilot agent)
    scripts\
        pptx_template.py      # Python starter template
    references\
        design-guide.md       # Color themes and layout rules
```

The `SKILL.md` file has been updated to include both the Python and PowerShell COM generation methods, so Copilot will automatically choose the right one based on your environment.

---

## Color Themes

| Theme          | Primary   | Accent    | Background |
|----------------|-----------|-----------|------------|
| Corporate Blue | `#1F3564` | `#2E75B6` | `#FFFFFF`  |
| Slate & Teal   | `#2C3E50` | `#1ABC9C` | `#F8F9FA`  |
| Crimson Pro    | `#C0392B` | `#E74C3C` | `#FFFFFF`  |
| Forest Green   | `#1E574B` | `#27AE60` | `#F5F5F5`  |
| Midnight Dark  | `#0D1117` | `#58A6FF` | `#161B22`  |

Specify a theme name or hex values when prompting Copilot:
```
Create a presentation using the Midnight Dark theme
```
