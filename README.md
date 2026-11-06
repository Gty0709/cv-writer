# CV-Writer: Professional Academic & Job-Seeking CV Skill

A production-grade Antigravity & AI Agent skill for drafting, polishing, and formatting professional Curriculum Vitae (CV) and Resumes for both **academic graduate applications (Master's / PhD)** and **job-seeking (求职 / 校招 / 社招)**.

---

## 🌟 Key Features

- **Dual-Track Adaptation**:
  - **Academic CV**: Highlights educational pedigree, research projects, publications, conferences, and academic honors.
  - **Job-Seeking CV**: Emphasizes professional experience, cross-functional leadership, quantitative business outcomes, and portfolio projects.
- **Studio-Grade Writing Principles**:
  - Rigorous grammar and tense rules (past tense for past roles, active voice, zero first-person pronouns).
  - Strong Action Verbs taxonomy with domain-specific vocabulary.
  - STAR / Action-Context-Result structure with bold lead-in tags (`• Bold Prefix:`).
- **Pixel-Perfect Typography & Layout**:
  - Classic Times New Roman styling with formal border separators.
  - Right-aligned tab stops (`pos="10204"`) for clean dual-column date and location headers.
- **Strict Page Budgeting**:
  - Enforces exact **Single-Page** or **Full Two-Page** formatting.
  - Zero orphan/widow line spillover.
- **Dual Deliverable Output**:
  - Automatically compiles standard Microsoft Word `.docx` and exports pristine `.pdf` documents simultaneously.

---

## 📁 Repository Structure

```
cv-writer/
├── SKILL.md                 # Antigravity Skill definition and operational workflow
├── README.md                # Project documentation
├── references/
│   ├── cv-principles.md     # Comprehensive CV writing standards & rules
│   └── action-verbs.md      # Action verbs taxonomy categorized by industry function
├── templates/
│   └── template.docx        # Pre-configured baseline OpenXML Word template
└── scripts/
    └── build_cv.ps1         # Automated PowerShell builder & PDF exporter
```

---

## 🚀 Installation

### 1. Install to Antigravity
Clone or copy this repository to your global or project Antigravity skills directory:
```bash
# Global Antigravity skills directory
git clone https://github.com/Gty0709/cv-writer.git ~/.gemini/antigravity/skills/cv-writer
```

### 2. Install via Skills-Manager
If using `skills-manager`:
```bash
skills-manager-cli.exe skills install --source https://github.com/Gty0709/cv-writer.git
skills-manager-cli.exe skills deploy
```

---

## 📄 License

MIT License. Designed for professional academic consulting and career advisory workflows.
