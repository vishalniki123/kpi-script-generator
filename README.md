# KPI Script Generator 🚀

A web-based application to automatically generate **PSQ & SQD KPI Excel scripts** from a master KPI Excel file.

This tool removes manual effort by:
- Reading KPI definitions from a source Excel
- Applying PSQ / SQD templates
- Generating individual KPI Excel scripts
- Providing a downloadable ZIP output

---

## ✨ Features

- Upload KPI master Excel file
- Upload PSQ & SQD script templates
- Automatic KPI-wise Excel generation
- Tracking-only logic handling
- Clean, user-friendly web interface
- Download generated scripts as ZIP
- Dark / Light mode ready UI
- Runs locally or on cloud (Render / Azure / AWS)

---

## 🧱 Project Structure

```text
kpi-script-generator/
├── app.py
├── kpi_generator.py
├── requirements.txt
├── README.md
├── templates/
│   └── index.html
├── static/
│   └── style.css
├── uploads/        (auto-created at runtime)
├── output/         (auto-created at runtime)
└── .gitignore