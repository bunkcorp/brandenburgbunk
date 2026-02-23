# ActuarialExams workspace

Quick map of where things live after reorganization.

## Audio (SuperCollider, Bela, DnB)
- **Audio/Brandenburg/** – Brandenburg MIDI + Liquid DnB patches (`.scd`), `run_brandenburg_record.sh`, BrandenburgDnB project
- **Audio/Bela/** – Bela scripts, docs (`BELA_*.md`), `bela_receiver_stable.py`, BelaNetworkStream, capture logs
- **Audio/copied_patches/** – Sample/break WAVs (e.g. amen.wav, fats.wav) used by the patches

**Run Brandenburg + record:**  
`./Audio/Brandenburg/run_brandenburg_record.sh`

## Exam prep
- **Exam_Prep/** – PA modules, solutions, study materials, PA_Module_Materials (rap lyrics, vocal scripts), PA task4 assets
- **ANKI_LOGIC/, ATPA/, Modules/, ISLP/** – Other exam/study resources

## Scripts
- **Scripts/PA_Anki/** – Anki/PA card scripts: `fix_*_card.py`, `get_*_cards.py`, `rhyme_*.py`, `tag_pa_cards_by_syllabus.py`, etc.
- **Scripts/Study/** – `create_study_schedule.py`, `create_study_schedule_v2.py`
- **Scripts/Bela/** – (Bela-specific Python lives under Audio/Bela)

## Data
- **Spreadsheets/** – All `.xlsx` files (forward selection, stepwise models, exercises, exam costs, etc.)
- **Exam_Prep/** – `cards_by_lo.json`, `lo1c_refrain_notes.json`, `PA_Study_Schedule.ics`

## Other
- **Resources/, Administrative/, gma policy/** – Unchanged
- Loose PDFs, `testi.ipynb` – Still at repo root (move into Resources or Exam_Prep if you want them grouped)
