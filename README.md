# DaveAI

**DaveAI** stands for **D**eterministic **A**vicollect **V**irtual **E**xpert, a
privacy‑preserving chatbot that answers natural‑language questions about historic flight
data. It uses a **text‑to‑SQL** architecture: your question is translated into a real SQL
query that runs against a DuckDB database, so all arithmetic is computed by the
database, not guessed by the language model.

The name is deliberate. "Deterministic" reflects the architecture: answers are *grounded
by SQL execution*, not produced by free‑form text generation. The LLM never invents a
number, the database always computes it. Nothing leaves your machine: no cloud APIs, no
external data sharing.

## Stack
- **LLM runtime:** Ollama, running Qwen2.5‑7B‑Instruct (Apache 2.0, open‑weight)
- **Database:** DuckDB (embedded analytical database)
- **Language:** Python
- **Excel:** pandas + openpyxl
- **GPU:** Intel Arc (Vulkan backend — no CUDA)

## Status
- [x] Part 1 — Environment, GitHub, and data pipeline
- [ ] Part 2 — Text‑to‑SQL core + MLflow experiment tracking
- [ ] Part 3 — Voice I/O (English, Yoruba, Igbo, Hausa) + UI + deployment

## Quick start
See the Part 1 guide. In short: create the venv, `pip install -r requirements.txt`,
put your Excel files in `data/raw/`, then `python src/db_loader.py`.