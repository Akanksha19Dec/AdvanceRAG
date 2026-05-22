# 🧠 RAG Explorer — VWO Login Dashboard

A **Retrieval-Augmented Generation (RAG)** system built with **ChromaDB**, **Nomic Embed Text** (via local Ollama), and **Groq LLM** — featuring two independent RAG applications with premium dark-themed UIs.

---

## 🏗️ Architecture

```
📄 PDF / Excel ──→ ✂️ Chunker ──→ 🧠 nomic-embed-text (Ollama) ──→ 🗄️ ChromaDB
                                                                         │
❓ User Query ──→ 🧠 Embed Query ──→ 🔍 ChromaDB Search ←───────────────┘
                                          │
                                     📎 Top-K Chunks ──→ ⚡ Groq LLM ──→ 💬 Answer + Sources
```

---

## 📦 Two RAG Applications

| | **PRD RAG** (`localhost:5000`) | **Test Case RAG** (`localhost:5001`) |
|---|---|---|
| **Source** | VWO PRD (PDF) | 100 Test Cases (Excel) |
| **Chunking** | 800 chars / 120 overlap | 1 test case = 1 chunk |
| **Retrieval** | Semantic top-3 | Smart module-filter + semantic top-5 |
| **Script** | `rag_app.py` | `rag_testcases.py` |
| **UI** | `index.html` | `testcases.html` |

---

## ⚡ Tech Stack

| Component | Technology |
|-----------|-----------|
| Embeddings | `nomic-embed-text` via local Ollama |
| Vector Store | ChromaDB (local persistent) |
| LLM | `openai/gpt-oss-120b` via Groq API |
| PDF Extraction | PyPDF2 |
| Excel I/O | openpyxl |
| Web Server | Flask |
| API Key Mgmt | python-dotenv (`.env`) |

---

## 🚀 Quick Start

### Prerequisites

- **Python 3.10+**
- **Ollama** running locally with `nomic-embed-text`:
  ```bash
  ollama pull nomic-embed-text
  ```
- **Groq API key** — free at [console.groq.com](https://console.groq.com)

### Install Dependencies

```bash
pip install flask chromadb requests python-dotenv PyPDF2 openpyxl
```

### Configure API Key

Create a `.env` file in the project root:
```
GROQ_API_KEY=gsk_your_key_here
```

### Run PRD RAG
```bash
python rag_app.py
# Open http://localhost:5000
```

### Run Test Case RAG
```bash
python rag_testcases.py
# Open http://localhost:5001
```

Both servers can run simultaneously.

---

## 📂 Project Structure

```
RAG/
├── .env                        # Groq API key
├── .gitignore                  # Git ignore rules
├── rag_app.py                  # PRD RAG backend (Flask)
├── index.html                  # PRD RAG frontend
├── rag_testcases.py            # Test Case RAG backend (Flask)
├── testcases.html              # Test Case RAG frontend
├── generate_test_cases.py      # Generates 100 test cases → Excel
├── Test_Cases_VWO_Login.xlsx   # 100 enterprise test cases
├── rag_system_walkthrough.md   # Detailed walkthrough
├── rag_pipeline_architecture.png
├── data/
│   └── Product Requirements Document_ VWO Login Dashboard.pdf
├── chroma_db/                  # ChromaDB storage (PRD) — auto-created
└── chroma_tc_db/               # ChromaDB storage (Test Cases) — auto-created
```

---

## 🔍 Key Features

- **Pipeline Visualization** — Animated 6-step RAG pipeline in the UI header
- **ChromaDB Caching** — Embeddings persist; subsequent runs skip re-embedding
- **Smart Module Detection** — Test Case RAG detects module keywords and retrieves ALL matching test cases via metadata filter
- **Chunk Explorer** — Expandable cards with embedding stats (dimension, magnitude, min/max/mean)
- **Source Attribution** — Every answer shows source chunks with similarity percentages
- **Module Filter Pills** — Click to filter test cases by module (Test Case RAG)
- **Dual Dark-Themed UIs** — Premium glassmorphism design with micro-animations

---

## 🧪 Test Case Coverage

| Module | Count |
|--------|-------|
| Login Authentication | 20 |
| Password Management | 15 |
| SSO Integration | 12 |
| Input Validation | 13 |
| UI/UX Design | 12 |
| Session Management | 10 |
| Accessibility | 10 |
| Security | 8 |
| **Total** | **100** |

Categories: Positive (43) · Negative (35) · Edge Case (17) · Boundary (5)

---

## 📝 License

This project is for educational and demonstration purposes.
