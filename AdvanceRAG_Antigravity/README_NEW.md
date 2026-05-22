# 🚀 Advanced RAG Pipeline — VWO Test Cases Explorer

<div align="center">

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Backend-Flask-green?logo=flask&logoColor=white)](https://flask.palletsprojects.com/)
[![Qdrant](https://img.shields.io/badge/VectorDB-Qdrant-purple?logo=qdrant&logoColor=white)](https://qdrant.tech/)
[![License](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

**Enterprise-Grade RAG System for Intelligent Test Case Retrieval & Augmented Analysis**

[Features](#-features) • [Quick Start](#-quick-start) • [Architecture](#-architecture) • [Tech Stack](#-tech-stack) • [Usage](#-usage) • [Contributing](#-contributing)

</div>

---

## ✨ Features

- 🎯 **Semantic Search** — BGE-M3 embeddings (1024-dim dense vectors) for precise test case retrieval
- 🔄 **Two-Stage Ranking** — Vector similarity + BGE Reranker v2-m3 for refined results (Top-5)
- 🧠 **LLM Integration** — Groq Llama 3.3 70B for context-aware answer generation
- ⚡ **Production-Ready** — Qdrant vector DB (Rust-based, ultra-fast)
- 🎨 **Web Interface** — Beautiful vanilla HTML/CSS/JS dashboard
- 📊 **5000+ Test Cases** — VWO Login Dashboard test coverage
- 🔐 **Environment-Safe** — Secure API key management with .env

---

## 🏗️ Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│                         USER QUERY                               │
└────────────────────────┬─────────────────────────────────────────┘
                         │
                         ▼
        ┌────────────────────────────────┐
        │  📝 Query Preprocessing        │
        │  (Text normalization)          │
        └────────────┬───────────────────┘
                     │
                     ▼
        ┌────────────────────────────────┐
        │  🧬 BGE-M3 Embedding           │
        │  (BAAI/bge-m3)                 │
        │  1024-dim Dense Vector         │
        └────────────┬───────────────────┘
                     │
                     ▼
        ┌────────────────────────────────┐
        │  🔍 Qdrant Vector Search       │
        │  Top-10 Similarity Results     │
        └────────────┬───────────────────┘
                     │
                     ▼
        ┌────────────────────────────────┐
        │  🎯 BGE-Reranker v2-m3         │
        │  Cross-Encoder Scoring         │
        │  Top-5 Refined Results         │
        └────────────┬───────────────────┘
                     │
                     ▼
        ┌────────────────────────────────┐
        │  🤖 Groq LLM (Llama 3.3 70B)   │
        │  Context-Aware Response        │
        └────────────┬───────────────────┘
                     │
                     ▼
        ┌────────────────────────────────┐
        │  💡 FINAL ANSWER               │
        │  + Source References           │
        └────────────────────────────────┘
```

---

## 🛠️ Tech Stack

| Layer | Technology | Purpose |
|-------|-----------|---------|
| **Embedding** | `BAAI/bge-m3` | 1024-dim dense vectors for semantic search |
| **Vector DB** | Qdrant | Rust-based, ultra-fast similarity search |
| **Reranking** | `BAAI/bge-reranker-v2-m3` | Cross-encoder for precise relevance scoring |
| **LLM** | Groq API + Llama 3.3 70B | Fast, powerful text generation |
| **Backend** | Flask (Python) | RESTful API server |
| **Frontend** | Vanilla HTML/CSS/JS | Responsive web interface |
| **Database** | SQLite + Qdrant | Local persistent storage |

---

## 📂 Project Structure

```
AdvanceRAG_Antigravity/
│
├── 🚀 Core Application
│   ├── app.py                      # Flask app & routes (entry point)
│   ├── config.py                   # Centralized configuration
│   ├── llm.py                      # Groq LLM integration
│   │
├── 🧬 Data Pipeline
│   ├── data_ingestion.py           # CSV → Qdrant data loading
│   ├── embeddings.py               # BGE-M3 embedding generator
│   ├── vector_store.py             # Qdrant operations
│   ├── reranker.py                 # BGE Reranker integration
│   │
├── 🔍 Query Processing
│   ├── query_pipeline.py           # End-to-end retrieval flow
│   │
├── 🎨 Web Interface
│   ├── index.html                  # Main dashboard
│   ├── styles.css                  # Styling
│   ├── script.js                   # Frontend logic
│   │
├── 📊 Testing & Generation
│   ├── generate_test_cases.py      # Test case generator
│   │
├── 📁 Directories
│   ├── data/                       # Raw test case data
│   ├── qdrant_db/                  # Vector database (local)
│   ├── uploads/                    # User uploads
│   │
├── 📦 Dependencies
│   ├── requirements.txt            # Python dependencies
│   ├── .env                        # Environment variables (local only)
│   ├── .gitignore                  # Git ignore rules
│   │
└── 📖 Documentation
    └── README.md                   # This file
```

---

## 🚀 Quick Start

### Prerequisites
- **Python 3.10+**
- **pip** or **conda**
- **Git**

### 1️⃣ Clone Repository
```bash
git clone https://github.com/Akanksha19Dec/AdvanceRAG.git
cd AdvanceRAG_Antigravity
```

### 2️⃣ Set Up Environment
```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### 3️⃣ Configure API Keys
Create a `.env` file in the root directory:
```env
GROQ_API_KEY=your_groq_api_key_here
QDRANT_PATH=./qdrant_db
```

Get your Groq API key from: [https://console.groq.com](https://console.groq.com)

### 4️⃣ Initialize Vector Database
```bash
python data_ingestion.py
```

### 5️⃣ Run the Application
```bash
python app.py
```

Visit: **http://localhost:5000** 🌐

---

## 📖 Usage

### Web Interface
1. Open dashboard at `http://localhost:5000`
2. Enter your query about VWO test cases
3. View semantically similar test cases + AI-generated insights
4. Click sources to view full test case details

### API Endpoints

#### Query Test Cases
```bash
curl -X POST http://localhost:5000/api/query \
  -H "Content-Type: application/json" \
  -d '{"query": "login form validation"}'
```

**Response:**
```json
{
  "query": "login form validation",
  "results": [
    {
      "test_case_id": 123,
      "title": "Verify Login Form Validation",
      "description": "...",
      "score": 0.92,
      "source": "VWO Dashboard"
    }
  ],
  "generated_answer": "Based on the retrieved test cases...",
  "execution_time_ms": 245
}
```

#### Health Check
```bash
curl http://localhost:5000/health
```

---

## 🔧 Configuration

Edit `config.py` to customize:

```python
# Embedding Model
EMBEDDING_MODEL = "BAAI/bge-m3"
EMBEDDING_DIM = 1024

# Reranking
RERANKER_MODEL = "BAAI/bge-reranker-v2-m3"
TOP_K_INITIAL = 10
TOP_K_FINAL = 5

# LLM
LLM_MODEL = "llama-3.3-70b-versatile"
LLM_MAX_TOKENS = 1024
LLM_TEMPERATURE = 0.7

# Vector DB
QDRANT_PATH = "./qdrant_db"
COLLECTION_NAME = "vwo_advanced_test_cases"
```

---

## 📊 Performance Metrics

| Metric | Value | Notes |
|--------|-------|-------|
| **Query Latency** | ~300-500ms | End-to-end (embedding + search + rerank + LLM) |
| **Embedding Throughput** | ~1000 docs/min | BGE-M3 (GPU optimized) |
| **Vector Similarity Speed** | ~5-10ms | Qdrant (top-10 retrieval) |
| **Reranking Speed** | ~50-100ms | BGE-Reranker (5 docs) |
| **Collection Size** | 5000+ | Test cases in vector DB |
| **Vector Dimensions** | 1024 | BGE-M3 dense embeddings |

---

## 🧪 Testing

### Run Test Suite
```bash
python -m pytest tests/
```

### Generate Sample Test Cases
```bash
python generate_test_cases.py
```

### Manual Testing
```bash
# Test embedding quality
python -c "from embeddings import embed_text; print(embed_text('test query')[:5])"

# Test Qdrant connection
python -c "from vector_store import check_qdrant; check_qdrant()"
```

---

## 🐛 Troubleshooting

| Issue | Solution |
|-------|----------|
| `ModuleNotFoundError: No module named 'torch'` | Run `pip install -r requirements.txt` with `--no-cache-dir` |
| `Qdrant connection failed` | Ensure `qdrant_db/` exists or run `python data_ingestion.py` |
| `GROQ_API_KEY not found` | Create `.env` file with valid API key |
| `Port 5000 already in use` | Change port in `app.py`: `app.run(port=5001)` |
| `Out of memory` | Reduce `batch_size` in `config.py` or use GPU |

---

## 📚 Model Details

### BGE-M3 Embedding
- **Dimensions**: 1024
- **Training Data**: Diverse cross-lingual corpus
- **Strengths**: Excellent for semantic search, multilingual support
- **Paper**: [BAAI/bge-m3](https://huggingface.co/BAAI/bge-m3)

### BGE-Reranker v2-m3
- **Type**: Cross-encoder
- **Input**: Query + Document pairs
- **Output**: Relevance score (0-1)
- **Strengths**: Significant ranking improvement over dense retrievers alone
- **Paper**: [BAAI/bge-reranker-v2-m3](https://huggingface.co/BAAI/bge-reranker-v2-m3)

### Groq Llama 3.3 70B
- **Type**: Large Language Model
- **Speed**: Fastest inference on Groq hardware
- **Context**: 8K tokens
- **Strengths**: High quality responses, code generation, reasoning

---

## 🤝 Contributing

Contributions are welcome! Here's how to contribute:

1. **Fork** the repository
2. **Create** a feature branch: `git checkout -b feature/amazing-feature`
3. **Commit** changes: `git commit -m 'Add amazing feature'`
4. **Push** to branch: `git push origin feature/amazing-feature`
5. **Open** a Pull Request

### Development Setup
```bash
# Install dev dependencies
pip install -r requirements.txt

# Pre-commit checks
python -m black .
python -m pylint *.py
```

---

## 📄 License

This project is licensed under the **MIT License** — see [LICENSE](LICENSE) file for details.

---

## 💡 Future Enhancements

- [ ] Hybrid search (dense + sparse BM25)
- [ ] Multi-modal embeddings (text + images)
- [ ] Real-time indexing with streaming updates
- [ ] Advanced filtering and metadata search
- [ ] Caching layer for frequently asked questions
- [ ] Analytics dashboard
- [ ] Docker containerization
- [ ] Kubernetes deployment configs

---

## 📞 Support & Contact

- **Issues**: [GitHub Issues](https://github.com/Akanksha19Dec/AdvanceRAG/issues)
- **Discussions**: [GitHub Discussions](https://github.com/Akanksha19Dec/AdvanceRAG/discussions)
- **Email**: [akanksha.19dec@gmail.com](mailto:akanksha.19dec@gmail.com)

---

<div align="center">

⭐ **If this project helped you, please consider giving it a star!** ⭐

**Made with ❤️ by [Akanksha19Dec](https://github.com/Akanksha19Dec)**

</div>
