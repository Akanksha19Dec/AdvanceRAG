"""
Configuration — Advanced RAG Pipeline
=======================================
Centralized settings for embedding models, vector DB, reranker, and LLM.
"""

import os

# ── Paths ──
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Test_Cases_VWO_Login_5000.xlsx")
QDRANT_PATH = os.path.join(BASE_DIR, "qdrant_db")

# ── Embedding Model (BAAI/bge-m3) ──
EMBED_MODEL_NAME = "BAAI/bge-m3"
EMBED_DIM = 1024
EMBED_MAX_SEQ_LENGTH = 8192

# ── Reranker Model (BAAI/bge-reranker-v2-m3) ──
RERANKER_MODEL_NAME = "BAAI/bge-reranker-v2-m3"
RERANKER_MAX_LENGTH = 8192

# ── Vector Database (Qdrant) ──
COLLECTION_NAME = "vwo_advanced_test_cases"
TOP_K_RETRIEVE = 10        # Initial vector search retrieval count
TOP_K_RERANK = 5           # Final count after cross-encoder reranking
QDRANT_BATCH_SIZE = 100    # Batch size for upserts

# ── LLM (Groq Cloud — Llama 3.3 70B) ──
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
GROQ_MODEL = "llama-3.3-70b-versatile"
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_TEMPERATURE = 0.2
GROQ_MAX_TOKENS = 2048

# ── Flask ──
FLASK_PORT = 5002
FLASK_DEBUG = False

# ── Module Keywords for Smart Detection ──
MODULE_KEYWORDS = {
    "Login Authentication": [
        "login", "authentication", "auth", "sign in", "signin", "credentials",
    ],
    "Password Management": [
        "password", "forgot password", "reset password", "change password",
    ],
    "Session Management": [
        "session", "timeout", "logout", "cookie", "remember me",
    ],
    "Input Validation": [
        "validation", "input", "sanitize", "injection", "script",
    ],
    "SSO Integration": [
        "sso", "oauth", "saml", "google login", "microsoft login",
        "social login", "single sign",
    ],
    "UI/UX Design": [
        "ui", "ux", "responsive", "mobile", "dark mode", "design", "branding",
    ],
    "Accessibility": [
        "accessibility", "wcag", "screen reader", "aria", "keyboard",
        "a11y", "contrast",
    ],
    "Security": [
        "security", "https", "brute force", "gdpr", "compliance",
        "encryption", "tls", "csrf",
    ],
    "Multi-Factor Authentication": [
        "mfa", "two factor", "2fa", "totp", "authenticator", "biometric",
    ],
    "Account Management": [
        "account", "profile", "delete account", "settings", "team", "member",
    ],
    "Notifications & Alerts": [
        "notification", "alert", "email", "sms", "push", "digest", "consent",
    ],
    "Audit & Logging": [
        "audit", "log", "logging", "event", "trail", "history",
    ],
    "API Authentication": [
        "api key", "api authentication", "bearer", "token", "rate limit",
    ],
    "Mobile App Login": [
        "mobile", "ios", "android", "face id", "touch id", "fingerprint",
    ],
    "Performance & Load": [
        "performance", "load", "concurrent", "latency", "speed", "benchmark",
    ],
    "Localization": [
        "localization", "locale", "language", "translation",
        "internationalization", "rtl",
    ],
    "Browser Compatibility": [
        "browser", "chrome", "firefox", "safari", "edge", "cross-browser",
    ],
    "Error Handling": [
        "error", "failure", "exception", "timeout", "recovery", "retry",
    ],
    "Data Privacy": [
        "privacy", "gdpr", "ccpa", "data export", "data deletion",
        "encryption", "pii",
    ],
    "Integration Testing": [
        "integration", "e2e", "end to end", "cross module", "workflow",
    ],
}
