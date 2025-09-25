## Test-Achats CrawlBot (public pages only)

This project uses Crawl4AI to crawl publicly accessible pages on `www.test-achats.be`, index content into a local Chroma vector store, and expose a FastAPI chatbot for retrieval-augmented answers.

Important: This scaffold intentionally avoids login and paywalled content. It respects `robots.txt` by default. Do not use it to bypass authentication.

### Prerequisites
- Python 3.10+
- `crawl4ai` browser setup: run the installer

```bash
python3 -m crawl4ai.install
```

### Setup
```bash
pip install -r requirements.txt
```

### Crawl (public pages)
```bash
python -m scripts.crawl_public --domain https://www.test-achats.be --limit 1000
```

Outputs raw markdown and metadata to `data/raw/`.

### Index
```bash
python -m scripts.index
```

Builds a Chroma collection at `data/chroma/`.

### Run API
```bash
uvicorn src.api:app --reload --port 8000
```

Query:
```bash
curl -X POST http://localhost:8000/qa -H 'Content-Type: application/json' -d '{"question":"best camera from last test in 2025"}'
```

### Environment
Create `.env` if needed:
```
OPENAI_API_KEY=...
```

### Notes on DB choice (free options)
- For local/free usage, Chroma (embedded) is zero-cost and simple. For remote DB with free tier, Supabase (Postgres) is a good default. See discussion in the main app response.

