from __future__ import annotations

from typing import List, Dict, Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import chromadb
from chromadb.utils import embedding_functions


class QARequest(BaseModel):
    question: str
    k: int = 4


class QAResponse(BaseModel):
    answer: str
    sources: List[Dict[str, Optional[str]]]


def build_collection(persist_dir: str = 'data/chroma'):
    client = chromadb.PersistentClient(path=persist_dir)
    embedder = embedding_functions.SentenceTransformerEmbeddingFunction(
        model_name="all-MiniLM-L6-v2"
    )
    return client.get_or_create_collection(
        name="test_achats_public",
        embedding_function=embedder,
    )


app = FastAPI(title="Test-Achats Public QA")
collection = build_collection()


@app.post('/qa', response_model=QAResponse)
def qa(req: QARequest) -> QAResponse:
    try:
        res = collection.query(query_texts=[req.question], n_results=req.k)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    docs = res.get('documents', [[]])[0]
    metas = res.get('metadatas', [[]])[0]
    sources: List[Dict[str, Optional[str]]] = []
    for i in range(len(docs)):
        meta = metas[i] if i < len(metas) else {}
        sources.append({
            'url': (meta or {}).get('url'),
            'title': (meta or {}).get('title'),
        })

    # Simple extractive answer: concatenate top passages
    context = "\n\n".join(docs)
    answer = context[:1200] if context else "No answer found. Try refining your question."
    return QAResponse(answer=answer, sources=sources)

