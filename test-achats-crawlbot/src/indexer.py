from __future__ import annotations

import json
import os
from typing import List, Dict

import chromadb
from chromadb.utils import embedding_functions


def load_documents(path: str) -> List[Dict]:
    docs: List[Dict] = []
    if not os.path.exists(path):
        return docs
    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
                if 'markdown' in obj and obj['markdown']:
                    docs.append(obj)
            except Exception:
                continue
    return docs


def build_chroma(docs: List[Dict], persist_dir: str = 'data/chroma') -> None:
    os.makedirs(persist_dir, exist_ok=True)
    client = chromadb.PersistentClient(path=persist_dir)

    embedder = embedding_functions.SentenceTransformerEmbeddingFunction(
        model_name="all-MiniLM-L6-v2"
    )

    coll = client.get_or_create_collection(
        name="test_achats_public",
        embedding_function=embedder,
        metadata={"hnsw:space": "cosine"},
    )

    ids: List[str] = []
    documents: List[str] = []
    metadatas: List[Dict] = []

    for i, d in enumerate(docs):
        ids.append(f"doc-{i}")
        documents.append(d.get('markdown', ''))
        metadatas.append({
            "url": d.get('url', ''),
            "title": d.get('title', ''),
        })

    if ids:
        coll.add(ids=ids, documents=documents, metadatas=metadatas)


def main() -> None:
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', default='data/raw/crawl.jsonl')
    parser.add_argument('--persist', default='data/chroma')
    args = parser.parse_args()

    docs = load_documents(args.input)
    if not docs:
        print("No documents found. Run the crawler first.")
        return
    build_chroma(docs, persist_dir=args.persist)
    print(f"Indexed {len(docs)} documents into {args.persist}")


if __name__ == '__main__':
    main()

