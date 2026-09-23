import logging
from config import Config

logger = logging.getLogger(__name__)

_chroma_client = None
_collection = None

def get_chroma_collection():
    global _chroma_client, _collection
    if _collection is not None:
        return _collection

    try:
        import chromadb
        _chroma_client = chromadb.PersistentClient(path=Config.CHROMA_PERSIST_DIRECTORY)
        _collection = _chroma_client.get_or_create_collection(
            name="oes_questions",
            metadata={"hnsw:space": "cosine"}
        )
        return _collection
    except Exception as e:
        logger.error(f"ChromaDB initialization failed: {e}")
        return None


def index_question(question):
    """
    Indexes a single Question model in ChromaDB for semantic search.
    """
    col = get_chroma_collection()
    if not col:
        return False
    try:
        col.upsert(
            documents=[question.text],
            metadatas=[{
                "question_id": question.id,
                "topic": question.topic,
                "difficulty": question.difficulty,
                "question_type": question.question_type
            }],
            ids=[f"q_{question.id}"]
        )
        return True
    except Exception as e:
        logger.error(f"Failed to index question {question.id} in ChromaDB: {e}")
        return False


def index_all_questions(questions):
    col = get_chroma_collection()
    if not col or not questions:
        return False
    try:
        docs = [q.text for q in questions]
        metas = [{
            "question_id": q.id,
            "topic": q.topic,
            "difficulty": q.difficulty,
            "question_type": q.question_type
        } for q in questions]
        ids = [f"q_{q.id}" for q in questions]
        
        col.upsert(documents=docs, metadatas=metas, ids=ids)
        return True
    except Exception as e:
        logger.error(f"Failed batch indexing in ChromaDB: {e}")
        return False


def search_similar_questions(query_text, n_results=5, topic_filter=None):
    """
    Performs semantic search against question database.
    Returns list of matched dict metadata and similarity score.
    """
    col = get_chroma_collection()
    if not col:
        return []
    try:
        where_clause = {"topic": topic_filter} if topic_filter else None
        results = col.query(
            query_texts=[query_text],
            n_results=n_results,
            where=where_clause
        )
        matched = []
        if results and 'metadatas' in results and len(results['metadatas']) > 0:
            for i, meta in enumerate(results['metadatas'][0]):
                doc_text = results['documents'][0][i] if 'documents' in results else ""
                dist = results['distances'][0][i] if 'distances' in results else 0.0
                matched.append({
                    "question_id": meta.get("question_id"),
                    "topic": meta.get("topic"),
                    "difficulty": meta.get("difficulty"),
                    "question_type": meta.get("question_type"),
                    "text": doc_text,
                    "similarity_score": round(1.0 - dist, 3)
                })
        return matched
    except Exception as e:
        logger.error(f"ChromaDB search error: {e}")
        return []
