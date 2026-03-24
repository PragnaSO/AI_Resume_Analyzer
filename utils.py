import re
from collections import Counter
from pathlib import Path
from typing import Dict, List

import pdfplumber
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

from models.skill_list import SKILLS, ROLE_PROFILES, CERTIFICATIONS, LANGUAGES


TERM_NORMALIZATION = {
    r"\bnodejs\b": "node.js",
    r"\bnode js\b": "node.js",
    r"\brestful\b": "rest api",
    r"\brest services\b": "rest api",
    r"\bmachine-learning\b": "machine learning",
    r"\bml\b": "machine learning",
    r"\bartificial intelligence\b": "ai",
}


FALLBACK_STOPWORDS = {
    "a",
    "an",
    "and",
    "are",
    "as",
    "at",
    "be",
    "by",
    "for",
    "from",
    "has",
    "he",
    "in",
    "is",
    "it",
    "its",
    "of",
    "on",
    "that",
    "the",
    "to",
    "was",
    "were",
    "will",
    "with",
}


JD_GENERIC_TERMS = {
    "need",
    "needs",
    "strong",
    "good",
    "great",
    "role",
    "candidate",
    "responsible",
    "responsibilities",
    "required",
    "preferred",
    "must",
    "plus",
    "ability",
    "experience",
    "project",
    "projects",
}


SECTION_PATTERNS = {
    "experience": r"\b(experience|work history|employment)\b",
    "projects": r"\b(projects|project experience)\b",
    "education": r"\b(education|academic)\b",
    "skills": r"\b(skills|technical skills|core skills)\b",
}


IMPACT_PATTERNS = [
    r"\b(increased|improved|reduced|optimized|delivered|achieved|boosted)\b",
    r"\b\d+%\b",
    r"\$\s?\d+[\d,]*(?:\.\d+)?",
    r"\b\d+[\d,]*(?:\.\d+)?\s?(users|customers|clients|requests|transactions)\b",
]


def get_stopwords() -> set:
    try:
        import nltk
        from nltk.corpus import stopwords

        nltk.download("stopwords", quiet=True)
        return set(stopwords.words("english"))
    except Exception:
        return FALLBACK_STOPWORDS


def extract_text_from_pdf(file_path: str) -> str:
    text_chunks: List[str] = []
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            if page_text.strip():
                text_chunks.append(page_text)
    return "\n".join(text_chunks)


def extract_text_from_docx(file_path: str) -> str:
    document = Document(file_path)
    paragraphs = [p.text.strip() for p in document.paragraphs if p.text and p.text.strip()]
    return "\n".join(paragraphs)


def extract_text_from_doc_windows(file_path: str) -> str:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise ValueError(
            "Legacy .doc parsing needs Microsoft Word on Windows with pywin32 available. "
            "Please upload DOCX or PDF instead."
        ) from exc

    word = None
    doc = None
    txt_path = Path(file_path).with_suffix(".txt")

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(Path(file_path).resolve()))
        # FileFormat 2 = plain text
        doc.SaveAs(str(txt_path.resolve()), FileFormat=2)

        text = txt_path.read_text(encoding="utf-8", errors="ignore")
        return text
    except Exception as exc:
        raise ValueError(
            "Unable to read .doc file automatically. Please save it as .docx and upload again."
        ) from exc
    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()
        if txt_path.exists():
            txt_path.unlink(missing_ok=True)


def extract_text_from_resume(file_path: str) -> str:
    extension = Path(file_path).suffix.lower()

    if extension == ".pdf":
        return extract_text_from_pdf(file_path)
    if extension == ".docx":
        return extract_text_from_docx(file_path)
    if extension == ".doc":
        return extract_text_from_doc_windows(file_path)

    raise ValueError("Unsupported file format. Please upload PDF, DOCX, or DOC.")


def normalize_technical_terms(text: str) -> str:
    normalized = text.lower()
    for pattern, replacement in TERM_NORMALIZATION.items():
        normalized = re.sub(pattern, replacement, normalized)
    return normalized


def preprocess_text(text: str) -> str:
    stop_words = get_stopwords()
    lowered = normalize_technical_terms(text)
    cleaned = re.sub(r"[^a-z0-9+#.\s]", " ", lowered)
    tokens = cleaned.split()
    filtered_tokens = [token for token in tokens if token not in stop_words and len(token) > 1]
    return " ".join(filtered_tokens)


def extract_skills(text: str, skill_list: List[str] = SKILLS) -> List[str]:
    processed_text = preprocess_text(text)
    found_skills = []

    for skill in skill_list:
        skill_lower = skill.lower()
        if " " in skill_lower:
            if skill_lower in processed_text:
                found_skills.append(skill)
        else:
            if re.search(rf"\b{re.escape(skill_lower)}\b", processed_text):
                found_skills.append(skill)

    return sorted(set(found_skills), key=str.lower)


def extract_priority_keywords(job_description: str, top_n: int = 20) -> List[str]:
    tokens = preprocess_text(job_description).split()
    filtered = [
        token
        for token in tokens
        if len(token) >= 3 and not token.isdigit() and token not in JD_GENERIC_TERMS
    ]
    most_common = Counter(filtered).most_common(top_n)
    return [keyword for keyword, _ in most_common]


def calculate_semantic_score(resume_text: str, job_description: str) -> float:
    processed_resume = preprocess_text(resume_text)
    processed_jd = preprocess_text(job_description)

    if not processed_resume.strip() or not processed_jd.strip():
        return 0.0

    word_vectorizer = TfidfVectorizer(ngram_range=(1, 2), sublinear_tf=True)
    word_vectors = word_vectorizer.fit_transform([processed_resume, processed_jd])
    word_similarity = float(cosine_similarity(word_vectors[0:1], word_vectors[1:2])[0][0]) * 100

    char_vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5), sublinear_tf=True)
    char_vectors = char_vectorizer.fit_transform([processed_resume, processed_jd])
    char_similarity = float(cosine_similarity(char_vectors[0:1], char_vectors[1:2])[0][0]) * 100

    score = (0.7 * word_similarity) + (0.3 * char_similarity)
    return round(max(0.0, min(score, 100.0)), 2)


def calculate_skill_score(resume_skills: List[str], jd_skills: List[str]) -> float | None:
    if not jd_skills:
        return None

    overlap = set(s.lower() for s in resume_skills) & set(s.lower() for s in jd_skills)
    score = (len(overlap) / max(len(jd_skills), 1)) * 100
    return round(score, 2)


def calculate_keyword_coverage_score(resume_text: str, job_description: str) -> Dict[str, object]:
    resume_tokens = set(preprocess_text(resume_text).split())
    jd_keywords = extract_priority_keywords(job_description)

    if not jd_keywords:
        return {"keyword_score": None, "missing_keywords": []}

    matched = [kw for kw in jd_keywords if kw in resume_tokens]
    missing = [kw for kw in jd_keywords if kw not in resume_tokens]
    score = (len(matched) / len(jd_keywords)) * 100

    return {
        "keyword_score": round(score, 2),
        "missing_keywords": missing[:12],
    }


def calculate_section_score(resume_text: str) -> Dict[str, object]:
    text = resume_text.lower()
    found_sections = []
    missing_sections = []

    for section, pattern in SECTION_PATTERNS.items():
        if re.search(pattern, text):
            found_sections.append(section)
        else:
            missing_sections.append(section)

    score = (len(found_sections) / len(SECTION_PATTERNS)) * 100
    return {
        "section_score": round(score, 2),
        "missing_sections": missing_sections,
    }


def calculate_impact_score(resume_text: str) -> float:
    text = resume_text.lower()
    hits = 0
    for pattern in IMPACT_PATTERNS:
        hits += len(re.findall(pattern, text))

    # Saturates quickly: 5+ strong impact indicators is considered excellent.
    score = min((hits / 5) * 100, 100)
    return round(score, 2)


def calculate_match_score(
    resume_text: str,
    job_description: str,
    resume_skills: List[str] | None = None,
    jd_skills: List[str] | None = None,
) -> float:
    semantic_score = calculate_semantic_score(resume_text, job_description)
    skill_score = calculate_skill_score(resume_skills or [], jd_skills or [])
    keyword_data = calculate_keyword_coverage_score(resume_text, job_description)
    keyword_score = keyword_data["keyword_score"]
    section_score = calculate_section_score(resume_text)["section_score"]
    impact_score = calculate_impact_score(resume_text)

    weighted_components = [
        (semantic_score, 0.45),
        (keyword_score, 0.20),
        (section_score, 0.20),
        (impact_score, 0.15),
    ]

    if skill_score is not None:
        weighted_components.append((skill_score, 0.20))

    numerator = sum(value * weight for value, weight in weighted_components if value is not None)
    denominator = sum(weight for value, weight in weighted_components if value is not None)

    if denominator == 0:
        return 0.0

    score = numerator / denominator
    return round(max(0.0, min(score, 100.0)), 2)


def generate_suggestion(
    missing_skills: List[str],
    missing_keywords: List[str],
    missing_sections: List[str],
    score: float,
) -> str:
    if not missing_skills and not missing_keywords and score >= 80:
        return "Your resume aligns well. Add measurable achievements to make it stronger."

    if missing_skills or missing_keywords or missing_sections:
        improvements = []

        if missing_skills:
            top_missing_skills = ", ".join(missing_skills[:3])
            improvements.append(f"show project/work evidence for: {top_missing_skills}")

        if missing_keywords:
            top_missing_keywords = ", ".join(missing_keywords[:5])
            improvements.append(f"include JD keywords naturally: {top_missing_keywords}")

        if missing_sections:
            top_missing_sections = ", ".join(missing_sections)
            improvements.append(f"add clear sections: {top_missing_sections}")

        guidance = "; ".join(improvements)
        return f"Improve match by updating resume content: {guidance}. Use quantified achievements in bullets."

    if missing_skills:
        top_missing = ", ".join(missing_skills[:3])
        return (
            f"Include project or work experience showing these skills: {top_missing}. "
            "Use concise bullet points with impact metrics."
        )

    return "Improve match by adding role-specific keywords and clearer technical project outcomes."


def extract_certifications(text: str) -> List[str]:
    text_lower = text.lower()
    found_certs = []
    for cert in CERTIFICATIONS:
        if cert.lower() in text_lower:
            found_certs.append(cert.title())
    return sorted(set(found_certs))


def extract_languages(text: str) -> List[str]:
    text_lower = text.lower()
    found_langs = []
    lang_keywords = {
        "english": ["english", "fluent"],
        "spanish": ["spanish"],
        "french": ["french"],
        "german": ["german"],
        "chinese": ["chinese", "mandarin"],
        "japanese": ["japanese"],
    }
    for lang, keywords in lang_keywords.items():
        for kw in keywords:
            if kw in text_lower:
                found_langs.append(lang.capitalize())
                break
    return sorted(set(found_langs))


def infer_job_role(job_description: str) -> str:
    jd_lower = job_description.lower()
    role_scores = {}

    for role, profile in ROLE_PROFILES.items():
        score = 0
        for keyword in profile["keywords"]:
            if keyword.lower() in jd_lower:
                score += 2
        role_scores[role] = score

    if max(role_scores.values()) > 0:
        return max(role_scores, key=role_scores.get)
    return "full_stack_engineer"


def calculate_role_weighted_score(
    resume_skills: List[str], role: str, jd_skills: List[str]
) -> float:
    if role not in ROLE_PROFILES:
        return 0.0

    profile = ROLE_PROFILES[role]
    weights = profile["weights"]

    resume_skills_lower = set(s.lower() for s in resume_skills)
    jd_skills_lower = set(s.lower() for s in jd_skills)

    weighted_sum = 0.0
    weight_count = 0.0

    for skill in jd_skills_lower:
        weight = weights.get(skill, 0.5)
        if skill in resume_skills_lower:
            weighted_sum += weight * 100
        weight_count += weight * 100

    if weight_count == 0:
        return 0.0

    return round(weighted_sum / weight_count, 2)


def analyze_resume_against_jd(resume_text: str, job_description: str) -> Dict[str, object]:
    resume_skills = extract_skills(resume_text)
    jd_skills = extract_skills(job_description)
    missing_skills = sorted(set(jd_skills) - set(resume_skills), key=str.lower)

    certifications = extract_certifications(resume_text)
    languages = extract_languages(resume_text)
    detected_role = infer_job_role(job_description)
    role_weighted_score = calculate_role_weighted_score(resume_skills, detected_role, jd_skills)

    keyword_data = calculate_keyword_coverage_score(resume_text, job_description)
    section_data = calculate_section_score(resume_text)
    semantic_score = calculate_semantic_score(resume_text, job_description)
    skill_score = calculate_skill_score(resume_skills, jd_skills)
    impact_score = calculate_impact_score(resume_text)

    score = calculate_match_score(resume_text, job_description, resume_skills, jd_skills)

    return {
        "skills_found": resume_skills,
        "missing_skills": missing_skills,
        "missing_keywords": keyword_data["missing_keywords"],
        "missing_sections": section_data["missing_sections"],
        "certifications": certifications,
        "languages": languages,
        "detected_role": detected_role.replace("_", " ").title(),
        "role_weighted_score": role_weighted_score,
        "score_breakdown": {
            "semantic": semantic_score,
            "skills": skill_score if skill_score is not None else "N/A",
            "keywords": keyword_data["keyword_score"] if keyword_data["keyword_score"] is not None else "N/A",
            "sections": section_data["section_score"],
            "impact": impact_score,
            "role_fit": role_weighted_score,
        },
        "match_score": score,
        "suggestion": generate_suggestion(
            missing_skills,
            keyword_data["missing_keywords"],
            section_data["missing_sections"],
            score,
        ),
    }
