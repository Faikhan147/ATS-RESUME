import os
import sys
import json
import re
import shutil
from pathlib import Path

from docx import Document
from openai import OpenAI


MODEL = "gpt-5.6-luna"


# ============================================================
# BASIC HELPERS
# ============================================================

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def get_client():
    api_key = os.environ.get("OPENAI_API_KEY")

    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY environment variable is missing."
        )

    return OpenAI(api_key=api_key)


# ============================================================
# DOCX EXTRACTION
# ============================================================

def extract_docx(docx_path, output_path):

    doc = Document(docx_path)

    paragraphs = []

    for index, paragraph in enumerate(doc.paragraphs):

        text = paragraph.text.strip()

        if not text:
            continue

        paragraphs.append({
            "index": index,
            "text": text,
            "style": paragraph.style.name if paragraph.style else "",
            "alignment": str(paragraph.alignment),
        })

    data = {
        "paragraphs": paragraphs
    }

    save_json(output_path, data)

    print(f"Extracted {len(paragraphs)} paragraphs.")


# ============================================================
# OPENAI JSON EXTRACTION
# ============================================================

def get_response_text(response):

    # Current Responses API SDK exposes output_text.
    if hasattr(response, "output_text"):
        return response.output_text

    # Fallback
    return str(response)


def clean_json_text(text):

    text = text.strip()

    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text)
        text = re.sub(r"```$", "", text)
        text = text.strip()

    return text


# ============================================================
# ATS ANALYSIS
# ============================================================

def ats_analysis(resume_json_path, jd_path, output_path):

    resume = load_json(resume_json_path)

    with open(jd_path, "r", encoding="utf-8") as f:
        jd = f.read()

    prompt = f"""
You are an ATS resume evaluation system.

Compare the resume against the job description.

IMPORTANT:

The resume is real candidate data.

DO NOT assume the candidate has skills that are not present.

Return ONLY valid JSON.

Required JSON:

{{
  "ats_score": 0,
  "matched_skills": [],
  "missing_skills": [],
  "suggestions": []
}}

ATS_SCORE must be an integer from 0 to 100.

Evaluate:
- Required technical skills
- Tools
- Cloud platforms
- Programming/scripting
- DevOps/CI/CD
- Experience relevance
- Job title relevance
- Keywords
- Responsibilities

JOB DESCRIPTION:

{jd}

RESUME:

{json.dumps(resume, indent=2, ensure_ascii=False)}
"""

    client = get_client()

    response = client.responses.create(
        model=MODEL,
        input=prompt
    )

    text = clean_json_text(get_response_text(response))

    result = json.loads(text)

    if "ats_score" not in result:
        raise RuntimeError("AI response does not contain ats_score.")

    score = int(result["ats_score"])

    if score < 0 or score > 100:
        raise RuntimeError("Invalid ATS score returned by AI.")

    result["ats_score"] = score

    save_json(output_path, result)

    print(json.dumps(result, indent=2))


# ============================================================
# RESUME REWRITE
# ============================================================

def rewrite_resume(
    original_docx,
    resume_json_path,
    jd_path,
    output_path
):

    resume = load_json(resume_json_path)

    with open(jd_path, "r", encoding="utf-8") as f:
        jd = f.read()

    prompt = f"""
You are a professional ATS resume writer.

Your job is to improve ONLY the wording of:

1. Professional Summary
2. Skills
3. Experience

The original DOCX is the source of truth.

STRICT RULES:

DO NOT change:

- Candidate name
- Contact information
- Education
- Degree
- University
- Company names
- Job titles
- Employment dates
- Actual experience facts
- Projects
- Certifications
- Any factual information

DO NOT invent:
- Skills
- Tools
- Technologies
- Achievements
- Metrics
- Companies
- Projects
- Responsibilities

IMPORTANT:

Professional Summary MUST ALWAYS be rewritten.

Skills section MUST ALWAYS be rewritten.

Experience section MUST ALWAYS be rewritten.

Project titles and project bullet points MUST ALWAYS be rewritten.

Even if the section already looks good,
rewrite it to better align with the target JD.

ATS OPTIMIZATION RULES:

1. Reorder existing skills based on JD relevance.

2. Highlight JD-relevant technologies already present
   anywhere in the resume.

3. You may add missing skills ONLY when they are
   supported by evidence found elsewhere in the resume.

Evidence may come from:
- Experience
- Projects
- Existing Skills
- Certifications

Examples of VALID additions:

- Jenkins -> CI/CD Pipelines
- Jenkins -> Deployment Automation
- Terraform -> Infrastructure as Code (IaC)
- Docker + Kubernetes -> Container Orchestration
- Prometheus + Grafana -> Monitoring & Observability
- Trivy + SonarQube -> DevSecOps
- GitHub Actions -> CI/CD

Examples of INVALID additions:

- Jenkins -> ArgoCD
- Jenkins -> GitLab CI
- Jenkins -> Vault
- Docker -> OpenShift
- Kubernetes -> KEDA

unless explicitly supported by resume content.

DO NOT invent:

- Companies
- Projects
- Experience
- Employment dates
- Metrics
- Responsibilities
- Achievements

DO NOT claim the candidate used a tool
unless there is evidence somewhere in the resume.

Keep rewritten text approximately the same length.

Do not increase section length by more than 20%.

Preserve ATS friendliness.

The generated text must fit naturally into the existing resume layout.

Return ONLY valid JSON.

Format:

{{
  "replacements": [
    {{
      "paragraph_index": 0,
      "new_text": "..."
    }}
  ]
}}

Only return paragraphs that belong to:

- Professional Summary
- Skills
- Experience
- Project Titles
- Project Bullet Points

Formatting Preservation Rules:

- Do not create new sections.
- Do not create new tables.
- Do not create new headings.
- Do not change document structure.
- Only replace paragraph text.
- Assume original font, size, spacing,
  alignment and formatting will be preserved.

JOB DESCRIPTION:

{jd}

RESUME STRUCTURE:

{json.dumps(resume, indent=2, ensure_ascii=False)}
"""

    client = get_client()

    response = client.responses.create(
        model=MODEL,
        input=prompt
    )

    text = clean_json_text(get_response_text(response))

    result = json.loads(text)

    save_json(output_path, result)

    print(json.dumps(result, indent=2, ensure_ascii=False))


# ============================================================
# APPLY AI TEXT TO ORIGINAL DOCX
# ============================================================

def replace_paragraph_text(paragraph, new_text):

    runs = paragraph.runs

    if not runs:

        paragraph.add_run(new_text)
        return

    # Keep the first run's character formatting.
    runs[0].text = new_text

    # Remove text from remaining runs while preserving
    # the paragraph's style, spacing, alignment, etc.
    for run in runs[1:]:
        run.text = ""


def apply_rewrite(
    original_docx,
    rewritten_json,
    output_docx
):

    doc = Document(original_docx)

    rewritten = load_json(rewritten_json)

    replacements = rewritten.get("replacements", [])

    for item in replacements:

        index = int(item["paragraph_index"])
        new_text = item["new_text"]

        if index < 0 or index >= len(doc.paragraphs):
            print(
                f"WARNING: Paragraph index {index} does not exist."
            )
            continue

        paragraph = doc.paragraphs[index]

        print(
            f"Updating paragraph {index}: "
            f"{paragraph.text[:60]}..."
        )

        replace_paragraph_text(
            paragraph,
            new_text
        )

    doc.save(output_docx)

    print(f"Created: {output_docx}")


# ============================================================
# MAIN
# ============================================================

def main():

    if len(sys.argv) < 2:
        print("""
Usage:

python3 resume_ai.py extract input.docx output.json

python3 resume_ai.py ats resume.json jd.txt ats.json

python3 resume_ai.py rewrite input.docx resume.json jd.txt rewritten.json

python3 resume_ai.py apply input.docx rewritten.json output.docx
""")
        sys.exit(1)

    command = sys.argv[1]

    if command == "extract":

        extract_docx(
            sys.argv[2],
            sys.argv[3]
        )

    elif command == "ats":

        ats_analysis(
            sys.argv[2],
            sys.argv[3],
            sys.argv[4]
        )

    elif command == "rewrite":

        rewrite_resume(
            sys.argv[2],
            sys.argv[3],
            sys.argv[4],
            sys.argv[5]
        )

    elif command == "apply":

        apply_rewrite(
            sys.argv[2],
            sys.argv[3],
            sys.argv[4]
        )

    else:

        raise RuntimeError(
            f"Unknown command: {command}"
        )


if __name__ == "__main__":
    main()
