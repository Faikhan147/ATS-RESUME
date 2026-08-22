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

Evaluate the resume exactly as written.

Consider:
- Skills section
- Experience section
- Projects section
- Certifications
- Keywords
- Responsibilities

Do not assume experience that is not present.

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

Your job is to improve ONLY the wording, ordering, and JD alignment of:

1. Professional Summary
2. Skills
3. Experience
4. Project Titles
5. Project Bullet Points

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
- Projects facts
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

IMPORTANT — MANDATORY REWRITING:

Professional Summary MUST ALWAYS be rewritten.

Skills section MUST ALWAYS be rewritten and optimized for the JD.

Experience section MUST ALWAYS be rewritten.

Project titles and project bullet points MUST ALWAYS be rewritten.

Even if a section already looks good,
rewrite it to better align with the target JD.

SKILLS SECTION — MANDATORY JD OPTIMIZATION:

The Skills section MUST NOT be left unchanged.

You MUST:

1. Extract ALL skills and technologies supported by evidence anywhere
   in the resume.

2. Compare those supported skills against the target JD.

3. Reorder the Skills section according to JD relevance.

4. Put the most JD-relevant supported skills first.

5. ADD supported skills that are missing from the current Skills section
   when those skills are explicitly supported elsewhere in the resume.

6. Evidence may come from:
   - Professional Summary
   - Experience
   - Projects
   - Existing Skills
   - Certifications

7. Examples:
   - Helm if used in Projects
   - Bash if used in Projects
   - Python if used in Projects
   - Loki if used in Projects
   - SonarQube if used in Projects
   - Trivy if used in Projects
   - CI/CD if Jenkins or GitHub Actions pipelines are present
   - Infrastructure as Code (IaC) if Terraform is present
   - Container Orchestration if Docker/Kubernetes/EKS/ECS are present
   - Monitoring & Observability if Prometheus/Grafana/CloudWatch are present

8. NEVER add a technology just because it appears in the JD.

9. NEVER invent unsupported skills, tools, technologies, experience,
   achievements, or responsibilities.

10. Preserve the existing Skills categories where possible.

11. Reorder skills inside the existing categories based on JD relevance.

12. Add supported skills to the most appropriate existing category.

13. Return the COMPLETE replacement text for the Skills section.

14. Do NOT return only individual skill additions.

15. The final Skills section must be meaningfully optimized for the
    target JD, not merely copied from the original resume.

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
      "section": "Professional Summary",
      "new_text": "..."
    }},
    {{
      "paragraph_index": 0,
      "section": "Skills",
      "new_text": "..."
    }}
  ]
}}

MANDATORY OUTPUT REQUIREMENT:

The "replacements" array MUST contain at least one replacement
for each of these sections:

- Professional Summary
- Skills
- Experience
- Project Titles
- Project Bullet Points

For the Skills section, return the COMPLETE optimized Skills text.

Do not omit the Skills replacement.

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
# OPTIMIZE RESUME
# ============================================================

def optimize_resume(
    original_docx,
    jd_path,
    target_score=80,
    max_attempts=3
):

    current_docx = original_docx

    best_score = 0
    best_resume = original_docx

    for attempt in range(1, max_attempts + 1):

        print(f"ATTEMPT {attempt}")

        extract_docx(
            current_docx,
            f"resume_{attempt}.json"
        )

        ats_analysis(
            f"resume_{attempt}.json",
            jd_path,
            f"ats_{attempt}.json"
        )

        result = load_json(
            f"ats_{attempt}.json"
        )

        score = result["ats_score"]

        print(f"ATS = {score}")

        if score > best_score:
            best_score = score
            best_resume = current_docx

        if score >= target_score:
            print("TARGET REACHED")
            break

        rewrite_resume(
            current_docx,
            f"resume_{attempt}.json",
            jd_path,
            f"rewrite_{attempt}.json"
        )

        next_docx = (
            f"resume_optimized_{attempt}.docx"
        )

        apply_rewrite(
            current_docx,
            f"rewrite_{attempt}.json",
            next_docx
        )

        current_docx = next_docx

    shutil.copy(
        best_resume,
        "best_resume.docx"
    )

    print(
        f"BEST ATS SCORE = {best_score}"
    )


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

python3 resume_ai.py optimize input.docx jd.txt
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

    elif command == "optimize":

        optimize_resume(
            sys.argv[2],
            sys.argv[3]
        )

    else:

        raise RuntimeError(
            f"Unknown command: {command}"
        )


if __name__ == "__main__":
    main()
