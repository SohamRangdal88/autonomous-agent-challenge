"""
Autonomous Document Agent
FastAPI application that accepts natural-language business requests,
decomposes them into structured execution plans, executes tasks sequentially,
and generates polished Microsoft Word (.docx) documents.
"""

from __future__ import annotations

import os
import re
import json
import time
import uuid
import logging
from datetime import datetime, timezone
from typing import List, Optional, Tuple

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request, status
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field, field_validator
from openai import OpenAI, APIError, RateLimitError, APIConnectionError

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# -----------------------------------------------------------------------------
# Configuration and Logging
# -----------------------------------------------------------------------------

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("autonomous_agent")

GROQ_API_KEY: str = os.getenv("GROQ_API_KEY", "").strip()
GROQ_BASE_URL: str = os.getenv("GROQ_BASE_URL", "https://api.groq.com/openai/v1").strip()
GROQ_MODEL: str = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile").strip()
LLM_TEMPERATURE: float = float(os.getenv("LLM_TEMPERATURE", "0.2"))
LLM_MAX_RETRIES: int = int(os.getenv("LLM_MAX_RETRIES", "3"))
OUTPUT_DIR: str = os.getenv("OUTPUT_DIR", "outputs").strip()

os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI(
    title="Autonomous Document Agent API",
    description="Multi-step autonomous AI agent for generating executive-grade business documents.",
    version="1.0.0"
)

# Allow all origins in development; restrict to your domain in production via env
ALLOWED_ORIGINS = os.getenv("ALLOWED_ORIGINS", "*").split(",")
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -----------------------------------------------------------------------------
# Pydantic Schemas
# -----------------------------------------------------------------------------

class AgentRequest(BaseModel):
    request: str = Field(
        ...,
        description="Natural-language business request or prompt for the document agent."
    )

    @field_validator("request")
    @classmethod
    def validate_request_content(cls, v: str) -> str:
        if not v or not v.strip():
            raise ValueError("Request string cannot be empty or whitespace only.")
        if len(v) > 5000:
            raise ValueError(f"Request exceeds maximum allowed length of 5,000 characters (received {len(v)}).")
        return v.strip()


class TaskPlan(BaseModel):
    title: str = Field(..., min_length=3, description="Concise, descriptive title of the task")
    objective: str = Field(..., min_length=10, description="Clear, actionable goal for this step")


class ExecutionPlan(BaseModel):
    document_title: str = Field(..., min_length=3, description="Formal document title")
    document_type: str = Field(..., min_length=3, description="Document type (e.g. Business Proposal, Project Plan)")
    assumptions: List[str] = Field(..., min_length=1, description="List of baseline planning assumptions")
    tasks: List[TaskPlan] = Field(
        ...,
        min_length=3,
        max_length=5,
        description="Ordered sequence of 3 to 5 discrete tasks"
    )

    @field_validator("assumptions", mode="before")
    @classmethod
    def coerce_assumptions_list(cls, v):
        if isinstance(v, str):
            return [line.strip("- *").strip() for line in v.splitlines() if line.strip()]
        return v


class TaskExecutionResult(BaseModel):
    task_id: int = Field(..., description="1-based index of task execution")
    title: str
    objective: str
    status: str = Field(..., description="'completed' or 'failed'")
    output: str = Field(..., description="Detailed content produced during this step")
    error: Optional[str] = None


class AgentResponse(BaseModel):
    run_id: str
    status: str = Field(..., description="'success', 'partial_failure', or 'failed'")
    plan: ExecutionPlan
    tasks: List[TaskExecutionResult]
    summary: str
    document_url: str
    filename: str
    created_at: str


# -----------------------------------------------------------------------------
# Core LLM Client Factory
# -----------------------------------------------------------------------------

def get_llm_client() -> OpenAI:
    """Initialize and return an OpenAI client configured for the Groq API."""
    api_key = GROQ_API_KEY or "not_configured"
    return OpenAI(
        api_key=api_key,
        base_url=GROQ_BASE_URL,
    )


# -----------------------------------------------------------------------------
# Modular Function 1: Calling the LLM with Bounded Retry
# -----------------------------------------------------------------------------

def call_llm(
    client: OpenAI,
    messages: List[dict],
    model: Optional[str] = None,
    temperature: Optional[float] = None,
    json_mode: bool = False,
    max_retries: Optional[int] = None,
) -> str:
    """
    Call the LLM with bounded retries and exponential backoff.
    Guarantees no infinite retry loops, cleans response formatting,
    and never leaks API keys into error logs.
    """
    target_model = model or GROQ_MODEL
    target_temp = temperature if temperature is not None else LLM_TEMPERATURE
    retries = max_retries if max_retries is not None else LLM_MAX_RETRIES

    if not GROQ_API_KEY or GROQ_API_KEY == "your_groq_api_key_here":
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=(
                "GROQ_API_KEY is not configured. Please supply a valid Groq API key in your .env file "
                "or environment variable to execute live AI generation."
            )
        )

    last_error: Optional[Exception] = None

    for attempt in range(1, retries + 1):
        try:
            kwargs = {
                "model": target_model,
                "messages": messages,
                "temperature": target_temp,
            }
            if json_mode:
                kwargs["response_format"] = {"type": "json_object"}

            response = client.chat.completions.create(**kwargs)
            content = response.choices[0].message.content or ""
            content = content.strip()

            if not content:
                raise ValueError("LLM returned an empty response.")

            return content

        except (RateLimitError, APIConnectionError) as e:
            last_error = e
            wait_time = 2 ** attempt
            logger.warning(
                "Transient LLM error on attempt %d/%d: %s. Retrying in %ds...",
                attempt, retries, type(e).__name__, wait_time
            )
            time.sleep(wait_time)
        except APIError as e:
            last_error = e
            logger.error("Non-recoverable LLM API error: %s (Status: %s)", type(e).__name__, getattr(e, "status_code", "N/A"))
            break
        except Exception as e:
            last_error = e
            logger.warning("Unexpected error during LLM invocation (attempt %d/%d): %s", attempt, retries, str(e))
            if attempt < retries:
                time.sleep(1.5)

    error_msg = f"LLM invocation failed after {retries} attempts: {str(last_error)}"
    logger.error(error_msg)
    raise RuntimeError(error_msg)


# -----------------------------------------------------------------------------
# Modular Function 2: Creating and Validating the Plan
# -----------------------------------------------------------------------------

def create_and_validate_plan(
    client: OpenAI,
    model: str,
    request_text: str,
) -> ExecutionPlan:
    """
    Generate and validate a structured execution plan using the LLM.
    Enforces between 3 and 5 tasks, explicit assumptions, and strict schema compliance.
    """
    system_prompt = (
        "You are an expert AI strategic planner and systems architect.\n"
        "Your task is to decompose the user's business request into a structured, sequential execution plan.\n"
        "You MUST output valid, parseable JSON conforming strictly to this schema:\n"
        "{\n"
        '  "document_title": "Descriptive Professional Title",\n'
        '  "document_type": "Business Proposal | Project Plan | Technical Architecture | Strategy Memo",\n'
        '  "assumptions": ["Explicit assumption 1", "Explicit assumption 2", ...],\n'
        '  "tasks": [\n'
        '    {"title": "Step 1: ...", "objective": "Detailed objective defining deliverables and scope"},\n'
        '    {"title": "Step 2: ...", "objective": "Detailed objective defining deliverables and scope"},\n'
        '    {"title": "Step 3: ...", "objective": "Detailed objective defining deliverables and scope"}\n'
        "  ]\n"
        "}\n\n"
        "Rules:\n"
        "1. Exactly between 3 and 5 tasks are permitted. Do NOT generate fewer than 3 or more than 5 tasks.\n"
        "2. The plan must adapt dynamically to the user's unique context, domain, and constraints.\n"
        "3. Cover requirements/scope, core architecture/strategy, budget/timelines with assumptions, risks, and review.\n"
        "4. Formulate realistic, explicit assumptions (e.g. currency in INR if requested, student scale, timeline).\n"
        "5. Do NOT assume external real-world actions, live network queries, or surveys took place.\n"
        "6. Return ONLY the raw JSON object, without introductory chatter."
    )

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": f"Create an execution plan for this business request:\n\n{request_text}"}
    ]

    raw_response = call_llm(client, messages, model=model, temperature=0.1, json_mode=True)

    # Clean markdown fences if model inadvertently wrapped json in ```json ... ```
    cleaned_json = raw_response.strip()
    if cleaned_json.startswith("```"):
        cleaned_json = re.sub(r"^```(?:json)?\s*", "", cleaned_json)
        cleaned_json = re.sub(r"\s*```$", "", cleaned_json)

    try:
        data = json.loads(cleaned_json)
    except json.JSONDecodeError as jde:
        logger.error("Failed to parse JSON execution plan: %s\nRaw output: %s", jde, raw_response)
        raise ValueError(f"Malformed LLM JSON response during planning: {str(jde)}")

    try:
        plan = ExecutionPlan.model_validate(data)
        return plan
    except Exception as pe:
        logger.error("Pydantic validation error for plan: %s", pe)
        raise ValueError(f"Plan validation failed: {str(pe)}")


# -----------------------------------------------------------------------------
# Modular Function 3: Executing One Task
# -----------------------------------------------------------------------------

def execute_task(
    client: OpenAI,
    model: str,
    original_request: str,
    assumptions: List[str],
    task: TaskPlan,
    previous_outputs: List[TaskExecutionResult],
) -> str:
    """
    Execute a single planned task sequentially using the LLM.
    Receives original request, planning assumptions, and results of all earlier steps.
    """
    system_prompt = (
        "You are an executive-level business consultant, technical architect, and strategic writer.\n"
        "You are executing one discrete task in an autonomous multi-step document generation pipeline.\n\n"
        "Strict Guidelines:\n"
        "1. Provide rich, exhaustive, professional content suitable for inclusion in an executive Word document.\n"
        "2. Build upon and maintain consistency with the outputs of all previous tasks provided in the prompt.\n"
        "3. Use structured formatting with clear headings, subheadings, bullet lists, and tabular summaries where appropriate.\n"
        "4. Guardrails: Do NOT claim that external real-world actions, live internet searches, vendor meetings, or "
        "live data collection took place unless actually executed. State assumptions and illustrative numbers clearly.\n"
        "5. If financial figures or budgets are included, explicitly label them as 'Illustrative Estimates'.\n"
        "6. Do not include meta-chatter (e.g. 'Here is your task output:'). Start directly with the deliverable content."
    )

    prior_context_blocks = []
    for prev in previous_outputs:
        status_marker = "COMPLETED" if prev.status == "completed" else "FAILED"
        preview = prev.output if len(prev.output) < 1800 else prev.output[:1800] + "\n...[truncated for length]"
        prior_context_blocks.append(
            f"--- [Earlier Task #{prev.task_id}: {prev.title} ({status_marker})] ---\n{preview}"
        )

    context_str = "\n\n".join(prior_context_blocks) if prior_context_blocks else "No prior tasks; this is Step 1."
    assumptions_str = "\n".join([f"- {a}" for a in assumptions])

    user_prompt = (
        f"### ORIGINAL BUSINESS REQUEST:\n{original_request}\n\n"
        f"### PLANNING ASSUMPTIONS:\n{assumptions_str}\n\n"
        f"### PREVIOUS TASK OUTPUTS (Accumulated Context):\n{context_str}\n\n"
        f"### CURRENT TASK TO EXECUTE:\n"
        f"Title: {task.title}\n"
        f"Objective: {task.objective}\n\n"
        f"Please execute this task thoroughly, providing formal deliverable content for this document section."
    )

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]

    return call_llm(client, messages, model=model, temperature=0.2, json_mode=False)


# -----------------------------------------------------------------------------
# Modular Function 4: Generating the Polished Word Document
# -----------------------------------------------------------------------------

def set_cell_background(cell, hex_color: str):
    """Set background color of a table cell via OXML."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set inner padding for a table cell in twips."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def add_markdown_content_to_doc(doc: docx.Document, text: str):
    """
    Parse markdown formatted text lines from LLM task outputs into styled docx elements.
    Converts markdown headers, bullets, and paragraphs into proper Word typography.
    """
    lines = text.split("\n")
    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        # Markdown Headings
        if line.startswith("### "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(line[4:].strip())
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(51, 65, 85)  # Slate
        elif line.startswith("## "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line[3:].strip())
            run.bold = True
            run.font.size = Pt(14)
            run.font.color.rgb = RGBColor(37, 99, 235)  # Blue
        elif line.startswith("# "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line[2:].strip())
            run.bold = True
            run.font.size = Pt(15)
            run.font.color.rgb = RGBColor(30, 58, 138)  # Navy
        elif line.startswith("- ") or line.startswith("* "):
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(3)
            bullet_text = line[2:].strip()
            # Handle inline bolding **term**
            _append_formatted_runs(p, bullet_text)
        elif re.match(r"^\d+\.\s+", line):
            p = doc.add_paragraph(style="List Number")
            p.paragraph_format.space_after = Pt(3)
            num_text = re.sub(r"^\d+\.\s+", "", line).strip()
            _append_formatted_runs(p, num_text)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.line_spacing = 1.15
            _append_formatted_runs(p, line)


def _append_formatted_runs(paragraph, text: str):
    """Helper to render inline **bold** tags into distinct Word runs."""
    parts = re.split(r"(\*\*.*?\*\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            paragraph.add_run(part)


def generate_word_document(
    run_id: str,
    request_text: str,
    plan: ExecutionPlan,
    task_results: List[TaskExecutionResult],
    output_dir: str = OUTPUT_DIR,
) -> Tuple[str, str]:
    """
    Generate a polished, executive-ready Microsoft Word (.docx) document
    featuring professional typography, custom color palette, structured tables,
    and all mandatory challenge sections.
    """
    doc = docx.Document()

    # Configure Standard Page Margins (1 inch all sides)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # ---------------------------------------------------------
    # 1. Professional Title Section
    # ---------------------------------------------------------
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(10)
    title_p.paragraph_format.space_after = Pt(4)
    title_run = title_p.add_run(plan.document_title)
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(26)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)  # Deep Navy (#1E3A8A)

    subtitle_p = doc.add_paragraph()
    subtitle_p.paragraph_format.space_after = Pt(16)
    sub_run = subtitle_p.add_run(f"{plan.document_type}  |  Autonomous AI Document Engine")
    sub_run.font.name = "Calibri"
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(100, 116, 139)  # Slate Gray

    # Metadata Summary Bar
    meta_p = doc.add_paragraph()
    meta_p.paragraph_format.space_after = Pt(16)
    timestamp_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    overall_status = "SUCCESS" if all(t.status == "completed" for t in task_results) else "PARTIAL FAILURE"

    runs_data = [
        ("Run ID: ", True), (f"{run_id}    ", False),
        ("Date: ", True), (f"{timestamp_str}    ", False),
        ("Status: ", True), (f"{overall_status}", False),
    ]
    for label, is_bold in runs_data:
        r = meta_p.add_run(label)
        r.font.size = Pt(9.5)
        r.font.bold = is_bold
        r.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ---------------------------------------------------------
    # 2. Executive Overview & Original Request Callout
    # ---------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)
    h1_run = h1.add_run("1. Executive Overview & Context")
    h1_run.font.color.rgb = RGBColor(30, 58, 138)

    req_intro = doc.add_paragraph()
    req_intro.paragraph_format.space_before = Pt(4)
    req_intro.paragraph_format.space_after = Pt(4)
    req_label = req_intro.add_run("Original Business Request: ")
    req_label.bold = True
    req_intro.add_run(request_text)

    # Callout Box for Request
    callout_tbl = doc.add_table(rows=1, cols=1)
    callout_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    callout_tbl.autofit = False
    callout_cell = callout_tbl.cell(0, 0)
    callout_cell.width = Inches(6.5)
    set_cell_background(callout_cell, "F1F5F9")  # Soft light slate
    set_cell_margins(callout_cell, top=140, bottom=140, left=180, right=180)

    box_p = callout_cell.paragraphs[0]
    box_p.paragraph_format.space_after = Pt(2)
    b_label = box_p.add_run("Original Business Request:\n")
    b_label.bold = True
    b_label.font.size = Pt(10)
    b_label.font.color.rgb = RGBColor(30, 41, 59)
    b_text = box_p.add_run(request_text)
    b_text.font.size = Pt(9.5)
    b_text.font.italic = True
    b_text.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ---------------------------------------------------------
    # 3. Planning Assumptions
    # ---------------------------------------------------------
    h2 = doc.add_heading(level=1)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(6)
    h2_run = h2.add_run("2. Planning Assumptions & Constraints")
    h2_run.font.color.rgb = RGBColor(30, 58, 138)

    assump_intro = doc.add_paragraph()
    assump_intro.paragraph_format.space_after = Pt(4)
    assump_intro.add_run(
        "The autonomous agent formulated the following baseline assumptions to establish project boundaries, "
        "cost models, and implementation targets:"
    )

    for item in plan.assumptions:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(3)
        p.add_run(item)

    notice_p = doc.add_paragraph()
    notice_p.paragraph_format.space_before = Pt(4)
    notice_p.paragraph_format.space_after = Pt(12)
    notice_run = notice_p.add_run(
        "Note: All cost calculations, staffing levels, and timelines are modeled illustrative estimates "
        "synthesized based on specified requirements. Verified empirical statistics should be benchmarked during the pilot phase."
    )
    notice_run.font.size = Pt(9)
    notice_run.font.italic = True
    notice_run.font.color.rgb = RGBColor(100, 116, 139)

    # ---------------------------------------------------------
    # 4. Structured Execution Plan Table
    # ---------------------------------------------------------
    h3 = doc.add_heading(level=1)
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(6)
    h3_run = h3.add_run("3. Autonomous Execution Plan")
    h3_run.font.color.rgb = RGBColor(30, 58, 138)

    tbl = doc.add_table(rows=1, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    col_widths = [Inches(0.6), Inches(2.0), Inches(2.7), Inches(1.2)]
    headers = ["Step", "Task Title", "Objective", "Execution Status"]

    # Header Row
    hdr_cells = tbl.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], "1E3A8A")  # Deep Navy
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=100, right=100)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in (0, 3) else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    # Populate Tasks in Table
    for idx, t_res in enumerate(task_results, start=1):
        row_cells = tbl.add_row().cells
        bg_color = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
        for i, width in enumerate(col_widths):
            row_cells[i].width = width
            set_cell_background(row_cells[i], bg_color)
            set_cell_margins(row_cells[i], top=90, bottom=90, left=100, right=100)

        # Step #
        p0 = row_cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(str(idx))
        r0.font.size = Pt(9)
        r0.bold = True

        # Title
        p1 = row_cells[1].paragraphs[0]
        r1 = p1.add_run(t_res.title)
        r1.font.size = Pt(9)
        r1.bold = True

        # Objective
        p2 = row_cells[2].paragraphs[0]
        r2 = p2.add_run(t_res.objective)
        r2.font.size = Pt(8.5)

        # Status
        p3 = row_cells[3].paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = p3.add_run(t_res.status.upper())
        r3.font.size = Pt(8.5)
        r3.bold = True
        if t_res.status == "completed":
            r3.font.color.rgb = RGBColor(22, 101, 52)  # Dark Green
        else:
            r3.font.color.rgb = RGBColor(185, 28, 28)  # Red

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ---------------------------------------------------------
    # 5. Detailed Deliverables per Task
    # ---------------------------------------------------------
    h4 = doc.add_heading(level=1)
    h4.paragraph_format.space_before = Pt(14)
    h4.paragraph_format.space_after = Pt(6)
    h4_run = h4.add_run("4. Detailed Task Deliverables & Analysis")
    h4_run.font.color.rgb = RGBColor(30, 58, 138)

    for idx, t_res in enumerate(task_results, start=1):
        sub_h = doc.add_heading(level=2)
        sub_h.paragraph_format.space_before = Pt(12)
        sub_h.paragraph_format.space_after = Pt(3)
        sub_h_run = sub_h.add_run(f"4.{idx} {t_res.title}")
        sub_h_run.font.color.rgb = RGBColor(37, 99, 235)  # Slate Blue

        obj_p = doc.add_paragraph()
        obj_p.paragraph_format.space_after = Pt(6)
        obj_run = obj_p.add_run(f"Objective: {t_res.objective}")
        obj_run.font.size = Pt(9.5)
        obj_run.font.italic = True
        obj_run.font.color.rgb = RGBColor(71, 85, 105)

        if t_res.status == "failed":
            fail_p = doc.add_paragraph()
            fail_p.paragraph_format.space_after = Pt(6)
            frun = fail_p.add_run(f"[TASK EXECUTION FAILED: {t_res.error or 'Unknown error'}]")
            frun.bold = True
            frun.font.color.rgb = RGBColor(185, 28, 28)

        add_markdown_content_to_doc(doc, t_res.output)

    # ---------------------------------------------------------
    # 6. Strategic Recommendations & Risk Governance
    # ---------------------------------------------------------
    h5 = doc.add_heading(level=1)
    h5.paragraph_format.space_before = Pt(16)
    h5.paragraph_format.space_after = Pt(6)
    h5_run = h5.add_run("5. Strategic Recommendations & Risk Governance")
    h5_run.font.color.rgb = RGBColor(30, 58, 138)

    recs = [
        ("Phased Rollout Protocol", "Initiate with a closed internal sandbox or 10% user cohort before widespread release."),
        ("Fallback and Graceful Degradation", "Ensure instant, deterministic escalation to human operators when confidence falls below 80%."),
        ("Data Isolation & Compliance", "Enforce strict zero-retention policies on sensitive personally identifiable information (PII)."),
        ("Model Latency & Cost Optimization", "Implement semantic response caching and token budgets to manage recurring inference expenses.")
    ]
    for r_title, r_desc in recs:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(3)
        rt = p.add_run(f"{r_title}: ")
        rt.bold = True
        p.add_run(r_desc)

    # ---------------------------------------------------------
    # 7. Final Review & Quality Sign-Off
    # ---------------------------------------------------------
    h6 = doc.add_heading(level=1)
    h6.paragraph_format.space_before = Pt(16)
    h6.paragraph_format.space_after = Pt(6)
    h6_run = h6.add_run("6. Final Review & Governance Sign-Off")
    h6_run.font.color.rgb = RGBColor(30, 58, 138)

    review_p = doc.add_paragraph()
    review_p.paragraph_format.space_after = Pt(8)
    review_p.add_run(
        "This document was autonomously compiled through multi-step plan decomposition. "
        "All sections have been structured to ensure executive traceability, clear assumptions, and verifiable deliverables."
    )

    # Sign-Off Table
    sign_tbl = doc.add_table(rows=3, cols=3)
    sign_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_widths = [Inches(2.1), Inches(2.2), Inches(2.2)]
    sign_headers = ["Review Milestone", "Verification Status", "Stakeholder Sign-Off"]

    for i, htitle in enumerate(sign_headers):
        sign_tbl.rows[0].cells[i].width = sign_widths[i]
        set_cell_background(sign_tbl.rows[0].cells[i], "334155")  # Slate Gray
        set_cell_margins(sign_tbl.rows[0].cells[i], top=80, bottom=80, left=80, right=80)
        p = sign_tbl.rows[0].cells[i].paragraphs[0]
        r = p.add_run(htitle)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    milestones = [
        ("Requirements & Scope Validation", "Automated Checks Passed", "Project Lead"),
        ("Technical & Risk Architecture Review", "Assumptions Verified", "Solutions Architect")
    ]
    for row_idx, (m, st, role) in enumerate(milestones, start=1):
        r_cells = sign_tbl.rows[row_idx].cells
        for col_idx, text_val in enumerate([m, st, role]):
            r_cells[col_idx].width = sign_widths[col_idx]
            set_cell_background(r_cells[col_idx], "F8FAFC")
            set_cell_margins(r_cells[col_idx], top=80, bottom=80, left=80, right=80)
            p = r_cells[col_idx].paragraphs[0]
            r = p.add_run(text_val)
            r.font.size = Pt(8.5)

    # Save to disk
    filename = f"document_{run_id}.docx"
    filepath = os.path.join(output_dir, filename)
    doc.save(filepath)
    logger.info("Successfully generated Word document at %s (size: %d bytes)", filepath, os.path.getsize(filepath))
    return filepath, filename


# -----------------------------------------------------------------------------
# Summary Generator
# -----------------------------------------------------------------------------

def generate_summary(plan: ExecutionPlan, task_results: List[TaskExecutionResult]) -> str:
    """Produce a concise, factual executive summary of run execution."""
    completed = sum(1 for t in task_results if t.status == "completed")
    total = len(task_results)
    assumptions_summary = "; ".join(plan.assumptions[:3])
    if len(plan.assumptions) > 3:
        assumptions_summary += f" (+{len(plan.assumptions)-3} more)"

    return (
        f"Generated '{plan.document_title}' ({plan.document_type}). "
        f"Completed {completed}/{total} planned tasks covering: "
        f"{', '.join([t.title for t in task_results])}. "
        f"Baseline assumptions established: {assumptions_summary}. "
        f"Comprehensive document formatted and ready for download."
    )


# -----------------------------------------------------------------------------
# Exception Handlers
# -----------------------------------------------------------------------------

from fastapi.exceptions import RequestValidationError

@app.exception_handler(RequestValidationError)
async def request_validation_exception_handler(request: Request, exc: RequestValidationError):
    errors = exc.errors()
    error_msg = errors[0].get("msg", "Validation error") if errors else "Validation error"
    if error_msg.startswith("Value error, "):
        error_msg = error_msg[len("Value error, "):]
    return JSONResponse(
        status_code=status.HTTP_400_BAD_REQUEST,
        content={"error": "Bad Request", "detail": error_msg, "errors": [
            {"loc": e.get("loc"), "msg": e.get("msg"), "type": e.get("type")} for e in errors
        ]}
    )

@app.exception_handler(ValueError)
async def value_error_handler(request: Request, exc: ValueError):
    return JSONResponse(
        status_code=status.HTTP_400_BAD_REQUEST,
        content={"error": "Invalid Request", "detail": str(exc)}
    )


# -----------------------------------------------------------------------------
# API Endpoints
# -----------------------------------------------------------------------------

@app.get("/health", status_code=status.HTTP_200_OK, tags=["System"])
def health_check():
    """Health check endpoint returning service status and active model configuration."""
    return {
        "status": "healthy",
        "service": "Autonomous Document Agent",
        "model": GROQ_MODEL,
        "base_url": GROQ_BASE_URL,
        "timestamp": datetime.now(timezone.utc).isoformat()
    }


@app.post("/agent", response_model=AgentResponse, status_code=status.HTTP_200_OK, tags=["Agent"])
def run_agent(body: AgentRequest):
    """
    Main Autonomous Agent Endpoint:
    1. Validates business request.
    2. Decomposes request into a 3-5 task execution plan via LLM.
    3. Executes each task sequentially, feeding earlier outputs into later steps.
    4. Compiles a professionally formatted Word document (.docx).
    5. Returns run status, plan, detailed task results, and document download URL.
    """
    request_text = body.request
    run_id = str(uuid.uuid4())
    logger.info("Initiating autonomous agent run %s for request (len=%d)", run_id, len(request_text))

    client = get_llm_client()

    # Stage B: Planning
    try:
        plan = create_and_validate_plan(client, GROQ_MODEL, request_text)
    except Exception as e:
        logger.error("Planning phase failed for run %s: %s", run_id, e)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Autonomous planning failed: {str(e)}"
        )

    # Stage C: Sequential Execution
    task_results: List[TaskExecutionResult] = []
    has_failure = False

    for idx, task in enumerate(plan.tasks, start=1):
        logger.info("Run %s: Executing Task %d/%d - '%s'", run_id, idx, len(plan.tasks), task.title)
        try:
            task_output = execute_task(
                client=client,
                model=GROQ_MODEL,
                original_request=request_text,
                assumptions=plan.assumptions,
                task=task,
                previous_outputs=task_results
            )
            task_results.append(
                TaskExecutionResult(
                    task_id=idx,
                    title=task.title,
                    objective=task.objective,
                    status="completed",
                    output=task_output
                )
            )
        except Exception as te:
            has_failure = True
            err_msg = str(te)
            logger.error("Run %s: Task %d '%s' failed: %s", run_id, idx, task.title, err_msg)
            task_results.append(
                TaskExecutionResult(
                    task_id=idx,
                    title=task.title,
                    objective=task.objective,
                    status="failed",
                    output=f"Task execution failed due to an error: {err_msg}",
                    error=err_msg
                )
            )

    # Stage D: Word Document Generation
    try:
        filepath, filename = generate_word_document(
            run_id=run_id,
            request_text=request_text,
            plan=plan,
            task_results=task_results,
            output_dir=OUTPUT_DIR
        )
    except Exception as de:
        logger.error("Document generation failed for run %s: %s", run_id, de)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Word document generation failed: {str(de)}"
        )

    # Overall Run Status determination
    if all(t.status == "completed" for t in task_results):
        overall_status = "success"
    elif any(t.status == "completed" for t in task_results):
        overall_status = "partial_failure"
    else:
        overall_status = "failed"

    summary = generate_summary(plan, task_results)

    return AgentResponse(
        run_id=run_id,
        status=overall_status,
        plan=plan,
        tasks=task_results,
        summary=summary,
        document_url=f"/documents/{filename}",
        filename=filename,
        created_at=datetime.now(timezone.utc).isoformat()
    )


@app.get("/documents/{filename}", tags=["Documents"])
def download_document(filename: str):
    """
    Download a generated Word document by filename.
    Includes strict regex validation and canonical path verification to prevent path traversal.
    """
    # Strict filename pattern matching: only alphanumeric, hyphen, underscore, ending in .docx
    if not re.match(r"^[a-zA-Z0-9_\-]+\.docx$", filename):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Invalid filename format. Only alphanumeric characters, hyphens, and underscores ending with .docx are allowed."
        )

    output_abs = os.path.abspath(OUTPUT_DIR)
    target_path = os.path.abspath(os.path.join(output_abs, filename))

    # Path traversal check: verify target path stays within OUTPUT_DIR
    if not target_path.startswith(output_abs + os.sep) and target_path != output_abs:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Path traversal detected. Access denied."
        )

    if not os.path.isfile(target_path):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Document '{filename}' not found."
        )

    return FileResponse(
        path=target_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename
    )


if __name__ == "__main__":
    import uvicorn
    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "8000"))
    uvicorn.run("main:app", host=host, port=port, reload=True)
