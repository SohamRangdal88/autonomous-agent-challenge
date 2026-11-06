"""
Automated Test Suite for Autonomous Document Agent
Covers all challenge requirements with 100% offline LLM mocking:
- Health check
- Empty request rejection
- Excessively long request rejection
- Valid plan parsing & schema validation
- Malformed plan rejection (invalid JSON, too few/too many tasks)
- Word document (.docx) creation, file existence, and structure
- Path traversal & invalid filename rejection
- Missing document 404 handling
- End-to-end multi-step agent execution
- Partial failure resilience
"""

import os
import sys
import json
import pytest
from unittest.mock import patch, MagicMock

# Add project root directory to sys.path so 'main' is always importable
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from fastapi.testclient import TestClient
import docx

from main import (
    app,
    create_and_validate_plan,
    generate_word_document,
    ExecutionPlan,
    TaskPlan,
    TaskExecutionResult,
    OUTPUT_DIR,
)

client = TestClient(app)

# -----------------------------------------------------------------------------
# Sample Fixtures
# -----------------------------------------------------------------------------

SAMPLE_VALID_PLAN_JSON = {
    "document_title": "AI-Powered Customer Support Chatbot Business Proposal",
    "document_type": "Business Proposal",
    "assumptions": [
        "Implementation timeline is strictly 6 weeks",
        "Target client is a mid-sized e-commerce retailer",
        "All cost estimates are in INR and labeled as illustrative"
    ],
    "tasks": [
        {
            "title": "Requirements Gathering and Scope Definition",
            "objective": "Identify primary customer support pain points, chat channels, and SLA targets."
        },
        {
            "title": "Technical Architecture and Groq LLM Selection",
            "objective": "Design system integration with Shopify, Groq inference pipeline, and CRM webhook routing."
        },
        {
            "title": "Implementation Roadmap and Budget in INR",
            "objective": "Build a week-by-week delivery schedule with illustrative cost projections in INR."
        },
        {
            "title": "Risk Analysis and Mitigation Strategy",
            "objective": "Analyze hallucination risks, failover mechanisms, data privacy, and human escalation."
        }
    ]
}


@pytest.fixture
def sample_plan() -> ExecutionPlan:
    return ExecutionPlan.model_validate(SAMPLE_VALID_PLAN_JSON)


@pytest.fixture
def sample_task_results() -> list[TaskExecutionResult]:
    return [
        TaskExecutionResult(
            task_id=1,
            title="Requirements Gathering and Scope Definition",
            objective="Identify primary customer support pain points, chat channels, and SLA targets.",
            status="completed",
            output="### Executive Scope\nThe chatbot will resolve Tier-1 customer queries with sub-2s response latency."
        ),
        TaskExecutionResult(
            task_id=2,
            title="Technical Architecture and Groq LLM Selection",
            objective="Design system integration with Shopify, Groq inference pipeline, and CRM webhook routing.",
            status="completed",
            output="### System Architecture\nFastAPI backend interfacing with Groq API utilizing `llama-3.3-70b-versatile`."
        ),
        TaskExecutionResult(
            task_id=3,
            title="Implementation Roadmap and Budget in INR",
            objective="Build a week-by-week delivery schedule with illustrative cost projections in INR.",
            status="completed",
            output="### 6-Week Delivery Schedule\n- Weeks 1-2: Core Integration\n- Weeks 3-4: Pilot Testing\n\n*Estimated Budget: ₹4,50,000 (Illustrative)*"
        ),
        TaskExecutionResult(
            task_id=4,
            title="Risk Analysis and Mitigation Strategy",
            objective="Analyze hallucination risks, failover mechanisms, data privacy, and human escalation.",
            status="completed",
            output="### Risk Governance\nHuman operator fallback enabled whenever confidence score < 80%."
        )
    ]


# -----------------------------------------------------------------------------
# Test 1: Successful Health Check
# -----------------------------------------------------------------------------

def test_health_check():
    """Verify GET /health returns 200 with service metadata and healthy status."""
    response = client.get("/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "healthy"
    assert "service" in data
    assert "model" in data
    assert "timestamp" in data


# -----------------------------------------------------------------------------
# Test 2: Empty Request Rejection
# -----------------------------------------------------------------------------

def test_empty_request_rejection():
    """Verify POST /agent rejects empty or whitespace-only request strings with HTTP 400."""
    # Test completely empty string
    res_empty = client.post("/agent", json={"request": ""})
    assert res_empty.status_code == 400
    assert "empty or whitespace" in res_empty.json()["detail"].lower()

    # Test whitespace-only string
    res_whitespace = client.post("/agent", json={"request": "     \n\t  "})
    assert res_whitespace.status_code == 400
    assert "empty or whitespace" in res_whitespace.json()["detail"].lower()

    # Test missing request field entirely
    res_missing = client.post("/agent", json={})
    assert res_missing.status_code == 400


# -----------------------------------------------------------------------------
# Test 3: Excessively Long Request Rejection
# -----------------------------------------------------------------------------

def test_excessively_long_request_rejection():
    """Verify POST /agent rejects requests exceeding 5,000 characters with HTTP 400."""
    long_request = "A" * 5001
    response = client.post("/agent", json={"request": long_request})
    assert response.status_code == 400
    detail = response.json()["detail"].lower()
    assert "exceeds maximum allowed length" in detail or "5,000" in detail


# -----------------------------------------------------------------------------
# Test 4: Valid Plan Parsing
# -----------------------------------------------------------------------------

def test_valid_plan_parsing():
    """Verify create_and_validate_plan correctly parses and validates a well-formed LLM response."""
    mock_client = MagicMock()
    
    with patch("main.call_llm", return_value=json.dumps(SAMPLE_VALID_PLAN_JSON)):
        plan = create_and_validate_plan(mock_client, "test-model", "Build a chatbot proposal")

    assert isinstance(plan, ExecutionPlan)
    assert plan.document_title == "AI-Powered Customer Support Chatbot Business Proposal"
    assert plan.document_type == "Business Proposal"
    assert len(plan.assumptions) == 3
    assert len(plan.tasks) == 4
    assert plan.tasks[0].title == "Requirements Gathering and Scope Definition"


def test_valid_plan_parsing_with_markdown_fences():
    """Verify create_and_validate_plan strips markdown ```json ``` code fences before parsing."""
    mock_client = MagicMock()
    fenced_content = f"```json\n{json.dumps(SAMPLE_VALID_PLAN_JSON, indent=2)}\n```"

    with patch("main.call_llm", return_value=fenced_content):
        plan = create_and_validate_plan(mock_client, "test-model", "Build a chatbot proposal")

    assert isinstance(plan, ExecutionPlan)
    assert len(plan.tasks) == 4


# -----------------------------------------------------------------------------
# Test 5: Malformed Plan Rejection
# -----------------------------------------------------------------------------

def test_malformed_plan_rejection_invalid_json():
    """Verify create_and_validate_plan rejects non-JSON / corrupted strings."""
    mock_client = MagicMock()
    bad_json = "I am an AI, here is your plan: { broken json ... "

    with patch("main.call_llm", return_value=bad_json):
        with pytest.raises(ValueError) as excinfo:
            create_and_validate_plan(mock_client, "test-model", "test prompt")
        assert "Malformed LLM JSON" in str(excinfo.value)


def test_malformed_plan_rejection_too_few_tasks():
    """Verify rejection when LLM generates fewer than 3 tasks (violates min_length=3)."""
    mock_client = MagicMock()
    bad_plan = {
        "document_title": "Too Short Plan",
        "document_type": "Brief",
        "assumptions": ["Assumption 1"],
        "tasks": [
            {"title": "Task 1", "objective": "Objective with sufficient length."},
            {"title": "Task 2", "objective": "Objective with sufficient length."}
        ]
    }

    with patch("main.call_llm", return_value=json.dumps(bad_plan)):
        with pytest.raises(ValueError) as excinfo:
            create_and_validate_plan(mock_client, "test-model", "test prompt")
        assert "Plan validation failed" in str(excinfo.value)


def test_malformed_plan_rejection_too_many_tasks():
    """Verify rejection when LLM generates more than 5 tasks (violates max_length=5)."""
    mock_client = MagicMock()
    bad_plan = {
        "document_title": "Too Long Plan",
        "document_type": "Report",
        "assumptions": ["Assumption 1"],
        "tasks": [
            {"title": f"Task {i}", "objective": f"Valid objective description for step {i}."}
            for i in range(1, 7)
        ]
    }

    with patch("main.call_llm", return_value=json.dumps(bad_plan)):
        with pytest.raises(ValueError) as excinfo:
            create_and_validate_plan(mock_client, "test-model", "test prompt")
        assert "Plan validation failed" in str(excinfo.value)


# -----------------------------------------------------------------------------
# Test 6: Document Creation and File Existence
# -----------------------------------------------------------------------------

def test_document_creation_and_file_existence(sample_plan, sample_task_results, tmp_path):
    """Verify generate_word_document writes a valid, non-empty .docx file with all required sections."""
    test_run_id = "test-run-12345"
    test_request = "Create a business proposal for an AI customer support chatbot."

    filepath, filename = generate_word_document(
        run_id=test_run_id,
        request_text=test_request,
        plan=sample_plan,
        task_results=sample_task_results,
        output_dir=str(tmp_path)
    )

    # 1. Verify file exists on disk
    assert os.path.isfile(filepath)
    assert filename == f"document_{test_run_id}.docx"
    assert os.path.getsize(filepath) > 0

    # 2. Open generated file with python-docx and inspect structural contents
    doc = docx.Document(filepath)

    full_text = "\n".join([p.text for p in doc.paragraphs])
    assert sample_plan.document_title in full_text
    assert "Original Business Request" in full_text
    assert "Planning Assumptions & Constraints" in full_text
    assert "Detailed Task Deliverables & Analysis" in full_text
    assert "Strategic Recommendations & Risk Governance" in full_text
    assert "Final Review & Governance Sign-Off" in full_text

    # 3. Verify tables exist (Execution Plan Table + Sign-off Table + Callout)
    assert len(doc.tables) >= 2
    # Verify execution table contains step titles
    plan_table_text = "\n".join([cell.text for row in doc.tables[1].rows for cell in row.cells])
    assert "Requirements Gathering" in plan_table_text


# -----------------------------------------------------------------------------
# Test 7: Invalid Document Filename Rejection (Path Traversal Protection)
# -----------------------------------------------------------------------------

def test_invalid_document_filename_rejection():
    """Verify GET /documents/{filename} blocks path traversal, non-docx extensions, and invalid characters."""
    # Non-.docx extension
    res1 = client.get("/documents/malicious_payload.exe")
    assert res1.status_code == 400
    assert "invalid filename format" in res1.json()["detail"].lower()

    # Text file
    res2 = client.get("/documents/secret.txt")
    assert res2.status_code == 400

    # Illegal special characters
    res3 = client.get("/documents/evil*file.docx")
    assert res3.status_code == 400

    # Path traversal patterns with dots/slashes
    res4 = client.get("/documents/..%2F..%2Fetc%2Fpasswd.docx")
    assert res4.status_code in (400, 404)


# -----------------------------------------------------------------------------
# Test 8: Missing Document Returning HTTP 404
# -----------------------------------------------------------------------------

def test_missing_document_returns_404():
    """Verify GET /documents/{filename} returns HTTP 404 when validly formatted filename does not exist."""
    response = client.get("/documents/document_00000000-0000-0000-0000-000000000000.docx")
    assert response.status_code == 404
    assert "not found" in response.json()["detail"].lower()


# -----------------------------------------------------------------------------
# Test 9: Full End-to-End Autonomous Agent Execution (Mocked LLM)
# -----------------------------------------------------------------------------

def test_full_agent_workflow_success_mocked():
    """
    Test complete POST /agent pipeline with mocked LLM:
    - Planning -> Sequential Execution -> Word Doc Generation -> Response Assembly
    """
    mock_llm_responses = [
        json.dumps(SAMPLE_VALID_PLAN_JSON),  # Planning call
        "### Task 1 Output\nIdentified 3 core chat channels and SLA target of 2 minutes.",
        "### Task 2 Output\nSelected Groq Llama 3.3 70B Versatile for high speed and accuracy.",
        "### Task 3 Output\n6-week delivery roadmap structured with ₹4,50,000 budget.",
        "### Task 4 Output\nRisk mitigations defined for hallucinations and fallback to human reps."
    ]

    with patch("main.call_llm", side_effect=mock_llm_responses):
        response = client.post(
            "/agent",
            json={"request": "Create a business proposal for an AI customer support chatbot."}
        )

    assert response.status_code == 200
    data = response.json()

    assert data["status"] == "success"
    assert "run_id" in data
    assert data["plan"]["document_title"] == "AI-Powered Customer Support Chatbot Business Proposal"
    assert len(data["tasks"]) == 4
    for t in data["tasks"]:
        assert t["status"] == "completed"
        assert len(t["output"]) > 0

    assert "document_url" in data
    assert data["filename"].endswith(".docx")
    assert "summary" in data

    # Verify download of the newly generated document
    doc_response = client.get(data["document_url"])
    assert doc_response.status_code == 200
    assert len(doc_response.content) > 0
    assert doc_response.headers["content-type"] == (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    # Clean up generated test file
    gen_path = os.path.join(OUTPUT_DIR, data["filename"])
    if os.path.exists(gen_path):
        os.remove(gen_path)


# -----------------------------------------------------------------------------
# Test 10: Partial Failure Resilience
# -----------------------------------------------------------------------------

def test_partial_failure_handling_mocked():
    """
    Verify that if a single task fails during sequential execution:
    - The failure is recorded in the task result
    - The overall run status is marked 'partial_failure'
    - Word document is still generated with failure callout
    """
    def mock_llm_side_effect(client, messages, **kwargs):
        if kwargs.get("json_mode"):
            return json.dumps(SAMPLE_VALID_PLAN_JSON)
        # Check current task in prompt to isolate Task 2 failure
        user_msg = messages[-1]["content"]
        if "CURRENT TASK TO EXECUTE:\nTitle: Technical Architecture" in user_msg:
            raise RuntimeError("Simulated Groq timeout error on task 2")
        return "### Deliverable Content\nStandard task execution completed successfully."

    with patch("main.call_llm", side_effect=mock_llm_side_effect):
        response = client.post(
            "/agent",
            json={"request": "Create a business proposal for an AI customer support chatbot."}
        )

    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "partial_failure"
    assert data["tasks"][0]["status"] == "completed"
    assert data["tasks"][1]["status"] == "failed"
    assert "Simulated Groq timeout error" in data["tasks"][1]["error"]
    assert data["tasks"][2]["status"] == "completed"
    assert data["tasks"][3]["status"] == "completed"

    # Verify document was still generated and downloadable
    gen_path = os.path.join(OUTPUT_DIR, data["filename"])
    assert os.path.exists(gen_path)
    os.remove(gen_path)
