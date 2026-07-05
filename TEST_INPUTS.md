# Test Inputs

Use these inputs to test the `POST /agent` endpoint from Swagger UI.

Swagger URL:

```text
http://127.0.0.1:8000/docs
```

Endpoint:

```http
POST /agent
```

---

## Test Input 1: Standard Business Request

This test checks whether the agent can create a normal business proposal from a clear request.

```json
{
  "request": "Create a project proposal for an AI-powered customer support chatbot for a fintech company. Include business objective, proposed solution, implementation plan, risks, mitigation, and success metrics."
}
```

### Expected Behavior

The agent should:

- Identify the document type as a business proposal
- Create an autonomous task plan
- Retrieve relevant fintech/chatbot context from the local knowledge base
- Generate document sections such as:
  - Executive Summary
  - Business Objective
  - Problem Statement
  - Proposed Solution
  - Implementation Plan
  - Risks and Mitigation
  - Success Metrics
  - Conclusion
- Perform reflection/self-check
- Generate a `.docx` file
- Return a `document_url`

### Expected Response Fields

The response should include fields like:

```json
{
  "message": "Agent executed successfully and generated Word document.",
  "document_type": "Business Proposal",
  "title": "Business Proposal: AI-Powered Customer Support Chatbot",
  "document_url": "/documents/generated_file_name.docx",
  "rag_context": [],
  "reflection_result": {},
  "execution_log": [],
  "status": "document_generated"
}
```

---

## Test Input 2: Complex / Ambiguous Request

This test checks whether the agent can handle vague requirements, make assumptions, resolve conflicts, use Simple RAG, and produce a phased plan.

```json
{
  "request": "Create a technical and business plan for launching an AI document automation system for a mid-size NBFC. The request is intentionally vague: budget is limited, timeline is aggressive, compliance is important, and stakeholders want both high accuracy and fast delivery. The agent should make reasonable assumptions, resolve conflicts, and produce a phased plan."
}
```

### Expected Behavior

The agent should:

- Identify the document type as a technical and business plan
- Detect that the request is complex or ambiguous
- Add special sections:
  - Assumptions Made by the Agent
  - Conflict Resolution Approach
- Retrieve relevant context from Simple RAG, such as:
  - NBFC Compliance and Data Privacy Notes
  - AI Document Automation Risk Checklist
  - MVP and Phased Delivery Strategy
  - Business Document Quality Checklist
- Generate a phased implementation roadmap
- Include compliance, risk, mitigation, and success metric sections
- Perform reflection/self-check
- Generate a `.docx` file
- Return a `document_url`

### Expected Response Fields

The response should include fields like:

```json
{
  "message": "Agent executed successfully and generated Word document.",
  "document_type": "Technical and Business Plan",
  "title": "Technical and Business Plan: AI Document Automation System",
  "rag_context": [
    {
      "id": "nbfc_compliance",
      "title": "NBFC Compliance and Data Privacy Notes",
      "score": 27
    }
  ],
  "reflection_result": {
    "quality_score": 94,
    "missing_sections": [],
    "rag_chunks_used": 4
  },
  "document_url": "/documents/generated_file_name.docx",
  "status": "document_generated"
}
```

---

## What to Check in the Response

For both tests, verify that:

- `status` is `document_generated`
- `document_url` is returned
- `plan` contains an autonomous task list
- `rag_context` contains retrieved knowledge chunks
- `reflection_result` contains a quality score
- `llm_sources` shows `"source": "groq"` when the Groq call succeeds
- `execution_log` shows each major step performed by the agent

---

## What to Check in the Generated Word Document

Open the returned `document_url`.

The `.docx` file should contain:

- Document title
- Document type
- Original user request
- Autonomous Agent Plan
- Agent Assumptions
- Retrieved Context from Simple RAG
- Execution Log
- Agent Reflection / Self-Check
- Generated Business Document
- Final Note

---

## Notes

The generated document filename changes on every run because the system appends a unique ID to avoid overwriting previous documents.

Example:

```text
technical_and_business_plan_ai_document_automation_system_5e6d8cb5.docx
```

The exact generated content may vary slightly because the document sections are generated using an LLM.
