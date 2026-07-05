# Autonomous Document Agent

A FastAPI-based autonomous AI agent that accepts a natural language request, creates its own task plan, retrieves supporting context using a simple local RAG layer, generates business document sections using an LLM, reflects on the generated output, and produces a polished Microsoft Word `.docx` document.

This project was built for the **Python AI Engineer – Autonomous Agents – 60-Minute Build Challenge**.

---

## Features

- FastAPI backend with `POST /agent`
- Accepts natural language request as JSON
- Creates an autonomous task/TODO plan
- Decides document type and document sections
- Uses Simple Local RAG for domain-specific context retrieval
- Uses Groq LLM for section-wise document generation
- Includes retry and fallback logic for LLM calls
- Performs reflection/self-check before final output
- Generates a Microsoft Word `.docx` file
- Provides document download URL
- Supports standard and complex/ambiguous business requests

---

## Tech Stack

- Python
- FastAPI
- Groq LLM API
- python-docx
- Pydantic
- Uvicorn
- python-dotenv

---

## Project Structure

```text
autonomous-doc-agent/
|
├── main.py
├── requirements.txt
├── README.md
├── .env
├── .gitignore
└── generated_docs/
```

---

## API Overview

### Health Check

```http
GET /
```

Returns API health status.

### Run Autonomous Agent

```http
POST /agent
```

Request body:

```json
{
  "request": "Create a project proposal for an AI-powered customer support chatbot for a fintech company."
}
```

The agent will:

1. Understand the user request
2. Decide the document type
3. Create the document structure
4. Retrieve relevant context from the local knowledge base
5. Generate each section using the LLM
6. Reflect on the generated output
7. Generate a Word document
8. Return the document URL and execution details

### Download Generated Document

```http
GET /documents/{filename}
```

Downloads the generated `.docx` file.

---

## Setup Instructions

### 1. Create Virtual Environment

```bash
python -m venv venv
```

Activate on Windows:

```bash
venv\Scripts\activate
```

Activate on macOS/Linux:

```bash
source venv/bin/activate
```

---

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

---

### 3. Create `.env` File

Create a `.env` file in the project root:

```env
GROQ_API_KEY=your_groq_api_key_here
GROQ_MODEL=llama-3.3-70b-versatile
```

Do not commit `.env` to GitHub.

---

### 4. Run the FastAPI Server

```bash
uvicorn main:app --reload
```

Open Swagger UI:

```text
http://127.0.0.1:8000/docs
```

---

## Test Input 1: Standard Business Request

```json
{
  "request": "Create a project proposal for an AI-powered customer support chatbot for a fintech company. Include business objective, proposed solution, implementation plan, risks, mitigation, and success metrics."
}
```

Expected behavior:

- Creates a Business Proposal
- Generates sections like Executive Summary, Business Objective, Problem Statement, Proposed Solution, Implementation Plan, Risks and Mitigation, Success Metrics, and Conclusion
- Retrieves relevant fintech chatbot context
- Generates a Word document

---

## Test Input 2: Complex / Ambiguous Request

```json
{
  "request": "Create a technical and business plan for launching an AI document automation system for a mid-size NBFC. The request is intentionally vague: budget is limited, timeline is aggressive, compliance is important, and stakeholders want both high accuracy and fast delivery. The agent should make reasonable assumptions, resolve conflicts, and produce a phased plan."
}
```

Expected behavior:

- Creates a Technical and Business Plan
- Detects ambiguity and adds:
  - Assumptions Made by the Agent
  - Conflict Resolution Approach
- Retrieves relevant RAG context:
  - NBFC Compliance and Data Privacy Notes
  - AI Document Automation Risk Checklist
  - MVP and Phased Delivery Strategy
  - Business Document Quality Checklist
- Produces a phased implementation plan
- Performs reflection/self-check
- Generates a Word document

---

## Simple RAG Implementation

This project includes a simple local RAG layer.

Instead of using an external vector database, the project uses a small in-code knowledge base containing domain guidance related to:

- Fintech chatbot metrics
- Fintech chatbot risks
- NBFC compliance
- AI document automation architecture
- AI document automation risks
- MVP/phased delivery strategy
- Business document quality checklist

The retrieval function compares the user request, document type, planned sections, and assumptions against the local knowledge base using keyword overlap and tag matching.

The most relevant chunks are passed into the LLM prompt before section generation.

This improves the generated document because the model receives domain-specific context instead of relying only on the raw user request.

---

## Reflection / Self-Check

The project also includes a reflection/self-check step.

After all document sections are generated, the agent checks:

- Whether all planned sections were generated
- Whether generated content has enough depth
- Whether assumptions were used where input was incomplete
- Whether RAG context was used
- Whether business evaluation sections such as risks, compliance, or success metrics are present
- Whether complex or ambiguous requests were handled properly

The agent then produces:

- Quality score
- Missing sections
- Total content length
- Number of RAG chunks used
- Strengths
- Improvement points
- Self-check summary

This improves reliability because the agent does not blindly generate the final document. It reviews its own output before creating the `.docx` file.

---

## Retry and Fallback Logic

The LLM call uses retry and fallback logic.

If the Groq API key is missing, the model name is invalid, or the LLM call fails, the system returns fallback content instead of crashing immediately.

This makes the API more reliable and demo-safe.

---

## Example Response Fields

A successful response includes:

```json
{
  "message": "Agent executed successfully and generated Word document.",
  "document_type": "Technical and Business Plan",
  "title": "Technical and Business Plan: AI Document Automation System",
  "document_filename": "technical_and_business_plan_ai_document_automation_system_xxxxxxxx.docx",
  "document_url": "/documents/technical_and_business_plan_ai_document_automation_system_xxxxxxxx.docx",
  "rag_context": [],
  "reflection_result": {
    "quality_score": 94,
    "missing_sections": [],
    "rag_chunks_used": 4
  },
  "execution_log": [],
  "status": "document_generated"
}
```

---

## Generated Word Document Includes

The generated `.docx` file contains:

- Document title
- Document type
- Original user request
- Autonomous agent plan
- Agent assumptions
- Retrieved context from Simple RAG
- Execution log
- Agent reflection/self-check
- Generated business document
- Final note

---

## Engineering Decisions

### Why FastAPI?

FastAPI was selected because it is lightweight, fast to build with, supports automatic Swagger documentation, and works well for JSON APIs.

### Why Simple Local RAG?

A local RAG implementation was chosen to keep the project realistic for a 60-minute build challenge. It avoids external vector database setup while still demonstrating retrieval-augmented generation.

### Why Reflection / Self-Check?

Reflection improves the agent's reliability. It allows the system to review whether the generated output satisfies the original request before producing the final Word document.

### Why Retry/Fallback?

LLM APIs can fail due to configuration errors, rate limits, or network issues. Retry and fallback logic make the system more robust and prevent complete failure during demo.

---

## Known Limitations

- The RAG layer uses keyword overlap instead of embeddings.
- The knowledge base is local and small.
- Generated documents use mock assumptions when exact business data is missing.
- The system is synchronous and intended for demo-scale usage.
- Authentication and persistent database storage are not included.

---

## Future Improvements

- Replace keyword-based RAG with embedding-based retrieval
- Add vector database support using Pinecone, ChromaDB, or FAISS
- Add user authentication
- Store generated documents and logs in a database
- Add background task processing for large documents
- Add multi-agent review for quality, compliance, and tone
- Add conversation memory for multi-turn document editing

---

## How to Run Demo

Start the server:

```bash
uvicorn main:app --reload
```

Open Swagger:

```text
http://127.0.0.1:8000/docs
```

Run both test inputs using `POST /agent`.

Open the returned `document_url` in the browser to download the generated `.docx` file.

---

## Assignment Requirement Mapping

| Requirement | Status |
|---|---|
| Python API | Completed |
| FastAPI preferred | Completed |
| POST `/agent` | Completed |
| Natural language request input | Completed |
| Agent-generated task/TODO list | Completed |
| Autonomous execution | Completed |
| Final `.docx` output | Completed |
| Free/free-tier LLM | Completed using Groq |
| One real engineering improvement | Completed: Simple RAG + Reflection/Self-check |
| Standard test input | Completed |
| Complex test input | Completed |
| Execution explanation | Supported through execution log and document sections |

---

## Author

Built by Abhishek Bala for the Python AI Engineer – Autonomous Agents assignment.
