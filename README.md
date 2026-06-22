# 🚀 Nasiko Autonomous Document Agent

An autonomous AI agent powered by GPT-4o that transforms simple text briefs into fully formatted, production-ready documents. The agent intelligently analyzes your prompt and autonomously routes it to generate either an **Excel Spreadsheet**, a **PowerPoint Presentation**, or a **Word Document**.

## 🌟 Overview

Instead of explicitly coding document templates, this agent uses LLM-driven planning to:
1. **Understand Intent:** Analyzes the prompt to determine the most appropriate file format.
2. **Plan Content:** Drafts a complete JSON content schema containing section headings, bullet points, aesthetic image search queries, and real historical data.
3. **Render Documents:** Uses native Python libraries (`openpyxl`, `python-pptx`, `python-docx`) to build beautifully styled, native Microsoft Office files.

## ✨ Features

* **Tri-Format Generation:** Supports `.xlsx`, `.pptx`, and `.docx`.
* **Autonomous Routing:** * Asks for "financials" or "sales"? Generates an Excel sheet.
  * Asks for an "essay" or "report"? Generates a Word document.
  * Asks for a "pitch" or vague topic? Generates a Slide Deck.
* **Real-World Data Injection:** Excel engine is strictly prompted to use actual historical data (no dummy placeholders).
* **Aesthetic Styling:** Presentations and Word documents are automatically styled with modern color palettes and high-quality stock photography via the Pexels API.
* **Self-Healing & Revision:** Built-in capability to revise generated documents based on user feedback.

## 🛠️ Prerequisites

Before you begin, ensure you have the following installed:
* [Docker](https://www.docker.com/) and Docker Compose
* An [OpenAI API Key](https://platform.openai.com/)
* A [Pexels API Key](https://www.pexels.com/api/) (for automated image fetching)

## 🚀 Getting Started

**1. Clone the repository and navigate to the directory:**
```bash
git clone <your-repo-url>
cd nasiko-document-agent