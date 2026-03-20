# Commvault POC Document Generator

AI-powered tool that generates professional Commvault Proof of Concept (POC) documents in `.docx` format. Enter a client name, workloads, and sales rep details — the app calls OpenAI to produce a structured, branded document ready for customer delivery.

## Features

- Branded cover page with Commvault background and logo
- Auto-generated sections: Executive Summary, Objective, Scope, Hardware Requirements, Networking, Prerequisites, Test Cases, Timeline, Closure
- Version History and POC Team tables from user input
- Per-workload test cases and timelines
- Clickable hyperlinks for documentation references
- Page headers (logo) and footers (page numbers)
- A4 page format
- Web-based frontend with Commvault branding

## Prerequisites

- Node.js 18+
- OpenAI API key

## Setup

```bash
npm install
```

Create a `.env` file in the project root:

```env
OPENAI_API_KEY=your_openai_api_key_here
OPENAI_MODEL=gpt-4o
PORT=4500
```

## Running

**Development:**

```bash
npm run dev
```

**Production:**

```bash
npm run build
npm start
```

Open `http://localhost:4500` in your browser.

## Usage

1. Fill in the Client Name
2. Add one or more Workloads (e.g. VMware, Azure, SAP HANA)
3. Enter Sales Rep Name, Role, and Email
4. Click "Generate POC Document"
5. The `.docx` file downloads automatically

## API

**POST** `/generate-poc`

```json
{
  "clientName": "Acme Corp",
  "workloads": ["VMware", "Azure"],
  "salesRepName": "John Doe",
  "salesRepRole": "Solutions Architect",
  "salesRepEmail": "john@example.com"
}
```

Returns a `.docx` file as binary response.

## Project Structure

```
├── assets/              # Background image and Commvault logo
├── public/              # Frontend HTML
├── src/
│   ├── index.ts         # Express server
│   ├── controller.ts    # Request handler
│   ├── schema.ts        # Zod validation schemas
│   ├── aiService.ts     # OpenAI integration
│   └── docxService.ts   # Document generation
├── .env                 # Environment variables
├── package.json
└── tsconfig.json
```
