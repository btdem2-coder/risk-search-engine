# Risk Search Console

Node.js web app for searching risk records using Azure AI Search + Azure OpenAI embeddings, with a simple UI and Excel export.

## Run locally

1. Install dependencies:
   - `npm install`
2. Copy environment template:
   - `copy .env.example .env` (Windows PowerShell)
3. Fill `.env` with real values.
4. Start:
   - `npm start`
5. Open:
   - `http://localhost:3000`

## Deploy to Render

### Option A: Blueprint (recommended)

1. Push this repo to GitHub.
2. In Render, click **New +** -> **Blueprint**.
3. Select your repo (Render will detect `render.yaml`).
4. Add required secret env vars in Render dashboard:
   - `AZURE_OPENAI_ENDPOINT`
   - `AZURE_OPENAI_KEY`
   - `AZURE_OPENAI_DEPLOYMENT`
   - `AZURE_SEARCH_ENDPOINT`
   - `AZURE_SEARCH_KEY`
   - `AZURE_SEARCH_INDEX`
   - `OPENAI_API_KEY`
   - `OPENAI_FILTER_MODEL` (optional, defaults to `gpt-4o-mini`)
5. Deploy.

### Option B: Manual service setup

1. In Render, create a **Web Service** from your GitHub repo.
2. Set:
   - **Environment**: Node
   - **Build Command**: `npm install`
   - **Start Command**: `npm start`
3. Add the same env vars listed above.
4. Deploy.

## Health check

- Endpoint: `GET /health`
- Returns: `{ "ok": true }`

## Security note

- Do not commit `.env`.
- If any real keys were ever committed/shared, rotate them in Azure/OpenAI and update Render env vars.
