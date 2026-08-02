# Deploy TradeFlow Atlas on Railway

1. Push this project folder to a private GitHub repository. Do not commit `outputs/.env` or `outputs/tradeflow-atlas.sqlite3`.
2. In Railway, create a **New Project** and select **Deploy from GitHub repo**.
3. Railway detects `requirements.txt` and `railway.json`. It installs the Python dependencies and starts `python3 outputs/tradeflow_server.py`.
4. Add these Railway service variables from your local `outputs/.env` (never paste them into source files):

   - `WATSONX_APIKEY`
   - `WATSONX_PROJECT_ID`
   - `WATSONX_URL`
   - `WATSONX_MODEL_ID` (optional; defaults to `ibm/granite-4-h-small`)
   - `NEWS_API` (optional; enables Industry Signals)

5. For persistent notes, checklists, time capsules, and theme settings, add a Railway Volume mounted at `/data`, then set:

   ```text
   TRADEFLOW_DATA_DIR=/data
   ```

Railway supplies `PORT` automatically. The server binds to `0.0.0.0` in deployed environments, while local use remains unchanged.

## Important

- The SQLite database is durable only when `TRADEFLOW_DATA_DIR` points to a Railway Volume.
- The web app keeps its notes local to the deployed server; they are not passed to Watsonx or public web search.
