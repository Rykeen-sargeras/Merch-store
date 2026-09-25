# Printify Audit Ledger

A private Railway-hosted web app for syncing Printify products and orders, assigning each product to an owner, tracking period payouts, and exporting an Excel payday report.

## Features

- Fetches products and paginated orders from connected Printify shops.
- Suggests an owner from the title prefix before the first dash. For example, Roxy - Mug suggests Roxy.
- Select multiple products and assign an owner in one step. The assignment applies to future sales and the open period.
- Tracks item retail, production cost, quantity, profit, the owner's 75% share, Rykeen's 25% share, a 25% reserve from Rykeen's share, and Rykeen's net.
- Closing a period downloads an XLSX workbook with a product payout summary and sale details, archives the period, and starts a new period. Previous period sales are not included in the new period.
- Protects the app with a password. The Printify token is stored in Railway Variables, not in the source or database.

## Deploy from GitHub to Railway

1. Upload the contents of this folder into a new GitHub repository. .gitignore excludes local secrets, database files, and payout workbooks.
2. In Railway, create a project and deploy from that GitHub repository.
3. Add a Railway Volume to the service and set its mount path to /data. This is required to keep the SQLite ledger when Railway replaces the running container.
4. Add these service variables in Railway:

| Variable | Value |
| --- | --- |
| APP_PASSWORD | A private password you choose for opening the app |
| SESSION_SECRET | A long random secret used to protect sign-in sessions |
| PRINTIFY_API_TOKEN | Your Printify personal access token |
| DATABASE_PATH | /data/ledger.sqlite3 |

5. Generate a Railway public domain and open it. Sign in with APP_PASSWORD, choose Sync products & sales, and assign owners on Products & owners.

Railway reads railway.json and starts the app with Gunicorn. The /health route is used for the deployment health check.

## Local run

Python 3.11 or newer is recommended. Install dependencies with pip install -r requirements.txt. Copy .env.example to .env, set the values in your shell, use a writable local DATABASE_PATH such as ./ledger.sqlite3, then run flask --app app run.

## Payout math

Profit is max(0, retail minus Printify production cost). The product owner receives 75% of profit. Rykeen receives the other 25%. The displayed reserve is 25% of Rykeen's share; Rykeen net is the remainder. Shipping is shown in Excel details but is not subtracted from profit. The reserve is a bookkeeping estimate, not a tax calculation or payment.

Canceled, cancelled, and refunded line items are excluded from payout totals. Printify may omit a retail price for some order types; reconcile those rows with the sales channel before closing a period.
