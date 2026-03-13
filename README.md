# Axxerion Operations Hub

Work Order Flow & Facilities Management dashboard for InSite Property Group. Pulls live work order and service request data from the Axxerion CAFM platform and displays it in an interactive web dashboard.

## Features

- **Live Data** — Fetches work orders and service requests from the Axxerion REST API on a 10-minute refresh cycle
- **Server-Side Caching** — Caches API responses to keep the dashboard fast without hammering the Axxerion API
- **Status API** — `/api/status` endpoint showing cache health, record counts, and next refresh time
- **Static Dashboard** — Single-page HTML/JS frontend with interactive charts and filters

## Tech Stack

- **Backend:** Node.js + Express
- **API:** Axxerion REST (CAFM/IWMS platform)
- **Frontend:** Static HTML/CSS/JS
- **Hosting:** Railway

## Getting Started

```bash
npm install
npm start
```

Runs on `http://localhost:3000` by default.

## API Endpoints

| Endpoint | Description |
|----------|-------------|
| `GET /api/workorders` | Cached work order data |
| `GET /api/requests` | Cached service request data |
| `GET /api/status` | Cache health and refresh status |

## Environment Variables

| Variable | Description |
|----------|-------------|
| `AX_USER` | Axxerion API username (required) |
| `AX_PASS` | Axxerion API password (required) |
| `PORT` | Server port (default: 3000) |

Copy `.env.example` to `.env` and fill in your credentials:

```bash
cp .env.example .env
```

## Project Structure

```
axxerion-workorder-flow/
├── server.js              # Express server + Axxerion API caching
├── package.json
├── railway.json           # Railway deployment config
└── public/
    ├── index.html         # Dashboard UI
    ├── data.js            # Work order data handling
    └── requests.js        # Service request data handling
```

## Deployment

Deployed on [Railway](https://railway.app). Push to `main` to trigger auto-deploy.

## Repository

Public repo: [tvt0002/axxerion-workorder-flow](https://github.com/tvt0002/axxerion-workorder-flow)
