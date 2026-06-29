# KARA TASK LIST

A pastel-themed full-stack task management app with **free** AI-powered PowerPoint generation using Google Gemini.

## Stack
- **Backend:** Flask (Python) on port **8888** REST API + Gemini integration + PPT builder
- **Frontend:** Flask + Jinja2 + Bootstrap 5 + Chart.js on port **8889**
- **Database:** PostgreSQL 15
- **AI:** Google Gemini 2.5 Flash (free tier, no credit card)
- **Container:** Docker + Docker Compose

## Features
- **Home** total tasks, status breakdown, 30-day creation trend chart
- **List Task** paginated table (10/page), live search by title & date, sort by created/modified/status, inline status update, edit modal, delete confirmation
- **Create Task** form with confirmation modal, auto status = Planned
- **Generate PPT** pick date range + status, Gemini analyzes tasks and produces a designed pastel PowerPoint, downloaded instantly

## Setup

### 1. Get a free Gemini API key
1. Go to <https://aistudio.google.com/app/apikey>
2. Sign in with any Google account
3. Click **"Create API Key"** pick or create a project
4. Copy the key (starts with `AIzaSy...`)

The free tier gives you **10 requests/minute, 1500/day** with `gemini-2.5-flash` โ€” more than enough for this app. No credit card required.

### 2. Configure and run
```bash
cp .env
# edit .env, paste your GEMINI_API_KEY

docker compose up -d --build
```

### 3. Open
- Frontend (UI): <http://localhost:8889>
- Backend health: <http://localhost:8888/api/health>

## Database access (Beekeeper / PgAdmin)

| Field    | Value           |
|----------|-----------------|
| Host     | `localhost`     |
| Port     | `5432`          |
| Database | `kara_taskdb`   |
| User     | `postgres`     |
| Password | `postgres`     |

## API Endpoints
- `GET   /api/tasks?page=&per_page=&search=&search_date=&status=&sort_by=`
- `POST  /api/tasks`                      
- `GET   /api/tasks/<id>`
- `PUT   /api/tasks/<id>`                  
- `PATCH /api/tasks/<id>/status`
- `DELETE /api/tasks/<id>`
- `GET   /api/dashboard/summary`
- `POST  /api/generate-ppt`

## Switching Gemini models
Default is `gemini-2.5-flash`. If you hit rate limits or want lighter:
```bash
# in .env
GEMINI_MODEL=gemini-2.5-flash-lite
```

## Reset / clean up
```bash
docker compose down          # stop
docker compose down -v       # also drop the database volume (reseeds from init.sql)
```
