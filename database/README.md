# ๐Ÿ—„๏ธ KARA TASK LIST โ€” Database Structure

## Database Info
| Item     | Value           |
|----------|-----------------|
| Engine   | PostgreSQL 15   |
| DB Name  | `kara_taskdb`   |
| User     | `postgres`     |
| Password | `postgres`     |
| Port     | `5432`          |

## Tables

### `tasks`
The single table that stores every task.

| Column        | Type            | Constraints                                                                 | Description                          |
|---------------|-----------------|-----------------------------------------------------------------------------|--------------------------------------|
| `id`          | `SERIAL`        | `PRIMARY KEY`                                                               | Auto-increment ID                    |
| `detail`      | `VARCHAR(500)`  | `NOT NULL`                                                                  | Task name / description              |
| `task_date`   | `DATE`          | `NOT NULL`                                                                  | Scheduled date for the task          |
| `status`      | `VARCHAR(20)`   | `NOT NULL`, default `'Planned'`, `CHECK IN ('Planned','Completed','Cancelled')` | Current task status        |
| `created_at`  | `TIMESTAMP`     | `NOT NULL`, default `CURRENT_TIMESTAMP`                                     | Row creation time                    |
| `modified_at` | `TIMESTAMP`     | `NOT NULL`, default `CURRENT_TIMESTAMP`, auto-updated by trigger            | Last modification time               |

### Indexes
| Index Name                 | Column(s)             | Purpose                                       |
|----------------------------|-----------------------|-----------------------------------------------|
| `tasks_pkey`               | `id`                  | Primary key                                   |
| `idx_tasks_status`         | `status`              | Filter & sort by status                       |
| `idx_tasks_task_date`      | `task_date`           | Date-range filter in PPT generator            |
| `idx_tasks_created_at`     | `created_at DESC`     | Default ordering on list page                 |
| `idx_tasks_modified_at`    | `modified_at DESC`    | Sort by modified                              |
| `idx_tasks_detail_lower`   | `LOWER(detail)`       | Case-insensitive title search                 |

### Trigger
- `trg_tasks_modified_at` โ€” BEFORE UPDATE โ†’ sets `modified_at = CURRENT_TIMESTAMP` automatically.

## ERD (single-table)

```
โ”Œโ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”
โ”‚ tasks                                         โ”‚
โ”œโ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”ค
โ”‚ PK  id            SERIAL                      โ”‚
โ”‚     detail        VARCHAR(500) NOT NULL       โ”‚
โ”‚     task_date     DATE         NOT NULL       โ”‚
โ”‚     status        VARCHAR(20)  NOT NULL       โ”‚
โ”‚                   CHECK Planned|Completed|    โ”‚
โ”‚                         Cancelled             โ”‚
โ”‚     created_at    TIMESTAMP    NOT NULL       โ”‚
โ”‚     modified_at   TIMESTAMP    NOT NULL       โ”‚
โ”‚                   (auto-updated via trigger)  โ”‚
โ””โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”˜
```

No foreign keys โ€” the schema is intentionally single-table because the app currently has no users/categories/projects entities. If you later add multi-user support, the natural extension is:

```
users (id, name, email, ...) 1 โ”€โ”€โ”€< tasks.user_id (FK)
```

## How the schema is created

The schema is created **automatically** in two ways:

1. **Docker (recommended)** โ€” `database/init.sql` is mounted into the Postgres container at `/docker-entrypoint-initdb.d/`. On the **first** container start (empty volume), Postgres runs every `.sql` file in that folder. This includes the seed data so you immediately see sample tasks in the app.

2. **SQLAlchemy fallback** โ€” the backend also calls `db.create_all()` on startup. This creates the table if it doesn't exist, but does **not** create the indexes, trigger, CHECK constraint, or seed data. Always prefer the SQL file for production.

## Reset / re-seed

```bash
docker compose down -v          # drop the volume
docker compose up -d --build    # rebuild, init.sql runs again with seed data
```

## Manual connect

```bash
docker exec -it kara_db psql -U postgres -d kara_taskdb

# inside psql:
\dt              -- list tables
\d tasks         -- describe tasks table
SELECT * FROM tasks ORDER BY created_at DESC LIMIT 10;
```
