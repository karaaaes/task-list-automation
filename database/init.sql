-- ============================================================
--  KARA TASK LIST - Database Schema (with notes)
--  PostgreSQL 15
-- ============================================================

CREATE TABLE IF NOT EXISTS tasks (
    id           SERIAL PRIMARY KEY,
    detail       VARCHAR(500) NOT NULL,
    notes        TEXT         NOT NULL DEFAULT '',
    task_date    DATE         NOT NULL,
    status       VARCHAR(20)  NOT NULL DEFAULT 'Planned'
                 CHECK (status IN ('Planned', 'Completed', 'Cancelled')),
    created_at   TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP,
    modified_at  TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- For existing databases: add notes column if missing
ALTER TABLE tasks ADD COLUMN IF NOT EXISTS notes TEXT NOT NULL DEFAULT '';

COMMENT ON TABLE  tasks               IS 'Stores all user tasks for KARA TASK LIST';
COMMENT ON COLUMN tasks.id            IS 'Primary key';
COMMENT ON COLUMN tasks.detail        IS 'Task name / short description';
COMMENT ON COLUMN tasks.notes         IS 'Long-form notes / context for the task';
COMMENT ON COLUMN tasks.task_date     IS 'Scheduled date';
COMMENT ON COLUMN tasks.status        IS 'Planned | Completed | Cancelled';
COMMENT ON COLUMN tasks.created_at    IS 'Row creation time';
COMMENT ON COLUMN tasks.modified_at   IS 'Last modification time';

-- Indexes
CREATE INDEX IF NOT EXISTS idx_tasks_status         ON tasks(status);
CREATE INDEX IF NOT EXISTS idx_tasks_task_date      ON tasks(task_date);
CREATE INDEX IF NOT EXISTS idx_tasks_created_at     ON tasks(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_tasks_modified_at    ON tasks(modified_at DESC);
CREATE INDEX IF NOT EXISTS idx_tasks_detail_lower   ON tasks(LOWER(detail));

-- Trigger to auto-update modified_at
CREATE OR REPLACE FUNCTION set_modified_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.modified_at = CURRENT_TIMESTAMP;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_tasks_modified_at ON tasks;
CREATE TRIGGER trg_tasks_modified_at
    BEFORE UPDATE ON tasks
    FOR EACH ROW
    EXECUTE FUNCTION set_modified_at();

-- Seed data with notes
INSERT INTO tasks (detail, notes, task_date, status, created_at) VALUES
    ('Prepare weekly team report',          'Cover sprint progress, blockers, and next-week priorities. Share with leadership.', CURRENT_DATE,      'Planned',   CURRENT_TIMESTAMP - INTERVAL '1 day'),
    ('Review Q3 marketing campaign deck',   'Focus on positioning slides and CTA flow. Vendor presents draft on Thursday.',     CURRENT_DATE + 1,  'Planned',   CURRENT_TIMESTAMP - INTERVAL '2 days'),
    ('Finalize budget proposal',            'Locked numbers with finance. Awaiting CFO sign-off before distribution.',          CURRENT_DATE - 1,  'Completed', CURRENT_TIMESTAMP - INTERVAL '3 days'),
    ('Onboard new designer',                'Day-one setup complete. Paired with senior for first sprint.',                     CURRENT_DATE - 3,  'Completed', CURRENT_TIMESTAMP - INTERVAL '5 days'),
    ('Migrate legacy database',             'Scope creep and dependencies forced postponement. Reassess in next quarter.',      CURRENT_DATE - 10, 'Cancelled', CURRENT_TIMESTAMP - INTERVAL '10 days'),
    ('Plan team offsite retreat',           'Venue shortlist ready. Need final headcount and dietary requirements.',            CURRENT_DATE + 7,  'Planned',   CURRENT_TIMESTAMP - INTERVAL '1 day'),
    ('Update internal wiki',                'Reorganized navigation. Added 12 new how-to pages.',                               CURRENT_DATE - 2,  'Completed', CURRENT_TIMESTAMP - INTERVAL '4 days'),
    ('Interview frontend candidates',       'Three candidates shortlisted. Technical round scheduled this week.',               CURRENT_DATE + 2,  'Planned',   CURRENT_TIMESTAMP - INTERVAL '2 days'),
    ('Quarterly performance reviews',       'All 8 reviews submitted. Feedback compiled and shared individually.',              CURRENT_DATE - 5,  'Completed', CURRENT_TIMESTAMP - INTERVAL '7 days'),
    ('Draft API v2 specification',          'Breaking changes documented. Open RFC for engineering review.',                    CURRENT_DATE + 5,  'Planned',   CURRENT_TIMESTAMP - INTERVAL '1 day'),
    ('Vendor contract renewal',             'Vendor unable to meet new SLA terms. Sourcing alternatives.',                      CURRENT_DATE - 7,  'Cancelled', CURRENT_TIMESTAMP - INTERVAL '8 days'),
    ('Customer feedback analysis',          'Pulled NPS data for last 90 days. Synthesizing top three themes.',                 CURRENT_DATE,      'Planned',   CURRENT_TIMESTAMP)
ON CONFLICT DO NOTHING;
