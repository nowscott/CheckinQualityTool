-- Neon Postgres schema used by /api/inspection/*.
-- The API also creates these objects lazily so a fresh Vercel project can start
-- without a separate migration runner.
CREATE TABLE IF NOT EXISTS inspection_batches (
  id text PRIMARY KEY,
  business_week_start date NOT NULL,
  business_week_end date NOT NULL,
  source_name text NOT NULL,
  source_sha256 text NOT NULL,
  roster_name text NOT NULL,
  roster_sha256 text NOT NULL,
  roster_snapshot_date date,
  sample_limit integer NOT NULL,
  eligible_count integer NOT NULL,
  selected_count integer NOT NULL,
  teacher_count integer NOT NULL,
  unsubmitted_count integer NOT NULL,
  rule_version text NOT NULL,
  attempt integer NOT NULL,
  batch_kind text NOT NULL DEFAULT 'formal',
  status text NOT NULL DEFAULT 'active',
  created_at timestamptz NOT NULL DEFAULT now(),
  voided_at timestamptz
);

CREATE UNIQUE INDEX IF NOT EXISTS inspection_active_week_kind_idx
  ON inspection_batches (business_week_start, batch_kind)
  WHERE status = 'active';

CREATE INDEX IF NOT EXISTS inspection_batches_filter_idx
  ON inspection_batches (status, batch_kind, business_week_start DESC, created_at DESC);

CREATE TABLE IF NOT EXISTS inspection_items (
  batch_id text NOT NULL REFERENCES inspection_batches(id),
  position integer NOT NULL,
  teacher_name text NOT NULL,
  teacher_email text NOT NULL,
  student_name text NOT NULL,
  student_id text NOT NULL,
  course_id text NOT NULL,
  lesson_start text NOT NULL,
  lesson_end text NOT NULL,
  submitted_value text NOT NULL,
  product_group text NOT NULL,
  campus text NOT NULL,
  project_group text NOT NULL,
  selection_reason text NOT NULL,
  PRIMARY KEY (batch_id, position)
);

CREATE INDEX IF NOT EXISTS inspection_items_teacher_idx
  ON inspection_items (LOWER(TRIM(teacher_email)), batch_id);

CREATE TABLE IF NOT EXISTS inspection_users (
  id text PRIMARY KEY,
  username text NOT NULL UNIQUE,
  display_name text NOT NULL,
  password_hash text NOT NULL,
  role text NOT NULL DEFAULT 'viewer' CHECK (role IN ('admin', 'operator', 'viewer')),
  is_active boolean NOT NULL DEFAULT true,
  session_version integer NOT NULL DEFAULT 1,
  last_login_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS inspection_sessions (
  token_hash text PRIMARY KEY,
  user_id text NOT NULL REFERENCES inspection_users(id) ON DELETE CASCADE,
  created_at timestamptz NOT NULL DEFAULT now(),
  expires_at timestamptz NOT NULL,
  revoked_at timestamptz
);

CREATE INDEX IF NOT EXISTS inspection_sessions_user_idx
  ON inspection_sessions (user_id, expires_at DESC)
  WHERE revoked_at IS NULL;

CREATE TABLE IF NOT EXISTS inspection_login_attempts (
  username text PRIMARY KEY,
  window_started_at timestamptz NOT NULL DEFAULT now(),
  failed_count integer NOT NULL DEFAULT 0,
  locked_until timestamptz
);

CREATE TABLE IF NOT EXISTS inspection_audit_log (
  id text PRIMARY KEY,
  actor_user_id text REFERENCES inspection_users(id) ON DELETE SET NULL,
  target_user_id text REFERENCES inspection_users(id) ON DELETE SET NULL,
  action text NOT NULL,
  metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS inspection_audit_created_idx
  ON inspection_audit_log (created_at DESC);
