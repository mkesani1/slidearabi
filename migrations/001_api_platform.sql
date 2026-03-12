-- ============================================================================
-- SlideArabi API Platform — Database Migration 001
-- ============================================================================
-- Supabase (PostgreSQL) migration for the API/MCP credit system.
-- Creates tables, indexes, RLS policies, functions, and triggers.
--
-- Run via: Supabase Dashboard → SQL Editor, or `supabase db push`.
-- Date: 2026-03-10
-- ============================================================================

BEGIN;

-- ────────────────────────────────────────────────────────────────────────────
-- 1. TABLES
-- ────────────────────────────────────────────────────────────────────────────

-- API accounts — one per Supabase auth user who enables API access
CREATE TABLE IF NOT EXISTS api_accounts (
    id                   UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id              UUID NOT NULL REFERENCES auth.users(id) UNIQUE,
    credits_available    INTEGER NOT NULL DEFAULT 0 CHECK (credits_available >= 0),
    credits_reserved     INTEGER NOT NULL DEFAULT 0 CHECK (credits_reserved >= 0),
    plan                 TEXT NOT NULL DEFAULT 'pay_as_you_go',
    stripe_customer_id   TEXT,
    auto_topup_enabled   BOOLEAN NOT NULL DEFAULT false,
    auto_topup_threshold INTEGER NOT NULL DEFAULT 20,
    auto_topup_amount    INTEGER NOT NULL DEFAULT 100,
    created_at           TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at           TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- API keys — hashed, never stored in plaintext
CREATE TABLE IF NOT EXISTS api_keys (
    id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    account_id   UUID NOT NULL REFERENCES api_accounts(id) ON DELETE CASCADE,
    key_hash     TEXT NOT NULL UNIQUE,                -- SHA-256 of full key
    key_prefix   VARCHAR(16) NOT NULL,                -- "sa_live_a1b2c3d4" for display
    name         TEXT NOT NULL DEFAULT 'default',
    is_test      BOOLEAN NOT NULL DEFAULT false,      -- sa_test_ prefix keys
    is_active    BOOLEAN NOT NULL DEFAULT true,
    last_used_at TIMESTAMPTZ,
    created_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- File uploads — staging area before conversion (24h TTL)
CREATE TABLE IF NOT EXISTS api_uploads (
    id           TEXT PRIMARY KEY,                    -- "upl_xxxxxxxx"
    account_id   UUID NOT NULL REFERENCES api_accounts(id) ON DELETE CASCADE,
    filename     TEXT NOT NULL,
    slide_count  INTEGER NOT NULL,
    size_bytes   BIGINT NOT NULL,
    storage_path TEXT NOT NULL,                       -- Supabase Storage path
    expires_at   TIMESTAMPTZ NOT NULL,
    created_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Conversion jobs — tracks each API/MCP conversion request
CREATE TABLE IF NOT EXISTS api_jobs (
    id               TEXT PRIMARY KEY,                -- "job_xxxxxxxx"
    account_id       UUID NOT NULL REFERENCES api_accounts(id) ON DELETE CASCADE,
    upload_id        TEXT NOT NULL REFERENCES api_uploads(id),
    source           TEXT NOT NULL DEFAULT 'api',     -- 'api', 'mcp', 'web'
    status           TEXT NOT NULL DEFAULT 'queued',
    progress_percent INTEGER NOT NULL DEFAULT 0,
    current_phase    TEXT,
    slide_count      INTEGER NOT NULL,
    credits_reserved INTEGER NOT NULL DEFAULT 0,
    credits_charged  INTEGER NOT NULL DEFAULT 0,
    options          JSONB NOT NULL DEFAULT '{}',
    webhook_url      TEXT,
    idempotency_key  TEXT,
    result_path      TEXT,                            -- Supabase Storage path of output
    error_code       TEXT,
    error_message    TEXT,
    created_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
    started_at       TIMESTAMPTZ,
    completed_at     TIMESTAMPTZ
);

-- Credit transaction ledger — immutable, append-only
CREATE TABLE IF NOT EXISTS credit_transactions (
    id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    account_id            UUID NOT NULL REFERENCES api_accounts(id) ON DELETE CASCADE,
    type                  TEXT NOT NULL,               -- purchase, reserve, settle, release, refund, bonus, promo
    amount                INTEGER NOT NULL,            -- positive = added, negative = deducted
    balance_after         INTEGER NOT NULL,
    description           TEXT NOT NULL,
    job_id                TEXT REFERENCES api_jobs(id),
    stripe_payment_intent TEXT,
    created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Webhooks — registered callback endpoints
CREATE TABLE IF NOT EXISTS webhooks (
    id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    account_id    UUID NOT NULL REFERENCES api_accounts(id) ON DELETE CASCADE,
    url           TEXT NOT NULL,
    secret        TEXT NOT NULL,                      -- HMAC signing secret
    events        TEXT[] NOT NULL DEFAULT ARRAY['conversion.completed','conversion.failed'],
    is_active     BOOLEAN NOT NULL DEFAULT true,
    failure_count INTEGER NOT NULL DEFAULT 0,
    created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);


-- ────────────────────────────────────────────────────────────────────────────
-- 2. INDEXES
-- ────────────────────────────────────────────────────────────────────────────

-- key_hash already has UNIQUE constraint → implicit unique index

CREATE INDEX IF NOT EXISTS idx_api_keys_account
    ON api_keys(account_id);

CREATE INDEX IF NOT EXISTS idx_api_jobs_account_created
    ON api_jobs(account_id, created_at DESC);

CREATE INDEX IF NOT EXISTS idx_api_jobs_active_status
    ON api_jobs(status)
    WHERE status IN ('queued', 'processing');

CREATE UNIQUE INDEX IF NOT EXISTS idx_api_jobs_idempotency
    ON api_jobs(idempotency_key)
    WHERE idempotency_key IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_credit_txn_account_created
    ON credit_transactions(account_id, created_at DESC);

CREATE INDEX IF NOT EXISTS idx_credit_txn_stripe_pi
    ON credit_transactions(stripe_payment_intent)
    WHERE stripe_payment_intent IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_api_uploads_account
    ON api_uploads(account_id);

CREATE INDEX IF NOT EXISTS idx_api_uploads_expires
    ON api_uploads(expires_at)
    WHERE expires_at > now();

CREATE INDEX IF NOT EXISTS idx_webhooks_account
    ON webhooks(account_id);


-- ────────────────────────────────────────────────────────────────────────────
-- 3. TRIGGER: auto-update updated_at on api_accounts
-- ────────────────────────────────────────────────────────────────────────────

CREATE OR REPLACE FUNCTION update_api_accounts_updated_at()
RETURNS TRIGGER
LANGUAGE plpgsql
AS $$
BEGIN
    NEW.updated_at = now();
    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS trg_api_accounts_updated_at ON api_accounts;

CREATE TRIGGER trg_api_accounts_updated_at
    BEFORE UPDATE ON api_accounts
    FOR EACH ROW
    EXECUTE FUNCTION update_api_accounts_updated_at();


-- ────────────────────────────────────────────────────────────────────────────
-- 4. ATOMIC CREDIT FUNCTIONS (SECURITY DEFINER — run as owner, not caller)
-- ────────────────────────────────────────────────────────────────────────────

-- reserve_credits: Atomically hold credits from available for a pending job.
-- Returns TRUE on success, FALSE if insufficient credits.
CREATE OR REPLACE FUNCTION reserve_credits(
    p_account_id UUID,
    p_amount     INTEGER,
    p_job_id     TEXT
) RETURNS BOOLEAN
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
    v_available     INTEGER;
    v_new_available INTEGER;
BEGIN
    -- Lock the row to prevent concurrent modifications
    SELECT credits_available INTO v_available
    FROM api_accounts
    WHERE id = p_account_id
    FOR UPDATE;

    IF v_available IS NULL THEN
        RAISE EXCEPTION 'Account not found: %', p_account_id;
    END IF;

    IF v_available < p_amount THEN
        RETURN FALSE;
    END IF;

    v_new_available := v_available - p_amount;

    UPDATE api_accounts
    SET credits_available = v_new_available,
        credits_reserved  = credits_reserved + p_amount
    WHERE id = p_account_id;

    INSERT INTO credit_transactions (account_id, type, amount, balance_after, description, job_id)
    VALUES (
        p_account_id, 'reserve', -p_amount, v_new_available,
        format('Reserved %s credits for job %s', p_amount, p_job_id),
        p_job_id
    );

    RETURN TRUE;
END;
$$;


-- settle_credits: Finalize reserved credits after successful conversion.
CREATE OR REPLACE FUNCTION settle_credits(
    p_account_id UUID,
    p_amount     INTEGER,
    p_job_id     TEXT
) RETURNS VOID
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
    v_available INTEGER;
BEGIN
    SELECT credits_available INTO v_available
    FROM api_accounts
    WHERE id = p_account_id
    FOR UPDATE;

    IF v_available IS NULL THEN
        RAISE EXCEPTION 'Account not found: %', p_account_id;
    END IF;

    UPDATE api_accounts
    SET credits_reserved = credits_reserved - p_amount
    WHERE id = p_account_id;

    INSERT INTO credit_transactions (account_id, type, amount, balance_after, description, job_id)
    VALUES (
        p_account_id, 'settle', -p_amount, v_available,
        format('Settled %s credits for completed job %s', p_amount, p_job_id),
        p_job_id
    );
END;
$$;


-- release_credits: Return reserved credits to available on failure/cancellation.
CREATE OR REPLACE FUNCTION release_credits(
    p_account_id UUID,
    p_amount     INTEGER,
    p_job_id     TEXT
) RETURNS VOID
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
    v_new_available INTEGER;
BEGIN
    SELECT credits_available + p_amount INTO v_new_available
    FROM api_accounts
    WHERE id = p_account_id
    FOR UPDATE;

    IF v_new_available IS NULL THEN
        RAISE EXCEPTION 'Account not found: %', p_account_id;
    END IF;

    UPDATE api_accounts
    SET credits_available = v_new_available,
        credits_reserved  = credits_reserved - p_amount
    WHERE id = p_account_id;

    INSERT INTO credit_transactions (account_id, type, amount, balance_after, description, job_id)
    VALUES (
        p_account_id, 'release', p_amount, v_new_available,
        format('Released %s credits — job %s failed/cancelled', p_amount, p_job_id),
        p_job_id
    );
END;
$$;


-- grant_credits: Add credits from purchase, promo, or signup bonus.
CREATE OR REPLACE FUNCTION grant_credits(
    p_account_id  UUID,
    p_amount      INTEGER,
    p_description TEXT,
    p_type        TEXT DEFAULT 'purchase',
    p_stripe_pi   TEXT DEFAULT NULL
) RETURNS VOID
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
    v_new_available INTEGER;
BEGIN
    SELECT credits_available + p_amount INTO v_new_available
    FROM api_accounts
    WHERE id = p_account_id
    FOR UPDATE;

    IF v_new_available IS NULL THEN
        RAISE EXCEPTION 'Account not found: %', p_account_id;
    END IF;

    UPDATE api_accounts
    SET credits_available = v_new_available
    WHERE id = p_account_id;

    INSERT INTO credit_transactions (account_id, type, amount, balance_after, description, stripe_payment_intent)
    VALUES (p_account_id, p_type, p_amount, v_new_available, p_description, p_stripe_pi);
END;
$$;


-- ────────────────────────────────────────────────────────────────────────────
-- 5. ROW LEVEL SECURITY (RLS)
-- ────────────────────────────────────────────────────────────────────────────
-- The API gateway uses SUPABASE_SERVICE_KEY (service_role) which bypasses RLS.
-- These policies protect data accessed via the Supabase client-side JS SDK
-- (e.g., from the Next.js dashboard).

ALTER TABLE api_accounts ENABLE ROW LEVEL SECURITY;
ALTER TABLE api_keys ENABLE ROW LEVEL SECURITY;
ALTER TABLE api_uploads ENABLE ROW LEVEL SECURITY;
ALTER TABLE api_jobs ENABLE ROW LEVEL SECURITY;
ALTER TABLE credit_transactions ENABLE ROW LEVEL SECURITY;
ALTER TABLE webhooks ENABLE ROW LEVEL SECURITY;

-- api_accounts: users can read/update only their own account
CREATE POLICY api_accounts_select ON api_accounts
    FOR SELECT USING (user_id = auth.uid());
CREATE POLICY api_accounts_insert ON api_accounts
    FOR INSERT WITH CHECK (user_id = auth.uid());
CREATE POLICY api_accounts_update ON api_accounts
    FOR UPDATE USING (user_id = auth.uid());

-- api_keys: users can CRUD their own keys
CREATE POLICY api_keys_select ON api_keys
    FOR SELECT USING (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));
CREATE POLICY api_keys_insert ON api_keys
    FOR INSERT WITH CHECK (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));
CREATE POLICY api_keys_update ON api_keys
    FOR UPDATE USING (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));
CREATE POLICY api_keys_delete ON api_keys
    FOR DELETE USING (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));

-- api_uploads: users can read their own uploads
CREATE POLICY api_uploads_select ON api_uploads
    FOR SELECT USING (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));

-- api_jobs: users can read their own jobs
CREATE POLICY api_jobs_select ON api_jobs
    FOR SELECT USING (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));

-- credit_transactions: users can read their own
CREATE POLICY credit_txn_select ON credit_transactions
    FOR SELECT USING (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));

-- webhooks: users can CRUD their own
CREATE POLICY webhooks_select ON webhooks
    FOR SELECT USING (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));
CREATE POLICY webhooks_insert ON webhooks
    FOR INSERT WITH CHECK (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));
CREATE POLICY webhooks_update ON webhooks
    FOR UPDATE USING (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));
CREATE POLICY webhooks_delete ON webhooks
    FOR DELETE USING (account_id IN (SELECT id FROM api_accounts WHERE user_id = auth.uid()));


-- ────────────────────────────────────────────────────────────────────────────
-- 6. TABLE COMMENTS
-- ────────────────────────────────────────────────────────────────────────────

COMMENT ON TABLE api_accounts IS 'API platform accounts linked to Supabase auth users. Tracks credit balance and plan tier.';
COMMENT ON TABLE api_keys IS 'API keys for REST and MCP access. Keys are SHA-256 hashed; only prefix stored in plaintext.';
COMMENT ON TABLE api_uploads IS 'Staged file uploads awaiting conversion. Expire after 24 hours.';
COMMENT ON TABLE api_jobs IS 'Conversion jobs created via REST API or MCP. Tracks status, progress, credits, and results.';
COMMENT ON TABLE credit_transactions IS 'Immutable ledger of all credit movements: purchases, reservations, settlements, releases, refunds.';
COMMENT ON TABLE webhooks IS 'Developer-registered webhook endpoints for conversion event notifications.';
COMMENT ON FUNCTION reserve_credits IS 'Atomically reserve credits from available balance for a pending conversion job.';
COMMENT ON FUNCTION settle_credits IS 'Settle reserved credits after successful job completion.';
COMMENT ON FUNCTION release_credits IS 'Release reserved credits back to available after job failure or cancellation.';
COMMENT ON FUNCTION grant_credits IS 'Add credits to account from purchase, promo code, or signup bonus.';

COMMIT;
