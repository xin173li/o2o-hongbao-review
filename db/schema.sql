-- O2O 红包核销数据库（SQLite）
-- 设计原则：事实表只存"发生过什么"，核对结果单独存 audit_result，可重复核销、可留痕

PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

-- ========== 维度 ==========
CREATE TABLE IF NOT EXISTS dim_channel (
    channel_code TEXT PRIMARY KEY,
    channel_name TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS dim_store (
    store_code   TEXT NOT NULL,
    store_name   TEXT NOT NULL,
    channel_code TEXT NOT NULL,
    PRIMARY KEY (store_code, channel_code)
);

CREATE TABLE IF NOT EXISTS dim_model (
    model_code TEXT PRIMARY KEY,   -- 归一后（大写、去空格）
    model_name TEXT NOT NULL        -- 首次出现的原始写法
);

-- ========== 政策 ==========
-- 一行一条政策：某型号 + 某红包类型 + 某渠道（NULL 表示全渠道）+ 生效区间
CREATE TABLE IF NOT EXISTS policy (
    policy_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    model_code   TEXT NOT NULL,
    red_type     TEXT NOT NULL CHECK (red_type IN ('order', 'retail')),
    channel_code TEXT,                     -- NULL = 全渠道通用
    period_start TEXT NOT NULL,            -- YYYY-MM
    period_end   TEXT,                     -- NULL = 长期有效
    unit_amount  REAL NOT NULL,            -- 单台红包标准（不含税）
    doc_no       TEXT,
    src_file     TEXT,
    imported_at  TEXT DEFAULT (datetime('now')),
    UNIQUE (model_code, red_type, channel_code, period_start)
);

-- ========== 事实：下单 ==========
CREATE TABLE IF NOT EXISTS fact_order (
    period       TEXT NOT NULL,
    channel_code TEXT NOT NULL,
    store_code   TEXT NOT NULL,
    model_code   TEXT NOT NULL,
    qty          REAL NOT NULL DEFAULT 0,  -- 净数量：退货为负
    amount       REAL,
    src_file     TEXT,
    imported_at  TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (period, channel_code, store_code, model_code)
);

-- ========== 事实：红包发放 ==========
CREATE TABLE IF NOT EXISTS fact_issue (
    period       TEXT NOT NULL,
    channel_code TEXT NOT NULL,
    store_code   TEXT NOT NULL,
    model_code   TEXT NOT NULL,
    red_type     TEXT NOT NULL CHECK (red_type IN ('order', 'retail')),
    qty          REAL NOT NULL DEFAULT 0,
    unit_amount  REAL NOT NULL DEFAULT 0,
    total_amount REAL,
    cum_qty_src  REAL,                     -- 台账自带的累计列（可空，用于交叉校验）
    src_file     TEXT,
    imported_at  TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (period, channel_code, store_code, model_code, red_type)
);

-- ========== 期初累计 ==========
CREATE TABLE IF NOT EXISTS opening_balance (
    base_period    TEXT NOT NULL,
    store_code     TEXT NOT NULL,
    model_code     TEXT NOT NULL,
    channel_code   TEXT NOT NULL,
    cum_order_qty  REAL NOT NULL DEFAULT 0,
    cum_retail_qty REAL NOT NULL DEFAULT 0,
    cum_order_rqty REAL NOT NULL DEFAULT 0,
    src_file       TEXT,
    PRIMARY KEY (base_period, channel_code, store_code, model_code)
);

-- ========== 批次与结果 ==========
CREATE TABLE IF NOT EXISTS import_batch (
    batch_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at   TEXT DEFAULT (datetime('now')),
    finished_at  TEXT,
    file_name    TEXT,
    file_hash    TEXT,
    sheet_count  INTEGER,
    rows_written INTEGER,
    rows_failed  INTEGER,
    status       TEXT,
    summary      TEXT
);

CREATE TABLE IF NOT EXISTS import_log (
    batch_id   INTEGER,
    file_name  TEXT,
    sheet_name TEXT,
    data_type  TEXT,
    rows_written INTEGER,
    rows_failed  INTEGER,
    errors     TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS audit_run (
    run_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at   TEXT DEFAULT (datetime('now')),
    periods      TEXT,          -- 本次涉及的月份范围
    baseline     TEXT,
    judge_from   TEXT,
    fingerprint  TEXT,          -- 源数据指纹，未变则复用批次
    status       TEXT DEFAULT 'DONE'
);

CREATE TABLE IF NOT EXISTS audit_result (
    run_id       INTEGER NOT NULL,
    check_code   TEXT NOT NULL,      -- C1 政策 / C2 下单 / C3 零售 / C0 数据检查
    period       TEXT,
    channel_code TEXT,
    store_code   TEXT,
    store_name   TEXT,
    model_code   TEXT,
    red_type     TEXT,
    left_label   TEXT,
    left_value   REAL,
    right_label  TEXT,
    right_value  REAL,
    diff         REAL,
    status       TEXT NOT NULL,      -- PASS / FAIL / WARN
    note         TEXT,
    PRIMARY KEY (run_id, check_code, period, channel_code, store_code, model_code, red_type)
);

CREATE INDEX IF NOT EXISTS idx_result_check ON audit_result (run_id, check_code, status);
CREATE INDEX IF NOT EXISTS idx_result_period ON audit_result (run_id, period);
