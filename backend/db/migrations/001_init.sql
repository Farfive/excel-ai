CREATE EXTENSION IF NOT EXISTS "pgcrypto";

CREATE TABLE IF NOT EXISTS workbook_sessions (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    workbook_uuid VARCHAR(64) NOT NULL UNIQUE,
    filename VARCHAR(512) NOT NULL,
    cell_count INTEGER DEFAULT 0,
    cluster_count INTEGER DEFAULT 0,
    anomaly_count INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT NOW(),
    workbook_snapshot JSONB
);

CREATE INDEX IF NOT EXISTS idx_workbook_sessions_uuid ON workbook_sessions(workbook_uuid);

CREATE TABLE IF NOT EXISTS chat_messages (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    workbook_uuid VARCHAR(64) NOT NULL,
    role VARCHAR(32) NOT NULL,
    content TEXT,
    tool_calls JSONB,
    created_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_chat_messages_workbook_uuid ON chat_messages(workbook_uuid);
CREATE INDEX IF NOT EXISTS idx_chat_messages_created_at ON chat_messages(created_at);
