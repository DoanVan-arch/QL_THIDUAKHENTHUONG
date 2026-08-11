-- Migration: Add undo support to activity_log table
-- Run this to add snapshot_data, undone_at, and undone_by_id columns

-- Add snapshot_data column to store state before action
ALTER TABLE activity_log ADD COLUMN snapshot_data TEXT;

-- Add undone_at timestamp
ALTER TABLE activity_log ADD COLUMN undone_at DATETIME;

-- Add undone_by_id foreign key
ALTER TABLE activity_log ADD COLUMN undone_by_id INTEGER;

-- Add foreign key constraint
-- ALTER TABLE activity_log ADD CONSTRAINT fk_activity_log_undone_by 
--     FOREIGN KEY (undone_by_id) REFERENCES users(id) ON DELETE SET NULL;

-- Add index on created_at for faster 72h window queries (if not exists)
CREATE INDEX IF NOT EXISTS idx_activity_log_created_at ON activity_log(created_at);

-- Add index on undone_at for faster queries
CREATE INDEX IF NOT EXISTS idx_activity_log_undone_at ON activity_log(undone_at);
