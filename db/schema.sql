-- Timetable Scheduler — SQLite Schema
-- This file is documentation only.
-- The schema is embedded in timetable_scheduler.py for in-memory use.

PRAGMA foreign_keys = ON;

CREATE TABLE subjects (
    code        TEXT PRIMARY KEY,
    name_en     TEXT NOT NULL,
    name_cn     TEXT NOT NULL,
    loading_hrs INTEGER NOT NULL   -- 2 or 4
);

CREATE TABLE centres (
    code TEXT PRIMARY KEY          -- 'CS', 'WT', 'KT'...
);

CREATE TABLE rooms (
    code        TEXT PRIMARY KEY,  -- 'CSW - C1'
    centre      TEXT NOT NULL REFERENCES centres,
    capacity    INTEGER NOT NULL
);

CREATE TABLE class_groups (
    code            TEXT PRIMARY KEY,  -- 'CS1', 'WT2'
    centre          TEXT NOT NULL REFERENCES centres,
    is_police_cadet INTEGER DEFAULT 0
);

CREATE TABLE classes (
    code            TEXT PRIMARY KEY,  -- 'DAE101_CS1'
    subject_code    TEXT NOT NULL REFERENCES subjects,
    group_code      TEXT NOT NULL REFERENCES class_groups,
    student_count   INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE teachers (
    id      INTEGER PRIMARY KEY,
    name    TEXT NOT NULL UNIQUE
);

CREATE TABLE teacher_subjects (
    teacher_id      INTEGER NOT NULL REFERENCES teachers,
    subject_code    TEXT NOT NULL REFERENCES subjects,
    lec1_quota      INTEGER DEFAULT 0,
    backup_quota    INTEGER DEFAULT 0,
    PRIMARY KEY (teacher_id, subject_code)
);

-- Teacher marks slots where they are NOT available (default = available)
CREATE TABLE teacher_unavailability (
    teacher_id  INTEGER NOT NULL REFERENCES teachers,
    day         TEXT NOT NULL,
    start_time  TEXT NOT NULL,
    PRIMARY KEY (teacher_id, day, start_time)
);

-- One room per class group (enforces same-subject same-room)
CREATE TABLE group_rooms (
    group_code  TEXT NOT NULL REFERENCES class_groups,
    room_code   TEXT NOT NULL REFERENCES rooms,
    PRIMARY KEY (group_code)
);

-- The schedule: fixed Day/Time from Class list answer, teachers assigned by system
CREATE TABLE schedule (
    class_code      TEXT PRIMARY KEY REFERENCES classes,
    day             TEXT,
    time1           TEXT,   -- '0900 - 1100'
    time2           TEXT,   -- '1100 - 1300' or NULL for 2hr
    room_code       TEXT REFERENCES rooms,
    teacher1_id     INTEGER REFERENCES teachers,
    teacher2_id     INTEGER REFERENCES teachers,
    teacher3_id     INTEGER REFERENCES teachers
);

-- Business rules (configurable, not hardcoded)
-- rule_type: 'PREFERRED_DAYS', 'POLICE_CADET_SUBJECTS'
CREATE TABLE scheduling_rules (
    id          INTEGER PRIMARY KEY,
    rule_type   TEXT NOT NULL,
    target      TEXT,   -- JSON: subject codes or group codes
    value       TEXT,   -- JSON: days array or subject codes
    description TEXT
);

-- Seed: default rules
INSERT INTO scheduling_rules VALUES
    (1, 'PREFERRED_DAYS',
     '["DAE101","DAE102","DAE103","DAE106","DAE108"]',
     '["Monday","Tuesday","Thursday"]',
     'Core subjects prefer Mon/Tue/Thu'),
    (2, 'POLICE_CADET_DAYS',
     '["CS1","CS2","CS3"]',
     '["Monday","Tuesday","Wednesday","Thursday","Friday"]',
     'Police Cadet classes may use any day'),
    (3, 'POLICE_CADET_SUBJECTS',
     '["CS1","CS2","CS3"]',
     '["DAE101","DAE102","DAE103","DAE106"]',
     'Police Cadet classes: no MathsPlus (DAE108)');
