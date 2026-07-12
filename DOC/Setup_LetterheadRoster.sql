-- Setup_LetterheadRoster.sql
--
-- Makes the letterhead's attorney/staff roster fully database-driven instead
-- of baked into a static vector graphic. Run this once (it's safe to re-run -
-- the ALTER TABLE lines no-op if the columns already exist, and the UPDATEs
-- simply re-apply the same values).
--
-- HOW TO CHANGE THE LETTERHEAD NAMES FROM NOW ON:
--   You do NOT need to re-run this whole script. Just UPDATE the relevant
--   row(s) in the employee table directly, e.g.:
--
--     -- Someone leaves the firm:
--     UPDATE employee SET ON_LETTERHEAD = NULL WHERE CODE = 'AJS'
--
--     -- A new Legal Assistant joins and should appear in that group:
--     UPDATE employee SET ON_LETTERHEAD = 'Y', LH_ORDER = 95,
--         LH_HEADING = 'Legal Assistants' WHERE CODE = 'XYZ'
--
--     -- Someone's name changes (marriage, correction, etc.) - just fix FN/MI/LN
--     -- as you would for any other employee record; the letterhead picks it up
--     -- automatically the next time a PDF is generated.
--
-- Column meanings:
--   ON_LETTERHEAD     'Y' = shown on the letterhead sidebar, else not shown.
--   LH_ORDER          Sort order (top to bottom). Leave gaps (10, 20, 30, ...)
--                      so you can insert someone later without renumbering.
--   LH_HEADING        Heading text printed above this row's block, e.g.
--                      'Legal Assistants'. Consecutive rows sharing the same
--                      heading text are grouped under ONE printed heading
--                      (only the first one prints it). Leave blank/NULL for
--                      no heading (used for the attorneys at the top).
--   LH_SUFFIX         Optional text appended after the name, e.g. 'Esq' or
--                      '(1983-2015)'.
--   LH_SUFFIX_ITALIC  'Y' to italicize LH_SUFFIX (e.g. 'Senior Rep').
--   LH_BLANK_BEFORE   'Y' to insert a blank line above this row (used for the
--                      gap before the in-memoriam entry).

IF COL_LENGTH('employee', 'ON_LETTERHEAD') IS NULL
    ALTER TABLE employee ADD ON_LETTERHEAD CHAR(1) NULL;
IF COL_LENGTH('employee', 'LH_ORDER') IS NULL
    ALTER TABLE employee ADD LH_ORDER INT NULL;
IF COL_LENGTH('employee', 'LH_HEADING') IS NULL
    ALTER TABLE employee ADD LH_HEADING NVARCHAR(60) NULL;
IF COL_LENGTH('employee', 'LH_SUFFIX') IS NULL
    ALTER TABLE employee ADD LH_SUFFIX NVARCHAR(30) NULL;
IF COL_LENGTH('employee', 'LH_SUFFIX_ITALIC') IS NULL
    ALTER TABLE employee ADD LH_SUFFIX_ITALIC CHAR(1) NULL;
IF COL_LENGTH('employee', 'LH_BLANK_BEFORE') IS NULL
    ALTER TABLE employee ADD LH_BLANK_BEFORE CHAR(1) NULL;
GO

-- Clear any previous roster flags before re-applying (safe for re-runs).
UPDATE employee SET ON_LETTERHEAD = NULL, LH_ORDER = NULL, LH_HEADING = NULL,
    LH_SUFFIX = NULL, LH_SUFFIX_ITALIC = NULL, LH_BLANK_BEFORE = NULL
WHERE ON_LETTERHEAD = 'Y';

-- Attorneys (no heading; "Esq" suffix)
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=10, LH_SUFFIX='Esq' WHERE CODE='ACS';   -- Amanda C. Shaw
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=20, LH_SUFFIX='Esq' WHERE CODE='JIR';   -- Joel I. Ross
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=30, LH_SUFFIX='Esq' WHERE CODE='NJT';   -- Nicholas J. Tammaro

-- In-memoriam entry (blank line above, dates instead of "Esq")
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=40, LH_SUFFIX='(1983-2015)', LH_BLANK_BEFORE='Y' WHERE CODE='ZMP'; -- Zachary M. Primrose

-- Solo roles (each is its own one-person heading)
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=50, LH_HEADING='Accounts /Office Manager' WHERE CODE='VMF'; -- Vickie M. Flanagan
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=60, LH_HEADING='Senior Legal Assistant' WHERE CODE='MKR';   -- Meghan K. Riese
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=70, LH_HEADING='Accounts Supervisor' WHERE CODE='AWF';      -- Andrew W. Fisher
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=80, LH_HEADING='Accounting Manager' WHERE CODE='CBC';       -- Chris B. Chapman

-- Legal Assistants group
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=90,  LH_HEADING='Legal Assistants' WHERE CODE='NAC'; -- Newsy A. Chapman
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=100, LH_HEADING='Legal Assistants' WHERE CODE='LED'; -- Laureen E. Doody
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=110, LH_HEADING='Legal Assistants' WHERE CODE='JRW'; -- Jill R. Washburn
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=120, LH_HEADING='Legal Assistants' WHERE CODE='RMM'; -- Rachael M. Meixell
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=130, LH_HEADING='Legal Assistants' WHERE CODE='CWP'; -- Charlene W. Prosser
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=140, LH_HEADING='Legal Assistants' WHERE CODE='APF'; -- Aleah Freeman
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=150, LH_HEADING='Legal Assistants' WHERE CODE='AJS'; -- Austyn Scott

-- Administrative Assistants group
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=160, LH_HEADING='Administrative Assistants' WHERE CODE='BJB'; -- Billie Jo Bushaw
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=170, LH_HEADING='Administrative Assistants' WHERE CODE='AZS'; -- Andrea Ziccarello

-- Account Representatives group (White shown first, with "Senior Rep" suffix)
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=180, LH_HEADING='Account Representatives', LH_SUFFIX='Senior Rep', LH_SUFFIX_ITALIC='Y' WHERE CODE='DAW'; -- Deborah A. White
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=190, LH_HEADING='Account Representatives' WHERE CODE='SLL'; -- Sara Ladd
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=200, LH_HEADING='Account Representatives' WHERE CODE='SOL'; -- Shannon Long
UPDATE employee SET ON_LETTERHEAD='Y', LH_ORDER=210, LH_HEADING='Account Representatives' WHERE CODE='KJB'; -- Kevin Byrnes

-- Verify:
-- SELECT CODE, FN, MI, LN, LH_ORDER, LH_HEADING, LH_SUFFIX, LH_SUFFIX_ITALIC, LH_BLANK_BEFORE
-- FROM employee WHERE ON_LETTERHEAD = 'Y' ORDER BY LH_ORDER;
