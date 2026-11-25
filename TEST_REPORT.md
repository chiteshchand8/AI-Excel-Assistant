# V3.6.5 beta — test report

Evidence: developer/user Windows CMD logs and user-confirmed Excel checks; source inspection and syntax compilation in a separate environment. Windows Excel COM was not independently rerun in that environment.

| Milestone | Result |
|---|---|
| Installation/startup | Passed on user machine |
| Chart formatting regression | Functional pass; minor axis formatting issues |
| HR dashboard | User-confirmed pass, including slicer combinations |
| Sales dashboard | User-confirmed pass, including no-data and reset scenarios |
| Workbook preservation | Functional pass; no exhaustive cell/object comparison |
| Persistent file lock | Five retries after initial failure; recovery and original hash checks confirmed by user |
| Failed-operation rollback | All five diagnostic checks passed |

Rollback diagnostic: first KPI sheet created in stage; second operation rejected unknown column; original SHA256 unchanged; staging removed; matching backup found. The diagnostic bypassed Gemini intentionally.

Not established: successful replacement after a transient lock is released, exhaustive formula security, arbitrary workbook compatibility, full workbook preservation, comprehensive cleaning/validation/filter regression, clean-machine installation. GitHub publication itself is pending.
