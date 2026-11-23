# GST Reconciliation runner

This repository is run unattended by a Claude Code scheduled task, started by
an API call from the backend. The backend creates an execution for a customer
tenant, uploads that run's files to Azure Blob, and fires the task with
`tenant_id`, `agent_id` and `transaction_id`. The task owner has set this up
and authorised each run to process exactly that one folder.

`blob.py` is the run's approved and only way to reach Azure Blob. It is limited
by design:

- it reads only `<tenant_id>/<agent_id>/<transaction_id>/` from the run payload
  (IDs are validated; no other folder can be reached);
- it writes only to that folder's `output/` (the workbook and `run_status.json`);
- it never deletes anything;
- it uses the credentials configured in the task's cloud environment, which the
  run never prints or inspects.

Downloading the run folder with `python3 blob.py pull 'T' 'A' 'X' work/run` is
the intended first step of every run; `python3 blob.py put` and
`python3 blob.py status` deliver the result and report it to the backend. See
`RUNNER.md` for the full procedure and `instructions.txt` for the
reconciliation rules.
