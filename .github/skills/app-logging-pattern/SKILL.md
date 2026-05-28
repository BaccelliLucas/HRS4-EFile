---
name: app-logging-pattern
description: 'Understand, extend, or debug the file-based logging system in this project. Use when dealing with app.log, streamlit.log, logging.basicConfig, logger.exception, uncaught exceptions, Streamlit runtime errors, or Microsoft Graph request tracing.'
argument-hint: 'Describe the logging change, failure, or flow you want to inspect'
user-invocable: true
---

# App Logging Pattern

Use this skill when you need to understand, change, or standardize how this workspace records operational logs and surfaces failures to the operator.

## What This Project Uses

- The application uses Python's standard `logging` module, not `print(...)` and not a custom logger wrapper.
- The main runtime log is `app.log`, written next to the executable in packaged mode or next to the script in development mode.
- The Streamlit subprocess writes its stdout and stderr to a separate `streamlit.log` file.
- Modules use named loggers such as `ui` and `msgraph`, but handler setup is currently driven by `logging.basicConfig(...)` in entry-point code.
- Fatal errors are not only logged. They are also surfaced to the operator through `tkinter.messagebox` or `streamlit.error`, depending on the execution path.

## Logging Topology In This Repository

### Launcher flow

- `main.py` is the launcher and primary logging bootstrap.
- `setup_logging()` resolves the runtime root directory, configures `app.log`, and installs a custom `sys.excepthook`.
- The format used there is:
  `%(asctime)s [%(levelname)s] %(message)s`
- `main.py` also logs launcher lifecycle events such as:
  - application startup
  - Python interpreter discovery
  - Streamlit process start and exit
  - early process termination

### Streamlit UI flow

- `ui.py` configures logging again for the Streamlit runtime, also targeting `app.log`.
- The format used there is:
  `%(asctime)s [%(levelname)s] %(name)s - %(message)s`
- `ui.py` uses `logger = logging.getLogger("ui")` and records:
  - module startup
  - outbound event submission metadata
  - success after HTTP 201
  - failure bodies returned by Microsoft Graph
  - uncaught exceptions around `msgraph.create_event(...)`

### Microsoft Graph flow

- `msgraph.py` does not configure handlers directly.
- It relies on the active logging configuration and uses `logger = logging.getLogger("msgraph")`.
- It logs:
  - whether required environment variables are present, without printing the secret value
  - login start and success
  - token acquisition failures
  - outbound POST target URL
  - HTTP response status from Graph

## File Responsibilities

- `main.py`: bootstrap logging, install uncaught exception hook, supervise the Streamlit subprocess, and guide the operator to the right log file.
- `ui.py`: log user-triggered actions and convert backend failures into Streamlit-visible messages.
- `msgraph.py`: log authentication and HTTP request lifecycle without owning global logging configuration.

## When To Use

- You need to add logs to a new flow and want them to land in `app.log` consistently.
- You are debugging why a user saw a popup or a Streamlit error and need the matching log path.
- You want to trace Microsoft Graph authentication or event creation failures.
- You need to understand why there are two log files, `app.log` and `streamlit.log`.
- You want to refactor logging without breaking the packaged executable behavior.
- You need to review whether a module should call `logging.basicConfig(...)` or only `getLogger(...)`.

## Procedure

1. Identify which runtime owns the failure.
   Use `main.py` for launcher and process-start problems, `ui.py` for Streamlit interactions, and `msgraph.py` for Graph authentication or event POST issues.

2. Confirm which file should contain the evidence.
   Use `app.log` for application-level logging and `streamlit.log` for raw subprocess stdout or stderr from the Streamlit server.

3. Preserve the existing root-path convention.
   When adding or moving logging setup, continue resolving the runtime directory from `sys.executable` when frozen and from `__file__` during development.

4. Use named loggers in modules.
   Prefer `logging.getLogger("ui")` or `logging.getLogger("msgraph")` in leaf modules instead of re-creating custom handlers there.

5. Log operator-relevant milestones.
   Record starts, external calls, response codes, and exceptions. Do not log every trivial variable assignment.

6. Keep secrets out of the log.
   Presence checks like `'<set>'` versus `'<VAZIO>'` are acceptable. Secret values themselves are not.

7. Mirror failures in the correct UI surface.
   Launcher failures should continue to show `messagebox` dialogs. Streamlit interaction failures should continue to show `st.error(...)` while preserving the stack trace in `app.log`.

8. Validate both packaged and development assumptions.
   Any change to path resolution or logging bootstrap must still work when the app runs from a bundled executable.

## Message Conventions Seen In The Project

Prefer messages shaped like these when extending the system:

- launcher start: `Iniciando aplicação. root=<path>`
- dependency discovery: `Python encontrado: <path>`
- process lifecycle: `Streamlit iniciado. PID=<pid>`
- operator-visible fatal error: `Exceção não capturada`
- UI action: `Enviando evento: subject=..., start=..., end=..., attendees=<n>`
- backend success: `Evento criado com sucesso (201)`
- backend failure: `Erro ao criar evento: status=<code> body=<text>`
- auth lifecycle: `Iniciando login via MSAL (Client Credentials)`
- HTTP lifecycle: `POST https://graph.microsoft.com/...`

## Implementation Rules

- Keep using the standard `logging` module unless the project explicitly decides to migrate.
- Do not replace file logging with console-only logging; this project depends on persistent local log files.
- Use `logger.exception(...)` inside `except` blocks when the traceback matters.
- Use `logger.error(...)` when you already have a meaningful failure message and do not need an automatic traceback.
- Keep log entries short, technical, and directly actionable.
- If you add a new module, prefer a named logger and let bootstrap code own the global configuration.
- Treat `streamlit.log` as subprocess diagnostics, not as the main application audit trail.

## Risks And Caveats

- This repository currently configures `logging.basicConfig(...)` in more than one place. Any refactor should account for the fact that `basicConfig` only takes effect once unless handlers are reset or `force=True` is used.
- `main.py` and `ui.py` use slightly different format strings. If you standardize them, do it intentionally and verify downstream troubleshooting still works.
- Errors may be reported twice by design: once in the UI surface and once in `app.log`. That duplication is operationally useful here.

## Review Checklist

- New logs go to the correct file for the owning runtime.
- New module code uses `getLogger(...)` instead of configuring handlers ad hoc.
- Exceptions visible to the user still leave enough detail in `app.log`.
- Secrets, tokens, and raw credentials are not written to disk.
- Packaged execution still writes logs next to the executable.

## Expected Outcomes

- You can tell where to add logs without broad repo exploration.
- You know which file to inspect for launcher, Streamlit, and Graph issues.
- You can extend the logging system while preserving operator-facing error behavior.