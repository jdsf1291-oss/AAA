# AGENTS.md

## Cursor Cloud specific instructions

### Repository state
- The `main` branch is a placeholder: it contains only `README.md` (`# AAA`). There is no application, service, web/API server, database, build step, test suite, or dependency manifest on `main`.
- The only functional code lives on the feature branch `origin/cursor/tavr-ecg-pacemaker-ppt-02bf`: `generate_tavr_ppt.py`, a standalone Python script that uses `python-pptx` to generate `TAVR_ECG_Pacemaker.pptx` (a 15-slide Chinese-language medical deck about TAVR periprocedural ECG assessment and pacemaker risk). It is a one-off document generator, not a long-running service.

### Toolchain
- Python 3.12 and Node 22 are preinstalled.
- Non-obvious caveat: the base image ships without `ensurepip`, so `python3 -m venv` fails until `python3.12-venv` is installed (`sudo apt-get install -y python3.12-venv`). Alternatively, use `pip install --break-system-packages` against the system interpreter.

### Running the only application (PPT generator)
- The script's sole third-party dependency is `python-pptx`; there is no `requirements.txt`, so install it directly.
- Example:
  - `python3 -m venv .venv && ./.venv/bin/pip install python-pptx`
  - `./.venv/bin/python generate_tavr_ppt.py` → writes `TAVR_ECG_Pacemaker.pptx` and prints `Saved TAVR_ECG_Pacemaker.pptx with 15 slides`.
- There are no lint, test, or build commands configured anywhere in the repository.
