# Python Environment

This workspace has two virtual environments with different platform targets.

- Use `.venv-linux` when running from this shell/WSL/Linux environment.
- Do not use `.venv/Scripts/python.exe` from WSL. That is a Windows venv and is not reliably executable here.

Recommended commands from the workspace root:

```bash
. .venv-linux/bin/activate
python --version
python -c "import openpyxl, pandas; print(openpyxl.__version__, pandas.__version__)"
```

Direct execution without activation:

```bash
./.venv-linux/bin/python your_script.py
./.venv-linux/bin/pip list
```

Packages currently confirmed in `.venv-linux`:

- `openpyxl 3.1.5`
- `pandas 3.0.2`

If an AI or script needs Python package work in this workspace, default to `.venv-linux/bin/python`.
