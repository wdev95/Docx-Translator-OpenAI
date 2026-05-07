from __future__ import annotations

import shlex
import subprocess
import sys
from pathlib import Path


def main() -> int:
    project_dir = Path(__file__).resolve().parent
    translate_script = project_dir / "translate.py"

    if not translate_script.exists():
        print("Error: translate.py not found next to this script.")
        return 1

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        "DOCX-Translator",
        "--collect-all",
        "svc_ttk",
        "--hidden-import",
        "darkdetect",
        translate_script.name,
    ]

    print("Building executable with command:")
    print(" ".join(shlex.quote(part) for part in cmd))

    completed = subprocess.run(cmd, cwd=project_dir)

    if completed.returncode == 0:
        dist_exe = project_dir / "dist" / "DOCX-Translator" / "DOCX-Translator.exe"
        print("\nBuild finished successfully.")
        print(f"Executable: {dist_exe}")
        print("JSON files are not bundled; they will be created at first app start.")
    else:
        print("\nBuild failed.")

    return int(completed.returncode)


if __name__ == "__main__":
    raise SystemExit(main())
