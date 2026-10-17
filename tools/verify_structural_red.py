"""Run on Modal: prove the structural boundary tests detect disabled mutations."""
from pathlib import Path
import re
import subprocess

path = Path("crates/duke-sheets-core/src/workbook/structural.rs")
original = path.read_text()
try:
    disabled = original
    signatures = list(re.finditer(r"fn (edit_axis|rewrite_references|remap_name_scopes)\b[^\{]*\{", disabled))
    for match in reversed(signatures):
        start = match.end()
        depth = 1
        end = start
        while depth:
            if disabled[end] == "{": depth += 1
            if disabled[end] == "}": depth -= 1
            end += 1
        disabled = disabled[:start] + (" Ok(()) " if match.group(1) == "edit_axis" else " ") + disabled[end - 1:]
    path.write_text(disabled)
    for package, expected in [("duke-sheets-core", 6), ("duke-sheets", 1)]:
        result = subprocess.run(["cargo", "test", "-p", package, "--test", "structural_edits"], capture_output=True, text=True)
        log = result.stdout + result.stderr
        Path(f"structural-red-{package}.log").write_text(log)
        failures = re.findall(r"^test (\S+) \.\.\. FAILED$", result.stdout, re.M)
        print(f"{package}: {len(failures)}/{expected} tests failed with the mutations disabled")
        for test in failures: print(f"  {test}")
        if result.returncode == 0 or len(failures) != expected:
            print(log)
            raise SystemExit("Negative control did not fail as expected")
finally:
    path.write_text(original)
