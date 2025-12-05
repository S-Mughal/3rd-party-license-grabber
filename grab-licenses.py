"""
Export package homepage and LICENSE text from node_modules into Excel.

Enhancements in this version:
- If homepage is missing, derive a candidate from repository.url.
- If a package has NO LICENSE, still add a row for that package.
- Column A (Path) now shows the package directory path when NO LICENSE exists.

Columns:
- A: Path to LICENSE file (if present) OR package directory path (if LICENSE missing)
- B: Homepage (homepage or derived from repository.url)
- C..N: License contents split into 32,767-character chunks

Usage:
  python export_licenses_to_excel.py --root /path/to/node_modules --out licenses.xlsx
"""

import argparse
import json
import os
import re
import sys
import unicodedata
import urllib.parse

import pandas as pd

# Excel cell content limit
EXCEL_CELL_LIMIT = 32767

# Recognized license file names (case-insensitive)
LICENSE_BASENAMES = {
    "license", "license.txt", "license.md", "licence", "licence.txt", "licence.md",
    "copying", "copying.txt", "copying.md"  # some projects use COPYING
}

# Recognized manifest file names (case-insensitive)
MANIFEST_BASENAMES = {"package.json", "package"}  # user mentioned 'package' that is a package.json


def is_probably_text(data: bytes) -> bool:
    """Heuristic: treat as text if decodes with limited errors and has no NULs."""
    if b"\x00" in data:
        return False
    try:
        text = data.decode("utf-8", errors="replace")
    except Exception:
        return False
    replacements = text.count("\ufffd")
    return replacements < max(5, len(text) * 0.002)


def read_text_file(path: str) -> str:
    """Read file as UTF-8 (fallback to latin-1) with best-effort decoding."""
    with open(path, "rb") as f:
        raw = f.read()
    if not is_probably_text(raw):
        raise ValueError("Binary or non-text file")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1", errors="replace")


def chunk_text(s: str, size: int):
    """Split text into size-limited chunks, respecting Excel cell limits."""
    s = unicodedata.normalize("NFC", s)
    return [s[i:i+size] for i in range(0, len(s), size)] or [""]


def find_manifests(root: str):
    """Yield paths to manifests (package.json or 'package') under node_modules."""
    root = os.path.abspath(root)
    for dirpath, _, filenames in os.walk(root, followlinks=True):
        if not filenames:
            continue
        lower_map = {fn.lower(): fn for fn in filenames}
        for candidate in MANIFEST_BASENAMES:
            if candidate in lower_map:
                yield os.path.join(dirpath, lower_map[candidate])


def find_first_license(package_dir: str) -> str | None:
    """
    Search for the first LICENSE-like file under the package directory (deep search).
    Returns the absolute path or None if not found.
    """
    for dirpath, _, filenames in os.walk(package_dir, followlinks=True):
        for name in filenames:
            lower = name.lower()
            if lower in LICENSE_BASENAMES or lower.startswith("license"):
                return os.path.join(dirpath, name)
    return None


def parse_manifest(manifest_path: str) -> dict | None:
    """Parse manifest JSON file into dict (best-effort)."""
    try:
        with open(manifest_path, "rb") as f:
            raw = f.read()
        text = raw.decode("utf-8", errors="replace")
        return json.loads(text)
    except Exception:
        return None


def derive_homepage_from_repository(repo_val) -> str | None:
    """
    Derive a human-friendly homepage URL from a repository value.
    Handles:
      - dict with {"url": "..."} or string
      - git+, ssh://, git@host:path, http(s)://
      - short forms like 'github:user/repo'
    """
    if repo_val is None:
        return None

    # Pull URL string from either dict or string
    if isinstance(repo_val, dict):
        url = (repo_val.get("url") or "").strip()
    elif isinstance(repo_val, str):
        url = repo_val.strip()
    else:
        return None

    if not url:
        return None

    # Remove 'git+' prefix when present
    if url.startswith("git+"):
        url = url[4:]

    # ssh://git@github.com/user/repo.git -> https://github.com/user/repo
    if url.startswith("ssh://"):
        parsed = urllib.parse.urlparse(url)
        host = parsed.hostname
        path = parsed.path or ""
        path = path[1:] if path.startswith("/") else path
        if path.endswith(".git"):
            path = path[:-4]
        if host and path:
            return f"https://{host}/{path}"

    # git@github.com:user/repo.git -> https://github.com/user/repo
    m = re.match(r"^git@([^:]+):(.+)$", url)
    if m:
        host = m.group(1)
        path = m.group(2)
        if path.endswith(".git"):
            path = path[:-4]
        return f"https://{host}/{path}"

    # http(s)://...(.git) -> same without .git
    if url.startswith("http://") or url.startswith("https://"):
        if url.endswith(".git"):
            url = url[:-4]
        return url

    # Short forms: github:user/repo, gitlab:user/repo, bitbucket:user/repo
    if ":" in url and not url.startswith(("http://", "https://")):
        host_key, path = url.split(":", 1)
        host_map = {"github": "github.com", "gitlab": "gitlab.com", "bitbucket": "bitbucket.org"}
        domain = host_map.get(host_key.lower())
        if domain and path:
            return f"https://{domain}/{path}"

    return None


def get_homepage(data: dict) -> str | None:
    """Return homepage, falling back to repository-derived URL."""
    if not data:
        return None
    homepage = data.get("homepage")
    if isinstance(homepage, str) and homepage.strip():
        return homepage.strip()

    # Fallback to repository.url or repository string
    repo_val = data.get("repository")
    derived = derive_homepage_from_repository(repo_val)
    if derived:
        return derived

    # Optional further fallback (e.g., bugs.url) intentionally omitted per requirements
    return None


def build_rows(manifests, node_modules_root):
    """
    For each manifest (i.e., each package), produce rows:
    [Path, Homepage, chunk1, chunk2, ...]
    Where Path = LICENSE file path if found, otherwise the PACKAGE DIRECTORY path.
    """
    rows = []
    for manifest_path in manifests:
        pkg_dir = os.path.dirname(manifest_path)
        data = parse_manifest(manifest_path)
        homepage = get_homepage(data) or ""

        # Find LICENSE (if any)
        license_path = find_first_license(pkg_dir)
        if license_path:
            try:
                text = read_text_file(license_path)
                chunks = chunk_text(text, EXCEL_CELL_LIMIT)
                rows.append([license_path, homepage, *chunks])
            except Exception as e:
                # Unreadable license; still add row with path and error note
                rows.append([license_path, homepage, f"[Skipped: {e}]"])
        else:
            # No LICENSE: Path should be the package directory path
            rows.append([pkg_dir, homepage])  # license columns remain empty
    return rows


def to_dataframe(rows):
    """Create a DataFrame with dynamic number of columns."""
    if not rows:
        return pd.DataFrame(columns=["Path", "Homepage", "License_Chunk_1"])

    # Max number of license chunks across rows
    max_chunks = 0
    for r in rows:
        chunks_count = max(0, len(r) - 2)  # minus Path + Homepage
        if chunks_count > max_chunks:
            max_chunks = chunks_count

    columns = ["Path", "Homepage"] + [f"License_Chunk_{i+1}" for i in range(max_chunks)]

    # Pad rows so they all have the same number of chunk columns
    normalized = []
    for r in rows:
        path = r[0] if len(r) > 0 else ""
        homepage = r[1] if len(r) > 1 else ""
        chunks = r[2:] if len(r) > 2 else []
        padded = chunks + [""] * (max_chunks - len(chunks))
        normalized.append([path, homepage] + padded)

    return pd.DataFrame(normalized, columns=columns)


def main():
    parser = argparse.ArgumentParser(description="Export homepage and LICENSE contents from node_modules to Excel.")
    parser.add_argument("--root", default="node_modules", help="Path to node_modules root (default: ./node_modules)")
    parser.add_argument("--out", default="licenses.xlsx", help="Output Excel filename (default: licenses.xlsx)")
    args = parser.parse_args()

    node_modules_root = os.path.abspath(args.root)
    if not os.path.isdir(node_modules_root):
        print(f"ERROR: Directory not found: {args.root}", file=sys.stderr)
        sys.exit(1)

    print(f"Scanning for package manifests under: {node_modules_root}")
    manifests = list(find_manifests(node_modules_root))
    print(f"Found {len(manifests)} package manifest(s). Building workbook...")

    rows = build_rows(manifests, node_modules_root)
    df = to_dataframe(rows)

    # Write using openpyxl engine
    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Licenses")
        ws = writer.sheets["Licenses"]
        ws.column_dimensions["A"].width = 120  # Path (license file OR package dir)
        ws.column_dimensions["B"].width = 60   # Homepage
        for col_idx in range(3, ws.max_column + 1):
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 60

    print(f"Done. Wrote {len(rows)} row(s) to {args.out}")


if __name__ == "__main__":
    main()
