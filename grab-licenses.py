
#!/usr/bin/env python3
"""
Export package homepage, name, declared license, version, and LICENSE text from node_modules into Excel.

Columns:
- A: Path -> LICENSE file path (if present) OR package directory path (if LICENSE missing)
- B: Homepage -> from manifest 'homepage' or derived from 'repository.url'
- C: Name -> from manifest 'name' or derived from node_modules path (supports @scope/pkg)
- D: License -> declared license from manifest (SPDX string or best-effort)
- E: Version -> from manifest 'version'
- F..N: License_Chunk_i -> LICENSE contents split into 32,767-character chunks

Enhancements:
- If homepage is missing, derive a candidate from repository.url.
- If a package has NO LICENSE file, still add a row for that package (Path = package dir).
- Robust handling of manifest formats, long license text, and nested node_modules.

Usage:
  python grab-licenses.py --root /path/to/node_modules --out licenses.xlsx
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
MANIFEST_BASENAMES = {"package.json", "package"}  # some systems or tooling may use 'package' without .json


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


def extract_license_value(data: dict) -> str | None:
    """
    Return the declared license for the package (best-effort).
    Supports:
      - license: "MIT"
      - license: { type/name/spdx: "MIT" }
      - licenses: [ { type/name/spdx: "MIT" }, ... ]  (deprecated)
    """
    if not data:
        return None

    lic = data.get("license") or data.get("licence")
    if isinstance(lic, str):
        cleaned = lic.strip()
        if cleaned:
            return cleaned
    elif isinstance(lic, dict):
        for key in ("type", "name", "spdx", "value"):
            v = lic.get(key)
            if isinstance(v, str) and v.strip():
                return v.strip()

    licenses = data.get("licenses")
    if isinstance(licenses, list) and licenses:
        vals = []
        for item in licenses:
            if isinstance(item, str):
                s = item.strip()
                if s:
                    vals.append(s)
            elif isinstance(item, dict):
                v = item.get("type") or item.get("name") or item.get("spdx") or item.get("value")
                if isinstance(v, str) and v.strip():
                    vals.append(v.strip())
        if vals:
            return "; ".join(vals)

    return None


def derive_package_name(pkg_dir: str) -> str | None:
    """
    Derive package name from its location under .../node_modules/.
    Handles nested node_modules and scopes, e.g.:
      /path/node_modules/@scope/pkg/... -> @scope/pkg
      /path/node_modules/pkg/...        -> pkg
    """
    norm = os.path.abspath(pkg_dir)
    parts = norm.split(os.sep)
    idx = None
    for i, p in enumerate(parts):
        if p == "node_modules":
            idx = i
    if idx is None or idx + 1 >= len(parts):
        return os.path.basename(pkg_dir)

    first = parts[idx + 1]
    if first.startswith("@") and idx + 2 < len(parts):
        return f"{first}/{parts[idx + 2]}"
    return first


def get_homepage(data: dict) -> str | None:
    """Return homepage, falling back to repository-derived URL."""
    if not data:
        return None
    homepage = data.get("homepage")
    if isinstance(homepage, str) and homepage.strip():
        return homepage.strip()
    repo_val = data.get("repository")
    derived = derive_homepage_from_repository(repo_val)
    if derived:
        return derived
    return None


def build_rows(manifests, node_modules_root):
    """
    For each manifest (i.e., each package), produce rows:
    [Path, Homepage, Name, License, Version, chunk1, chunk2, ...]
    Where Path = LICENSE file path if found, otherwise the PACKAGE DIRECTORY path.
    """
    rows = []
    for manifest_path in manifests:
        pkg_dir = os.path.dirname(manifest_path)
        data = parse_manifest(manifest_path)

        homepage = get_homepage(data) or ""
        declared_license = extract_license_value(data) or ""
        manifest_name = (data.get("name") or "").strip() if data else ""
        name = manifest_name or (derive_package_name(pkg_dir) or "")
        version = (data.get("version") or "").strip() if data else ""

        # Find LICENSE (if any)
        license_path = find_first_license(pkg_dir)
        if license_path:
            try:
                text = read_text_file(license_path)
                chunks = chunk_text(text, EXCEL_CELL_LIMIT)
                rows.append([license_path, homepage, name, declared_license, version, *chunks])
            except Exception as e:
                rows.append([license_path, homepage, name, declared_license, version, f"[Skipped: {e}]"])
        else:
            # No LICENSE: Path should be the package directory path
            rows.append([pkg_dir, homepage, name, declared_license, version])  # license columns remain empty
    return rows


def to_dataframe(rows):
    """Create a DataFrame with dynamic number of columns."""
    if not rows:
        return pd.DataFrame(columns=["Path", "Homepage", "Name", "License", "Version", "License_Chunk_1"])

    max_chunks = 0
    for r in rows:
        chunks_count = max(0, len(r) - 5)  # minus Path + Homepage + Name + License + Version
        if chunks_count > max_chunks:
            max_chunks = chunks_count

    columns = ["Path", "Homepage", "Name", "License", "Version"] + [f"License_Chunk_{i+1}" for i in range(max_chunks)]

    normalized = []
    for r in rows:
        path = r[0] if len(r) > 0 else ""
        homepage = r[1] if len(r) > 1 else ""
        name = r[2] if len(r) > 2 else ""
        declared_license = r[3] if len(r) > 3 else ""
        version = r[4] if len(r) > 4 else ""
        chunks = r[5:] if len(r) > 5 else []
        padded = chunks + [""] * (max_chunks - len(chunks))
        normalized.append([path, homepage, name, declared_license, version] + padded)

    return pd.DataFrame(normalized, columns=columns)


def main():
    parser = argparse.ArgumentParser(description="Export homepage, name, license, version, and LICENSE contents from node_modules to Excel.")
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
        ws.column_dimensions["C"].width = 40   # Name
        ws.column_dimensions["D"].width = 24   # License (declared)
        ws.column_dimensions["E"].width = 16   # Version
        for col_idx in range(6, ws.max_column + 1):
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 60

    print(f"Done. Wrote {len(rows)} row(s) to {args.out}")


if __name__ == "__main__":
    main()
