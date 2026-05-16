#!/usr/bin/env python3
"""
image_to_docx.py
────────────────
Convert images (single file / folder / ZIP) of Word-document content
into a real .docx file using:
  1. Google Gemini API  → generates npm-docx JS code
  2. Node.js on your PC → runs the JS → output.docx

Requirements
  pip install google-genai Pillow
  node + npm  (https://nodejs.org)

Usage
  python image_to_docx.py scan.jpg
  python image_to_docx.py screenshots/
  python image_to_docx.py pages.zip
  python image_to_docx.py pages.zip -o results/ -n report.docx --save-js

API key
  Set the GEMINI_API_KEY environment variable, or pass -k YOUR_KEY.
"""

import os, sys, io, zipfile, shutil, subprocess, argparse, re
from pathlib import Path

# ── auto-install Python deps ───────────────────────────────────────────────────
def _pip(*pkgs):
    subprocess.run([sys.executable, "-m", "pip", "install", *pkgs, "-q"], check=True)

try:
    from google import genai
    from google.genai import types as gtypes
except ImportError:
    print("Installing google-genai …")
    _pip("google-genai")
    from google import genai
    from google.genai import types as gtypes

try:
    from PIL import Image
except ImportError:
    print("Installing Pillow …")
    _pip("Pillow")
    from PIL import Image


# ── constants ──────────────────────────────────────────────────────────────────
MIME_MAP = {
    ".png":  "image/png",
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".heic": "image/heic",
    ".heif": "image/heif",
}
EXTRA_EXT = {".bmp", ".tif", ".tiff"}   # auto-converted to PNG before sending

GEMINI_MODEL = "gemini-3-flash-preview"

GEMINI_PROMPT = r"""
You are a document-reconstruction assistant.

TASK
----
I am giving you one or more images that each show one page of a document
(possibly a Microsoft Word document). Your job is to output a **single,
complete, runnable Node.js script** that uses the `docx` npm package to
recreate that document as accurately as possible, then saves it as
output.docx in the current working directory.

RULES
-----
1. Determine the correct page order from the images before writing anything.
2. If the source is already formatted (headings, bullets, tables, bold/italic,
   indents, alignment …) reproduce that formatting as faithfully as possible.
3. If the source is plain text, apply standard Word formatting:
   A4 page, Times New Roman 14 pt body, 2.5 cm margins.
4. The script MUST end exactly with:
       Packer.toBuffer(doc).then(buf => {
         fs.writeFileSync("output.docx", buf);
         console.log("output.docx written successfully");
       });
5. Only import from 'docx' and 'fs'.
6. Return ONLY the JavaScript — no markdown fences, no prose explanation.

REFERENCE SKELETON (adapt as needed)
--------------------------------------
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, HeadingLevel, LevelFormat,
  Header, Footer, PageNumber
} = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 28 } }   // 14 pt
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1",
        basedOn: "Normal", next: "Normal", quickFormat: true,
        run:       { size: 32, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2",
        basedOn: "Normal", next: "Normal", quickFormat: true,
        run:       { size: 28, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 }
      }
    ]
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2013",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }]
    }]
  },
  sections: [{
    properties: {
      page: {
        size:   { width: 11906, height: 16838 },   // A4 twips
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
      }
    },
    children: [
      /* all Paragraph / Table nodes go here */
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("output.docx", buf);
  console.log("output.docx written successfully");
});
"""


# ── image helpers ──────────────────────────────────────────────────────────────
def _load_image(path: Path) -> tuple[bytes, str] | None:
    """Return (raw_bytes, mime_type).  Unsupported formats are converted to PNG."""
    ext = path.suffix.lower()
    if ext in MIME_MAP:
        return path.read_bytes(), MIME_MAP[ext]
    if ext in EXTRA_EXT:
        try:
            buf = io.BytesIO()
            Image.open(path).convert("RGB").save(buf, format="PNG")
            return buf.getvalue(), "image/png"
        except Exception as e:
            print(f"  ⚠  Cannot convert {path.name}: {e}")
    return None


def collect_images(input_path: str) -> list[tuple[str, bytes, str]]:
    """
    Return sorted list of (filename, bytes, mime_type) from:
      • a single image file
      • a folder of images (recursive)
      • a ZIP archive of images
    """
    root = Path(input_path)
    if not root.exists():
        raise FileNotFoundError(f"Not found: {input_path}")

    entries: list[tuple[str, bytes, str]] = []

    def _scan(directory: Path):
        for p in sorted(directory.rglob("*")):
            if p.is_file() and p.suffix.lower() in (set(MIME_MAP) | EXTRA_EXT):
                result = _load_image(p)
                if result:
                    entries.append((p.name, *result))

    if root.is_dir():
        _scan(root)
    elif root.suffix.lower() == ".zip":
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(root) as zf:
                zf.extractall(tmp)
            _scan(Path(tmp))
    else:
        result = _load_image(root)
        if result:
            entries.append((root.name, *result))

    if not entries:
        raise ValueError("No supported image files found.")

    print(f"  Found {len(entries)} image(s): {[e[0] for e in entries]}")
    return entries


# ── Gemini ─────────────────────────────────────────────────────────────────────
def call_gemini(images: list[tuple[str, bytes, str]], api_key: str) -> str:
    """Send images + prompt to Gemini; return cleaned JS source."""
    client = genai.Client(api_key=api_key)

    parts: list = [GEMINI_PROMPT]
    for name, data, mime in images:
        print(f"    Attaching: {name}  ({mime})")
        parts.append(gtypes.Part.from_bytes(data=data, mime_type=mime))

    print("  Calling Gemini …")
    resp = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=parts,
        config=gtypes.GenerateContentConfig(max_output_tokens=8192),
    )

    js = resp.text.strip()
    # Strip markdown fences the model sometimes adds despite instructions
    js = re.sub(r"^```[a-zA-Z]*\n?", "", js)
    js = re.sub(r"\n?```\s*$",        "", js)
    return js.strip()


# ── Node.js runner ─────────────────────────────────────────────────────────────

# Common Windows installation directories for Node.js
_WIN_NODE_DIRS = [
    r"C:\Program Files\nodejs",
    r"C:\Program Files (x86)\nodejs",
    r"D:\Program Files\nodejs",
    r"D:\Program Files (x86)\nodejs",
]

def _find_cmd(name: str) -> str:
    """
    Return the full path to `node` or `npm`, searching:
      1. The system PATH (works on Linux/macOS and correctly-configured Windows).
      2. Common Windows install directories (fallback when PATH is stale).
    Raises RuntimeError if neither works.
    """
    # 1 — try PATH first
    try:
        r = subprocess.run([name, "--version"], capture_output=True)
        if r.returncode == 0:
            return name   # found on PATH
    except FileNotFoundError:
        pass

    # 2 — search known Windows directories
    if sys.platform == "win32":
        exe = name + ".cmd" if name == "npm" else name + ".exe"
        for d in _WIN_NODE_DIRS:
            candidate = os.path.join(d, exe)
            if os.path.isfile(candidate):
                # Verify it actually runs
                try:
                    r = subprocess.run([candidate, "--version"], capture_output=True)
                    if r.returncode == 0:
                        print(f"  Found {name} at: {candidate}")
                        return candidate
                except Exception:
                    pass

    raise RuntimeError(
        f"'{name}' not found.\n"
        "Install Node.js (includes npm) from https://nodejs.org\n"
        "After installing, either:\n"
        "  • Restart your terminal, or\n"
        "  • Pass the install folder via --node-dir  e.g.  --node-dir \"D:\\Program Files\\nodejs\""
    )


def _resolve_node_paths(node_dir: str | None = None) -> tuple[str, str]:
    """
    Return (node_exe, npm_exe) paths.
    If node_dir is given, build paths directly from it (bypasses PATH entirely).
    """
    if node_dir:
        if sys.platform == "win32":
            node_exe = os.path.join(node_dir, "node.exe")
            npm_exe  = os.path.join(node_dir, "npm.cmd")
        else:
            node_exe = os.path.join(node_dir, "node")
            npm_exe  = os.path.join(node_dir, "npm")
        for p in (node_exe, npm_exe):
            if not os.path.isfile(p):
                raise RuntimeError(f"Not found in --node-dir: {p}")
        return node_exe, npm_exe
    return _find_cmd("node"), _find_cmd("npm")


# Persistent folder that survives between runs — npm install only happens once
WORK_DIR = Path.home() / ".img2docx"


def _docx_installed(node_exe: str, npm_exe: str) -> bool:
    """
    Return True if the `docx` package is already available to Node — either
    globally (npm list -g docx) or locally in WORK_DIR.
    """
    # 1. local WORK_DIR node_modules
    if (WORK_DIR / "node_modules" / "docx").is_dir():
        return True

    # 2. global npm install
    r = subprocess.run(
        [npm_exe, "list", "-g", "docx", "--depth=0"],
        capture_output=True, text=True,
    )
    if r.returncode == 0 and "docx" in r.stdout:
        return True

    return False


def _ensure_docx(node_exe: str, npm_exe: str):
    """Install docx into WORK_DIR once; skip if already present."""
    if _docx_installed(node_exe, npm_exe):
        print("  docx already installed — skipping npm install")
        return

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    print(f"  npm install docx  (one-time, into {WORK_DIR}) …")
    r = subprocess.run(
        [npm_exe, "install", "docx"],
        cwd=str(WORK_DIR),
        capture_output=True,
        text=True,
    )
    if r.returncode != 0:
        raise RuntimeError(f"npm install failed:\n{r.stderr.strip()}")
    print("  docx installed successfully")


def run_node(js_code: str, output_path: str, keep_workdir: bool = False,
             node_dir: str | None = None) -> str:
    """
    1. Ensure docx is available (install once into ~/.img2docx if needed).
    2. Write the generated JS to ~/.img2docx/docx_gen.js.
    3. node docx_gen.js  →  output.docx  (written into ~/.img2docx/).
    4. Copy the .docx to output_path.

    Parameters
    ----------
    js_code       JS source returned by Gemini.
    output_path   Destination path for the finished .docx.
    keep_workdir  Print the work dir path so you can inspect it manually.
    node_dir      Explicit Node.js install folder (auto-detected when omitted).

    Returns
    -------
    Absolute path of the output .docx.
    """
    node_exe, npm_exe = _resolve_node_paths(node_dir)
    WORK_DIR.mkdir(parents=True, exist_ok=True)

    # ── ensure docx package ──────────────────────────────────────────────────
    _ensure_docx(node_exe, npm_exe)

    # ── write JS ─────────────────────────────────────────────────────────────
    js_file = WORK_DIR / "docx_gen.js"
    js_file.write_text(js_code, encoding="utf-8")
    if keep_workdir:
        print(f"  Work dir: {WORK_DIR}")

    # ── run node ─────────────────────────────────────────────────────────────
    print("  node docx_gen.js …")
    r = subprocess.run(
        [node_exe, "docx_gen.js"],
        cwd=str(WORK_DIR),
        capture_output=True,
        text=True,
        timeout=120,
    )
    if r.stdout.strip():
        print("  node:", r.stdout.strip())
    if r.returncode != 0:
        raise RuntimeError(
            f"node exited with code {r.returncode}:\n"
            f"{r.stderr.strip()}\n{r.stdout.strip()}\n\n"
            "Tip: run with --save-js to inspect the generated script."
        )

    # ── collect output ───────────────────────────────────────────────────────
    # The JS writes output.docx into cwd (WORK_DIR)
    generated = WORK_DIR / "output.docx"
    if not generated.exists():
        docx_files = list(WORK_DIR.glob("*.docx"))
        if not docx_files:
            raise RuntimeError(
                "Node.js ran without error but produced no .docx.\n"
                "Run with --save-js and check the generated JS manually."
            )
        generated = docx_files[0]

    shutil.copy2(str(generated), output_path)
    return str(Path(output_path).resolve())


# ── public API ─────────────────────────────────────────────────────────────────
def convert_images_to_docx(
    input_path: str,
    output_dir: str = ".",
    api_key: str | None = None,
    output_filename: str = "output.docx",
    save_js: bool = False,
    js_filename: str = "generated.js",
    gemini_model: str | None = None,
    keep_workdir: bool = False,
    node_dir: str | None = None,
) -> str:
    """
    Full pipeline: images → Gemini JS → Node.js → .docx

    Parameters
    ----------
    input_path       Path to an image, a folder of images, or a ZIP file.
    output_dir       Directory to write the .docx (and optionally the .js).
    api_key          Gemini API key; falls back to $GEMINI_API_KEY env var.
    output_filename  Name of the output Word file (default: output.docx).
    save_js          Also save the generated JS next to the .docx.
    js_filename      Filename for the saved JS (default: generated.js).
    gemini_model     Override which Gemini model to call.
    keep_workdir     Keep the Node.js temp folder (useful to debug JS errors).
    node_dir         Explicit Node.js install folder, e.g. D:\\Program Files\\nodejs
                     Only needed when node/npm are not on PATH.

    Returns
    -------
    Absolute path to the generated .docx.
    """
    global GEMINI_MODEL
    if gemini_model:
        GEMINI_MODEL = gemini_model

    api_key = api_key or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError(
            "Gemini API key required.\n"
            "  • Pass  api_key='...'  OR\n"
            "  • Export  GEMINI_API_KEY=your_key  in your shell."
        )

    os.makedirs(output_dir, exist_ok=True)
    output_path = str(Path(output_dir) / output_filename)
    sep = "─" * 52

    # ── 1. collect images ─────────────────────────────────────────────────────
    print(f"\n{sep}\nStep 1 │ Loading images  ←  {input_path}")
    images = collect_images(input_path)

    # ── 2. Gemini → JS ────────────────────────────────────────────────────────
    print(f"\n{sep}\nStep 2 │ Gemini ({GEMINI_MODEL}) → JS")
    js_code = call_gemini(images, api_key)
    print(f"  JS length: {len(js_code):,} chars")

    if len(js_code) < 100:
        print(f"  ⚠  Very short response:\n{js_code}")

    if save_js:
        js_path = str(Path(output_dir) / js_filename)
        Path(js_path).write_text(js_code, encoding="utf-8")
        print(f"  JS saved  →  {js_path}")

    # ── 3. Node.js → .docx ───────────────────────────────────────────────────
    print(f"\n{sep}\nStep 3 │ Node.js  →  {output_filename}")
    out = run_node(js_code, output_path, keep_workdir=keep_workdir, node_dir=node_dir)

    print(f"\n{sep}\n✓  Done!  →  {out}\n{sep}\n")
    return out


# ── CLI ────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description="Convert document images → .docx  (Gemini + Node.js npm docx)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  python image_to_docx.py scan.jpg
  python image_to_docx.py pages/
  python image_to_docx.py scans.zip  -o output/  -n report.docx
  python image_to_docx.py scans.zip  --save-js
  python image_to_docx.py scans.zip  --keep-workdir     # debug JS errors
        """,
    )
    ap.add_argument("input",
        help="Image file, folder of images, or ZIP archive")
    ap.add_argument("-o", "--output-dir",  default=".",   metavar="DIR",
        help="Output directory  (default: current directory)")
    ap.add_argument("-n", "--output-name", default="output.docx", metavar="FILE",
        help="Output .docx filename  (default: output.docx)")
    ap.add_argument("-k", "--api-key",     default=None,  metavar="KEY",
        help="Gemini API key  (or set $GEMINI_API_KEY)")
    ap.add_argument("--save-js",           action="store_true",
        help="Save the generated JavaScript alongside the .docx")
    ap.add_argument("--js-name",           default="generated.js", metavar="FILE",
        help="Filename for the saved JS  (default: generated.js)")
    ap.add_argument("--model",             default=None,  metavar="MODEL",
        help=f"Gemini model  (default: {GEMINI_MODEL})")
    ap.add_argument("--keep-workdir",      action="store_true",
        help="Keep the Node.js temp folder after running  (debug aid)")
    ap.add_argument("--node-dir",          default=None,  metavar="PATH",
        help=r'Explicit Node.js install folder, e.g. "D:\Program Files\nodejs"  '
             "(only needed when node/npm are not on PATH)")

    args = ap.parse_args()

    try:
        convert_images_to_docx(
            input_path      = args.input,
            output_dir      = args.output_dir,
            api_key         = args.api_key,
            output_filename = args.output_name,
            save_js         = args.save_js,
            js_filename     = args.js_name,
            gemini_model    = args.model,
            keep_workdir    = args.keep_workdir,
            node_dir        = args.node_dir,
        )
    except Exception as exc:
        print(f"\n❌  {exc}", file=sys.stderr)
        sys.exit(1)