import gradio as gr
import tempfile
import os
import shutil
import subprocess
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent

def run_cmd(cmd, cwd=None, env=None):
    """Run a command, print nice logs, and also save them to run.log in cwd."""
    cwd = str(cwd or os.getcwd())
    print(f"🟦 Running: {' '.join(cmd)}  (cwd={cwd})")
    proc = subprocess.run(
        cmd,
        cwd=cwd,
        env=env,
        capture_output=True,
        text=True
    )
    if proc.stdout:
        print("🟩 STDOUT:")
        print(proc.stdout)
    if proc.stderr:
        print("🟥 STDERR:")
        print(proc.stderr)
    # Save to run.log for debugging
    try:
        runlog = Path(cwd) / "run.log"
        with open(runlog, "a", encoding="utf-8") as f:
            f.write(f"$ {' '.join(cmd)}\n")
            if proc.stdout:
                f.write(proc.stdout + "\n")
            if proc.stderr:
                f.write(proc.stderr + "\n")
        print(f"🧾 Run log saved to: {runlog}")
    except Exception as e:
        print(f"⚠️ Could not write run.log: {e}")

    if proc.returncode != 0:
        # Let Gradio see the failure so it surfaces properly
        raise subprocess.CalledProcessError(proc.returncode, cmd, proc.stdout, proc.stderr)
    return proc

def _locate_pdf_json(temp_dir: str) -> str:
    """
    Your extractor writes a JSON like <pdf_stem>_comprehensive_data.json.
    Find it (and a few common fallbacks). Raise if not found.
    """
    td = Path(temp_dir)

    # Prefer exactly-named file if present
    candidates = [
        td / "pdf_data.json",                    # legacy name (if ever created)
        td / "input_comprehensive_data.json",    # most common from your logs
        td / "comprehensive_data.json",          # another common alias
        td / "output.json",                      # generic
    ]
    for p in candidates:
        if p.exists():
            print(f"✅ Using PDF JSON: {p}")
            return str(p)

    # Generic pattern: anything *_comprehensive_data.json
    globs = list(td.glob("*_comprehensive_data.json"))
    if globs:
        print(f"✅ Using PDF JSON (glob): {globs[0]}")
        return str(globs[0])

    # If still not found, surface a helpful error
    searched = ", ".join(str(p) for p in candidates) + ", " + str(td / "*_comprehensive_data.json")
    raise FileNotFoundError(
        f"PDF JSON not found. Looked for: {searched}\nTemp dir: {temp_dir}"
    )

def process_files(pdf_file, word_file):
    # Create a unique temporary directory for this run
    temp_dir = tempfile.mkdtemp(prefix="hf_redtext_")
    print(f"📂 Temp dir: {temp_dir}")

    # Define standard filenames for use in the pipeline
    pdf_path = os.path.join(temp_dir, "input.pdf")
    word_path = os.path.join(temp_dir, "input.docx")
    word_json_path = os.path.join(temp_dir, "word_data.json")
    updated_json_path = os.path.join(temp_dir, "updated_word_data.json")
    final_docx_path = os.path.join(temp_dir, "updated.docx")

    # Copy the uploaded files to the temp directory
    shutil.copy(pdf_file, pdf_path)
    print(f"📄 PDF copied to: {pdf_path}")
    shutil.copy(word_file, word_path)
    print(f"📝 DOCX copied to: {word_path}")

    # 1) PDF → JSON  (extractor writes <stem>_comprehensive_data.json into cwd)
    run_cmd(["python", str(SCRIPT_DIR / "extract_pdf_data.py"), pdf_path], cwd=temp_dir)

    # Find the JSON produced by the extractor
    pdf_json_path = _locate_pdf_json(temp_dir)

    # 2) DOCX red text → JSON
    run_cmd(["python", str(SCRIPT_DIR / "extract_red_text.py"), word_path, word_json_path], cwd=temp_dir)

    # 3) Merge JSON (uses the resolved pdf_json_path)
    run_cmd(["python", str(SCRIPT_DIR / "update_docx_with_pdf.py"), word_json_path, pdf_json_path, updated_json_path], cwd=temp_dir)

    # 4) Apply updates to DOCX
    run_cmd(["python", str(SCRIPT_DIR / "updated_word.py"), word_path, updated_json_path, final_docx_path], cwd=temp_dir)

    # Return the final .docx file
    return final_docx_path

iface = gr.Interface(
    fn=process_files,
    inputs=[
        gr.File(label="Upload PDF File", type="filepath"),
        gr.File(label="Upload Word File", type="filepath")
    ],
    outputs=gr.File(label="Download Updated Word File"),
    title="Red Text Replacer",
    description="Upload a PDF and Word document. Red-colored text in the Word doc will be replaced by matching content from the PDF."
)

if __name__ == "__main__":
    iface.launch(share=True)
