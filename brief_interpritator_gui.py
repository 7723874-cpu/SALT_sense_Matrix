# brief_interpritator_gui.py
# SAFE brief -> clean Word (DOCX)
# Input:  .docx / .xlsx / .txt (anonymized SAFE brief)
# Output: .docx (clean structured brief)
#
# Dependencies:
#   pip install python-docx openpyxl openai
#
# Works with different OpenAI SDK generations:
# - New: from openai import OpenAI -> client.responses.create(...)
# - Legacy: import openai -> openai.ChatCompletion.create(...)

import os
import re
import json
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from docx import Document
from openpyxl import load_workbook

try:
    from openai import OpenAI  # new SDK
except Exception:
    OpenAI = None

try:
    import openai as openai_legacy  # legacy SDK
except Exception:
    openai_legacy = None


# -----------------------------
# Helpers: read input
# -----------------------------
def _norm_text(s: str) -> str:
    s = (s or "").replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def read_docx(path: str) -> str:
    doc = Document(path)
    parts = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t:
            parts.append(t)

    for table in doc.tables:
        for row in table.rows:
            row_vals = []
            for cell in row.cells:
                c = (cell.text or "").strip()
                if c:
                    row_vals.append(c)
            if row_vals:
                parts.append(" | ".join(row_vals))

    return _norm_text("\n".join(parts))


def read_txt(path: str) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1251"):
        try:
            with open(path, "r", encoding=enc) as f:
                return _norm_text(f.read())
        except UnicodeDecodeError:
            continue
    with open(path, "rb") as f:
        return _norm_text(f.read().decode("utf-8", errors="replace"))


def read_xlsx_as_text(path: str) -> str:
    wb = load_workbook(path, data_only=True)
    out = []
    for sh in wb.worksheets:
        out.append(f"[SHEET] {sh.title}")
        max_r = sh.max_row or 0
        max_c = sh.max_column or 0
        for r in range(1, max_r + 1):
            row_vals = []
            empty = True
            for c in range(1, max_c + 1):
                v = sh.cell(r, c).value
                if v is None:
                    row_vals.append("")
                else:
                    empty = False
                    s = str(v).strip()
                    if len(s) > 600:
                        s = s[:600] + "…"
                    row_vals.append(s)
            if not empty:
                while row_vals and row_vals[-1] == "":
                    row_vals.pop()
                if row_vals:
                    out.append(" | ".join(row_vals))
    return _norm_text("\n".join(out))


def try_parse_legitimizer_excel(path: str) -> dict | None:
    try:
        wb = load_workbook(path, data_only=True)
        if "Brief" not in wb.sheetnames:
            return None

        sh = wb["Brief"]
        data = {}
        for r in range(2, (sh.max_row or 0) + 1):
            k = sh.cell(r, 1).value
            v = sh.cell(r, 2).value
            if k is None:
                continue
            key = str(k).strip()
            if not key:
                continue
            val = "" if v is None else str(v).strip()
            data[key] = val

        safe_text = ""
        if "SAFE_TEXT" in wb.sheetnames:
            st = wb["SAFE_TEXT"]
            v = st["A2"].value
            if v:
                safe_text = str(v)

        return {
            "source_type": "legitimizer_excel",
            "brief_table": data,
            "safe_text": _norm_text(safe_text),
        }
    except Exception:
        return None


def build_safe_brief_payload(source_path: str) -> dict:
    ext = os.path.splitext(source_path)[1].lower()

    if ext == ".docx":
        text = read_docx(source_path)
        return {"source_file": os.path.basename(source_path), "safe_raw_text": text, "safe_structured": None}

    if ext == ".txt":
        text = read_txt(source_path)
        return {"source_file": os.path.basename(source_path), "safe_raw_text": text, "safe_structured": None}

    if ext == ".xlsx":
        structured = try_parse_legitimizer_excel(source_path)
        if structured:
            raw = structured["safe_text"] or read_xlsx_as_text(source_path)
            return {
                "source_file": os.path.basename(source_path),
                "safe_raw_text": raw,
                "safe_structured": structured["brief_table"],
            }
        text = read_xlsx_as_text(source_path)
        return {"source_file": os.path.basename(source_path), "safe_raw_text": text, "safe_structured": None}

    raise ValueError("Supported inputs: .docx, .xlsx, .txt")


# -----------------------------
# Prompting + JSON extraction
# -----------------------------
def make_target_template() -> dict:
    return {
        "title": "",
        "executive_summary": "",
        "context": "",
        "goal": "",
        "tasks": [],
        "research_questions": [],
        "hypotheses": [],
        "audience": "",
        "segments": [],
        "methodology": "",
        "scope": "",
        "constraints": [],
        "deliverables": [],
        "timeline": "",
        "success_criteria": "",
        "risks_open_questions": [],
        "appendix_notes": "",
    }


def _ensure_keys(obj: dict) -> dict:
    tpl = make_target_template()
    out = {k: obj.get(k, tpl[k]) for k in tpl.keys()}

    list_keys = ("tasks", "research_questions", "hypotheses", "segments", "constraints", "deliverables", "risks_open_questions")
    for lk in list_keys:
        v = out.get(lk, [])
        if v is None:
            v = []
        if isinstance(v, str):
            v = [v]
        if not isinstance(v, list):
            v = list(v)
        out[lk] = [str(x).strip() for x in v if str(x).strip()]

    str_keys = ("title", "executive_summary", "context", "goal", "audience", "methodology", "scope", "timeline", "success_criteria", "appendix_notes")
    for sk in str_keys:
        v = out.get(sk, "")
        out[sk] = str(v).strip() if v is not None else ""

    # enforce "не указано" where appropriate
    for sk in ("executive_summary", "context", "goal", "audience", "methodology", "scope", "timeline", "success_criteria"):
        if not out[sk]:
            out[sk] = "не указано"

    if not out["title"]:
        out["title"] = "Brief (SAFE)"

    return out


def extract_json_strict(text: str) -> dict:
    text = (text or "").strip()
    if not text:
        raise ValueError("Empty model output")

    try:
        return json.loads(text)
    except Exception:
        pass

    m = re.search(r"\{.*\}", text, flags=re.DOTALL)
    if not m:
        raise ValueError("Could not find JSON object in model output")
    return json.loads(m.group(0))


def build_instruction(tone: str) -> str:
    tpl = json.dumps(make_target_template(), ensure_ascii=False, indent=2)
    return (
        "Ты — старший методолог и редактор. Получаешь обезличенный SAFE-бриф.\n"
        "Задача: сделать чёткий, понятный, красиво структурированный документ.\n"
        "Строгие правила:\n"
        "1) НИЧЕГО не выдумывай. Если данных нет — пиши 'не указано' или оставляй список пустым.\n"
        "2) Сохраняй максимум смысла и деталей, которые есть в исходнике.\n"
        "3) Перефразируй, не копируй длинные куски дословно.\n"
        "4) Тон: " + (tone or "деловой, чёткий, без лишних слов") + "\n"
        "Верни РОВНО один JSON-объект без пояснений, без обёрток, без markdown.\n"
        "JSON должен иметь такую структуру (поля и типы должны совпадать):\n"
        f"{tpl}\n"
    )


# -----------------------------
# OpenAI call (multi-SDK)
# -----------------------------
def call_llm_to_json(api_key: str, model: str, payload: dict, tone: str) -> dict:
    instruction = build_instruction(tone)

    user_obj = {
        "source_file": payload["source_file"],
        "safe_raw_text": payload["safe_raw_text"],
        "safe_structured_table": payload["safe_structured"] or {},
    }
    user_text = json.dumps(user_obj, ensure_ascii=False)

    # New SDK
    if OpenAI is not None:
        client = OpenAI(api_key=api_key)

        # IMPORTANT FIX: use "input_text" instead of "text"
        try:
            resp = client.responses.create(
                model=model,
                input=[
                    {"role": "system", "content": instruction},
                    {"role": "user", "content": [{"type": "input_text", "text": user_text}]},
                ],
                response_format={"type": "json_object"},
            )
            out_text = getattr(resp, "output_text", None)
            if not out_text:
                out_text = json.dumps(resp.model_dump(), ensure_ascii=False)
            return _ensure_keys(extract_json_strict(out_text))

        except TypeError:
            # response_format not supported
            resp = client.responses.create(
                model=model,
                input=[
                    {"role": "system", "content": instruction},
                    {"role": "user", "content": [{"type": "input_text", "text": user_text}]},
                ],
            )
            out_text = getattr(resp, "output_text", None)
            if not out_text:
                out_text = json.dumps(resp.model_dump(), ensure_ascii=False)
            return _ensure_keys(extract_json_strict(out_text))

        except Exception:
            # fall through to legacy if possible
            pass

    # Legacy SDK
    if openai_legacy is None:
        raise RuntimeError("OpenAI SDK not available. Install: pip install openai")

    openai_legacy.api_key = api_key

    try:
        resp = openai_legacy.ChatCompletion.create(
            model=model,
            messages=[
                {"role": "system", "content": instruction},
                {"role": "user", "content": user_text},
            ],
            temperature=0.2,
        )
        text = resp["choices"][0]["message"]["content"]
        return _ensure_keys(extract_json_strict(text))
    except AttributeError:
        raise RuntimeError("Your openai package is too old/odd. Upgrade: pip install -U openai")


# -----------------------------
# Word rendering
# -----------------------------
def add_heading(doc: Document, text: str, level: int = 1):
    doc.add_heading(text, level=level)


def add_paragraph(doc: Document, text: str):
    doc.add_paragraph(text)


def add_bullets(doc: Document, items: list[str]):
    for it in items:
        t = (it or "").strip()
        if not t:
            continue
        doc.add_paragraph(t, style="List Bullet")


def render_docx(out_path: str, doc_json: dict, meta: dict):
    doc = Document()

    title = (doc_json.get("title") or "").strip() or "Brief (SAFE)"
    add_heading(doc, title, level=0)

    meta_line = f"Source: {meta.get('source_file','')} | Generated: {meta.get('generated_at','')} | Model: {meta.get('model','')}"
    p = doc.add_paragraph(meta_line)
    if p.runs:
        p.runs[0].italic = True

    def sec(name: str, value: str):
        add_heading(doc, name, level=1)
        add_paragraph(doc, (value or "не указано").strip() or "не указано")

    def seclist(name: str, items: list[str]):
        add_heading(doc, name, level=1)
        if not items:
            add_paragraph(doc, "не указано")
        else:
            add_bullets(doc, items)

    sec("Executive summary", doc_json.get("executive_summary", "не указано"))
    sec("Context", doc_json.get("context", "не указано"))
    sec("Goal", doc_json.get("goal", "не указано"))
    seclist("Tasks", doc_json.get("tasks", []))
    seclist("Research questions", doc_json.get("research_questions", []))
    seclist("Hypotheses", doc_json.get("hypotheses", []))
    sec("Audience", doc_json.get("audience", "не указано"))
    seclist("Segments", doc_json.get("segments", []))
    sec("Methodology", doc_json.get("methodology", "не указано"))
    sec("Scope", doc_json.get("scope", "не указано"))
    seclist("Constraints", doc_json.get("constraints", []))
    seclist("Deliverables", doc_json.get("deliverables", []))
    sec("Timeline", doc_json.get("timeline", "не указано"))
    sec("Success criteria", doc_json.get("success_criteria", "не указано"))
    seclist("Risks / open questions", doc_json.get("risks_open_questions", []))
    sec("Appendix notes", doc_json.get("appendix_notes", "не указано"))

    doc.save(out_path)


# -----------------------------
# GUI
# -----------------------------
class BriefInterpretatorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("brief_interpritator — SAFE brief → clean Word")
        self.geometry("1180x780")

        self.input_path = tk.StringVar(value="")
        self.api_key = tk.StringVar(value="")
        self.model = tk.StringVar(value="gpt-5.2")
        self.tone = tk.StringVar(value="деловой, чёткий, без лишних слов")

        self.payload = None
        self.generated_json = None

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}

        box = ttk.LabelFrame(self, text="Input & Model")
        box.pack(fill="x", **pad)

        ttk.Label(box, text="SAFE brief file (.docx / .xlsx / .txt):").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(box, textvariable=self.input_path, width=92, state="readonly").grid(row=0, column=1, sticky="we", padx=8, pady=6)
        ttk.Button(box, text="Choose…", command=self.pick_file).grid(row=0, column=2, padx=8, pady=6)

        ttk.Label(box, text="OpenAI API key:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(box, textvariable=self.api_key, width=70, show="•").grid(row=1, column=1, sticky="w", padx=8, pady=6)

        ttk.Label(box, text="Model:").grid(row=1, column=2, sticky="e", padx=8, pady=6)
        ttk.Combobox(
            box,
            textvariable=self.model,
            state="readonly",
            width=18,
            values=["gpt-5.2", "gpt-5-mini", "gpt-4.1", "gpt-4o", "gpt-4o-mini"],
        ).grid(row=1, column=3, sticky="w", padx=8, pady=6)

        ttk.Label(box, text="Tone (RU):").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(box, textvariable=self.tone, width=92).grid(row=2, column=1, columnspan=3, sticky="we", padx=8, pady=6)

        box.columnconfigure(1, weight=1)

        act = ttk.Frame(self)
        act.pack(fill="x", **pad)

        ttk.Button(act, text="Preview input", command=self.preview_input).pack(side="left", padx=8)
        ttk.Button(act, text="Generate clean DOC (JSON + Word)", command=self.generate).pack(side="left", padx=8)
        ttk.Button(act, text="Save Word…", command=self.save_word).pack(side="left", padx=8)

        self.status = ttk.Label(act, text="Ready.")
        self.status.pack(side="left", padx=12)

        out = ttk.Notebook(self)
        out.pack(fill="both", expand=True, **pad)

        self.txt_in = tk.Text(out, wrap="word")
        self.txt_json = tk.Text(out, wrap="none")
        self.txt_log = tk.Text(out, wrap="word")

        out.add(self.txt_in, text="Input preview (SAFE)")
        out.add(self.txt_json, text="Generated JSON")
        out.add(self.txt_log, text="Log")

        self._set_text(self.txt_in, "Choose a SAFE brief file to preview.")
        self._set_text(self.txt_json, "")
        self._set_text(self.txt_log, "")

    def _set_text(self, widget: tk.Text, text: str):
        widget.config(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.config(state="normal")

    def log(self, msg: str):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.config(state="normal")

    def pick_file(self):
        path = filedialog.askopenfilename(
            title="Choose SAFE brief",
            filetypes=[
                ("Word document", "*.docx"),
                ("Excel workbook", "*.xlsx"),
                ("Text file", "*.txt"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        self.input_path.set(path)
        self.status.config(text="File selected.")
        self.payload = None
        self.generated_json = None
        self._set_text(self.txt_json, "")
        self._set_text(self.txt_log, "")
        self.log(f"[OK] Selected: {path}")

    def preview_input(self):
        if not self.input_path.get():
            messagebox.showinfo("No file", "Choose a SAFE brief file first.")
            return
        try:
            self.payload = build_safe_brief_payload(self.input_path.get())
            preview = self.payload["safe_raw_text"]
            if len(preview) > 20000:
                preview = preview[:20000] + "\n\n…(truncated preview)…"
            self._set_text(self.txt_in, preview)
            self.log(f"[OK] Loaded input. chars={len(self.payload['safe_raw_text'])}")
            if self.payload["safe_structured"]:
                self.log("[OK] Detected structured 'Brief' table from legitimizer Excel.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log(f"[ERROR] preview_input: {e}")

    def generate(self):
        if not self.input_path.get():
            messagebox.showinfo("No file", "Choose a SAFE brief file first.")
            return
        if not self.api_key.get().strip():
            messagebox.showinfo("No API key", "Paste your OpenAI API key.")
            return

        if self.payload is None:
            try:
                self.payload = build_safe_brief_payload(self.input_path.get())
            except Exception as e:
                messagebox.showerror("Error", str(e))
                return

        self.status.config(text="Generating…")
        self.update_idletasks()
        self._set_text(self.txt_log, "")

        try:
            doc_json = call_llm_to_json(
                api_key=self.api_key.get().strip(),
                model=self.model.get().strip(),
                payload=self.payload,
                tone=self.tone.get().strip() or "деловой, чёткий, без лишних слов",
            )
            self.generated_json = doc_json
            self._set_text(self.txt_json, json.dumps(doc_json, ensure_ascii=False, indent=2))
            self.log("[OK] Generated JSON.")
            self.status.config(text="Generated. You can save Word.")
        except Exception as e:
            messagebox.showerror("Generation error", str(e))
            self.log(f"[ERROR] generate: {e}")
            self.status.config(text="Error.")

    def save_word(self):
        if not self.generated_json:
            messagebox.showinfo("No output", "Generate first.")
            return

        base = os.path.splitext(os.path.basename(self.input_path.get() or "brief"))[0]
        base = re.sub(r"\s+", " ", base).strip()
        default_name = f"{base}__clean.docx" if base else "brief__clean.docx"

        out_path = filedialog.asksaveasfilename(
            title="Save clean Word document",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word document", "*.docx")],
        )
        if not out_path:
            return

        try:
            meta = {
                "source_file": os.path.basename(self.input_path.get()),
                "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
                "model": self.model.get(),
            }
            render_docx(out_path, self.generated_json, meta)
            self.log(f"[OK] Saved Word: {out_path}")
            self.status.config(text="Saved.")
            messagebox.showinfo("Saved", out_path)
        except Exception as e:
            messagebox.showerror("Save error", str(e))
            self.log(f"[ERROR] save_word: {e}")


def main():
    app = BriefInterpretatorGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
