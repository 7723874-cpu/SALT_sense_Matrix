import tkinter as tk
from tkinter import filedialog, messagebox
import json
import os
import re
from datetime import datetime
from math import log1p
from docx import Document
from typing import Optional

WINDOW_WIDTH = 1100
WINDOW_HEIGHT = 720

MAX_EVIDENCE_PER_TAG = 5
MIN_TOKEN_LEN = 3

SCORE_MULTIPLIER = 3.8
PRIORITY_CAP = 10


# ---------- text utils ----------
def normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = s.replace("ё", "е")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def tokenize(s: str):
    s = normalize_text(s)
    tokens = re.findall(r"[a-zа-я0-9]+", s, flags=re.IGNORECASE)
    return [t for t in tokens if len(t) >= MIN_TOKEN_LEN]


def extract_sentences(text: str):
    text = (text or "").replace("\n", " ")
    parts = re.split(r"(?<=[\.\!\?])\s+", text)
    return [p.strip() for p in parts if len(p.strip()) >= 25]


def extract_project_from_filename(filename: str) -> Optional[str]:
    basename = os.path.basename(filename)
    name_no_ext, _ = os.path.splitext(basename)
    tokens = name_no_ext.split("_")

    suffix_patterns = [
        re.compile(r"^(fg|idi)\d+$", flags=re.IGNORECASE),
        re.compile(r"^(part|p|v)\d+$", flags=re.IGNORECASE),
    ]

    while tokens:
        tail = tokens[-1]
        if any(p.match(tail) for p in suffix_patterns):
            tokens.pop()
        else:
            break

    if len(tokens) < 2:
        return None

    return "_".join(tokens)


def pick_evidence(sentences, key_tokens):
    if not sentences or not key_tokens:
        return []

    picked = []
    # сначала требуем 2 токена
    for s in sentences:
        s_norm = normalize_text(s)
        hits = 0
        for t in key_tokens:
            if t in s_norm:
                hits += 1
                if hits >= 2:
                    break
        if hits >= 2:
            picked.append(s.strip())
            if len(picked) >= MAX_EVIDENCE_PER_TAG:
                return picked

    # если мало — добираем по 1 токену
    if len(picked) < MAX_EVIDENCE_PER_TAG:
        for s in sentences:
            if s in picked:
                continue
            s_norm = normalize_text(s)
            if any(t in s_norm for t in key_tokens):
                picked.append(s.strip())
                if len(picked) >= MAX_EVIDENCE_PER_TAG:
                    break

    return picked[:MAX_EVIDENCE_PER_TAG]


def score_to_priority(raw_count: int) -> int:
    return int(round(min(PRIORITY_CAP, log1p(raw_count) * SCORE_MULTIPLIER)))


# ---------- generator (no LLM yet) ----------
# ВАЖНО: это прототип-генератор без LLM.
# Он формирует "2-й уровень" из фиксированного шаблона по типу исследования,
# а потом подтягивает цитаты/приоритеты из текста.
TEMPLATE_L2 = {
    "PRODUCT": [
        ("Контекст использования", "Ситуация (когда/где)"),
        ("Контекст использования", "Триггер (почему начали)"),
        ("Контекст использования", "Регулярность (как часто)"),

        ("Задача и ценность", "Основная задача (что хотят сделать)"),
        ("Задача и ценность", "Ожидание результата (что считать успехом)"),

        ("Функциональность", "Что работает"),
        ("Функциональность", "Чего не хватает"),
        ("Функциональность", "Понятность/простота"),

        ("Проблемы", "Боли и неудобства"),
        ("Проблемы", "Причины отказа/неиспользования"),

        ("Выбор", "Причины выбора"),
        ("Выбор", "Сравнение с альтернативами"),

        ("Улучшения", "Предложения/хотелки")
    ],
    "UX": [
        ("Навигация", "Поиск нужного (где это)"),
        ("Навигация", "Логика структуры"),
        ("Понятность", "Термины/подписи"),
        ("Фрустрации", "Где бесит/мешает"),
        ("Ошибки", "Сбои/неожиданное поведение"),
        ("Обучаемость", "Порог входа"),
        ("Ценность", "Зачем это вообще"),
        ("Улучшения", "Что исправить/добавить")
    ],
    "BRAND": [
        ("Восприятие", "Ассоциации"),
        ("Восприятие", "Эмоции"),
        ("Доверие", "Почему верю/не верю"),
        ("Дифференциация", "Чем отличается"),
        ("Опыт", "Личный опыт/истории"),
        ("Триггеры", "Почему выбрал/не выбрал"),
        ("Риски", "Сомнения/опасения")
    ],
    "SERVICE": [
        ("Процесс", "Этапы обслуживания"),
        ("Скорость", "Ожидание/время"),
        ("Коммуникация", "Контакт с персоналом/поддержкой"),
        ("Качество", "Что хорошо/плохо"),
        ("Проблемы", "Сбои/конфликты"),
        ("Улучшения", "Что улучшить")
    ],
    "COMMUNICATION": [
        ("Понимание", "Что понял/не понял"),
        ("Вера", "Доверие/скепсис"),
        ("Запоминаемость", "Что осталось в голове"),
        ("Релевантность", "Про меня/не про меня"),
        ("Отстройка", "Чем отличается"),
        ("Вызов действия", "Хочется/не хочется сделать шаг")
    ]
}

# Простые keywords для вытягивания доказательств. Можно расширять.
KEYWORDS_HINTS = {
    "Ситуация (когда/где)": ["когда", "где", "дома", "в офисе", "утром", "вечером", "в дороге"],
    "Триггер (почему начали)": ["потому", "решил", "нужно было", "стало", "захотел", "начал"],
    "Регулярность (как часто)": ["каждый", "часто", "редко", "раз", "постоянно", "иногда"],
    "Причины выбора": ["выбрал", "понрав", "удоб", "дешев", "лучше", "подошло"],
    "Сравнение с альтернативами": ["другой", "раньше", "вместо", "аналог", "альтернатив", "конкурент"],
    "Боли и неудобства": ["неудоб", "бесит", "проблем", "сложно", "тяжело", "не получается", "ошибка"],
    "Предложения/хотелки": ["хотел", "нужно бы", "было бы", "добав", "улучш", "сделайте"]
}


class TagCandidatesGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SALT Tag Candidates Generator (Level-2)")
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")

        self.transcripts = []
        self.transcript_sources = []

        # CORE опционально
        self.core_db = []
        self.core_db_path = None

        self.project_type = tk.StringVar(value="PRODUCT")
        self.project_name = tk.StringVar(value="")

        self.project_tagmap = []   # когда есть CORE или когда строим по шаблону
        self.new_candidates = []   # если CORE нет или если нашли вне CORE

        self.build_ui()

    # ---------- UI ----------
    def build_ui(self):
        main = tk.Frame(self.root, padx=12, pady=12)
        main.pack(fill="both", expand=True)

        top = tk.Frame(main)
        top.pack(fill="x")

        # CORE optional
        core_row = tk.Frame(top)
        core_row.pack(fill="x")
        tk.Button(
            core_row,
            text="(опц.) Загрузить CORE базу (JSON)",
            height=2,
            command=self.load_core_db
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        tk.Button(
            core_row,
            text="Очистить CORE",
            height=2,
            command=self.clear_core
        ).pack(side="left")

        row = tk.Frame(top, pady=8)
        row.pack(fill="x")

        tk.Label(row, text="Тип:", width=10, anchor="w").pack(side="left")
        tk.OptionMenu(
            row,
            self.project_type,
            "PRODUCT", "UX", "BRAND", "SERVICE", "COMMUNICATION"
        ).pack(side="left")

        tk.Label(row, text="Проект:", width=10, anchor="w").pack(side="left", padx=(18, 0))
        tk.Entry(row, textvariable=self.project_name).pack(side="left", fill="x", expand=True)

        tk.Button(
            top,
            text="1) Загрузить транскрипты (DOCX / TXT)",
            height=2,
            command=self.load_transcripts
        ).pack(fill="x", pady=(6, 0))

        tk.Button(
            top,
            text="2) Сгенерировать (без CORE тоже работает)",
            height=2,
            command=self.generate
        ).pack(fill="x", pady=(10, 0))

        save_row = tk.Frame(top, pady=10)
        save_row.pack(fill="x")

        tk.Button(
            save_row,
            text="💾 Сохранить project_tagmap.json",
            height=2,
            command=self.save_project_tagmap
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        tk.Button(
            save_row,
            text="💾 Сохранить new_candidates.json",
            height=2,
            command=self.save_new_candidates
        ).pack(side="left", fill="x", expand=True, padx=(6, 0))

        self.status = tk.Text(main)
        self.status.pack(fill="both", expand=True, pady=(12, 0))
        self._log("Готов. CORE база — опционально. Можно сразу загрузить транскрипты и генерировать.")

    # ---------- CORE ----------
    def load_core_db(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать JSON:\n{e}")
            return

        if not isinstance(db, list):
            messagebox.showerror("Ошибка", "CORE база должна быть массивом объектов (list).")
            return

        parsed = []
        for item in db:
            if not isinstance(item, dict):
                continue
            tag_path = item.get("tag_path") or item.get("path") or item.get("tag")
            if not tag_path or "/" not in str(tag_path):
                continue
            parsed.append({
                "tag_path": str(tag_path).strip(),
                "description": str(item.get("description", "")).strip(),
                "keywords": item.get("keywords") if isinstance(item.get("keywords"), list) else None,
                "type": str(item.get("type", self.project_type.get())).upper()
            })

        self.core_db = parsed
        self.core_db_path = path
        self._log(f"CORE загружен: {len(self.core_db)} тегов. Файл: {path}")

        if len(self.core_db) == 0:
            messagebox.showwarning("CORE = 0", "В CORE не найдено tag_path формата 'Блок/Тег'.")

    def clear_core(self):
        self.core_db = []
        self.core_db_path = None
        self._log("CORE очищен. Теперь генерация будет идти по шаблону (и сохранится в new_candidates).")

    # ---------- TRANSCRIPTS ----------
    def load_transcripts(self):
        paths = filedialog.askopenfilenames(
            filetypes=[
                ("Word documents", "*.docx"),
                ("Text files", "*.txt"),
            ]
        )
        if not paths:
            return

        self.transcripts = []
        self.transcript_sources = []
        detected_projects = []

        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            try:
                if ext == ".txt":
                    text = self._read_txt(path)
                elif ext == ".docx":
                    text = self._read_docx(path)
                else:
                    continue
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось прочитать {path}:\n{e}")
                continue

            if text.strip():
                self.transcripts.append(text)
                self.transcript_sources.append(os.path.basename(path))

            project_guess = extract_project_from_filename(os.path.basename(path))
            if project_guess:
                detected_projects.append(project_guess)
            else:
                self._log(f"Не удалось авто-определить проект из имени файла: {os.path.basename(path)} (нестандартный формат)")

        self._log(f"Загружено транскриптов: {len(self.transcripts)}")

        project_candidates = sorted(set(p for p in detected_projects if p))
        current_project = self.project_name.get().strip()

        if len(project_candidates) == 1 and not current_project:
            self.project_name.set(project_candidates[0])
            self._log(f"Проект определён из имени файла: {project_candidates[0]}")
        elif len(project_candidates) > 1:
            if not current_project:
                self.project_name.set(project_candidates[0])
                self._log(
                    "⚠️ Найдено несколько проектов в именах файлов: "
                    + ", ".join(project_candidates)
                    + f". Выбран по алфавиту: {project_candidates[0]}"
                )
            elif current_project not in project_candidates:
                self._log(
                    "⚠️ Поле 'Проект' не совпадает с найденными в файлах: "
                    + ", ".join(project_candidates)
                )
        elif not project_candidates:
            self._log("Авто-определение проекта из имен файлов не сработало.")

    def _read_txt(self, path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()

    def _read_docx(self, path):
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text and p.text.strip())

    # ---------- GENERATION ----------
    def generate(self):
        if not self.transcripts:
            messagebox.showwarning("Ошибка", "Сначала загрузите транскрипты.")
            return

        project_type = self.project_type.get().upper()
        project_name = self.project_name.get().strip() or "UnnamedProject"

        merged_text = "\n".join(self.transcripts)
        merged_norm = normalize_text(merged_text)
        sentences = extract_sentences(merged_text)

        self.project_tagmap = []
        self.new_candidates = []

        # База тегов, от которой пляшем:
        if self.core_db:
            base_paths = [x["tag_path"] for x in self.core_db if x.get("type", project_type) == project_type]
            base_set = set(base_paths)
            mode = "CORE"
            self._log(f"Режим: по CORE ({len(base_paths)} тегов).")
        else:
            base_paths = [f"{b}/{t}" for (b, t) in TEMPLATE_L2.get(project_type, [])]
            base_set = set(base_paths)
            mode = "TEMPLATE"
            self._log(f"Режим: по шаблону ({len(base_paths)} тегов).")

        # Генерация по базовым путям (CORE или шаблон)
        for tag_path in base_paths:
            # ключи: CORE keywords если есть, иначе hints, иначе токены из названия
            key_tokens = []

            if self.core_db:
                hit = next((x for x in self.core_db if x["tag_path"] == tag_path), None)
                if hit and hit.get("keywords"):
                    key_tokens = [normalize_text(k) for k in hit["keywords"] if isinstance(k, str) and k.strip()]

            if not key_tokens:
                sub = tag_path.split("/", 1)[-1].strip()
                hints = KEYWORDS_HINTS.get(sub, [])
                key_tokens = [normalize_text(x) for x in hints if x]

            if not key_tokens:
                key_tokens = tokenize(tag_path.replace("/", " "))

            key_tokens = [t for t in key_tokens if len(t) >= MIN_TOKEN_LEN]
            key_tokens = list(dict.fromkeys(key_tokens))

            raw_count = 0
            for t in key_tokens:
                raw_count += merged_norm.count(t)

            priority = score_to_priority(raw_count)
            evidence = pick_evidence(sentences, key_tokens)

            item = {
                "tag_path": tag_path,
                "type": project_type,
                "project": project_name,
                "priority": priority,
                "evidence_count": len(evidence),
                "evidence": evidence,
                "source": mode
            }

            # если CORE нет — это по сути новые кандидаты (потому что базы нет)
            if not self.core_db:
                self.new_candidates.append(item)
            else:
                self.project_tagmap.append(item)

        # Без CORE: ещё и сохраняем new_candidates (project_tagmap может быть пустым)
        if not self.core_db:
            self.new_candidates.sort(key=lambda x: x["priority"], reverse=True)
            self._log("Готово (без CORE). Топ-20 кандидатов:")
            for item in self.new_candidates[:20]:
                self._log(f"  {item['priority']:>2}  {item['tag_path']}  (evidence: {item['evidence_count']})")
            self._log("Сохраняй new_candidates.json и дальше уже загружай в core_tags_manager.")
            return

        # С CORE
        self.project_tagmap.sort(key=lambda x: x["priority"], reverse=True)
        self._log("Готово (CORE). Топ-20 по приоритету:")
        for item in self.project_tagmap[:20]:
            self._log(f"  {item['priority']:>2}  {item['tag_path']}  (evidence: {item['evidence_count']})")

    # ---------- SAVE ----------
    def save_project_tagmap(self):
        if not self.project_tagmap:
            messagebox.showwarning(
                "Нет данных",
                "project_tagmap пуст.\n"
                "Он появляется в режиме с CORE.\n"
                "Если CORE не загружен — сохраняй new_candidates.json."
            )
            return

        project_type = self.project_type.get().upper()
        project_name = (self.project_name.get().strip() or "UnnamedProject").replace(" ", "_")
        default_name = f"project_tagmap_{project_type}_{project_name}_{datetime.now().date()}.json"

        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile=default_name
        )
        if not path:
            return

        payload = {
            "meta": {
                "created_at": datetime.now().isoformat(),
                "type": project_type,
                "project": self.project_name.get().strip() or "UnnamedProject",
                "core_db_file": self.core_db_path,
                "transcripts_files": self.transcript_sources
            },
            "items": self.project_tagmap
        }

        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        self._log(f"Сохранено: {path}")

    def save_new_candidates(self):
        if not self.new_candidates:
            messagebox.showwarning("Нет данных", "new_candidates пуст. Сгенерируй сначала.")
            return

        project_type = self.project_type.get().upper()
        project_name = (self.project_name.get().strip() or "UnnamedProject").replace(" ", "_")
        default_name = f"new_candidates_{project_type}_{project_name}_{datetime.now().date()}.json"

        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile=default_name
        )
        if not path:
            return

        payload = {
            "meta": {
                "created_at": datetime.now().isoformat(),
                "type": project_type,
                "project": self.project_name.get().strip() or "UnnamedProject",
                "core_db_file": self.core_db_path,
                "transcripts_files": self.transcript_sources
            },
            "items": self.new_candidates
        }

        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        self._log(f"Сохранено: {path}")

    # ---------- LOG ----------
    def _log(self, text):
        self.status.insert("end", text + "\n")
        self.status.see("end")


if __name__ == "__main__":
    def _self_check_extract_project():
        cases = {
            "Yandex58_Cloud_IDI1.docx": "Yandex58_Cloud",
            "Bukhanka1_Local_Bakery_FG3.docx": "Bukhanka1_Local_Bakery",
            "MTS1_NextGenProduct_NNovgorod_FG4.docx": "MTS1_NextGenProduct_NNovgorod",
            "WeirdName.docx": None,
        }
        for name, expected in cases.items():
            result = extract_project_from_filename(name)
            print(f"[self-check] {name} -> {result} (expected: {expected})")


    _self_check_extract_project()

    root = tk.Tk()
    app = TagCandidatesGeneratorGUI(root)
    root.mainloop()
