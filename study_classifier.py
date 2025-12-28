"""
SALT_study_classifier.py — классификатор проектов по типу исследования.

Идея:
- Подгружаем транскрипты одного проекта (несколько .txt/.docx).
- Программа собирает текстовый срез.
- Отправляет в LLM с описанием макрогрупп исследований.
- На выходе:
    group_code  (BRAND / PRODUCT / COMM / CULTURE / DESIGN / EXPERIENCE)
    group_name  (по-русски)
    confidence  (0..1)
    justification (текстовое обоснование)
    key_signals (список сигналов из транскриптов)
- Результат отображается в GUI.

Зависимости:
    pip install openai python-docx pandas openpyxl
Нужен ключ в переменной окружения OPENAI_API_KEY.
"""

import os
import json
import pathlib
import re
import shutil
from typing import List, Dict, Any

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import pandas as pd
from docx import Document
from openai import OpenAI


# =========================
#  НАСТРОЙКИ МОДЕЛИ
# =========================

BASE_MODEL = "gpt-4o-mini"   # дешёвый режим
PREMIUM_MODEL = "gpt-4o"     # премиум-режим
PROJECT_TRANSCRYPTS_ROOT = r"D:\YandexDisk\Files\Project Transcrypts"
SETTINGS_PATH = pathlib.Path.home() / ".study_classifier_settings.json"


def get_openai_client() -> OpenAI:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "Не найден ключ OpenAI. "
            "Задай переменную окружения OPENAI_API_KEY."
        )
    return OpenAI(api_key=api_key)


def _normalize_taxonomy_cell(value: Any, *, upper: bool = True) -> str:
    if pd.isna(value):
        return ""
    text = re.sub(r"\s+", " ", str(value).strip())
    return text.upper() if upper else text


def load_settings() -> Dict[str, Any]:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        with SETTINGS_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(settings: Dict[str, Any]) -> None:
    SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    with SETTINGS_PATH.open("w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


# =========================
#  ЗАГРУЗКА СТРУКТУРЫ ЦЕЛЕЙ
# =========================

def load_goal_structure(xlsx_path: str) -> List[Dict[str, str]]:
    """
    Загружает структуру целей из Excel.
    Обязательные колонки: L1, L2, L3 (без учёта регистра).
    Возвращает список словарей:
    {
        "l1": "...",
        "l2": "...",
        "l3": "...",
        "path": "L1 / L2 / L3"
    }
    """
    if not xlsx_path:
        raise ValueError("Не задан путь к Excel-файлу.")

    path = pathlib.Path(xlsx_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {xlsx_path}")

    df = pd.read_excel(path)
    if df.empty:
        return []

    normalized_columns = {col.lower().strip(): col for col in df.columns}
    required = {"l1", "l2", "l3"}
    missing = required - set(normalized_columns.keys())
    if missing:
        raise ValueError(
            "В Excel нет обязательных колонок: "
            + ", ".join(sorted(missing))
        )

    structure: List[Dict[str, str]] = []
    seen_paths = set()
    for _, row in df.iterrows():
        l1 = _normalize_taxonomy_cell(row[normalized_columns["l1"]], upper=True)
        l2 = _normalize_taxonomy_cell(row[normalized_columns["l2"]], upper=True)
        l3 = _normalize_taxonomy_cell(row[normalized_columns["l3"]], upper=True)

        if not l1 or not l2 or not l3:
            continue

        path_str = f"{l1} / {l2} / {l3}"
        if path_str in seen_paths:
            continue

        seen_paths.add(path_str)
        structure.append({
            "l1": l1,
            "l2": l2,
            "l3": l3,
            "path": path_str,
        })

    return structure


# =========================
#  ЧТЕНИЕ ТРАНСКРИПТОВ
# =========================

def read_docx(path: pathlib.Path) -> str:
    doc = Document(str(path))
    parts: List[str] = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            parts.append(txt)
    return "\n".join(parts)


def read_txt(path: pathlib.Path) -> str:
    with path.open("r", encoding="utf-8", errors="replace") as f:
        return f.read()


def read_any_transcript(path_str: str) -> str:
    path = pathlib.Path(path_str)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")

    suffix = path.suffix.lower()
    if suffix == ".docx":
        return read_docx(path)
    elif suffix == ".txt":
        return read_txt(path)
    else:
        raise ValueError(f"Пока поддерживаются только .docx и .txt (не {suffix})")


def build_text_sample(transcript_paths: List[str], max_chars: int = 20000) -> str:
    """
    Собираем текстовый срез по всем транскриптам проекта.
    Берём по очереди, пока не доберём примерно max_chars символов.
    """
    chunks: List[str] = []
    total = 0

    for p in transcript_paths:
        try:
            txt = read_any_transcript(p)
        except Exception as e:
            # Пропускаем файл, но не валим всё
            chunks.append(f"\n\n[ОШИБКА ЧТЕНИЯ ФАЙЛА {p}: {e}]\n\n")
            continue

        if not txt.strip():
            continue

        remaining = max_chars - total
        if remaining <= 0:
            break

        if len(txt) <= remaining:
            chunks.append(f"\n\n=== ТРАНСКРИПТ {p} ===\n\n" + txt)
            total += len(txt)
        else:
            # Возьмём кусочек: начало + середина
            head = txt[: remaining // 2]
            middle_start = max(len(txt) // 2 - remaining // 4, 0)
            middle_end = middle_start + remaining // 4
            middle = txt[middle_start:middle_end]
            sample = head + "\n\n[...] \n\n" + middle
            chunks.append(f"\n\n=== ТРАНСКРИПТ {p} (фрагмент) ===\n\n" + sample)
            total += len(sample)

    return "\n".join(chunks).strip()


def build_text_sample_single(text: str, max_chars: int = 12000) -> str:
    """
    Делает компактный срез одного транскрипта.
    """
    text = text.strip()
    if len(text) <= max_chars:
        return text
    head = text[: max_chars // 2]
    middle_start = max(len(text) // 2 - max_chars // 4, 0)
    middle_end = middle_start + max_chars // 4
    middle = text[middle_start:middle_end]
    return head + "\n\n[...] \n\n" + middle


# =========================
#  ВЫЗОВ МОДЕЛИ ДЛЯ КЛАССИФИКАЦИИ
# =========================

GROUP_DEFS = """
Ты должен отнести проект к ОДНОЙ основной группе исследования:

1) BRAND — Брендовое исследование.
   Признаки:
   - много разговоров про бренды, марки, сети, торговые точки;
   - отношение к брендам (нравится / не нравится / доверие / раздражение);
   - драйверы и барьеры выбора бренда;
   - цена и ценность бренда;
   - восприятие коммуникаций бренда;
   - иногда роль селебрити или амбассадоров.

2) PRODUCT — Продукт / концепт / U&A.
   Признаки:
   - фокус на конкретных продуктах, форматах, сервисах;
   - опыт использования, удобство/неудобство;
   - неудовлетворённые потребности, pain points;
   - сценарии и контексты использования;
   - сравнение с альтернативами и конкурентами.

3) COMM — Коммуникации / реклама / креатив.
   Признаки:
   - обсуждаются ролики, макеты, креативы, сообщения;
   - понятность / непонятность, запоминаемость;
   - эмоциональный отклик на коммуникацию;
   - ключевые инсайты и смысловые сообщения.

4) CULTURE — Культурные / социологические / глубинные инсайты.
   Признаки:
   - много разговоров про ценности, роли, идентичности;
   - жизненные сценарии, страхи, мечты, нарративы;
   - культурные коды и социальный контекст;
   - менее про конкретный бренд или продукт, больше про людей и мир.

5) DESIGN — Нейминг / упаковка / дизайн.
   Признаки:
   - обсуждают внешний вид, упаковку, шрифты, цвет, форму;
   - ассоциации от дизайна, заметность на полке;
   - читаемость информации, воспринимаемость элементов;
   - названия бренда/линейки, варианты нейминга.

6) EXPERIENCE — Опыт людей / сервисов / HR / journey.
   Признаки:
   - фокус на опыте клиента или сотрудника;
   - путь (journey), точки контакта, шаги процесса;
   - барьеры и тормоза в сервисе;
   - внутренний опыт работы в компании, мотивация и демотивация сотрудников.
"""


def classify_study(
    project_name: str,
    text_sample: str,
    use_premium: bool,
) -> Dict[str, Any]:
    """
    Отправляет срез транскриптов в модель и получает JSON с классификацией.
    Формат ожидаемого ответа:
    {
      "group_code": "BRAND",
      "group_name": "Брендовое исследование",
      "confidence": 0.87,
      "justification": "…",
      "key_signals": ["…", "…"]
    }
    """
    if not text_sample.strip():
        raise ValueError("Пустой текст транскриптов — нечего классифицировать.")

    system_prompt = (
        "Ты — старший методолог и директор по исследованиям в компании SALT.\n"
        "Тебе дали фрагменты транскриптов одного проекта (фокус-группы/глубинки и т.п.).\n"
        "Твоя задача — отнести проект к одной из макрогрупп исследований SALT.\n\n"
        "Группы:\n"
        f"{GROUP_DEFS}\n\n"
        "Важно:\n"
        "- Выбери РОВНО ОДНУ основную группу (group_code).\n"
        "- Если проект на стыке, выбери ту группу, которая лучше описывает ЦЕЛЬ ИССЛЕДОВАНИЯ.\n"
        "- Оцени уверенность (confidence) от 0 до 1.\n"
        "- Обоснуй выбор понятным текстом (justification).\n"
        "- Выдели 3–7 ключевых сигналов из транскриптов, которые подсказали тебе решение (key_signals).\n\n"
        "Формат ответа — строго JSON без комментариев."
    )

    user_prompt = (
        f"Название проекта: {project_name}\n\n"
        "Ниже идут фрагменты транскриптов этого проекта.\n"
        "По ним нужно определить тип исследования по одной из групп.\n\n"
        "=== ТРАНСКРИПТЫ ПРОЕКТА ===\n\n"
        f"{text_sample}"
    )

    client = get_openai_client()
    model_name = PREMIUM_MODEL if use_premium else BASE_MODEL

    resp = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.2,
    )

    content = resp.choices[0].message.content.strip()

    # На случай, если модель завернёт в ```json
    if content.startswith("```"):
        content = content.strip("`")
        if content.lstrip().lower().startswith("json"):
            content = content.split("\n", 1)[1]

    data = json.loads(content)

    if not isinstance(data, dict):
        raise ValueError("Ожидался JSON-объект с полями group_code, group_name и т.п.")

    # лёгкая нормализация
    result = {
        "group_code": str(data.get("group_code", "")).strip(),
        "group_name": str(data.get("group_name", "")).strip(),
        "confidence": float(data.get("confidence", 0.0)),
        "justification": str(data.get("justification", "")).strip(),
        "key_signals": data.get("key_signals", []),
    }

    # key_signals приводим к списку строк
    if not isinstance(result["key_signals"], list):
        result["key_signals"] = [str(result["key_signals"])]

    result["key_signals"] = [str(x).strip() for x in result["key_signals"] if str(x).strip()]

    return result


def _strip_json_fence(content: str) -> str:
    if content.startswith("```"):
        content = content.strip("`")
        if content.lstrip().lower().startswith("json"):
            content = content.split("\n", 1)[1]
    return content.strip()


def _safe_list(value: Any) -> List[str]:
    if isinstance(value, list):
        items = value
    elif value is None:
        items = []
    else:
        items = [value]
    return [str(item).strip() for item in items if str(item).strip()]


def _confidence_label(confidence: float) -> str:
    if confidence >= 0.75:
        return "High"
    if confidence >= 0.5:
        return "Med"
    return "Low"


def _parse_confidence(value: Any) -> float:
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"low", "med", "medium", "mid", "high"}:
            mapping = {
                "low": 0.3,
                "med": 0.6,
                "medium": 0.6,
                "mid": 0.6,
                "high": 0.85,
            }
            return mapping[lowered]
        try:
            value = float(lowered)
        except ValueError:
            return 0.0
    try:
        conf = float(value)
    except (TypeError, ValueError):
        return 0.0
    if conf > 1.0:
        conf = conf / 100 if conf <= 100 else 1.0
    return max(0.0, min(1.0, conf))


def _parse_score(value: Any) -> float:
    try:
        score = float(value)
    except (TypeError, ValueError):
        score = 0.0
    if score <= 1.0:
        score *= 100
    return max(0.0, min(100.0, score))


def classify_study_taxonomy(
    text: str,
    structure_paths: List[str],
    model: str,
    top_n: int = 3,
) -> Dict[str, Any]:
    """
    Мульти-лейбл классификация по структуре целей.
    """
    if not text.strip():
        raise ValueError("Пустой текст транскриптов — нечего классифицировать.")
    if not structure_paths:
        raise ValueError("Список веток структуры пуст.")

    structure_listing = "\n".join(f"- {path}" for path in structure_paths)
    system_prompt = (
        "Ты — старший методолог и директор по исследованиям в компании SALT.\n"
        "Тебе дали фрагменты транскриптов одного проекта.\n"
        "Нужно выбрать несколько веток структуры целей исследования.\n\n"
        "Ветки (можно выбирать ТОЛЬКО из этого списка):\n"
        f"{structure_listing}\n\n"
        "Важно:\n"
        "- Различай \"о чём говорят\" и \"зачем исследуют\" — вывод делай по ЦЕЛИ.\n"
        "- Если обсуждают термины/формулировки/как назвать/как объяснить/как люди понимают, "
        "усиливай HUMAN / COMMUNICATION / *.\n"
        "- Если обсуждают функционал/механику/использование — усиливай SOLUTION / PRODUCT / SERVICE / * "
        "или HUMAN / EXPERIENCE / * (если про опыт людей).\n"
        "- Если обсуждают оценку/проверку/риски/соответствие — усиливай SOLUTION / AUDIT / *.\n"
        "- Никаких новых веток, только из списка.\n"
        "- Верни TOP-N веток с оценками и доказательствами.\n"
        "- Evidence: 2–5 коротких цитат/фрагментов без воды.\n"
        "- score в диапазоне 0..100.\n"
        "- confidence: 0..1 или Low/Med/High.\n"
        "- Ответ строго JSON по схеме ниже, без комментариев.\n\n"
        "JSON-схема:\n"
        "{\n"
        '  "branches": [\n'
        "    {\n"
        '      "path": "L1 / L2 / L3",\n'
        '      "score": 0,\n'
        '      "confidence": "Low/Med/High или число",\n'
        '      "rationale": "...",\n'
        '      "evidence": ["...", "..."]\n'
        "    }\n"
        "  ],\n"
        '  "split_recommended": true/false,\n'
        '  "split_reason": "почему стоит разделить/почему нет"\n'
        "}"
    )

    user_prompt = (
        "Ниже идут фрагменты транскриптов проекта.\n"
        "Определи цели исследования и выбери ветки.\n\n"
        "=== ТРАНСКРИПТЫ ПРОЕКТА ===\n\n"
        f"{text}"
    )

    client = get_openai_client()
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.2,
    )

    content = _strip_json_fence(resp.choices[0].message.content or "")
    data = json.loads(content)
    if not isinstance(data, dict):
        raise ValueError("Ожидался JSON-объект с полями branches и split_recommended.")

    branches_raw = data.get("branches", [])
    if not isinstance(branches_raw, list):
        branches_raw = []

    branches: List[Dict[str, Any]] = []
    allowed = set(structure_paths)
    for item in branches_raw:
        if not isinstance(item, dict):
            continue
        path = str(item.get("path", "")).strip()
        if path not in allowed:
            continue
        score = _parse_score(item.get("score"))
        confidence = _parse_confidence(item.get("confidence"))
        branches.append({
            "path": path,
            "score": score,
            "confidence": confidence,
            "confidence_label": _confidence_label(confidence),
            "rationale": str(item.get("rationale", "")).strip(),
            "evidence": _safe_list(item.get("evidence")),
        })

    branches.sort(key=lambda x: x["score"], reverse=True)
    branches = branches[:max(top_n, 1)]

    return {
        "branches": branches,
        "split_recommended": bool(data.get("split_recommended", False)),
        "split_reason": str(data.get("split_reason", "")).strip(),
    }


def aggregate_taxonomy_results(
    per_file_results: List[Dict[str, Any]],
    top_n: int = 3,
) -> Dict[str, Any]:
    aggregates: Dict[str, Dict[str, Any]] = {}
    total_files = len(per_file_results)
    split_votes = 0
    split_reasons: List[str] = []

    for item in per_file_results:
        if item.get("split_recommended"):
            split_votes += 1
            reason = str(item.get("split_reason", "")).strip()
            if reason:
                split_reasons.append(reason)
        for branch in item.get("branches", []):
            path = branch["path"]
            agg = aggregates.setdefault(path, {
                "scores": [],
                "confidences": [],
                "evidence": [],
                "rationales": [],
                "files": set(),
            })
            agg["scores"].append(branch.get("score", 0.0))
            agg["confidences"].append(branch.get("confidence", 0.0))
            agg["evidence"].extend(branch.get("evidence", []))
            rationale = str(branch.get("rationale", "")).strip()
            if rationale:
                agg["rationales"].append(rationale)
            agg["files"].add(item.get("transcript", ""))

    aggregated_branches: List[Dict[str, Any]] = []
    for path, agg in aggregates.items():
        scores = agg["scores"] or [0.0]
        avg_score = sum(scores) / len(scores)
        max_score = max(scores)
        repeat_bonus = max(0, len(agg["files"]) - 1) * 5
        final_score = min(100.0, round((avg_score + max_score) / 2 + repeat_bonus, 1))
        confidences = agg["confidences"] or [0.0]
        avg_conf = sum(confidences) / len(confidences)

        evidence_unique: List[str] = []
        for quote in agg["evidence"]:
            if quote and quote not in evidence_unique:
                evidence_unique.append(quote)
            if len(evidence_unique) >= 5:
                break

        rationale_unique: List[str] = []
        for rationale in agg["rationales"]:
            if rationale and rationale not in rationale_unique:
                rationale_unique.append(rationale)
            if len(rationale_unique) >= 2:
                break

        aggregated_branches.append({
            "path": path,
            "score": final_score,
            "confidence": round(avg_conf, 2),
            "confidence_label": _confidence_label(avg_conf),
            "evidence": evidence_unique,
            "rationale": " / ".join(rationale_unique),
            "files_supported": len(agg["files"]),
        })

    aggregated_branches.sort(key=lambda x: x["score"], reverse=True)
    aggregated_branches = aggregated_branches[:max(top_n, 1)]

    should_split = False
    if total_files > 0:
        should_split = split_votes / total_files >= 0.5 and split_votes >= 2 or split_votes == total_files

    split_reason = ""
    if split_reasons:
        unique_reasons: List[str] = []
        for reason in split_reasons:
            if reason not in unique_reasons:
                unique_reasons.append(reason)
            if len(unique_reasons) >= 2:
                break
        split_reason = " / ".join(unique_reasons)

    return {
        "branches": aggregated_branches,
        "split_recommended": should_split,
        "split_reason": split_reason,
        "total_files": total_files,
    }


# =========================
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =========================

def normalize_project_name(stem: str) -> str:
    return re.sub(r"_(FG|IDI)\d+$", "", stem)


# =========================
#  АВТО-РАСКЛАДКА ТРАНСКРИПТОВ
# =========================

def sanitize_windows_name(name: str) -> str:
    """Очищает строку для безопасного использования в именах файлов/папок Windows."""
    sanitized = name.replace("/", "_").replace("\\", "_")
    sanitized = re.sub(r'[<>:"|?*]+', "", sanitized)
    sanitized = re.sub(r"\s+", " ", sanitized)
    sanitized = sanitized.strip(" .")
    return sanitized or "Unknown"


def extract_client_and_project(stem: str) -> Dict[str, Any]:
    """
    Извлекает client и project_part из stem файла по правилам.
    Возвращает словарь с полями client, project_part, base, marker_found, warning.
    """
    warning = ""
    marker_pos = -1
    marker_match = re.search(r"_(IDI|FG)\d+$", stem)
    if marker_match:
        marker_pos = marker_match.start()

    if marker_pos == -1:
        base = stem
        warning = "Предупреждение: не найден маркер _IDI/_FG — используется полное имя файла."
    else:
        base = stem[:marker_pos]

    if "_" in base:
        client, project_part = base.split("_", 1)
    else:
        client = base
        project_part = ""
        if not warning:
            warning = "Предупреждение: не найден разделитель '_' в имени проекта."

    client_match = re.match(r"^([^\d]+)", client)
    if client_match:
        client = client_match.group(1)

    return {
        "client": client,
        "project_part": project_part,
        "base": base,
        "marker_found": marker_pos != -1,
        "warning": warning,
    }


def parse_taxonomy_path(path: str) -> List[str]:
    return [part.strip() for part in path.split("/") if part.strip()]


def organize_transcripts(
    root_dir: str,
    type_code: str,
    transcript_paths: List[str],
    project_name_hint: str | None = None,
    project_suffix: str | None = None,
    taxonomy_branch: Dict[str, Any] | None = None,
    secondary_branches: List[Dict[str, Any]] | None = None,
    multi_reason: str | None = None,
) -> Dict[str, Any]:
    """
    Создаёт структуру папок Project Transcrypts и копирует транскрипты.
    Возвращает словарь с подробным логом операций.
    """
    result: Dict[str, Any] = {
        "root_dir": root_dir,
        "type_code": type_code,
        "project_dir": "",
        "project_folder_name": "",
        "client": "",
        "project_part": "",
        "base": "",
        "marker_found": False,
        "copied": 0,
        "skipped": 0,
        "errors": 0,
        "log_lines": [],
        "warning": "",
        "multi_context_path": "",
    }

    if not transcript_paths:
        result["log_lines"].append("Нет файлов для копирования.")
        return result

    if not type_code and not taxonomy_branch:
        result["errors"] += 1
        result["log_lines"].append("Не задан type_code — раскладка не выполнена.")
        return result

    root_path = pathlib.Path(root_dir or PROJECT_TRANSCRYPTS_ROOT)
    result["root_dir"] = str(root_path)
    try:
        root_path.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        result["errors"] += 1
        result["log_lines"].append(f"Ошибка создания/доступа к корню библиотеки: {e}")
        return result

    if taxonomy_branch:
        branch_path = str(taxonomy_branch.get("path", "")).strip()
        path_parts = parse_taxonomy_path(branch_path)
        if not path_parts:
            result["errors"] += 1
            result["log_lines"].append("Не задан taxonomy branch path — раскладка не выполнена.")
            return result
        type_dir = root_path
        for part in path_parts:
            type_dir = type_dir / sanitize_windows_name(part)
        try:
            type_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            result["errors"] += 1
            result["log_lines"].append(f"Ошибка создания папки taxonomy ({type_dir}): {e}")
            return result
        type_code = "_".join(path_parts)
        project_suffix = type_code
    else:
        type_dir = root_path / sanitize_windows_name(type_code)
        try:
            type_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            result["errors"] += 1
            result["log_lines"].append(f"Ошибка создания папки типа ({type_dir}): {e}")
            return result

    first_stem = pathlib.Path(transcript_paths[0]).stem
    extraction = extract_client_and_project(first_stem)
    if not extraction["marker_found"] and project_name_hint:
        hint_extraction = extract_client_and_project(pathlib.Path(project_name_hint).stem)
        if hint_extraction["marker_found"]:
            extraction = hint_extraction

    result.update({
        "client": extraction["client"],
        "project_part": extraction["project_part"],
        "base": extraction["base"],
        "marker_found": extraction["marker_found"],
        "warning": extraction["warning"],
    })

    if extraction["warning"]:
        result["log_lines"].append(extraction["warning"])

    client_sanitized = sanitize_windows_name(extraction["client"])
    project_part_sanitized = sanitize_windows_name(extraction["project_part"] or "Project")
    if project_suffix:
        project_folder_name = f"{client_sanitized}_{project_part_sanitized}__{sanitize_windows_name(project_suffix)}"
    else:
        project_folder_name = f"{client_sanitized}_{project_part_sanitized}_{sanitize_windows_name(type_code)}"
    project_folder_name = sanitize_windows_name(project_folder_name)

    client_dir = type_dir / client_sanitized
    project_dir = client_dir / project_folder_name

    try:
        project_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        result["errors"] += 1
        result["log_lines"].append(f"Ошибка создания папки проекта ({project_dir}): {e}")
        return result

    result["project_dir"] = str(project_dir)
    result["project_folder_name"] = project_folder_name

    if taxonomy_branch and secondary_branches:
        lines = [
            "Destination uses PRIMARY branch; secondary recorded.",
            f"Primary: {taxonomy_branch.get('path', '')} (score {taxonomy_branch.get('score', 0)}, confidence {taxonomy_branch.get('confidence', 0)})",
        ]
        if multi_reason:
            lines.append(f"Reason: {multi_reason}")
        lines.append("Secondary:")
        for branch in secondary_branches:
            lines.append(
                f"- {branch.get('path', '')} (score {branch.get('score', 0)}, confidence {branch.get('confidence', 0)})"
            )
            rationale = branch.get("rationale", "")
            if rationale:
                lines.append(f"  rationale: {rationale}")
        multi_path = project_dir / "_multi.txt"
        try:
            multi_path.write_text("\n".join(lines), encoding="utf-8")
            result["multi_context_path"] = str(multi_path)
        except Exception as e:
            result["errors"] += 1
            result["log_lines"].append(f"ERROR writing _multi.txt: {e}")

    for src in transcript_paths:
        src_path = pathlib.Path(src)
        dst_path = project_dir / src_path.name
        if dst_path.exists():
            result["skipped"] += 1
            result["log_lines"].append(f"SKIP exists: {dst_path.name}")
            continue

        try:
            shutil.copy2(src_path, dst_path)
            result["copied"] += 1
            result["log_lines"].append(f"COPIED: {dst_path.name}")
        except Exception as e:
            result["errors"] += 1
            result["log_lines"].append(f"ERROR copying {src_path.name}: {e}")

    return result



# =========================
#  GUI
# =========================

class StudyClassifierGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("SALT Study Classifier — классификация проекта по типу исследования")
        self.geometry("950x700")

        self.settings = load_settings()
        self.transcript_paths: List[str] = []
        self.var_premium = tk.BooleanVar(value=False)
        self.var_auto_organize = tk.BooleanVar(value=True)
        self.var_library_root = tk.StringVar(
            value=self.settings.get("library_root", PROJECT_TRANSCRYPTS_ROOT)
        )
        self.structure_path = tk.StringVar(value=self.settings.get("structure_path", ""))
        self.goal_structure: List[Dict[str, str]] = []

        self._build_ui()
        self._maybe_autoload_structure()

    def _build_ui(self):
        root = tk.Frame(self, padx=10, pady=10)
        root.pack(fill=tk.BOTH, expand=True)

        # Блок: проект
        frame_proj = tk.LabelFrame(root, text="Проект", padx=10, pady=10)
        frame_proj.pack(fill=tk.X, expand=False)

        tk.Label(frame_proj, text="Название проекта:").grid(row=0, column=0, sticky="w")
        self.entry_project_name = tk.Entry(frame_proj)
        self.entry_project_name.grid(row=0, column=1, sticky="we", padx=5)

        frame_proj.columnconfigure(1, weight=1)

        # Блок: транскрипты
        frame_tr = tk.LabelFrame(root, text="Транскрипты проекта", padx=10, pady=10)
        frame_tr.pack(fill=tk.BOTH, expand=False, pady=(10, 0))

        btn_tr = tk.Button(
            frame_tr,
            text="Выбрать транскрипты проекта (.txt/.docx)...",
            command=self.choose_transcripts
        )
        btn_tr.grid(row=0, column=0, sticky="w")

        btn_clear_tr = tk.Button(
            frame_tr,
            text="Очистить список",
            command=self.clear_transcripts
        )
        btn_clear_tr.grid(row=0, column=1, sticky="w", padx=5)

        self.lst_transcripts = tk.Listbox(frame_tr, height=6)
        self.lst_transcripts.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(5, 0))

        frame_tr.columnconfigure(0, weight=1)
        frame_tr.columnconfigure(1, weight=0)
        frame_tr.rowconfigure(1, weight=1)

        # Блок: структура целей
        frame_structure = tk.LabelFrame(root, text="Структура целей (Excel)", padx=10, pady=10)
        frame_structure.pack(fill=tk.X, expand=False, pady=(10, 0))

        tk.Label(frame_structure, text="Файл структуры:").grid(row=0, column=0, sticky="w")
        entry_structure = tk.Entry(frame_structure, textvariable=self.structure_path, state="readonly")
        entry_structure.grid(row=0, column=1, sticky="we", padx=5)

        btn_choose_structure = tk.Button(
            frame_structure,
            text="Выбрать / задать файл структуры (.xlsx)",
            command=self.choose_structure
        )
        btn_choose_structure.grid(row=0, column=2, sticky="w")

        btn_load_structure = tk.Button(
            frame_structure,
            text="Загрузить/Обновить структуру",
            command=self.load_structure
        )
        btn_load_structure.grid(row=1, column=0, sticky="w", pady=(5, 0))

        self.lbl_structure_count = tk.Label(frame_structure, text="Веток загружено: 0", anchor="w")
        self.lbl_structure_count.grid(row=1, column=1, sticky="w", padx=5, pady=(5, 0))

        frame_structure.columnconfigure(1, weight=1)

        # Блок: настройки
        frame_opts = tk.LabelFrame(root, text="Настройки", padx=10, pady=10)
        frame_opts.pack(fill=tk.X, expand=False, pady=(10, 0))

        chk_premium = tk.Checkbutton(
            frame_opts,
            text="Премиум-режим (использовать gpt-4o — дороже, но точнее)",
            variable=self.var_premium
        )
        chk_premium.grid(row=0, column=0, sticky="w")

        chk_organize = tk.Checkbutton(
            frame_opts,
            text="Авто-раскладка транскриптов в Project Transcrypts",
            variable=self.var_auto_organize
        )
        chk_organize.grid(row=1, column=0, sticky="w", pady=(5, 0))

        tk.Label(frame_opts, text="Корень библиотеки Project Transcrypts:").grid(row=2, column=0, sticky="w", pady=(5, 0))
        entry_root = tk.Entry(frame_opts, textvariable=self.var_library_root)
        entry_root.grid(row=2, column=1, sticky="we", padx=5, pady=(5, 0))

        btn_choose_root = tk.Button(
            frame_opts,
            text="Выбрать папку библиотеки...",
            command=self.choose_library_root
        )
        btn_choose_root.grid(row=2, column=2, sticky="w", pady=(5, 0))

        frame_opts.columnconfigure(1, weight=1)

        # Кнопка запуска
        frame_run = tk.Frame(root)
        frame_run.pack(fill=tk.X, expand=False, pady=(10, 0))

        self.btn_run = tk.Button(
            frame_run,
            text="Классифицировать проект",
            command=self.run_classification,
            height=2
        )
        self.btn_run.pack(side=tk.LEFT)

        # Блок: результат
        frame_res = tk.LabelFrame(root, text="Результат классификации", padx=10, pady=10)
        frame_res.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.txt_result = scrolledtext.ScrolledText(frame_res, height=10, state=tk.DISABLED)
        self.txt_result.pack(fill=tk.BOTH, expand=True)

    # ---------- обработчики ----------

    def choose_transcripts(self):
        paths = filedialog.askopenfilenames(
            title="Выбери транскрипты проекта (.txt/.docx)",
            filetypes=[
                ("Документы", "*.txt;*.docx"),
                ("Только .txt", "*.txt"),
                ("Только .docx", "*.docx"),
                ("Все файлы", "*.*"),
            ]
        )
        if not paths:
            return
        self.transcript_paths = list(paths)
        self.lst_transcripts.delete(0, tk.END)
        for p in self.transcript_paths:
            self.lst_transcripts.insert(tk.END, p)

        # если название проекта пустое — подставим из первого файла
        if self.transcript_paths and not self.entry_project_name.get().strip():
            first_stem = pathlib.Path(self.transcript_paths[0]).stem
            project_base = normalize_project_name(first_stem)
            self.entry_project_name.insert(0, project_base)

    def clear_transcripts(self):
        self.transcript_paths.clear()
        self.lst_transcripts.delete(0, tk.END)

    def choose_library_root(self):
        path = filedialog.askdirectory(
            title="Корневая папка Project Transcrypts",
            initialdir=self.var_library_root.get() or PROJECT_TRANSCRYPTS_ROOT,
            mustexist=False,
        )
        if path:
            self.var_library_root.set(path)
            self._save_settings()

    def choose_structure(self):
        path = filedialog.askopenfilename(
            title="Структура целей исследования (.xlsx)",
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if path:
            self.structure_path.set(path)
            self._save_settings()

    def load_structure(self):
        path = self.structure_path.get().strip()
        if not path:
            messagebox.showwarning("Структура не задана", "Сначала выбери файл .xlsx.")
            return

        try:
            structure = load_goal_structure(path)
        except Exception as e:
            messagebox.showerror("Ошибка загрузки структуры", str(e))
            return

        self.goal_structure = structure
        self.lbl_structure_count.config(text=f"Веток загружено: {len(self.goal_structure)}")
        messagebox.showinfo("Структура загружена", f"Загружено веток: {len(self.goal_structure)}")

    def _save_settings(self) -> None:
        self.settings["structure_path"] = self.structure_path.get().strip()
        self.settings["library_root"] = self.var_library_root.get().strip()
        save_settings(self.settings)

    def _maybe_autoload_structure(self) -> None:
        path = self.structure_path.get().strip()
        if not path:
            return
        if not pathlib.Path(path).exists():
            return
        try:
            self.goal_structure = load_goal_structure(path)
            self.lbl_structure_count.config(
                text=f"Веток загружено: {len(self.goal_structure)}"
            )
        except Exception:
            self.goal_structure = []
            self.lbl_structure_count.config(text="Веток загружено: 0")

    def run_classification(self):
        if not self.transcript_paths:
            messagebox.showwarning("Нет транскриптов", "Выбери хотя бы один транскрипт проекта.")
            return

        project_name = self.entry_project_name.get().strip()
        if not project_name:
            project_name = normalize_project_name(pathlib.Path(self.transcript_paths[0]).stem)

        try:
            # тестируем наличие ключа заранее
            _ = get_openai_client()
        except Exception as e:
            messagebox.showerror("Ошибка OpenAI", str(e))
            return

        self.btn_run.config(state=tk.DISABLED)
        try:
            self._set_result("Собираю текстовый срез по транскриптам...\n")
            text_sample = build_text_sample(self.transcript_paths, max_chars=20000)

            self._append_result(f"Общий объём среза: {len(text_sample)} символов.\n")
            self._append_result("Отправляю в модель для классификации...\n\n")

            use_premium = self.var_premium.get()
            sections: List[str] = []
            warnings: List[str] = []
            classification_mode = "coarse"
            model_name = PREMIUM_MODEL if use_premium else BASE_MODEL
            taxonomy_loaded = bool(self.goal_structure)
            top_branches: List[Dict[str, Any]] = []
            primary_branch: Dict[str, Any] | None = None
            secondary_branches: List[Dict[str, Any]] = []
            split_reason = ""
            chosen_folder_path = ""
            copied_count = 0
            skipped_count = 0
            error_count = 0

            if taxonomy_loaded:
                structure_paths = [item["path"] for item in self.goal_structure]
                per_file_results: List[Dict[str, Any]] = []
                for transcript_path in self.transcript_paths:
                    try:
                        raw_text = read_any_transcript(transcript_path)
                        sample = build_text_sample_single(raw_text, max_chars=12000)
                        file_result = classify_study_taxonomy(
                            text=sample,
                            structure_paths=structure_paths,
                            model=model_name,
                            top_n=3,
                        )
                        file_result["transcript"] = transcript_path
                        per_file_results.append(file_result)
                    except Exception as e:
                        warnings.append(f"Ошибка классификации файла {transcript_path}: {e}")

                if per_file_results:
                    aggregate_result = aggregate_taxonomy_results(per_file_results, top_n=3)
                    classification_mode = "taxonomy"
                    top_branches = aggregate_result.get("branches", [])
                    sections.append(
                        self._format_taxonomy_result(
                            project_name,
                            aggregate_result,
                            warnings,
                        )
                    )
                    cls_for_registry = self._convert_taxonomy_to_registry(aggregate_result)
                    primary_branch = top_branches[0] if top_branches else None
                    secondary_branches = top_branches[1:] if len(top_branches) > 1 else []
                    split_reason = aggregate_result.get("split_reason", "") or ""
                else:
                    warnings.append("Переход в упрощённый режим: не удалось классифицировать файлы по структуре.")

            if classification_mode == "coarse":
                cls = classify_study(
                    project_name=project_name,
                    text_sample=text_sample,
                    use_premium=use_premium,
                )
                sections.append(self._format_coarse_result(project_name, cls, warnings))
                cls_for_registry = cls
                primary_branch = None

            if self.var_auto_organize.get():
                org_lines = ["Раскладка:"]
                root_dir = self.var_library_root.get().strip() or PROJECT_TRANSCRYPTS_ROOT
                org_lines.append(f"mode: {classification_mode}")
                org_lines.append(f"ROOT: {root_dir}")
                try:
                    if classification_mode == "taxonomy" and primary_branch:
                        if secondary_branches:
                            org_lines.append("Destination uses PRIMARY branch; secondary recorded.")
                        organize_result = organize_transcripts(
                            root_dir=root_dir,
                            type_code="",
                            transcript_paths=self.transcript_paths,
                            project_name_hint=project_name,
                            taxonomy_branch=primary_branch,
                            secondary_branches=secondary_branches,
                            multi_reason=split_reason,
                        )
                    else:
                        type_code = cls_for_registry.get("group_code", "")
                        if not type_code:
                            org_lines.append("Ошибка: group_code пустой — раскладка не выполнена.")
                            organize_result = None
                        else:
                            organize_result = organize_transcripts(
                                root_dir=root_dir,
                                type_code=type_code,
                                transcript_paths=self.transcript_paths,
                                project_name_hint=project_name,
                            )

                    if organize_result:
                        org_lines.append(f"type_code: {organize_result.get('type_code', '')}")
                        org_lines.append(f"client: {organize_result.get('client', '')}")
                        org_lines.append(f"project_part: {organize_result.get('project_part', '')}")
                        org_lines.append(f"project_folder_name: {organize_result.get('project_folder_name', '')}")
                        org_lines.append(f"project_dir: {organize_result.get('project_dir', '')}")
                        chosen_folder_path = organize_result.get("project_dir", "")
                        copied_count = organize_result.get("copied", 0)
                        skipped_count = organize_result.get("skipped", 0)
                        error_count = organize_result.get("errors", 0)
                        org_lines.append(
                            f"copied: {copied_count}, skipped: {skipped_count}, errors: {error_count}"
                        )
                        if organize_result.get("multi_context_path"):
                            org_lines.append(f"multi_context: {organize_result.get('multi_context_path')}")
                        if organize_result.get("warning"):
                            org_lines.append(organize_result["warning"])

                        log_lines = organize_result.get("log_lines", [])
                        org_lines.append("Операции:")
                        max_lines = 50
                        if log_lines:
                            org_lines.extend(log_lines[:max_lines])
                            extra = len(log_lines) - max_lines
                            if extra > 0:
                                org_lines.append(f"...и ещё {extra} строк(и).")
                        else:
                            org_lines.append("(нет операций)")
                except Exception as e:
                    org_lines.append(f"Ошибка автосортировки: {e}")

                sections.append("\n".join(org_lines))

            taxonomy_loaded_text = "yes" if taxonomy_loaded else "no"
            top_branches_text = ", ".join(
                f"{b.get('path', '')} ({b.get('score', 0)})" for b in top_branches
            ) or "(нет)"
            sections.append(
                "\n".join([
                    "Лог классификации:",
                    f"project_name: {project_name}",
                    f"taxonomy_loaded: {taxonomy_loaded_text}",
                    f"leaves_count: {len(self.goal_structure) if taxonomy_loaded else 0}",
                    f"top_branches: {top_branches_text}",
                    f"chosen_folder_path: {chosen_folder_path or '(не задан)'}",
                    f"copied: {copied_count}, skipped: {skipped_count}, errors: {error_count}",
                ])
            )

            self._set_result("\n\n".join(sections))
        except Exception as e:
            self._append_result(f"\nОШИБКА: {e}\n")
            messagebox.showerror("Ошибка классификации", str(e))
        finally:
            self.btn_run.config(state=tk.NORMAL)

    # ---------- работа с полем результата ----------

    def _set_result(self, text: str):
        self.txt_result.config(state=tk.NORMAL)
        self.txt_result.delete("1.0", tk.END)
        self.txt_result.insert(tk.END, text)
        self.txt_result.see(tk.END)
        self.txt_result.config(state=tk.DISABLED)

    def _append_result(self, text: str):
        self.txt_result.config(state=tk.NORMAL)
        self.txt_result.insert(tk.END, text)
        self.txt_result.see(tk.END)
        self.txt_result.config(state=tk.DISABLED)

    def _format_coarse_result(
        self,
        project_name: str,
        cls: Dict[str, Any],
        warnings: List[str],
    ) -> str:
        lines = []
        lines.append(f"Проект: {project_name}")
        lines.append(f"Группа: {cls['group_code']} — {cls['group_name']}")
        lines.append(f"Уверенность: {cls['confidence']:.2f}")
        lines.append("")
        lines.append("Обоснование:")
        lines.append(cls["justification"] or "(пусто)")
        lines.append("")
        if cls["key_signals"]:
            lines.append("Ключевые сигналы из транскриптов:")
            for i, sig in enumerate(cls["key_signals"], start=1):
                lines.append(f"  {i}. {sig}")
        else:
            lines.append("Ключевые сигналы: (нет)")

        if warnings:
            lines.append("")
            lines.append("ПРЕДУПРЕЖДЕНИЯ:")
            lines.extend(f"- {warning}" for warning in warnings)

        return "\n".join(lines)

    def _format_taxonomy_result(
        self,
        project_name: str,
        aggregate_result: Dict[str, Any],
        warnings: List[str],
    ) -> str:
        lines = []
        lines.append(f"Проект: {project_name}")
        lines.append("Top branches:")
        branches = aggregate_result.get("branches", [])
        total_files = aggregate_result.get("total_files", 0)

        if branches:
            primary = branches[0]
            lines.append(
                "PRIMARY: "
                f"{primary.get('path', '')} — score {primary.get('score', 0):.1f}, "
                f"confidence {primary.get('confidence', 0):.2f} ({primary.get('confidence_label', '')})"
            )
            if len(branches) > 1:
                lines.append("SECONDARY:")
                for branch in branches[1:]:
                    lines.append(
                        f"- {branch.get('path', '')} — score {branch.get('score', 0):.1f}, "
                        f"confidence {branch.get('confidence', 0):.2f} ({branch.get('confidence_label', '')})"
                    )
            lines.append("")
            for idx, branch in enumerate(branches, start=1):
                lines.append(
                    f"{idx}. {branch.get('path', '')} — score {branch.get('score', 0):.1f}, "
                    f"confidence {branch.get('confidence', 0):.2f} ({branch.get('confidence_label', '')}), "
                    f"files {branch.get('files_supported', 0)}/{total_files}"
                )
                rationale = branch.get("rationale", "")
                if rationale:
                    lines.append(f"   rationale: {rationale}")
                evidence = branch.get("evidence", [])
                if evidence:
                    lines.append("   evidence:")
                    for quote in evidence:
                        lines.append(f"     - {quote}")
        else:
            lines.append("(нет)")

        split_yes_no = "Да" if aggregate_result.get("split_recommended") else "Нет"
        split_reason = aggregate_result.get("split_reason", "") or "(нет)"
        lines.append("")
        lines.append(f"Split recommendation: {split_yes_no}")
        lines.append(f"Reason: {split_reason}")

        if warnings:
            lines.append("")
            lines.append("ПРЕДУПРЕЖДЕНИЯ:")
            lines.extend(f"- {warning}" for warning in warnings)

        return "\n".join(lines)

    def _convert_taxonomy_to_registry(
        self,
        aggregate_result: Dict[str, Any],
    ) -> Dict[str, Any]:
        branches = aggregate_result.get("branches", [])
        top_branch = branches[0] if branches else {}
        confidence = float(top_branch.get("score", 0.0)) / 100 if branches else 0.0
        return {
            "group_code": top_branch.get("path", ""),
            "group_name": "Taxonomy",
            "confidence": confidence,
            "justification": top_branch.get("rationale", ""),
            "key_signals": top_branch.get("evidence", []),
        }


if __name__ == "__main__":
    app = StudyClassifierGUI()
    app.mainloop()
