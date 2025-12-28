import tkinter as tk
from tkinter import filedialog, messagebox
import json
import difflib
from datetime import datetime


SIMILARITY_THRESHOLD = 0.8


class CoreTagsManagerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SALT Core Tags Manager")

        self.core_tags = []
        self.candidate_tags = []

        self.build_ui()

    # ---------- UI ----------
    def build_ui(self):
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(fill="both", expand=True)

        tk.Button(frame, text="Загрузить базу обязательных тегов",
                  command=self.load_core_tags).pack(fill="x")

        tk.Button(frame, text="Загрузить кандидатов (JSON)",
                  command=self.load_candidates).pack(fill="x", pady=(5, 10))

        lists_frame = tk.Frame(frame)
        lists_frame.pack(fill="both", expand=True)

        # Existing core
        left = tk.Frame(lists_frame)
        left.pack(side="left", fill="both", expand=True, padx=(0, 5))

        tk.Label(left, text="Обязательные теги").pack()
        self.core_list = tk.Listbox(left)
        self.core_list.pack(fill="both", expand=True)

        # Candidates
        right = tk.Frame(lists_frame)
        right.pack(side="right", fill="both", expand=True, padx=(5, 0))

        tk.Label(right, text="Новые кандидаты").pack()
        self.candidate_list = tk.Listbox(right)
        self.candidate_list.pack(fill="both", expand=True)

        tk.Button(frame, text="➕ Добавить выбранный тег в ядро",
                  command=self.add_selected_tag).pack(fill="x", pady=5)

        tk.Button(frame, text="💾 Сохранить базу",
                  command=self.save_core_tags).pack(fill="x")

    # ---------- Logic ----------
    def load_core_tags(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return

        with open(path, "r", encoding="utf-8") as f:
            self.core_tags = json.load(f)

        self.refresh_core_list()

    def load_candidates(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return

        with open(path, "r", encoding="utf-8") as f:
            raw_candidates = json.load(f)

        self.candidate_tags = self.filter_candidates(raw_candidates)
        self.refresh_candidate_list()

    def filter_candidates(self, candidates):
        filtered = []

        for cand in candidates:
            name = cand["tag"]
            if not self.is_similar_to_core(name):
                filtered.append(cand)

        return filtered

    def is_similar_to_core(self, candidate_name):
        for core in self.core_tags:
            ratio = difflib.SequenceMatcher(
                None, candidate_name.lower(), core["tag"].lower()
            ).ratio()
            if ratio >= SIMILARITY_THRESHOLD:
                return True
        return False

    def add_selected_tag(self):
        idx = self.candidate_list.curselection()
        if not idx:
            return

        tag = self.candidate_tags.pop(idx[0])
        tag["mandatory"] = True
        tag["added_at"] = datetime.now().isoformat()

        self.core_tags.append(tag)

        self.refresh_core_list()
        self.refresh_candidate_list()

    def save_core_tags(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")]
        )
        if not path:
            return

        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.core_tags, f, ensure_ascii=False, indent=2)

        messagebox.showinfo("Готово", "База обязательных тегов сохранена.")

    # ---------- Helpers ----------
    def refresh_core_list(self):
        self.core_list.delete(0, tk.END)
        for tag in self.core_tags:
            self.core_list.insert(tk.END, tag["tag"])

    def refresh_candidate_list(self):
        self.candidate_list.delete(0, tk.END)
        for tag in self.candidate_tags:
            self.candidate_list.insert(tk.END, tag["tag"])


if __name__ == "__main__":
    root = tk.Tk()
    app = CoreTagsManagerGUI(root)
    root.mainloop()
