import json
import os
import re
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import openpyxl

APP_NAME = "organizzer"
STATE_FILE = Path.home() / ".organizzer_state.json"
TEMP_DOWNLOAD_EXTS = {".tmp", ".crdownload", ".part", ".download"}
IGNORED_NAMES = {"desktop.ini", "thumbs.db"}

# ── Paleta: base branca com acentos em azul ────────────────────────────────
BG_MAIN = "#F5F8FD"
BG_CARD = "#FFFFFF"
BG_SUBTLE = "#EEF4FF"
BG_FLASH = "#DBEAFE"
BORDER_BLUE = "#C7DAF7"
BORDER_SOFT = "#E3EAF5"
BLUE_PRIMARY = "#1D4ED8"
BLUE_HOVER = "#1E40AF"
BLUE_SOFT = "#3B82F6"
TEXT_DARK = "#0F172A"
TEXT_MUTED = "#64748B"
TEXT_FOOTER = "#94A3B8"
GREEN = "#16A34A"
AMBER = "#D97706"
RED = "#DC2626"
RED_SOFT = "#FEE2E2"

FONT_FAMILY = "Segoe UI"

STATUS_COLORS = {"info": BLUE_SOFT, "ok": GREEN, "warn": AMBER, "error": RED}


def font(size: int = 12, weight: str = "normal") -> ctk.CTkFont:
    return ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)


def sanitize_folder_name(name) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', " ", str(name or "")).strip()
    return re.sub(r"\s+", " ", cleaned)


def lerp_color(c1: str, c2: str, t: float) -> str:
    """Interpola entre duas cores hex (t=0 -> c1, t=1 -> c2)."""
    t = max(0.0, min(1.0, t))
    a = [int(c1[i : i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i : i + 2], 16) for i in (1, 3, 5)]
    r = [round(x + (y - x) * t) for x, y in zip(a, b)]
    return "#{:02X}{:02X}{:02X}".format(*r)


def human_size(num: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            return f"{num:.0f} {unit}" if unit == "B" else f"{num:.1f} {unit}"
        num /= 1024
    return ""


def shorten(text: str, limit: int) -> str:
    return text if len(text) <= limit else text[: limit - 1] + "…"


class OrganizzerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("light")
        self.title(APP_NAME)
        self.geometry("700x760")
        self.minsize(640, 680)
        self.configure(fg_color=BG_MAIN)

        self.insurer_var = ctk.StringVar(value="")
        self.excel_path_var = ctk.StringVar(value="")
        self.source_dir_var = ctk.StringVar(value=str(Path.home() / "Downloads"))

        self.queue_items = []
        self.current_index = 0
        self.known_files_map = {}  # path -> (mtime, size)
        self._pending_files = {}  # path -> (mtime, size) aguardando estabilizar
        self.captured_files = []
        self.base_server_dir = None
        self.is_running_queue = False
        self.cfg_open = True

        self.progress_target = 0.0
        self.progress_current = 0.0

        self._load_saved_prefs()
        self._build_ui()
        self._update_buttons_state()
        self._set_status("Configure a seguradora, a planilha e a pasta de origem para iniciar.")

        self.bind("<Control-Return>", lambda _e: self._cut_and_next())

        self._poll_source_directory()
        self._animate_ui()

    # ── Preferências ───────────────────────────────────────────────────────
    def _load_saved_prefs(self):
        if STATE_FILE.exists():
            try:
                data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
                if data.get("last_source_dir") and os.path.isdir(data["last_source_dir"]):
                    self.source_dir_var.set(data["last_source_dir"])
                if data.get("last_insurer"):
                    self.insurer_var.set(data["last_insurer"])
                if data.get("last_excel") and os.path.isfile(data["last_excel"]):
                    self.excel_path_var.set(data["last_excel"])
            except Exception:
                pass

    def _save_prefs(self):
        try:
            payload = {
                "last_source_dir": self.source_dir_var.get().strip(),
                "last_insurer": self.insurer_var.get().strip(),
                "last_excel": self.excel_path_var.get().strip(),
            }
            STATE_FILE.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception:
            pass

    # ── Helpers de UI ──────────────────────────────────────────────────────
    def _card(self, parent, **kw) -> ctk.CTkFrame:
        return ctk.CTkFrame(
            parent,
            fg_color=BG_CARD,
            corner_radius=16,
            border_width=1,
            border_color=BORDER_SOFT,
            **kw,
        )

    def _secondary_btn(self, parent, text, command, width=92, height=32):
        return ctk.CTkButton(
            parent,
            text=text,
            width=width,
            height=height,
            corner_radius=9,
            fg_color=BG_SUBTLE,
            hover_color=BORDER_BLUE,
            text_color=BLUE_PRIMARY,
            font=font(12, "bold"),
            command=command,
        )

    def _entry(self, parent, var, placeholder):
        return ctk.CTkEntry(
            parent,
            textvariable=var,
            placeholder_text=placeholder,
            fg_color=BG_MAIN,
            border_color=BORDER_BLUE,
            border_width=1,
            text_color=TEXT_DARK,
            placeholder_text_color=TEXT_FOOTER,
            corner_radius=9,
            height=34,
            font=font(12),
        )

    def _set_status(self, text: str, kind: str = "info"):
        self.status_dot.configure(text_color=STATUS_COLORS.get(kind, BLUE_SOFT))
        self.status_label.configure(text=text)

    # ── Construção da interface ────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Cabeçalho
        header = self._card(self)
        header.grid(row=0, column=0, padx=20, pady=(18, 10), sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        badge = ctk.CTkFrame(header, fg_color=BLUE_PRIMARY, corner_radius=12, width=44, height=44)
        badge.grid(row=0, column=0, padx=(16, 12), pady=(14, 8))
        badge.grid_propagate(False)
        ctk.CTkLabel(badge, text="OG", font=font(15, "bold"), text_color="#FFFFFF").place(
            relx=0.5, rely=0.5, anchor="center"
        )

        titles = ctk.CTkFrame(header, fg_color="transparent")
        titles.grid(row=0, column=1, sticky="w", pady=(12, 6))
        ctk.CTkLabel(titles, text=APP_NAME, font=font(21, "bold"), text_color=TEXT_DARK).pack(
            anchor="w"
        )
        ctk.CTkLabel(
            titles,
            text="Organizador de pastas de renovação e arquivos do Quiver",
            font=font(12),
            text_color=TEXT_MUTED,
        ).pack(anchor="w")

        pill = ctk.CTkFrame(header, fg_color=BG_SUBTLE, corner_radius=99)
        pill.grid(row=0, column=2, padx=(8, 16), pady=(14, 8))
        self.progress_label = ctk.CTkLabel(
            pill, text="0 / 0", font=font(13, "bold"), text_color=BLUE_PRIMARY
        )
        self.progress_label.pack(padx=14, pady=4)

        self.progress_bar = ctk.CTkProgressBar(
            header,
            height=6,
            corner_radius=99,
            fg_color=BG_SUBTLE,
            progress_color=BLUE_PRIMARY,
        )
        self.progress_bar.grid(row=1, column=0, columnspan=3, padx=16, pady=(0, 14), sticky="ew")
        self.progress_bar.set(0)

        # Configuração do lote (recolhível)
        self.config_card = self._card(self)
        self.config_card.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.config_card.grid_columnconfigure(0, weight=1)

        top_cfg = ctk.CTkFrame(self.config_card, fg_color="transparent")
        top_cfg.grid(row=0, column=0, padx=16, pady=(12, 4), sticky="ew")
        top_cfg.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            top_cfg, text="CONFIGURAÇÃO DO LOTE", font=font(11, "bold"), text_color=BLUE_PRIMARY
        ).grid(row=0, column=0, sticky="w")

        self.cfg_summary = ctk.CTkLabel(
            top_cfg, text="", font=font(11), text_color=TEXT_MUTED, anchor="w"
        )
        self.cfg_summary.grid(row=0, column=1, padx=12, sticky="ew")

        self.toggle_cfg_btn = self._secondary_btn(
            top_cfg, "Recolher", self._toggle_config_card, width=78, height=26
        )
        self.toggle_cfg_btn.configure(font=font(11, "bold"))
        self.toggle_cfg_btn.grid(row=0, column=2, sticky="e")

        self.config_body = ctk.CTkFrame(self.config_card, fg_color="transparent")
        self.config_body.grid(row=1, column=0, padx=16, pady=(2, 14), sticky="ew")
        self.config_body.grid_columnconfigure(1, weight=1)

        def label(text, row):
            ctk.CTkLabel(
                self.config_body, text=text, font=font(12, "bold"), text_color=TEXT_DARK
            ).grid(row=row, column=0, padx=(0, 10), pady=5, sticky="w")

        label("Seguradora", 0)
        self._entry(self.config_body, self.insurer_var, "Ex.: PORTO SEGURO GERAIS").grid(
            row=0, column=1, columnspan=2, pady=5, sticky="ew"
        )

        label("Planilha", 1)
        self._entry(
            self.config_body, self.excel_path_var, "Selecione a planilha (aba RENOVAÇÃO)"
        ).grid(row=1, column=1, padx=(0, 8), pady=5, sticky="ew")
        self._secondary_btn(self.config_body, "Procurar", self._choose_excel_file).grid(
            row=1, column=2, pady=5
        )

        label("Pasta origem", 2)
        self._entry(
            self.config_body, self.source_dir_var, "Pasta onde caem os downloads do Quiver"
        ).grid(row=2, column=1, padx=(0, 8), pady=5, sticky="ew")
        self._secondary_btn(self.config_body, "Escolher", self._choose_source_folder).grid(
            row=2, column=2, pady=5
        )

        self.start_btn = ctk.CTkButton(
            self.config_body,
            text="Ler planilha, criar pastas e iniciar fila",
            height=40,
            corner_radius=11,
            fg_color=BLUE_PRIMARY,
            hover_color=BLUE_HOVER,
            text_color="#FFFFFF",
            font=font(13, "bold"),
            command=self._start_batch,
        )
        self.start_btn.grid(row=3, column=0, columnspan=3, pady=(10, 0), sticky="ew")

        # Área de trabalho
        work = self._card(self)
        work.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="nsew")
        work.grid_columnconfigure(0, weight=1)
        work.grid_rowconfigure(2, weight=1)

        # Banner da pasta atual (com faixa de destaque lateral)
        self.current_banner = ctk.CTkFrame(
            work,
            fg_color=BG_SUBTLE,
            corner_radius=12,
            border_width=1,
            border_color=BORDER_BLUE,
        )
        self.current_banner.grid(row=0, column=0, padx=16, pady=(14, 8), sticky="ew")
        self.current_banner.grid_columnconfigure(1, weight=1)

        accent = ctk.CTkFrame(self.current_banner, fg_color=BLUE_PRIMARY, width=5, corner_radius=3)
        accent.grid(row=0, column=0, rowspan=3, padx=(8, 0), pady=10, sticky="ns")

        ctk.CTkLabel(
            self.current_banner,
            text="PASTA ATUAL DE DESTINO",
            font=font(10, "bold"),
            text_color=BLUE_PRIMARY,
        ).grid(row=0, column=1, columnspan=3, padx=12, pady=(10, 0), sticky="w")

        self.current_item_label = ctk.CTkLabel(
            self.current_banner,
            text="Nenhum lote iniciado",
            font=font(15, "bold"),
            text_color=TEXT_DARK,
            wraplength=380,
            justify="left",
            anchor="w",
        )
        self.current_item_label.grid(row=1, column=1, padx=12, pady=(2, 0), sticky="w")

        self.next_up_label = ctk.CTkLabel(
            self.current_banner, text="", font=font(11), text_color=TEXT_MUTED, anchor="w"
        )
        self.next_up_label.grid(row=2, column=1, padx=12, pady=(0, 10), sticky="w")

        btns = ctk.CTkFrame(self.current_banner, fg_color="transparent")
        btns.grid(row=0, column=2, rowspan=3, padx=(0, 12), sticky="e")

        def ghost_btn(text, cmd, width):
            return ctk.CTkButton(
                btns,
                text=text,
                width=width,
                height=28,
                corner_radius=8,
                fg_color=BG_CARD,
                hover_color=BORDER_BLUE,
                text_color=BLUE_PRIMARY,
                border_width=1,
                border_color=BORDER_BLUE,
                font=font(11, "bold"),
                command=cmd,
            )

        self.copy_btn = ghost_btn("Copiar", self._copy_current_name, 74)
        self.copy_btn.pack(side="left", padx=(0, 6))
        self.open_folder_btn = ghost_btn("Abrir pasta", self._open_current_folder, 90)
        self.open_folder_btn.pack(side="left")

        self.current_banner.bind(
            "<Configure>",
            lambda e: self.current_item_label.configure(wraplength=max(180, e.width - 260)),
        )

        # Cabeçalho da lista de arquivos
        files_header = ctk.CTkFrame(work, fg_color="transparent")
        files_header.grid(row=1, column=0, padx=16, pady=(4, 4), sticky="ew")
        files_header.grid_columnconfigure(0, weight=1)

        title_box = ctk.CTkFrame(files_header, fg_color="transparent")
        title_box.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            title_box, text="ARQUIVOS DETECTADOS", font=font(11, "bold"), text_color=TEXT_MUTED
        ).pack(side="left")
        self.counter_label = ctk.CTkLabel(
            title_box, text="", font=font(11), text_color=BLUE_PRIMARY
        )
        self.counter_label.pack(side="left", padx=10)

        def link_btn(text, cmd, col, width):
            ctk.CTkButton(
                files_header,
                text=text,
                width=width,
                height=24,
                corner_radius=7,
                fg_color="transparent",
                hover_color=BG_SUBTLE,
                text_color=BLUE_PRIMARY,
                font=font(11, "bold"),
                command=cmd,
            ).grid(row=0, column=col, sticky="e")

        link_btn("Marcar todos", lambda: self._set_all_checked(True), 1, 92)
        link_btn("Desmarcar", lambda: self._set_all_checked(False), 2, 72)
        link_btn("Limpar lista", self._clear_captured_list, 3, 84)

        self.files_scroll = ctk.CTkScrollableFrame(
            work,
            fg_color=BG_MAIN,
            corner_radius=12,
            border_width=1,
            border_color=BORDER_SOFT,
        )
        self.files_scroll.grid(row=2, column=0, padx=16, pady=(0, 12), sticky="nsew")
        self.files_scroll.grid_columnconfigure(0, weight=1)

        self.empty_box = ctk.CTkFrame(self.files_scroll, fg_color="transparent")
        ctk.CTkLabel(self.empty_box, text="⬇", font=font(30), text_color=BORDER_BLUE).pack(
            pady=(30, 0)
        )
        ctk.CTkLabel(
            self.empty_box,
            text="Aguardando novos arquivos",
            font=font(13, "bold"),
            text_color=TEXT_MUTED,
        ).pack()
        ctk.CTkLabel(
            self.empty_box,
            text="Tudo que cair na pasta de origem aparece aqui automaticamente.",
            font=font(11),
            text_color=TEXT_FOOTER,
        ).pack(pady=(0, 30))
        self.empty_box.pack(fill="x")

        # Barra de ações
        action_bar = ctk.CTkFrame(work, fg_color="transparent")
        action_bar.grid(row=3, column=0, padx=16, pady=(0, 10), sticky="ew")
        action_bar.grid_columnconfigure(2, weight=1)

        self.back_btn = self._secondary_btn(action_bar, "← Voltar", self._go_back, 90, 42)
        self.back_btn.grid(row=0, column=0, padx=(0, 8))
        self.skip_btn = self._secondary_btn(action_bar, "Pular", self._skip_item, 90, 42)
        self.skip_btn.grid(row=0, column=1, padx=(0, 10))

        self.next_btn = ctk.CTkButton(
            action_bar,
            text="Recortar e próximo  →",
            height=42,
            corner_radius=11,
            fg_color=BLUE_PRIMARY,
            hover_color=BLUE_HOVER,
            text_color="#FFFFFF",
            font=font(13, "bold"),
            command=self._cut_and_next,
        )
        self.next_btn.grid(row=0, column=2, sticky="ew")

        # Barra de status
        status_row = ctk.CTkFrame(work, fg_color="transparent")
        status_row.grid(row=4, column=0, padx=16, pady=(0, 12), sticky="ew")
        self.status_dot = ctk.CTkLabel(
            status_row, text="●", font=font(11), text_color=BLUE_SOFT, width=14
        )
        self.status_dot.pack(side="left")
        self.status_label = ctk.CTkLabel(
            status_row,
            text="",
            font=font(11),
            text_color=TEXT_MUTED,
            wraplength=580,
            justify="left",
            anchor="w",
        )
        self.status_label.pack(side="left", padx=(4, 0), fill="x", expand=True)

        ctk.CTkLabel(
            self,
            text="Criado por Matheus Carvalho  •  Direitos reservados à Totumseg",
            font=font(10),
            text_color=TEXT_FOOTER,
        ).grid(row=3, column=0, pady=(0, 8))

    # ── Configuração recolhível ────────────────────────────────────────────
    def _toggle_config_card(self):
        if self.cfg_open:
            self.config_body.grid_remove()
            self.toggle_cfg_btn.configure(text="Expandir")
            insurer = self.insurer_var.get().strip()
            sheet = Path(self.excel_path_var.get().strip()).name
            self.cfg_summary.configure(text=shorten(f"{insurer}  •  {sheet}", 60))
        else:
            self.config_body.grid()
            self.toggle_cfg_btn.configure(text="Recolher")
            self.cfg_summary.configure(text="")
        self.cfg_open = not self.cfg_open

    def _choose_excel_file(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha de renovação",
            filetypes=[("Planilhas Excel", "*.xlsx *.xlsm")],
        )
        if path:
            self.excel_path_var.set(path)

    def _choose_source_folder(self):
        current = self.source_dir_var.get().strip()
        initial = current if os.path.isdir(current) else str(Path.home())
        folder = filedialog.askdirectory(
            title="Escolha a pasta de onde puxar os downloads", initialdir=initial
        )
        if folder:
            self.source_dir_var.set(folder)
            self._snapshot_source_dir()
            self._save_prefs()

    # ── Monitoramento da pasta de origem ───────────────────────────────────
    @staticmethod
    def _is_candidate(entry: Path) -> bool:
        name = entry.name.lower()
        if name.startswith((".", "~$")) or name in IGNORED_NAMES:
            return False
        return entry.suffix.lower() not in TEMP_DOWNLOAD_EXTS

    def _scan_source(self):
        src = Path(self.source_dir_var.get().strip())
        found = []
        if not src.is_dir():
            return found
        try:
            for entry in src.iterdir():
                try:
                    if entry.is_file() and self._is_candidate(entry):
                        st = entry.stat()
                        found.append((st.st_mtime, st.st_size, entry))
                except OSError:
                    continue
        except OSError:
            return []
        found.sort(key=lambda t: t[0])
        return found

    def _snapshot_source_dir(self):
        self.known_files_map = {
            str(entry.resolve()): (mtime, size) for mtime, size, entry in self._scan_source()
        }
        self._pending_files.clear()

    def _poll_source_directory(self):
        existing = {f["path"] for f in self.captured_files}
        for mtime, size, entry in self._scan_source():
            resolved = str(entry.resolve())
            sig = (mtime, size)
            if self.known_files_map.get(resolved) == sig:
                continue
            # só captura quando o arquivo parou de crescer (download concluído)
            if self._pending_files.get(resolved) != sig:
                self._pending_files[resolved] = sig
                continue
            self._pending_files.pop(resolved, None)
            self.known_files_map[resolved] = sig
            if resolved in existing:
                continue
            info = {
                "path": resolved,
                "name": entry.name,
                "size": size,
                "time": datetime.now().strftime("%H:%M"),
                "var": ctk.BooleanVar(value=True),
                "highlight": 1.0,
            }
            self.captured_files.append(info)
            self._add_file_row(info)
            existing.add(resolved)

        self._update_empty_state()
        self.after(1200, self._poll_source_directory)

    # ── Lista de arquivos (atualização incremental, sem recriar tudo) ──────
    def _add_file_row(self, info):
        frame = ctk.CTkFrame(
            self.files_scroll,
            fg_color=lerp_color(BG_CARD, BG_FLASH, info["highlight"]),
            corner_radius=10,
            border_width=1,
            border_color=BORDER_SOFT,
        )
        frame.pack(fill="x", padx=4, pady=3)
        frame.grid_columnconfigure(1, weight=1)

        ext = Path(info["name"]).suffix.lstrip(".").upper()[:4] or "ARQ"
        ctk.CTkLabel(
            frame,
            text=ext,
            width=44,
            height=22,
            corner_radius=6,
            fg_color=BG_SUBTLE,
            text_color=BLUE_PRIMARY,
            font=font(10, "bold"),
        ).grid(row=0, column=0, padx=(10, 8), pady=9)

        ctk.CTkCheckBox(
            frame,
            text=shorten(info["name"], 42),
            variable=info["var"],
            command=self._update_counter,
            checkbox_width=18,
            checkbox_height=18,
            corner_radius=5,
            fg_color=BLUE_PRIMARY,
            hover_color=BLUE_HOVER,
            border_color=BLUE_SOFT,
            text_color=TEXT_DARK,
            font=font(12),
        ).grid(row=0, column=1, sticky="w")

        ctk.CTkLabel(
            frame,
            text=f"{human_size(info['size'])}  •  {info['time']}",
            font=font(10),
            text_color=TEXT_FOOTER,
        ).grid(row=0, column=2, padx=8)

        ctk.CTkButton(
            frame,
            text="✕",
            width=26,
            height=24,
            corner_radius=6,
            fg_color="transparent",
            hover_color=RED_SOFT,
            text_color=RED,
            font=font(12, "bold"),
            command=lambda inf=info: self._remove_file(inf),
        ).grid(row=0, column=3, padx=(0, 8))

        info["frame"] = frame
        self._update_counter()

    def _drop_row(self, info):
        frame = info.pop("frame", None)
        if frame is not None:
            frame.destroy()
        if info in self.captured_files:
            self.captured_files.remove(info)

    def _remove_file(self, info):
        self._drop_row(info)
        self._update_empty_state()
        self._update_counter()

    def _clear_captured_list(self):
        for info in list(self.captured_files):
            self._drop_row(info)
        self._update_empty_state()
        self._update_counter()

    def _set_all_checked(self, value: bool):
        for info in self.captured_files:
            info["var"].set(value)
        self._update_counter()

    def _update_counter(self):
        total = len(self.captured_files)
        if not total:
            self.counter_label.configure(text="")
            return
        marked = sum(1 for f in self.captured_files if f["var"].get())
        self.counter_label.configure(text=f"{marked} de {total} marcado(s)")

    def _update_empty_state(self):
        if self.captured_files:
            self.empty_box.pack_forget()
        elif not self.empty_box.winfo_manager():
            self.empty_box.pack(fill="x")

    # ── Fila ───────────────────────────────────────────────────────────────
    def _start_batch(self):
        insurer = sanitize_folder_name(self.insurer_var.get())
        excel_path = Path(self.excel_path_var.get().strip())
        source_dir = Path(self.source_dir_var.get().strip())

        if not insurer:
            messagebox.showwarning("Atenção", "Informe o nome da seguradora.")
            return
        if not excel_path.is_file():
            messagebox.showwarning("Atenção", "Selecione uma planilha válida.")
            return
        if not source_dir.is_dir():
            messagebox.showwarning("Atenção", "Escolha uma pasta de origem válida.")
            return

        self._set_status("Lendo planilha...")
        self.update_idletasks()

        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        except Exception as exc:
            messagebox.showerror("Erro ao abrir planilha", f"Não foi possível ler o arquivo:\n{exc}")
            self._set_status("Falha ao abrir a planilha.", "error")
            return

        if "RENOVAÇÃO" not in wb.sheetnames:
            wb.close()
            messagebox.showerror(
                "Aba não encontrada", "A aba 'RENOVAÇÃO' não foi encontrada na planilha."
            )
            self._set_status("Aba 'RENOVAÇÃO' não encontrada.", "error")
            return

        ws = wb["RENOVAÇÃO"]
        base_dir = excel_path.parent
        items = []
        seen = set()

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, min_col=3, max_col=4, values_only=True), start=2):
            locador = sanitize_folder_name(row[0] if len(row) > 0 else "")
            inquilino = sanitize_folder_name(row[1] if len(row) > 1 else "")
            if not locador and not inquilino:
                continue

            pair_label = " x ".join(p for p in (locador, inquilino) if p)
            folder_name = shorten(sanitize_folder_name(f"{insurer} - {pair_label}"), 150).rstrip(". ")
            dest_path = base_dir / folder_name

            if folder_name not in seen:
                try:
                    dest_path.mkdir(parents=True, exist_ok=True)
                except OSError as exc:
                    wb.close()
                    messagebox.showerror("Erro ao criar pasta", f"{folder_name}\n\n{exc}")
                    self._set_status("Falha ao criar pastas.", "error")
                    return
                seen.add(folder_name)

            items.append(
                {
                    "row": row_idx,
                    "locador": locador,
                    "inquilino": inquilino,
                    "folder_name": folder_name,
                    "dest_path": dest_path,
                    "status": "pendente",
                }
            )

        wb.close()

        if not items:
            messagebox.showinfo(
                "Sem registros",
                "Nenhum locador/inquilino encontrado a partir da linha 2 nas colunas C e D.",
            )
            self._set_status("Nenhum registro encontrado na planilha.", "warn")
            return

        self.base_server_dir = base_dir
        self.queue_items = items
        self.current_index = 0
        self.is_running_queue = True
        self._clear_captured_list()
        self._snapshot_source_dir()
        self._save_prefs()

        if self.cfg_open:
            self._toggle_config_card()

        self._refresh_current_view()
        self._update_buttons_state()
        self._set_status(f"{len(items)} pastas preparadas em: {base_dir}", "ok")

    def _update_buttons_state(self):
        active = self.is_running_queue and self.current_index < len(self.queue_items)
        state = "normal" if active else "disabled"
        for btn in (self.next_btn, self.skip_btn, self.copy_btn, self.open_folder_btn):
            btn.configure(state=state)
        self.back_btn.configure(
            state="normal" if self.is_running_queue and self.current_index > 0 else "disabled"
        )

    def _refresh_current_view(self):
        total = len(self.queue_items)
        if not self.is_running_queue or total == 0:
            self.progress_label.configure(text="0 / 0")
            self.progress_target = 0.0
            self.current_item_label.configure(text="Nenhum lote iniciado")
            self.next_up_label.configure(text="")
            return

        if self.current_index >= total:
            done = sum(1 for i in self.queue_items if i["status"] == "concluido")
            pending = total - done
            self.progress_label.configure(text=f"{total} / {total}")
            self.progress_target = 1.0
            self.current_item_label.configure(text="Lote concluído! 🎉")
            self.next_up_label.configure(
                text=f"{done} concluído(s)" + (f"  •  {pending} pendente(s)" if pending else "")
            )
            self._set_status("Fila finalizada.", "ok")
            return

        item = self.queue_items[self.current_index]
        self.progress_label.configure(text=f"{self.current_index + 1:02d} / {total:02d}")
        self.progress_target = self.current_index / total
        self.current_item_label.configure(text=item["folder_name"])

        if self.current_index + 1 < total:
            nxt = self.queue_items[self.current_index + 1]["folder_name"]
            self.next_up_label.configure(text=shorten(f"Próxima: {nxt}", 70))
        else:
            self.next_up_label.configure(text="Última pasta da fila")

    def _copy_current_name(self):
        if self.is_running_queue and self.current_index < len(self.queue_items):
            name = self.queue_items[self.current_index]["folder_name"]
            self.clipboard_clear()
            self.clipboard_append(name)
            self._set_status(f"Copiado: {name}", "ok")

    def _open_current_folder(self):
        if not (self.is_running_queue and self.current_index < len(self.queue_items)):
            return
        folder = self.queue_items[self.current_index]["dest_path"]
        if not folder.exists():
            self._set_status("A pasta de destino não existe mais.", "error")
            return
        try:
            if hasattr(os, "startfile"):
                os.startfile(str(folder))
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as exc:
            self._set_status(f"Não foi possível abrir a pasta: {exc}", "error")

    def _go_back(self):
        if not self.is_running_queue or self.current_index <= 0:
            return
        self.current_index -= 1
        self.queue_items[self.current_index]["status"] = "pendente"
        self._refresh_current_view()
        self._update_buttons_state()
        self._set_status("Voltou para o item anterior.")

    def _skip_item(self):
        if not self.is_running_queue or self.current_index >= len(self.queue_items):
            return
        self.queue_items[self.current_index]["status"] = "pendente"
        self.current_index += 1
        self._refresh_current_view()
        self._update_buttons_state()
        self._set_status("Item marcado como pendente; avançando para o próximo.", "warn")

    def _cut_and_next(self):
        if not self.is_running_queue or self.current_index >= len(self.queue_items):
            return

        selected = [f for f in self.captured_files if f["var"].get()]
        if not selected and not messagebox.askyesno(
            "Nenhum arquivo marcado",
            "Não há arquivos marcados para mover.\nDeseja avançar mesmo assim?",
        ):
            return

        item = self.queue_items[self.current_index]
        dest_dir = item["dest_path"]
        dest_dir.mkdir(parents=True, exist_ok=True)

        moved, failed = 0, 0
        for f in selected:
            src_file = Path(f["path"])
            if not src_file.exists():
                self._drop_row(f)
                continue

            target = dest_dir / src_file.name
            if target.exists():
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                target = dest_dir / f"{src_file.stem}_{stamp}{src_file.suffix}"

            try:
                shutil.move(str(src_file), str(target))
                self.known_files_map.pop(str(src_file.resolve()), None)
                self._drop_row(f)
                moved += 1
            except Exception as exc:
                failed += 1
                self._set_status(f"Falha ao mover {src_file.name}: {exc}", "error")

        self._update_empty_state()
        self._update_counter()

        if failed:
            # mantém o item atual para o usuário resolver os arquivos que falharam
            return

        item["status"] = "concluido"
        self.current_index += 1
        self._refresh_current_view()
        self._update_buttons_state()

        self.next_btn.configure(text="✓  Movido!", fg_color=GREEN, hover_color=GREEN)
        self.after(
            450,
            lambda: self.next_btn.configure(
                text="Recortar e próximo  →", fg_color=BLUE_PRIMARY, hover_color=BLUE_HOVER
            ),
        )
        self._set_status(f"{moved} arquivo(s) movido(s) para '{item['folder_name']}'.", "ok")

        if self.current_index >= len(self.queue_items) and not self.cfg_open:
            self._toggle_config_card()

    # ── Animações leves (só alteram cores, sem recriar widgets) ────────────
    def _animate_ui(self):
        for f in self.captured_files:
            hl = f.get("highlight", 0.0)
            if hl > 0.0:
                hl = max(0.0, hl - 0.04)
                f["highlight"] = hl
                frame = f.get("frame")
                if frame is not None:
                    frame.configure(fg_color=lerp_color(BG_CARD, BG_FLASH, hl))

        diff = self.progress_target - self.progress_current
        if abs(diff) > 0.002:
            self.progress_current += diff * 0.22
            self.progress_bar.set(self.progress_current)
        elif self.progress_current != self.progress_target:
            self.progress_current = self.progress_target
            self.progress_bar.set(self.progress_current)

        self.after(40, self._animate_ui)


if __name__ == "__main__":
    app = OrganizzerApp()
    app.mainloop()
