import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import openpyxl

APP_NAME = "organizzer"
STATE_FILE = Path.home() / ".organizzer_state.json"
TEMP_EXTS = {".tmp", ".crdownload", ".part", ".download"}
IGNORED = {"desktop.ini", "thumbs.db"}

BG, CARD, SUBTLE = "#F5F8FD", "#FFFFFF", "#EEF4FF"
LINE, LINE_SOFT = "#C7DAF7", "#E3EAF5"
BLUE, BLUE_DK, BLUE_LT = "#1D4ED8", "#1E40AF", "#3B82F6"
DARK, MUTED, FAINT = "#0F172A", "#64748B", "#94A3B8"
GREEN, AMBER, RED = "#16A34A", "#D97706", "#DC2626"
STATUS = {"info": BLUE_LT, "ok": GREEN, "warn": AMBER, "error": RED}
FAMILY = "Segoe UI"


def resource_path(rel: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel


def F(size=11, weight="normal"):
    return ctk.CTkFont(family=FAMILY, size=size, weight=weight)


def clean(name) -> str:
    s = re.sub(r'[\\/:*?"<>|]+', " ", str(name or "")).strip()
    return re.sub(r"\s+", " ", s)


def short(text: str, n: int) -> str:
    return text if len(text) <= n else text[: n - 1] + "…"


def size_str(n: float) -> str:
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.1f} {u}"
        n /= 1024


def read_renovacao_rows(xl_path: Path):
    """Retorna (locador, inquilino) da aba 'RENOVAÇÃO', colunas C e D, a partir da linha 2."""
    suffix = xl_path.suffix.lower()

    if suffix == ".ods":
        try:
            from odf.opendocument import load as odf_load
            from odf.table import Table, TableCell, TableRow
            from odf.text import P

            doc = odf_load(str(xl_path))
            target_table = None
            for tbl in doc.getElementsByType(Table):
                if (tbl.getAttribute("name") or "").strip().upper() == "RENOVAÇÃO":
                    target_table = tbl
                    break
            if target_table is None:
                raise KeyError("Aba 'RENOVAÇÃO' não encontrada no arquivo .ods.")

            rows_out = []
            row_idx = 0
            for tr in target_table.getElementsByType(TableRow):
                row_repeat = int(tr.getAttribute("numberrowsrepeated") or 1)
                cells = []
                for tc in tr.getElementsByType(TableCell):
                    col_repeat = int(tc.getAttribute("numbercolumnsrepeated") or 1)
                    txt = "".join(
                        "".join(node.data for node in p.childNodes if hasattr(node, "data"))
                        for p in tc.getElementsByType(P)
                    ).strip()
                    for _ in range(min(col_repeat, 20)):
                        cells.append(txt)
                    if len(cells) >= 4:
                        break
                for _ in range(min(row_repeat, 5000)):
                    row_idx += 1
                    if row_idx >= 2:
                        c_val = cells[2] if len(cells) > 2 else ""
                        d_val = cells[3] if len(cells) > 3 else ""
                        rows_out.append((c_val, d_val))
            return rows_out
        except ImportError:
            # Fallback leve lendo content.xml do .ods sem dependências extras
            ns = {
                "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
                "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
            }
            with zipfile.ZipFile(xl_path, "r") as zf:
                xml_bytes = zf.read("content.xml")
            root = ET.fromstring(xml_bytes)
            target_tbl = None
            for tbl in root.findall(".//table:table", ns):
                if (tbl.attrib.get(f"{{{ns['table']}}}name") or "").strip().upper() == "RENOVAÇÃO":
                    target_tbl = tbl
                    break
            if target_tbl is None:
                raise KeyError("Aba 'RENOVAÇÃO' não encontrada no arquivo .ods.")

            rows_out = []
            row_idx = 0
            for tr in target_tbl.findall("table:table-row", ns):
                r_rep = int(tr.attrib.get(f"{{{ns['table']}}}number-rows-repeated") or 1)
                cells = []
                for tc in tr.findall("table:table-cell", ns):
                    c_rep = int(tc.attrib.get(f"{{{ns['table']}}}number-columns-repeated") or 1)
                    parts = []
                    for p in tc.findall(".//text:p", ns):
                        parts.append("".join(p.itertext()))
                    txt = " ".join(parts).strip()
                    for _ in range(min(c_rep, 20)):
                        cells.append(txt)
                    if len(cells) >= 4:
                        break
                for _ in range(min(r_rep, 5000)):
                    row_idx += 1
                    if row_idx >= 2:
                        c_val = cells[2] if len(cells) > 2 else ""
                        d_val = cells[3] if len(cells) > 3 else ""
                        rows_out.append((c_val, d_val))
            return rows_out

    wb = openpyxl.load_workbook(xl_path, data_only=True, read_only=True)
    try:
        sheet_name = next((s for s in wb.sheetnames if s.strip().upper() == "RENOVAÇÃO"), None)
        if not sheet_name:
            raise KeyError("Aba 'RENOVAÇÃO' não encontrada na planilha.")
        ws = wb[sheet_name]
        return [
            (row[0] if len(row) > 0 else "", row[1] if len(row) > 1 else "")
            for row in ws.iter_rows(min_row=2, min_col=3, max_col=4, values_only=True)
        ]
    finally:
        wb.close()


def btn(parent, text, cmd, w=70, h=26, kind="soft", **kw):
    styles = {
        "soft": dict(fg_color=SUBTLE, hover_color=LINE, text_color=BLUE),
        "ghost": dict(
            fg_color=CARD, hover_color=LINE, text_color=BLUE, border_width=1, border_color=LINE
        ),
        "link": dict(fg_color="transparent", hover_color=SUBTLE, text_color=BLUE),
        "primary": dict(fg_color=BLUE, hover_color=BLUE_DK, text_color="#FFFFFF"),
    }[kind]
    return ctk.CTkButton(
        parent,
        text=text,
        command=cmd,
        width=w,
        height=h,
        corner_radius=8,
        font=F(11, "bold"),
        **styles,
        **kw,
    )


class OrganizzerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        self.title(APP_NAME)
        self.geometry("500x540")
        self.minsize(460, 480)
        self.configure(fg_color=BG)

        self.insurer_var = ctk.StringVar()
        self.excel_var = ctk.StringVar()
        self.source_var = ctk.StringVar(value=str(Path.home() / "Downloads"))

        self.queue = []
        self.idx = 0
        self.running = False
        self.cfg_open = True
        self.known = {}
        self.pending = {}
        self.files = []

        self._load_prefs()
        self._build()
        self._refresh()
        self._status("Informe seguradora, planilha e pasta de origem.")
        self.bind("<Control-Return>", lambda _e: self._cut_next())
        self._poll()
        self.after(200, self._apply_icon)

    def _apply_icon(self):
        ico = resource_path("assets/icon.ico")
        if ico.is_file():
            try:
                self.iconbitmap(str(ico))
            except Exception:
                pass

    # ── prefs ──────────────────────────────────────────────────────────────
    def _load_prefs(self):
        try:
            d = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            if os.path.isdir(d.get("last_source_dir", "")):
                self.source_var.set(d["last_source_dir"])
            self.insurer_var.set(d.get("last_insurer", ""))
            if os.path.isfile(d.get("last_excel", "")):
                self.excel_var.set(d["last_excel"])
        except Exception:
            pass

    def _save_prefs(self):
        try:
            STATE_FILE.write_text(
                json.dumps(
                    {
                        "last_source_dir": self.source_var.get().strip(),
                        "last_insurer": self.insurer_var.get().strip(),
                        "last_excel": self.excel_var.get().strip(),
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
        except Exception:
            pass

    # ── interface ──────────────────────────────────────────────────────────
    def _card(self, parent):
        return ctk.CTkFrame(
            parent, fg_color=CARD, corner_radius=12, border_width=1, border_color=LINE_SOFT
        )

    def _entry(self, parent, var, ph):
        return ctk.CTkEntry(
            parent,
            textvariable=var,
            placeholder_text=ph,
            height=28,
            corner_radius=8,
            fg_color=BG,
            border_color=LINE,
            text_color=DARK,
            placeholder_text_color=FAINT,
            font=F(11),
        )

    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        head = self._card(self)
        head.grid(row=0, column=0, padx=10, pady=(10, 6), sticky="ew")
        head.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(head, text=APP_NAME, font=F(15, "bold"), text_color=DARK).grid(
            row=0, column=0, padx=12, pady=(6, 0), sticky="w"
        )
        ctk.CTkLabel(
            head, text="Pastas de renovação • Quiver", font=F(10), text_color=MUTED
        ).grid(row=1, column=0, padx=12, sticky="w")
        self.counter = ctk.CTkLabel(head, text="0 / 0", font=F(12, "bold"), text_color=BLUE)
        self.counter.grid(row=0, column=1, rowspan=2, padx=12)
        self.bar = ctk.CTkProgressBar(
            head, height=4, corner_radius=99, fg_color=SUBTLE, progress_color=BLUE
        )
        self.bar.grid(row=2, column=0, columnspan=2, padx=12, pady=(4, 8), sticky="ew")
        self.bar.set(0)

        self.cfg = self._card(self)
        self.cfg.grid(row=1, column=0, padx=10, pady=(0, 6), sticky="ew")
        self.cfg.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(self.cfg, fg_color="transparent")
        top.grid(row=0, column=0, padx=10, pady=(6, 2), sticky="ew")
        top.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(top, text="LOTE", font=F(10, "bold"), text_color=BLUE).grid(row=0, column=0)
        self.summary = ctk.CTkLabel(top, text="", font=F(10), text_color=MUTED, anchor="w")
        self.summary.grid(row=0, column=1, padx=8, sticky="ew")
        self.toggle_btn = btn(top, "Recolher", self._toggle_cfg, w=64, h=22)
        self.toggle_btn.grid(row=0, column=2)

        self.body = ctk.CTkFrame(self.cfg, fg_color="transparent")
        self.body.grid(row=1, column=0, padx=10, pady=(0, 8), sticky="ew")
        self.body.grid_columnconfigure(1, weight=1)
        rows = [
            ("Seguradora", self.insurer_var, "Ex.: PORTO SEGURO GERAIS", None),
            ("Planilha", self.excel_var, "Aba RENOVAÇÃO (.xlsx, .xlsm, .ods)", self._pick_excel),
            ("Origem", self.source_var, "Pasta de downloads do Quiver", self._pick_source),
        ]
        for r, (lab, var, ph, cmd) in enumerate(rows):
            ctk.CTkLabel(self.body, text=lab, font=F(11, "bold"), text_color=DARK).grid(
                row=r, column=0, padx=(0, 8), pady=2, sticky="w"
            )
            self._entry(self.body, var, ph).grid(
                row=r,
                column=1,
                columnspan=1 if cmd else 2,
                padx=(0, 6) if cmd else 0,
                pady=2,
                sticky="ew",
            )
            if cmd:
                btn(self.body, "...", cmd, w=34, h=28).grid(row=r, column=2, pady=2)
        btn(self.body, "Ler planilha e iniciar fila", self._start, h=30, kind="primary").grid(
            row=3, column=0, columnspan=3, pady=(6, 0), sticky="ew"
        )

        work = self._card(self)
        work.grid(row=2, column=0, padx=10, pady=(0, 4), sticky="nsew")
        work.grid_columnconfigure(0, weight=1)
        work.grid_rowconfigure(2, weight=1)

        ban = ctk.CTkFrame(work, fg_color=SUBTLE, corner_radius=10, border_width=1, border_color=LINE)
        ban.grid(row=0, column=0, padx=8, pady=(8, 4), sticky="ew")
        ban.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(ban, text="DESTINO ATUAL", font=F(9, "bold"), text_color=BLUE).grid(
            row=0, column=0, padx=10, pady=(5, 0), sticky="w"
        )
        self.cur_label = ctk.CTkLabel(
            ban,
            text="Nenhum lote iniciado",
            font=F(12, "bold"),
            text_color=DARK,
            anchor="w",
            justify="left",
            wraplength=280,
        )
        self.cur_label.grid(row=1, column=0, padx=10, sticky="w")
        self.next_label = ctk.CTkLabel(ban, text="", font=F(10), text_color=MUTED, anchor="w")
        self.next_label.grid(row=2, column=0, padx=10, pady=(0, 5), sticky="w")
        bb = ctk.CTkFrame(ban, fg_color="transparent")
        bb.grid(row=0, column=1, rowspan=3, padx=8)
        self.copy_btn = btn(bb, "Copiar", self._copy_name, w=56, h=24, kind="ghost")
        self.copy_btn.pack(pady=(0, 3))
        self.open_btn = btn(bb, "Abrir", self._open_folder, w=56, h=24, kind="ghost")
        self.open_btn.pack()
        ban.bind(
            "<Configure>", lambda e: self.cur_label.configure(wraplength=max(150, e.width - 100))
        )

        fh = ctk.CTkFrame(work, fg_color="transparent")
        fh.grid(row=1, column=0, padx=10, sticky="ew")
        fh.grid_columnconfigure(0, weight=1)
        self.files_title = ctk.CTkLabel(fh, text="ARQUIVOS", font=F(10, "bold"), text_color=MUTED)
        self.files_title.grid(row=0, column=0, sticky="w")
        btn(fh, "Todos", lambda: self._check_all(True), w=44, h=20, kind="link").grid(row=0, column=1)
        btn(fh, "Nenhum", lambda: self._check_all(False), w=50, h=20, kind="link").grid(
            row=0, column=2
        )
        btn(fh, "Limpar", self._clear, w=46, h=20, kind="link").grid(row=0, column=3)

        self.scroll = ctk.CTkScrollableFrame(
            work, fg_color=BG, corner_radius=10, border_width=1, border_color=LINE_SOFT
        )
        self.scroll.grid(row=2, column=0, padx=8, pady=(2, 6), sticky="nsew")
        self.scroll.grid_columnconfigure(0, weight=1)
        self.empty = ctk.CTkLabel(
            self.scroll,
            text="Aguardando arquivos na pasta de origem...",
            font=F(11),
            text_color=FAINT,
        )
        self.empty.pack(pady=24)

        act = ctk.CTkFrame(work, fg_color="transparent")
        act.grid(row=3, column=0, padx=8, pady=(0, 4), sticky="ew")
        act.grid_columnconfigure(2, weight=1)
        self.back_btn = btn(act, "←", self._back, w=36, h=34)
        self.back_btn.grid(row=0, column=0, padx=(0, 6))
        self.skip_btn = btn(act, "Pular", self._skip, w=56, h=34)
        self.skip_btn.grid(row=0, column=1, padx=(0, 6))
        self.next_btn = btn(act, "Recortar e próximo  →", self._cut_next, h=34, kind="primary")
        self.next_btn.configure(font=F(12, "bold"))
        self.next_btn.grid(row=0, column=2, sticky="ew")

        st = ctk.CTkFrame(work, fg_color="transparent")
        st.grid(row=4, column=0, padx=10, pady=(0, 6), sticky="ew")
        self.dot = ctk.CTkLabel(st, text="●", font=F(9), text_color=BLUE_LT, width=12)
        self.dot.pack(side="left")
        self.status_lbl = ctk.CTkLabel(
            st,
            text="",
            font=F(10),
            text_color=MUTED,
            anchor="w",
            justify="left",
            wraplength=420,
        )
        self.status_lbl.pack(side="left", padx=4, fill="x", expand=True)

        ctk.CTkLabel(
            self,
            text="Criado por Matheus Carvalho • Direitos reservados à Totumseg",
            font=F(9),
            text_color=FAINT,
        ).grid(row=3, column=0, pady=(0, 4))

    def _status(self, text, kind="info"):
        self.dot.configure(text_color=STATUS.get(kind, BLUE_LT))
        self.status_lbl.configure(text=text)

    def _toggle_cfg(self):
        if self.cfg_open:
            self.body.grid_remove()
            self.toggle_btn.configure(text="Expandir")
            name = Path(self.excel_var.get().strip()).name
            self.summary.configure(
                text=short(" • ".join(p for p in (self.insurer_var.get().strip(), name) if p), 48)
            )
        else:
            self.body.grid()
            self.toggle_btn.configure(text="Recolher")
            self.summary.configure(text="")
        self.cfg_open = not self.cfg_open

    def _pick_excel(self):
        p = filedialog.askopenfilename(
            title="Planilha de renovação",
            filetypes=[("Planilhas (Excel / LibreOffice)", "*.xlsx *.xlsm *.ods")],
        )
        if p:
            self.excel_var.set(p)

    def _pick_source(self):
        cur = self.source_var.get().strip()
        d = filedialog.askdirectory(
            title="Pasta de downloads",
            initialdir=cur if os.path.isdir(cur) else str(Path.home()),
        )
        if d:
            self.source_var.set(d)
            self._snapshot()
            self._save_prefs()

    # ── monitoramento da pasta ─────────────────────────────────────────────
    def _scan(self):
        src = Path(self.source_var.get().strip())
        out = []
        if not src.is_dir():
            return out
        try:
            for e in src.iterdir():
                try:
                    n = e.name.lower()
                    if (
                        e.is_file()
                        and not n.startswith((".", "~$"))
                        and n not in IGNORED
                        and e.suffix.lower() not in TEMP_EXTS
                    ):
                        s = e.stat()
                        out.append((s.st_mtime, s.st_size, e))
                except OSError:
                    continue
        except OSError:
            return []
        return sorted(out, key=lambda t: t[0])

    def _snapshot(self):
        self.known = {str(e.resolve()): (m, s) for m, s, e in self._scan()}
        self.pending.clear()

    def _poll(self):
        have = {f["path"] for f in self.files}
        for mtime, size, e in self._scan():
            p, sig = str(e.resolve()), (mtime, size)
            if self.known.get(p) == sig:
                continue
            if self.pending.get(p) != sig:
                self.pending[p] = sig
                continue
            self.pending.pop(p, None)
            self.known[p] = sig
            if p not in have:
                info = {"path": p, "name": e.name, "size": size, "var": ctk.BooleanVar(value=True)}
                self.files.append(info)
                self._add_row(info)
                have.add(p)
        self.after(1200, self._poll)

    # ── lista de arquivos ──────────────────────────────────────────────────
    def _add_row(self, info):
        self.empty.pack_forget()
        row = ctk.CTkFrame(
            self.scroll, fg_color=CARD, corner_radius=8, border_width=1, border_color=LINE_SOFT
        )
        row.pack(fill="x", padx=2, pady=2)
        row.grid_columnconfigure(0, weight=1)
        ctk.CTkCheckBox(
            row,
            text=short(info["name"], 36),
            variable=info["var"],
            command=self._count,
            checkbox_width=16,
            checkbox_height=16,
            corner_radius=4,
            fg_color=BLUE,
            hover_color=BLUE_DK,
            border_color=BLUE_LT,
            text_color=DARK,
            font=F(11),
        ).grid(row=0, column=0, padx=8, pady=4, sticky="w")
        ctk.CTkLabel(row, text=size_str(info["size"]), font=F(9), text_color=FAINT).grid(
            row=0, column=1, padx=4
        )
        btn(
            row,
            "✕",
            lambda i=info: self._remove(i),
            w=22,
            h=20,
            kind="link",
            text_color=RED,
            hover_color="#FEE2E2",
        ).grid(row=0, column=2, padx=(0, 6))
        info["row"] = row
        self._count()

    def _drop(self, info):
        row = info.pop("row", None)
        if row is not None:
            row.destroy()
        if info in self.files:
            self.files.remove(info)

    def _remove(self, info):
        self._drop(info)
        self._after_change()

    def _clear(self):
        for f in list(self.files):
            self._drop(f)
        self._after_change()

    def _check_all(self, v):
        for f in self.files:
            f["var"].set(v)
        self._count()

    def _after_change(self):
        if not self.files:
            self.empty.pack(pady=24)
        self._count()

    def _count(self):
        n = len(self.files)
        m = sum(1 for f in self.files if f["var"].get())
        self.files_title.configure(text=f"ARQUIVOS  {m}/{n}" if n else "ARQUIVOS")

    # ── fila ───────────────────────────────────────────────────────────────
    def _start(self):
        insurer = clean(self.insurer_var.get())
        xl = Path(self.excel_var.get().strip())
        if not insurer:
            return messagebox.showwarning("Atenção", "Informe a seguradora.")
        if not xl.is_file():
            return messagebox.showwarning("Atenção", "Selecione uma planilha válida.")
        if not Path(self.source_var.get().strip()).is_dir():
            return messagebox.showwarning("Atenção", "Escolha uma pasta de origem válida.")

        self._status("Lendo planilha...")
        self.update_idletasks()
        try:
            raw_rows = read_renovacao_rows(xl)
        except KeyError as exc:
            self._status("Aba 'RENOVAÇÃO' não encontrada.", "error")
            return messagebox.showerror("Aba não encontrada", str(exc))
        except Exception as exc:
            self._status("Falha ao abrir a planilha.", "error")
            return messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{exc}")

        base, items, seen = xl.parent, [], set()
        for c_val, d_val in raw_rows:
            loc, inq = clean(c_val), clean(d_val)
            if not loc and not inq:
                continue
            name = short(
                clean(f"{insurer} - {' x '.join(p for p in (loc, inq) if p)}"), 150
            ).rstrip(". ")
            path = base / name
            if name not in seen:
                try:
                    path.mkdir(parents=True, exist_ok=True)
                except OSError as exc:
                    self._status("Falha ao criar pastas.", "error")
                    return messagebox.showerror("Erro ao criar pasta", f"{name}\n\n{exc}")
                seen.add(name)
            items.append({"name": name, "path": path, "status": "pendente"})

        if not items:
            self._status("Nenhum registro nas colunas C e D.", "warn")
            return messagebox.showinfo(
                "Sem registros", "Nada encontrado a partir da linha 2 (colunas C e D)."
            )

        self.queue, self.idx, self.running = items, 0, True
        self._clear()
        self._snapshot()
        self._save_prefs()
        if self.cfg_open:
            self._toggle_cfg()
        self._refresh()
        self._status(f"{len(items)} pastas prontas em: {base}", "ok")

    def _refresh(self):
        total = len(self.queue)
        active = self.running and self.idx < total
        for b in (self.next_btn, self.skip_btn, self.copy_btn, self.open_btn):
            b.configure(state="normal" if active else "disabled")
        self.back_btn.configure(state="normal" if self.running and self.idx > 0 else "disabled")

        if not self.running or not total:
            self.counter.configure(text="0 / 0")
            self.bar.set(0)
            self.cur_label.configure(text="Nenhum lote iniciado")
            self.next_label.configure(text="")
        elif self.idx >= total:
            done = sum(1 for i in self.queue if i["status"] == "concluido")
            self.counter.configure(text=f"{total} / {total}")
            self.bar.set(1)
            self.cur_label.configure(text="Lote concluído!")
            self.next_label.configure(text=f"{done} concluído(s), {total - done} pendente(s)")
            self._status("Fila finalizada.", "ok")
        else:
            self.counter.configure(text=f"{self.idx + 1:02d} / {total:02d}")
            self.bar.set(self.idx / total)
            self.cur_label.configure(text=self.queue[self.idx]["name"])
            nxt = self.queue[self.idx + 1]["name"] if self.idx + 1 < total else None
            self.next_label.configure(text=short(f"Próxima: {nxt}", 60) if nxt else "Última da fila")

    def _current(self):
        return self.queue[self.idx] if self.running and self.idx < len(self.queue) else None

    def _copy_name(self):
        if item := self._current():
            self.clipboard_clear()
            self.clipboard_append(item["name"])
            self._status(f"Copiado: {short(item['name'], 60)}", "ok")

    def _open_folder(self):
        item = self._current()
        if not item:
            return
        try:
            if hasattr(os, "startfile"):
                os.startfile(str(item["path"]))
            else:
                subprocess.Popen(["xdg-open", str(item["path"])])
        except Exception as exc:
            self._status(f"Não foi possível abrir a pasta: {exc}", "error")

    def _back(self):
        if self.running and self.idx > 0:
            self.idx -= 1
            self.queue[self.idx]["status"] = "pendente"
            self._refresh()
            self._status("Voltou ao item anterior.")

    def _skip(self):
        if self._current():
            self.queue[self.idx]["status"] = "pendente"
            self.idx += 1
            self._refresh()
            self._status("Item deixado como pendente.", "warn")

    def _cut_next(self):
        item = self._current()
        if not item:
            return
        selected = [f for f in self.files if f["var"].get()]
        if not selected and not messagebox.askyesno(
            "Nenhum arquivo marcado", "Não há arquivos marcados.\nAvançar mesmo assim?"
        ):
            return

        item["path"].mkdir(parents=True, exist_ok=True)
        moved = failed = 0
        for f in selected:
            src = Path(f["path"])
            if not src.exists():
                self._drop(f)
                continue
            dst = item["path"] / src.name
            if dst.exists():
                dst = item["path"] / f"{src.stem}_{datetime.now():%Y%m%d_%H%M%S}{src.suffix}"
            try:
                shutil.move(str(src), str(dst))
                self.known.pop(str(src.resolve()), None)
                self._drop(f)
                moved += 1
            except Exception as exc:
                failed += 1
                self._status(f"Falha ao mover {src.name}: {exc}", "error")
        self._after_change()
        if failed:
            return

        item["status"] = "concluido"
        self.idx += 1
        self._refresh()
        self.next_btn.configure(text="✓  Movido!", fg_color=GREEN, hover_color=GREEN)
        self.after(
            400,
            lambda: self.next_btn.configure(
                text="Recortar e próximo  →", fg_color=BLUE, hover_color=BLUE_DK
            ),
        )
        self._status(f"{moved} arquivo(s) movido(s) para '{short(item['name'], 50)}'.", "ok")
        if self.idx >= len(self.queue) and not self.cfg_open:
            self._toggle_cfg()


if __name__ == "__main__":
    OrganizzerApp().mainloop()
