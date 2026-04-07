import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import requests
import json
import csv
import urllib3
import os
import base64
from datetime import datetime

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ─────────────────────────────────────────────
#  CONFIGURAÇÕES PADRÃO
# ─────────────────────────────────────────────
DEFAULT_BASE_URL = "https://server:50000/b1s/v1"
DEFAULT_BPLID    = "60"
DEFAULT_LOTE     = "100"
CONFIG_FILE      = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# ─────────────────────────────────────────────
#  PERSISTÊNCIA DE CONFIGURAÇÕES
# ─────────────────────────────────────────────
def _encode(value):
    """Ofusca senha em base64 (não é criptografia — apenas evita texto plano óbvio)."""
    return base64.b64encode(value.encode()).decode()

def _decode(value):
    try:
        return base64.b64decode(value.encode()).decode()
    except Exception:
        return value

def load_config():
    defaults = {
        "base_url": DEFAULT_BASE_URL,
        "company_db": "",
        "username": "",
        "password": "",
        "bpl_id": DEFAULT_BPLID,
        "lote_size": DEFAULT_LOTE,
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if "password" in saved:
                saved["password"] = _decode(saved["password"])
            defaults.update(saved)
        except Exception:
            pass
    return defaults

def save_config(data):
    to_save = dict(data)
    if "password" in to_save:
        to_save["password"] = _encode(to_save["password"])
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(to_save, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar config: {e}")

# ─────────────────────────────────────────────
#  LÓGICA SAP
# ─────────────────────────────────────────────
class SAPClient:
    def __init__(self, base_url):
        self.base_url   = base_url.rstrip("/")
        self.session_id = None
        self.session    = requests.Session()
        self.session.verify = False

    def login(self, company_db, username, password):
        payload = {"CompanyDB": company_db, "UserName": username, "Password": password}
        r = self.session.post(f"{self.base_url}/Login", json=payload, timeout=30)
        r.raise_for_status()
        self.session_id = r.json()["SessionId"]
        self.session.headers.update({"Cookie": f"B1SESSION={self.session_id}"})
        return self.session_id

    def logout(self):
        try:
            self.session.post(f"{self.base_url}/Logout", timeout=10)
        except Exception:
            pass

    def baixar_ativos(self, assets, bpl_id, doc_date, posting_date, value_date):
        payload = {
            "DocumentDate":  doc_date,
            "PostingDate":   posting_date,
            "AssetValueDate": value_date,
            "BPLId": int(bpl_id),
            "AssetDocumentLineCollection": [
                {"AssetNumber": a, "Quantity": 1, "TotalLC": 0}
                for a in assets
            ]
        }
        r = self.session.post(
            f"{self.base_url}/AssetRetirement",
            json=payload,
            timeout=600
        )
        return r


# ─────────────────────────────────────────────
#  INTERFACE GRÁFICA
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SAP B1 — Baixa de Ativos em Lote")
        self.resizable(True, True)
        self.minsize(900, 600)
        self.geometry("1200x750")
        self.configure(bg="#1a1a2e")
        self.config_data = load_config()
        self._build_ui()
        self._load_config_to_ui()
        self.csv_path   = None
        self.sap_client = None
        self.running    = False

    # ── Construção da UI ──────────────────────
    def _build_ui(self):
        # Paleta
        BG       = "#1a1a2e"
        PANEL    = "#16213e"
        ACCENT   = "#0f3460"
        BLUE     = "#4cc9f0"
        GREEN    = "#06d6a0"
        RED      = "#ef233c"
        YELLOW   = "#ffd166"
        FG       = "#e0e0e0"
        FG2      = "#a0a0c0"
        FONT     = ("Consolas", 10)
        FONT_LBL = ("Consolas", 9)
        FONT_HDR = ("Consolas", 13, "bold")

        self.BG = BG; self.PANEL = PANEL; self.ACCENT = ACCENT
        self.BLUE = BLUE; self.GREEN = GREEN; self.RED = RED
        self.YELLOW = YELLOW; self.FG = FG; self.FG2 = FG2
        self.FONT = FONT

        # ── Header ──
        hdr = tk.Frame(self, bg=ACCENT, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⬛ SAP B1 — Baixa de Ativos em Lote",
                 font=FONT_HDR, bg=ACCENT, fg=BLUE).pack()
        tk.Label(hdr, text="Service Layer",
                 font=FONT_LBL, bg=ACCENT, fg=FG2).pack()

        # ── Body ──
        body = tk.Frame(self, bg=BG, padx=20, pady=15)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=0)   # esquerda — largura fixa
        body.columnconfigure(1, weight=1)   # direita — expande
        body.rowconfigure(0, weight=1)

        # Coluna esquerda — configurações
        left = tk.Frame(body, bg=BG)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 20))

        def section(parent, title):
            f = tk.LabelFrame(parent, text=f"  {title}  ",
                              bg=PANEL, fg=BLUE, font=FONT_LBL,
                              bd=1, relief="flat",
                              labelanchor="nw", pady=8, padx=10)
            f.pack(fill="x", pady=(0, 10))
            return f

        def field(parent, label, default="", show=None):
            tk.Label(parent, text=label, bg=PANEL, fg=FG2,
                     font=FONT_LBL, anchor="w").pack(fill="x")
            var = tk.StringVar(value=default)
            e = tk.Entry(parent, textvariable=var, bg=ACCENT, fg=FG,
                         insertbackground=BLUE, font=FONT,
                         relief="flat", bd=4, show=show or "")
            e.pack(fill="x", pady=(0, 6))
            return var

        # Conexão
        sec_conn = section(left, "🔌 Conexão SAP")
        self.var_url  = field(sec_conn, "Service Layer", "")
        self.var_db   = field(sec_conn, "Company DB", "")
        self.var_user = field(sec_conn, "Usuário", "")
        self.var_pass = field(sec_conn, "Senha", show="*")

        # Parâmetros
        sec_par = section(left, "⚙️  Parâmetros")
        self.var_bpl  = field(sec_par, "BPLId (Filial)", "")
        self.var_lote = field(sec_par, "Tamanho do lote", "")

        # Datas
        today = datetime.today().strftime("%Y-%m-%d")
        sec_dt = section(left, "📅 Datas")
        self.var_docdate  = field(sec_dt, "Document Date", today)
        self.var_postdate = field(sec_dt, "Posting Date",  today)
        self.var_valdate  = field(sec_dt, "Value Date",    today)

        # CSV
        sec_csv = section(left, "📂 Arquivo CSV")
        self.lbl_csv = tk.Label(sec_csv, text="Nenhum arquivo selecionado",
                                bg=PANEL, fg=YELLOW, font=FONT_LBL,
                                wraplength=300, anchor="w")
        self.lbl_csv.pack(fill="x", pady=(0, 6))
        tk.Button(sec_csv, text="Selecionar CSV",
                  command=self._select_csv,
                  bg=ACCENT, fg=BLUE, font=FONT,
                  relief="flat", cursor="hand2", pady=4).pack(fill="x")

        # Botões de ação
        sec_act = tk.Frame(left, bg=BG)
        sec_act.pack(fill="x", pady=(5, 0))

        tk.Button(sec_act, text="💾  Salvar Configurações",
                  command=self._save_config,
                  bg=ACCENT, fg=YELLOW,
                  font=("Consolas", 10, "bold"),
                  relief="flat", cursor="hand2", pady=6).pack(fill="x", pady=(0, 8))

        self.btn_run = tk.Button(sec_act, text="▶  EXECUTAR",
                                 command=self._run,
                                 bg=GREEN, fg="#0a0a0a",
                                 font=("Consolas", 11, "bold"),
                                 relief="flat", cursor="hand2", pady=8)
        self.btn_run.pack(fill="x", pady=(0, 6))

        self.btn_stop = tk.Button(sec_act, text="⏹  PARAR",
                                  command=self._stop,
                                  bg=RED, fg="white",
                                  font=("Consolas", 11, "bold"),
                                  relief="flat", cursor="hand2", pady=8,
                                  state="disabled")
        self.btn_stop.pack(fill="x")

        # Coluna direita — log + progresso
        right = tk.Frame(body, bg=BG)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        # Progresso
        sec_prog = section(right, "📊 Progresso")
        prog_row = tk.Frame(sec_prog, bg=PANEL)
        prog_row.pack(fill="x", pady=(0, 6))

        self.var_total     = tk.StringVar(value="—")
        self.var_lotes_tot = tk.StringVar(value="—")
        self.var_ok        = tk.StringVar(value="0")
        self.var_err       = tk.StringVar(value="0")

        def stat(parent, label, var, color):
            f = tk.Frame(parent, bg=PANEL, padx=8)
            f.pack(side="left", expand=True)
            tk.Label(f, text=label, bg=PANEL, fg=FG2, font=FONT_LBL).pack()
            tk.Label(f, textvariable=var, bg=PANEL, fg=color,
                     font=("Consolas", 18, "bold")).pack()

        stat(prog_row, "Total",    self.var_total,     BLUE)
        stat(prog_row, "Lotes",    self.var_lotes_tot, YELLOW)
        stat(prog_row, "✔ OK",     self.var_ok,        GREEN)
        stat(prog_row, "✘ Erros",  self.var_err,       RED)

        self.progressbar = ttk.Progressbar(sec_prog, mode="determinate",
                                           length=400)
        self.progressbar.pack(fill="x", pady=(4, 0))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TProgressbar",
                        troughcolor=ACCENT, background=GREEN, thickness=12)

        # Log
        sec_log = section(right, "📋 Log de Execução")
        sec_log.pack(fill="both", expand=True)
        self.log = scrolledtext.ScrolledText(
            sec_log, width=60, height=28,
            bg="#0d0d1a", fg=FG, font=("Consolas", 9),
            relief="flat", bd=0,
            insertbackground=BLUE
        )
        self.log.pack(fill="both", expand=True)
        self.log.tag_config("ok",    foreground=GREEN)
        self.log.tag_config("erro",  foreground=RED)
        self.log.tag_config("info",  foreground=BLUE)
        self.log.tag_config("warn",  foreground=YELLOW)
        self.log.tag_config("plain", foreground=FG)

        # Botão exportar log
        tk.Button(right, text="💾  Exportar Log",
                  command=self._export_log,
                  bg=ACCENT, fg=FG2, font=FONT_LBL,
                  relief="flat", cursor="hand2", pady=4).pack(fill="x", pady=(6, 0))

    # ── Config ────────────────────────────────
    def _load_config_to_ui(self):
        c = self.config_data
        self.var_url.set(c.get("base_url",   DEFAULT_BASE_URL))
        self.var_db.set(c.get("company_db",  ""))
        self.var_user.set(c.get("username",  ""))
        self.var_pass.set(c.get("password",  ""))
        self.var_bpl.set(c.get("bpl_id",     DEFAULT_BPLID))
        self.var_lote.set(c.get("lote_size", DEFAULT_LOTE))

    def _save_config(self):
        save_config({
            "base_url":   self.var_url.get(),
            "company_db": self.var_db.get(),
            "username":   self.var_user.get(),
            "password":   self.var_pass.get(),
            "bpl_id":     self.var_bpl.get(),
            "lote_size":  self.var_lote.get(),
        })
        self._log(f"✔ Configurações salvas em {CONFIG_FILE}", "ok")

    # ── Helpers de UI ─────────────────────────
    def _log(self, msg, tag="plain"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.insert("end", f"[{ts}] {msg}\n", tag)
        self.log.see("end")

    def _select_csv(self):
        path = filedialog.askopenfilename(
            title="Selecionar CSV de ativos",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.csv_path = path
            self.lbl_csv.config(text=os.path.basename(path))
            self._log(f"CSV selecionado: {os.path.basename(path)}", "info")

    def _export_log(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
            initialfile=f"log_baixa_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.log.get("1.0", "end"))
            self._log(f"Log exportado: {path}", "info")

    def _set_running(self, state):
        self.running = state
        self.btn_run.config(state="disabled" if state else "normal")
        self.btn_stop.config(state="normal" if state else "disabled")

    def _stop(self):
        self.running = False
        self._log("⏹ Interrompido pelo usuário.", "warn")

    # ── Leitura CSV ───────────────────────────
    def _read_csv(self):
        assets = []
        with open(self.csv_path, newline="", encoding="utf-8-sig") as f:
            # detecta separador
            sample = f.read(1024)
            f.seek(0)
            sep = ";" if sample.count(";") > sample.count(",") else ","
            reader = csv.DictReader(f, delimiter=sep)
            for row in reader:
                code = (row.get("AssetNumber") or "").strip().strip('"').strip(";")
                if code:
                    assets.append(code)
        return assets

    # ── Execução principal ────────────────────
    def _run(self):
        # Validações básicas
        if not self.csv_path:
            messagebox.showwarning("Atenção", "Selecione um arquivo CSV.")
            return
        if not all([self.var_db.get(), self.var_user.get(), self.var_pass.get()]):
            messagebox.showwarning("Atenção", "Preencha Company DB, Usuário e Senha.")
            return

        self._set_running(True)
        self._save_config()
        threading.Thread(target=self._execute, daemon=True).start()

    def _execute(self):
        try:
            # 1. Leitura CSV
            self._log("📂 Lendo CSV...", "info")
            assets = self._read_csv()
            total = len(assets)
            if total == 0:
                self._log("CSV vazio ou sem coluna AssetNumber.", "erro")
                self._set_running(False)
                return

            lote_size = int(self.var_lote.get() or DEFAULT_LOTE)
            lotes = [assets[i:i+lote_size] for i in range(0, total, lote_size)]
            n_lotes = len(lotes)

            self.var_total.set(str(total))
            self.var_lotes_tot.set(str(n_lotes))
            self.progressbar["maximum"] = n_lotes
            self._log(f"✔ {total} ativos | {n_lotes} lotes de até {lote_size}", "info")

            # 2. Login
            self._log("🔐 Conectando ao SAP...", "info")
            self.sap_client = SAPClient(self.var_url.get())
            self.sap_client.login(
                self.var_db.get(),
                self.var_user.get(),
                self.var_pass.get()
            )
            self._log("✔ Login efetuado com sucesso.", "ok")

            ok_count  = 0
            err_count = 0

            # 3. Loop de lotes
            for idx, lote in enumerate(lotes, start=1):
                if not self.running:
                    break

                self._log(f"── Lote {idx}/{n_lotes} ({len(lote)} ativos)...", "plain")

                try:
                    r = self.sap_client.baixar_ativos(
                        lote,
                        self.var_bpl.get(),
                        self.var_docdate.get(),
                        self.var_postdate.get(),
                        self.var_valdate.get()
                    )

                    if r.status_code in (200, 201, 202):
                        doc_entry = r.json().get("DocEntry", "?")
                        self._log(
                            f"   ✔ OK — DocEntry: {doc_entry} | "
                            f"{lote[0]} → {lote[-1]}",
                            "ok"
                        )
                        ok_count += len(lote)
                    else:
                        err_msg = ""
                        try:
                            err_msg = r.json()["error"]["message"]["value"]
                        except Exception:
                            err_msg = r.text[:200]
                        self._log(
                            f"   ✘ ERRO lote {idx}: {err_msg}",
                            "erro"
                        )
                        self._log(
                            f"     Ativos: {lote[0]} → {lote[-1]}",
                            "warn"
                        )
                        err_count += len(lote)

                except Exception as e:
                    self._log(f"   ✘ Exceção lote {idx}: {e}", "erro")
                    err_count += len(lote)

                self.var_ok.set(str(ok_count))
                self.var_err.set(str(err_count))
                self.progressbar["value"] = idx

            # 4. Logout
            self.sap_client.logout()
            self._log("🔒 Sessão encerrada.", "info")

            # 5. Resumo
            self._log("─" * 50, "plain")
            self._log(
                f"CONCLUÍDO — ✔ {ok_count} ativos baixados | "
                f"✘ {err_count} com erro",
                "ok" if err_count == 0 else "warn"
            )

        except Exception as e:
            self._log(f"✘ Erro crítico: {e}", "erro")

        finally:
            self._set_running(False)


# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
