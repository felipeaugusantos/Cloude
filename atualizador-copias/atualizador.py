"""
Atualizador de Copias - Autpas
==============================

Ferramenta interna para o time de teste. Centraliza numa unica tela a
atualizacao das copias (pasta Autpas de cada tester) e a abertura do Autcom,
eliminando a necessidade de entrar pasta por pasta para rodar os .bat.

Fluxo:
  - O Jenkins compila a ultima versao em uma pasta de saida.
  - Cada tester tem uma pasta C:\\<usuario>\\Autpas (conteudo igual para todos).
  - Um .bat atualiza a Autpas com a ultima compilacao.
  - O autcom.exe fica dentro da propria Autpas.

Este programa apenas orquestra: roda os .bat de varios usuarios de uma vez e
abre o autcom.exe do usuario escolhido. Toda a configuracao de caminhos fica
no arquivo config.json (veja config.example.json).

Empacotamento (gera um unico .exe):
    pip install pyinstaller
    pyinstaller --onefile --noconsole --name AtualizadorCopias atualizador.py
"""

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox


APP_TITLE = "Atualizador de Copias - Autpas"
CONFIG_NAME = "config.json"


def base_dir():
    """Pasta onde o programa (ou .exe empacotado) esta rodando."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_config():
    """Le config.json ao lado do executavel. Retorna dict de config."""
    path = os.path.join(base_dir(), CONFIG_NAME)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Arquivo de configuracao nao encontrado:\n{path}\n\n"
            "Copie config.example.json para config.json e ajuste os caminhos."
        )
    with open(path, "r", encoding="utf-8") as fh:
        cfg = json.load(fh)

    cfg.setdefault("autcom_exe", "autcom.exe")
    cfg.setdefault("bat_nome", "atualizar.bat")
    cfg.setdefault("usuarios", [])

    normalizados = []
    for u in cfg["usuarios"]:
        if not isinstance(u, dict) or not u.get("pasta"):
            continue
        nome = u.get("nome") or os.path.basename(os.path.dirname(u["pasta"].rstrip("\\/")))
        normalizados.append({
            "nome": nome,
            "pasta": u["pasta"],
            # bat e autcom podem ser sobrescritos por usuario; senao usam o padrao
            "bat": u.get("bat"),
            "autcom": u.get("autcom"),
        })
    cfg["usuarios"] = normalizados
    return cfg


class AtualizadorApp:
    STATUS_PENDENTE = "Pendente"
    STATUS_RODANDO = "Atualizando..."
    STATUS_OK = "OK"
    STATUS_ERRO = "Erro"

    def __init__(self, root, cfg):
        self.root = root
        self.cfg = cfg
        self.msg_queue = queue.Queue()
        self.trabalhando = False

        root.title(APP_TITLE)
        root.geometry("820x560")
        root.minsize(700, 460)

        self._montar_ui()
        self._popular_usuarios()
        self.root.after(100, self._processar_fila)

    # ------------------------------------------------------------------ UI
    def _montar_ui(self):
        topo = ttk.Frame(self.root, padding=(12, 10))
        topo.pack(fill="x")

        ttk.Label(topo, text="Atualizador de Copias", font=("Segoe UI", 15, "bold")).pack(anchor="w")
        ttk.Label(
            topo,
            text="Selecione os testers e atualize as copias de uma vez. "
                 "Use 'Abrir Autcom' para iniciar o teste.",
            foreground="#555",
        ).pack(anchor="w")

        # Tabela de usuarios
        corpo = ttk.Frame(self.root, padding=(12, 0))
        corpo.pack(fill="both", expand=True)

        cols = ("usuario", "pasta", "status")
        self.tree = ttk.Treeview(corpo, columns=cols, show="headings", selectmode="extended", height=10)
        self.tree.heading("usuario", text="Tester")
        self.tree.heading("pasta", text="Pasta Autpas")
        self.tree.heading("status", text="Status")
        self.tree.column("usuario", width=150, anchor="w")
        self.tree.column("pasta", width=420, anchor="w")
        self.tree.column("status", width=120, anchor="center")

        self.tree.tag_configure("ok", foreground="#1a7f37")
        self.tree.tag_configure("erro", foreground="#b42318")
        self.tree.tag_configure("rodando", foreground="#9a6700")

        vsb = ttk.Scrollbar(corpo, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Botoes de acao
        acoes = ttk.Frame(self.root, padding=(12, 8))
        acoes.pack(fill="x")

        self.btn_sel_todos = ttk.Button(acoes, text="Selecionar todos", command=self._selecionar_todos)
        self.btn_sel_todos.pack(side="left")

        self.btn_atualizar_sel = ttk.Button(acoes, text="Atualizar selecionados", command=self._atualizar_selecionados)
        self.btn_atualizar_sel.pack(side="left", padx=(8, 0))

        self.btn_atualizar_todos = ttk.Button(acoes, text="Atualizar todos", command=self._atualizar_todos)
        self.btn_atualizar_todos.pack(side="left", padx=(8, 0))

        self.btn_autcom = ttk.Button(acoes, text="Abrir Autcom", command=self._abrir_autcom)
        self.btn_autcom.pack(side="right")

        # Barra de progresso
        prog = ttk.Frame(self.root, padding=(12, 0))
        prog.pack(fill="x")
        self.progress = ttk.Progressbar(prog, mode="determinate")
        self.progress.pack(fill="x")

        # Log
        logf = ttk.LabelFrame(self.root, text="Registro", padding=(8, 6))
        logf.pack(fill="both", expand=True, padx=12, pady=(8, 12))
        self.log = tk.Text(logf, height=8, wrap="word", state="disabled",
                           font=("Consolas", 9), background="#0f172a", foreground="#e2e8f0")
        logsb = ttk.Scrollbar(logf, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=logsb.set)
        self.log.pack(side="left", fill="both", expand=True)
        logsb.pack(side="right", fill="y")

    def _popular_usuarios(self):
        self.iid_por_usuario = {}
        for u in self.cfg["usuarios"]:
            iid = self.tree.insert("", "end", values=(u["nome"], u["pasta"], self.STATUS_PENDENTE))
            self.iid_por_usuario[iid] = u
        if not self.cfg["usuarios"]:
            self._log("Nenhum tester configurado. Edite o config.json e reinicie.")

    # -------------------------------------------------------------- helpers
    def _bat_do_usuario(self, u):
        return u.get("bat") or os.path.join(u["pasta"], self.cfg["bat_nome"])

    def _autcom_do_usuario(self, u):
        return u.get("autcom") or os.path.join(u["pasta"], self.cfg["autcom_exe"])

    def _selecionar_todos(self):
        self.tree.selection_set(self.tree.get_children())

    def _set_status(self, iid, status, tag=""):
        vals = list(self.tree.item(iid, "values"))
        vals[2] = status
        self.tree.item(iid, values=vals, tags=(tag,) if tag else ())

    def _log(self, texto):
        self.log.configure(state="normal")
        self.log.insert("end", texto + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _travar_botoes(self, travar):
        estado = "disabled" if travar else "normal"
        for b in (self.btn_atualizar_sel, self.btn_atualizar_todos, self.btn_autcom, self.btn_sel_todos):
            b.configure(state=estado)

    # ----------------------------------------------------------- acoes UI
    def _atualizar_selecionados(self):
        iids = list(self.tree.selection())
        if not iids:
            messagebox.showinfo(APP_TITLE, "Selecione ao menos um tester na lista.")
            return
        self._iniciar_atualizacao(iids)

    def _atualizar_todos(self):
        self._iniciar_atualizacao(list(self.tree.get_children()))

    def _abrir_autcom(self):
        iids = list(self.tree.selection())
        if len(iids) != 1:
            messagebox.showinfo(APP_TITLE, "Selecione exatamente um tester para abrir o Autcom.")
            return
        u = self.iid_por_usuario[iids[0]]
        exe = self._autcom_do_usuario(u)
        if not os.path.exists(exe):
            messagebox.showerror(APP_TITLE, f"Autcom nao encontrado:\n{exe}")
            self._log(f"[{u['nome']}] Autcom nao encontrado: {exe}")
            return
        try:
            subprocess.Popen([exe], cwd=u["pasta"])
            self._log(f"[{u['nome']}] Autcom iniciado.")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"Falha ao abrir o Autcom:\n{exc}")
            self._log(f"[{u['nome']}] Falha ao abrir Autcom: {exc}")

    # ------------------------------------------------------- atualizacao
    def _iniciar_atualizacao(self, iids):
        if self.trabalhando:
            return
        self.trabalhando = True
        self._travar_botoes(True)
        self.progress.configure(maximum=len(iids), value=0)
        for iid in iids:
            self._set_status(iid, self.STATUS_PENDENTE)
        t = threading.Thread(target=self._worker_atualizar, args=(iids,), daemon=True)
        t.start()

    def _worker_atualizar(self, iids):
        """Executado em thread separada. Comunica com a UI via fila."""
        for iid in iids:
            u = self.iid_por_usuario[iid]
            bat = self._bat_do_usuario(u)
            self.msg_queue.put(("status", iid, self.STATUS_RODANDO, "rodando"))
            self.msg_queue.put(("log", f"[{u['nome']}] Rodando: {bat}"))

            if not os.path.exists(bat):
                self.msg_queue.put(("status", iid, self.STATUS_ERRO, "erro"))
                self.msg_queue.put(("log", f"[{u['nome']}] .bat nao encontrado: {bat}"))
                self.msg_queue.put(("step",))
                continue
            try:
                proc = subprocess.run(
                    ["cmd", "/c", bat],
                    cwd=u["pasta"],
                    capture_output=True,
                    text=True,
                    errors="replace",
                )
                if proc.returncode == 0:
                    self.msg_queue.put(("status", iid, self.STATUS_OK, "ok"))
                    self.msg_queue.put(("log", f"[{u['nome']}] Atualizado com sucesso."))
                else:
                    self.msg_queue.put(("status", iid, self.STATUS_ERRO, "erro"))
                    detalhe = (proc.stderr or proc.stdout or "").strip().splitlines()
                    resumo = detalhe[-1] if detalhe else f"codigo {proc.returncode}"
                    self.msg_queue.put(("log", f"[{u['nome']}] Erro (cod {proc.returncode}): {resumo}"))
            except Exception as exc:  # noqa: BLE001
                self.msg_queue.put(("status", iid, self.STATUS_ERRO, "erro"))
                self.msg_queue.put(("log", f"[{u['nome']}] Falha ao executar: {exc}"))
            self.msg_queue.put(("step",))

        self.msg_queue.put(("fim",))

    def _processar_fila(self):
        """Roda no thread da UI; consome mensagens do worker."""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                tipo = msg[0]
                if tipo == "status":
                    _, iid, status, tag = msg
                    self._set_status(iid, status, tag)
                elif tipo == "log":
                    self._log(msg[1])
                elif tipo == "step":
                    self.progress.configure(value=self.progress["value"] + 1)
                elif tipo == "fim":
                    self.trabalhando = False
                    self._travar_botoes(False)
                    self._log("Concluido.")
        except queue.Empty:
            pass
        self.root.after(100, self._processar_fila)


def main():
    try:
        cfg = load_config()
    except Exception as exc:  # noqa: BLE001
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, str(exc))
        return

    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")  # tema nativo no Windows; ignora se indisponivel
    except tk.TclError:
        pass
    AtualizadorApp(root, cfg)
    root.mainloop()


if __name__ == "__main__":
    main()
