#!/usr/bin/env python3
"""
FolderSync - synchronisation sûre entre un dossier local (C:\\...) et OneDrive.

* Aucune dépendance : Python 3.9+ avec tkinter (inclus avec l'installeur Windows).
* Lancer l'interface :   python foldersync.py
* Mode console / tâche planifiée :
      python foldersync.py --cli            (analyse seulement)
      python foldersync.py --cli --apply    (synchronise)
* Arrière-plan (Windows) : python foldersync.py --install-task --interval 15

Principes anti-perte de données
-------------------------------
1. Aperçu obligatoire : on voit chaque action avant de l'exécuter (et on peut la changer).
2. Jamais d'écrasement ni de suppression directe : l'ancienne version est d'abord
   archivée dans <versions>/<AAAA-MM-JJ_HHMMSS>/<source|destination>/<chemin>.
3. Mémoire du dernier état synchronisé ("baseline") : permet de savoir qui a changé
   quoi, de détecter les vrais conflits et de distinguer "supprimé" de "jamais copié".
4. Conflit (modifié des deux côtés) : les deux versions sont conservées.
5. Copies atomiques (fichier temporaire + remplacement) vérifiées par SHA-256.
6. Garde-fous : refus d'analyser un dossier introuvable/illisible (disque débranché),
   confirmation si beaucoup de suppressions.
7. Journal de chaque opération + purge automatique des vieilles versions.
"""
import argparse
import difflib
import fnmatch
import hashlib
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta
from typing import Optional

APP = "FolderSync"
IS_WIN = os.name == "nt"
MTIME_TOLERANCE = 2.0            # secondes (précision variable selon les systèmes de fichiers)
RECALL_ATTRS = 0x00400000 | 0x00040000   # fichiers OneDrive "en ligne seulement"
SIDE_LABEL = {"s": "source", "d": "destination"}


# ----------------------------------------------------------------------------
# Utilitaires
# ----------------------------------------------------------------------------
def app_dir() -> str:
    base = os.environ.get("APPDATA") if IS_WIN else os.path.join(os.path.expanduser("~"), ".config")
    d = os.path.join(base or os.path.expanduser("~"), APP)
    os.makedirs(d, exist_ok=True)
    return d


def lp(path: str) -> str:
    """Chemin 'long' Windows (\\\\?\\) pour dépasser la limite de 260 caractères."""
    if not IS_WIN or path.startswith("\\\\?\\"):
        return path
    path = os.path.abspath(path)
    if path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def nk(rel: str) -> str:
    """Clé de comparaison : insensible à la casse sous Windows, séparateur '/'."""
    return os.path.normcase(rel).replace("\\", "/") if IS_WIN else rel


def sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(lp(path), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def fmt_size(n: float) -> str:
    for unit in ("o", "Ko", "Mo", "Go", "To"):
        if n < 1024 or unit == "To":
            return f"{n:.0f} {unit}" if unit == "o" else f"{n:.1f} {unit}"
        n /= 1024


def fmt_time(ts: float) -> str:
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")


def open_in_explorer(path: str):
    os.makedirs(path, exist_ok=True)
    if IS_WIN:
        os.startfile(path)  # noqa
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
DEFAULT_EXCLUDES = ["~$*", "*.tmp", "*.fsync-tmp", "Thumbs.db", "desktop.ini", ".DS_Store",
                    "$RECYCLE.BIN", "System Volume Information"]


@dataclass
class Config:
    source: str = ""
    dest: str = ""
    mode: str = "two_way"            # two_way | mirror (source -> destination)
    propagate_deletes: bool = True   # les suppressions sont propagées (toujours archivées)
    always_hash: bool = False        # comparer le contenu (SHA-256) même si taille/date identiques
    skip_online_only: bool = True    # ne pas télécharger les fichiers OneDrive "en ligne seulement"
    verify: bool = True              # vérifier chaque copie par SHA-256
    excludes: list = field(default_factory=lambda: list(DEFAULT_EXCLUDES))
    versions_dir: str = ""
    keep_days: int = 90
    auto_enabled: bool = False       # synchronisation automatique pendant que l'application est ouverte
    interval_min: int = 15           # délai entre deux synchronisations automatiques
    auto_on_start: bool = True       # synchroniser dès l'ouverture de l'application
    start_with_windows: bool = False # ouvrir l'application (réduite) au démarrage de Windows

    @staticmethod
    def path() -> str:
        return os.path.join(app_dir(), "config.json")

    @classmethod
    def load(cls) -> "Config":
        cfg = cls()
        try:
            with open(cls.path(), encoding="utf-8") as f:
                data = json.load(f)
            for k, v in data.items():
                if hasattr(cfg, k):
                    setattr(cfg, k, v)
        except (OSError, ValueError):
            pass
        if not cfg.dest:
            cfg.dest = os.environ.get("OneDrive", "")
        return cfg

    def save(self):
        with open(self.path(), "w", encoding="utf-8") as f:
            json.dump(asdict(self), f, indent=2, ensure_ascii=False)

    def versions_root(self) -> str:
        return self.versions_dir or os.path.join(app_dir(), "versions")

    def state_path(self) -> str:
        ident = hashlib.sha1(f"{os.path.normcase(self.source)}|{os.path.normcase(self.dest)}".encode()).hexdigest()[:10]
        return os.path.join(app_dir(), f"state_{ident}.json")

    def log_path(self) -> str:
        d = os.path.join(app_dir(), "logs")
        os.makedirs(d, exist_ok=True)
        return os.path.join(d, f"sync-{datetime.now():%Y-%m-%d}.log")

    def validate(self):
        if not self.source or not self.dest:
            raise ValueError("Choisissez le dossier source et le dossier destination.")
        a, b = os.path.normcase(os.path.abspath(self.source)), os.path.normcase(os.path.abspath(self.dest))
        if a == b:
            raise ValueError("La source et la destination sont identiques.")
        if b.startswith(a + os.sep) or a.startswith(b + os.sep):
            raise ValueError("Un des dossiers est à l'intérieur de l'autre : risque de boucle.")
        v = os.path.normcase(os.path.abspath(self.versions_root()))
        if v.startswith(a + os.sep) or v.startswith(b + os.sep) or v in (a, b):
            raise ValueError("Le dossier des versions ne doit pas être dans la source ni la destination.")


# ----------------------------------------------------------------------------
# Analyse des dossiers
# ----------------------------------------------------------------------------
@dataclass
class FInfo:
    rel: str
    path: str
    size: int
    mtime: float
    online: bool = False   # placeholder OneDrive (contenu pas sur le disque)


def is_excluded(rel: str, name: str, patterns) -> bool:
    relp = rel.replace("\\", "/")
    return any(fnmatch.fnmatch(name, p) or fnmatch.fnmatch(relp, p) for p in patterns)


def scan(root: str, excludes, progress=None) -> dict:
    base = lp(root)
    if not os.path.isdir(base):
        raise FileNotFoundError(f"Dossier introuvable : {root}")

    def onerr(e):                      # un dossier illisible ne doit JAMAIS passer pour "vide"
        raise e

    out = {}
    for dp, dns, fns in os.walk(base, onerror=onerr):
        rel_dir = os.path.relpath(dp, base)
        rel_dir = "" if rel_dir == "." else rel_dir
        dns[:] = [d for d in dns if not is_excluded(os.path.join(rel_dir, d), d, excludes)]
        for fn in fns:
            rel = os.path.join(rel_dir, fn)
            if is_excluded(rel, fn, excludes):
                continue
            full = os.path.join(dp, fn)
            try:
                st = os.stat(full)
            except FileNotFoundError:
                continue
            attrs = getattr(st, "st_file_attributes", 0)
            out[nk(rel)] = FInfo(rel, full, st.st_size, st.st_mtime, bool(attrs & RECALL_ATTRS))
        if progress:
            progress(len(out))
    return out


# ----------------------------------------------------------------------------
# Plan de synchronisation (comparaison)
# ----------------------------------------------------------------------------
@dataclass
class Action:
    key: str
    rel: str
    op: str              # push | pull | del_d | del_s | conflict | skip
    new: bool            # True = création, False = mise à jour
    s: Optional[FInfo]
    d: Optional[FInfo]
    note: str = ""
    selected: bool = True

    def label(self) -> str:
        return {
            "push": "→ Copier vers destination" if self.new else "→ Mettre à jour la destination",
            "pull": "← Copier vers source" if self.new else "← Mettre à jour la source",
            "del_d": "✕ Supprimer de la destination (archivé)",
            "del_s": "✕ Supprimer de la source (archivé)",
            "conflict": "⚠ Conflit : garder les deux versions",
            "skip": "⏸ Ignoré",
        }[self.op]


@dataclass
class Plan:
    actions: list
    same: list           # clés identiques des deux côtés
    gone: list           # clés connues de la baseline mais disparues des deux côtés
    n_total: int
    _S: dict = field(default_factory=dict)
    _D: dict = field(default_factory=dict)


def load_state(cfg: Config) -> dict:
    try:
        with open(cfg.state_path(), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_state(cfg: Config, state: dict):
    tmp = cfg.state_path() + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f)
    os.replace(tmp, cfg.state_path())


def _changed(info: FInfo, base_entry) -> bool:
    size, mtime = base_entry
    return info.size != size or abs(info.mtime - mtime) > MTIME_TOLERANCE


def _identical(cfg: Config, s: FInfo, d: FInfo):
    """True / False / None (indéterminé : fichier en ligne seulement)."""
    if s.size != d.size:
        return False
    close = abs(s.mtime - d.mtime) <= MTIME_TOLERANCE
    if s.online or d.online:
        return True if close else None
    if close and not cfg.always_hash:
        return True
    return sha256(s.path) == sha256(d.path)


def make_plan(cfg: Config, S: dict, D: dict, state: dict, progress=None) -> Plan:
    actions, same = [], []
    keys = sorted(set(S) | set(D))
    for i, k in enumerate(keys):
        if progress and i % 200 == 0:
            progress(i, len(keys))
        s, d, b = S.get(k), D.get(k), state.get(k)
        rel = (s or d).rel

        def act(op, new, note="", s=s, d=d):
            actions.append(Action(k, rel, op, new, s, d, note))

        if s and d:
            eq = _identical(cfg, s, d)
            if eq:
                same.append(k)
                continue
            if eq is None:
                act("skip", False, "fichier OneDrive 'en ligne seulement' : contenu non comparé")
                continue
            if cfg.mode == "mirror":
                act("push", False, "" if s.mtime >= d.mtime else "la destination était plus récente (archivée)")
            elif b:
                sc, dc = _changed(s, b["s"]), _changed(d, b["d"])
                if sc and dc:
                    act("conflict", False, "modifié des deux côtés depuis la dernière synchro")
                elif sc:
                    act("push", False)
                elif dc:
                    act("pull", False)
                else:
                    act("push" if s.mtime >= d.mtime else "pull", False, "le plus récent gagne")
            else:
                act("push" if s.mtime >= d.mtime else "pull", False,
                    "pas d'historique : le plus récent gagne, l'autre est archivé")
        elif s:       # seulement dans la source
            if cfg.mode == "two_way" and b and cfg.propagate_deletes:
                if _changed(s, b["s"]):
                    act("push", True, "supprimé à destination mais modifié depuis : conservé")
                else:
                    act("del_s", False, "supprimé à destination")
            else:
                act("push", True)
        else:         # seulement dans la destination
            if cfg.mode == "mirror":
                if cfg.propagate_deletes:
                    act("del_d", False, "absent de la source")
            elif b and cfg.propagate_deletes:
                if _changed(d, b["d"]):
                    act("pull", True, "supprimé dans la source mais modifié depuis : conservé")
                else:
                    act("del_d", False, "supprimé dans la source")
            else:
                act("pull", True)
    # Évite de copier depuis un fichier "en ligne seulement" (déclencherait un téléchargement)
    if cfg.skip_online_only:
        for a in actions:
            src = a.s if a.op == "push" else a.d if a.op == "pull" else None
            if src is not None and src.online:
                a.op, a.note = "skip", "fichier OneDrive 'en ligne seulement' (non téléchargé)"
    gone = [k for k in state if k not in S and k not in D]
    return Plan(actions, same, gone, len(keys))


def force_action(a: Action, op: str) -> bool:
    """Change manuellement l'action d'une ligne. Retourne False si impossible."""
    if op == "push" and a.s:
        a.op, a.new = "push", a.d is None
    elif op == "pull" and a.d:
        a.op, a.new = "pull", a.s is None
    elif op == "conflict" and a.s and a.d:
        a.op, a.new = "conflict", False
    elif op == "skip":
        a.op = "skip"
    else:
        return False
    a.note = "choix manuel"
    return True


# ----------------------------------------------------------------------------
# Exécution
# ----------------------------------------------------------------------------
class Syncer:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.vroot = os.path.join(cfg.versions_root(), self.stamp)
        self.root = {"s": cfg.source, "d": cfg.dest}
        self._log = open(cfg.log_path(), "a", encoding="utf-8")
        self.lock_path = None

    def log(self, msg: str):
        self._log.write(f"{datetime.now():%H:%M:%S} {msg}\n")
        self._log.flush()

    # -- primitives ---------------------------------------------------------
    def _archive_path(self, side: str, rel: str) -> str:
        return os.path.join(self.vroot, SIDE_LABEL[side], rel)

    def archive_copy(self, side: str, info: FInfo):
        dst = self._archive_path(side, info.rel)
        os.makedirs(os.path.dirname(lp(dst)), exist_ok=True)
        shutil.copy2(lp(info.path), lp(dst))
        self.log(f"ARCHIVE  {SIDE_LABEL[side]}: {info.rel} -> {dst}")

    def archive_move(self, side: str, info: FInfo):
        dst = self._archive_path(side, info.rel)
        os.makedirs(os.path.dirname(lp(dst)), exist_ok=True)
        shutil.move(lp(info.path), lp(dst))
        self.log(f"SUPPRIMÉ {SIDE_LABEL[side]}: {info.rel} (archivé dans {dst})")

    def copy_verified(self, src: str, dst: str):
        os.makedirs(os.path.dirname(lp(dst)), exist_ok=True)
        tmp = dst + ".fsync-tmp"
        try:
            shutil.copy2(lp(src), lp(tmp))
            if os.path.getsize(lp(tmp)) != os.path.getsize(lp(src)):
                raise IOError("taille différente après copie")
            if self.cfg.verify and sha256(tmp) != sha256(src):
                raise IOError("vérification SHA-256 échouée")
            os.replace(lp(tmp), lp(dst))
        except BaseException:
            try:
                os.remove(lp(tmp))
            except OSError:
                pass
            raise
        self.log(f"COPIE    {src} -> {dst}")

    def _target(self, side: str, info: Optional[FInfo], other: FInfo) -> str:
        return info.path if info else os.path.join(self.root[side], other.rel)

    # -- actions ------------------------------------------------------------
    def do(self, a: Action):
        if a.op == "push":
            if a.d:
                self.archive_copy("d", a.d)
            self.copy_verified(a.s.path, self._target("d", a.d, a.s))
        elif a.op == "pull":
            if a.s:
                self.archive_copy("s", a.s)
            self.copy_verified(a.d.path, self._target("s", a.s, a.d))
        elif a.op == "del_d":
            self.archive_move("d", a.d)
        elif a.op == "del_s":
            self.archive_move("s", a.s)
        elif a.op == "conflict":
            win_side, lose_side = ("s", "d") if a.s.mtime >= a.d.mtime else ("d", "s")
            win, lose = (a.s, a.d) if win_side == "s" else (a.d, a.s)
            base, ext = os.path.splitext(a.rel)
            cname = f"{base} (conflit {SIDE_LABEL[lose_side]} {datetime.now():%Y-%m-%d %H%M%S}){ext}"
            self.copy_verified(lose.path, os.path.join(self.root[lose_side], cname))
            self.copy_verified(lose.path, os.path.join(self.root[win_side], cname))
            self.archive_copy(lose_side, lose)
            self.copy_verified(win.path, lose.path)

    def apply(self, plan: Plan, actions: list, state: dict, progress=None) -> dict:
        res = {"done": 0, "errors": [], "archive": self.vroot}
        self.log(f"=== Synchronisation {self.cfg.source} <-> {self.cfg.dest} ({self.cfg.mode}) ===")
        for k in plan.gone:
            state.pop(k, None)
        for k in plan.same:                      # identiques : on mémorise cet état
            self._remember(state, k, plan, None)
        todo = [a for a in actions if a.op != "skip"]
        for i, a in enumerate(todo, 1):
            if progress:
                progress(i, len(todo), a.rel)
            if self.lock_path:
                try:
                    os.utime(self.lock_path)     # signale que la synchro est toujours vivante
                except OSError:
                    pass
            try:
                self.do(a)
                self._remember(state, a.key, plan, a)
                res["done"] += 1
            except Exception as e:  # noqa
                res["errors"].append((a.rel, str(e)))
                self.log(f"ERREUR   {a.rel}: {e}")
        save_state(self.cfg, state)
        prune_versions(self.cfg)
        self.log(f"=== Terminé : {res['done']} actions, {len(res['errors'])} erreurs ===")
        self._log.close()
        return res

    def _remember(self, state: dict, key: str, plan: Plan, a: Optional[Action]):
        if a and a.op in ("del_d", "del_s"):
            state.pop(key, None)
            return
        try:
            rel = a.rel if a else (plan._S.get(key) or plan._D.get(key)).rel
            ss = os.stat(lp(os.path.join(self.root["s"], rel)))
            ds = os.stat(lp(os.path.join(self.root["d"], rel)))
            state[key] = {"s": [ss.st_size, ss.st_mtime], "d": [ds.st_size, ds.st_mtime]}
        except OSError:
            state.pop(key, None)


def prune_versions(cfg: Config):
    root = cfg.versions_root()
    if cfg.keep_days <= 0 or not os.path.isdir(root):
        return
    limit = datetime.now() - timedelta(days=cfg.keep_days)
    for name in os.listdir(root):
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}_\d{6}", name):
            try:
                if datetime.strptime(name, "%Y-%m-%d_%H%M%S") < limit:
                    shutil.rmtree(os.path.join(root, name), ignore_errors=True)
            except ValueError:
                pass


def run_analysis(cfg: Config, progress=None):
    cfg.validate()
    S = scan(cfg.source, cfg.excludes, lambda n: progress and progress(f"Source : {n} fichiers"))
    D = scan(cfg.dest, cfg.excludes, lambda n: progress and progress(f"Destination : {n} fichiers"))
    state = load_state(cfg)
    plan = make_plan(cfg, S, D, state, lambda i, n: progress and progress(f"Comparaison {i}/{n}"))
    plan._S, plan._D = S, D
    return plan, state


def run_sync(cfg: Config, plan: Plan, actions: list, state: dict, progress=None) -> dict:
    lock = os.path.join(app_dir(), "sync.lock")
    if os.path.exists(lock) and time.time() - os.path.getmtime(lock) < 600:
        raise RuntimeError("Une synchronisation est déjà en cours (autre fenêtre ou tâche d'arrière-plan).")
    with open(lock, "w") as f:
        f.write(str(os.getpid()))
    try:
        sy = Syncer(cfg)
        sy.lock_path = lock
        return sy.apply(plan, actions, state, progress)
    finally:
        try:
            os.remove(lock)
        except OSError:
            pass


def deletion_warning(plan: Plan, actions: list) -> Optional[str]:
    dels = sum(1 for a in actions if a.op in ("del_d", "del_s"))
    if dels >= 10 and dels > 0.3 * max(plan.n_total, 1):
        return (f"{dels} suppressions sur {plan.n_total} fichiers ({dels * 100 // plan.n_total} %). "
                "Vérifiez que le bon dossier est branché/sélectionné.")
    return None


# ----------------------------------------------------------------------------
# Mode console
# ----------------------------------------------------------------------------
def cli(apply: bool, force: bool) -> int:
    cfg = Config.load()
    try:
        plan, state = run_analysis(cfg, lambda m: print("\r" + m, end="", file=sys.stderr))
    except Exception as e:  # noqa
        print(f"\nErreur : {e}", file=sys.stderr)
        return 2
    print(file=sys.stderr)
    todo = [a for a in plan.actions if a.op != "skip"]
    for a in plan.actions:
        print(f"{a.label():48} {a.rel}  {('[' + a.note + ']') if a.note else ''}")
    print(f"\n{len(plan.same)} identiques, {len(todo)} actions, {len(plan.actions) - len(todo)} ignorées.")
    if not apply:
        return 0
    warn = deletion_warning(plan, todo)
    if warn and not force:
        print(f"REFUSÉ : {warn} (relancez avec --force pour passer outre)", file=sys.stderr)
        return 3
    res = run_sync(cfg, plan, todo, state)
    print(f"{res['done']} actions effectuées, {len(res['errors'])} erreurs. Versions : {res['archive']}")
    for rel, err in res["errors"]:
        print(f"  ERREUR {rel}: {err}", file=sys.stderr)
    return 1 if res["errors"] else 0


# ----------------------------------------------------------------------------
# Arrière-plan : boucle automatique + tâche planifiée Windows
# ----------------------------------------------------------------------------
def background(interval_min: int) -> int:
    """Synchronise en boucle, sans fenêtre. Les conflits gardent les deux versions ;
    si le garde-fou de suppressions se déclenche, le cycle est sauté (voir journal)."""
    import time
    while True:
        cfg = Config.load()
        try:
            plan, state = run_analysis(cfg)
            todo = [a for a in plan.actions if a.op != "skip"]
            warn = deletion_warning(plan, todo)
            if warn:
                msg = f"CYCLE SAUTÉ : {warn}"
            elif todo:
                res = run_sync(cfg, plan, todo, state)
                msg = f"{res['done']} actions, {len(res['errors'])} erreurs"
            else:
                msg = "rien à faire"
        except Exception as e:  # noqa  (disque débranché, OneDrive en pause, etc.)
            msg = f"ERREUR : {e}"
        with open(cfg.log_path(), "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%H:%M:%S} [arrière-plan] {msg}\n")
        time.sleep(max(1, interval_min) * 60)


def launch_cmd(args: str) -> str:
    exe = sys.executable
    if getattr(sys, "frozen", False):                            # version .exe (PyInstaller)
        return f'"{exe}" {args}'
    pyw = os.path.join(os.path.dirname(exe), "pythonw.exe")      # pythonw = aucune fenêtre console
    return f'"{pyw if os.path.exists(pyw) else exe}" "{os.path.abspath(__file__)}" {args}'


def set_startup(enable: bool):
    """Windows : ouvre l'application (réduite) à l'ouverture de session, via la clé Run de l'utilisateur."""
    if not IS_WIN:
        return
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run",
                        0, winreg.KEY_SET_VALUE) as k:
        if enable:
            winreg.SetValueEx(k, APP, 0, winreg.REG_SZ, launch_cmd("--minimized"))
        else:
            try:
                winreg.DeleteValue(k, APP)
            except FileNotFoundError:
                pass


def install_task(interval_min: int):
    """Crée une tâche Windows qui démarre FolderSync en arrière-plan à l'ouverture de session."""
    if not IS_WIN:
        print("Disponible seulement sous Windows.")
        return 1
    cmd = launch_cmd(f"--background --interval {interval_min}")
    r = subprocess.run(["schtasks", "/Create", "/F", "/TN", APP, "/SC", "ONLOGON", "/RL", "LIMITED", "/TR", cmd])
    if r.returncode == 0:
        subprocess.run(["schtasks", "/Run", "/TN", APP])
        print(f"Tâche '{APP}' créée et démarrée (toutes les {interval_min} min, à chaque ouverture de session).")
    return r.returncode


def remove_task():
    if IS_WIN:
        subprocess.run(["schtasks", "/End", "/TN", APP])
        return subprocess.run(["schtasks", "/Delete", "/F", "/TN", APP]).returncode
    return 1


# ----------------------------------------------------------------------------
# Interface graphique
# ----------------------------------------------------------------------------
def gui(minimized: bool = False):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    class App:
        TAGS = {"push": "#e8f1ff", "pull": "#e9f8ea", "del_d": "#ffe9e9", "del_s": "#ffe9e9",
                "conflict": "#fff1d6", "skip": "#eeeeee"}

        def __init__(self, root):
            self.root = root
            root.title("FolderSync — Source ⇄ OneDrive")
            root.geometry("1150x720")
            self.cfg = Config.load()
            self.plan = None
            self.state = {}
            self.q = queue.Queue()
            self.busy = False
            self.next_run = None
            self._build()
            self._poll()
            self._auto_tick()
            root.protocol("WM_DELETE_WINDOW", self._close)
            if minimized:
                root.iconify()
            if self.cfg.auto_enabled:
                self._schedule()
                if self.cfg.auto_on_start:
                    root.after(1500, self._auto_run)

        # -- construction ----------------------------------------------------
        def _build(self):
            c = self.cfg
            self.v_src = tk.StringVar(value=c.source)
            self.v_dst = tk.StringVar(value=c.dest)
            self.v_mode = tk.StringVar(value="Bidirectionnel (recommandé)" if c.mode == "two_way"
                                       else "Miroir : source → destination")
            self.v_del = tk.BooleanVar(value=c.propagate_deletes)
            self.v_hash = tk.BooleanVar(value=c.always_hash)
            self.v_online = tk.BooleanVar(value=c.skip_online_only)
            self.v_verify = tk.BooleanVar(value=c.verify)
            self.v_ver = tk.StringVar(value=c.versions_dir)
            self.v_days = tk.IntVar(value=c.keep_days)
            self.v_excl = tk.StringVar(value="; ".join(c.excludes))

            top = ttk.Frame(self.root, padding=8)
            top.pack(fill="x")
            top.columnconfigure(1, weight=1)

            def row(r, text, var, cmd):
                ttk.Label(top, text=text).grid(row=r, column=0, sticky="w", pady=2)
                ttk.Entry(top, textvariable=var).grid(row=r, column=1, sticky="ew", padx=6)
                ttk.Button(top, text="Parcourir…", command=cmd).grid(row=r, column=2)

            row(0, "Source (ex. C:\\Documents\\Projet)", self.v_src, lambda: self._browse(self.v_src))
            row(1, "Destination (OneDrive)", self.v_dst, lambda: self._browse(self.v_dst))
            row(2, "Dossier des versions archivées", self.v_ver, lambda: self._browse(self.v_ver))
            ttk.Label(top, text="(vide = %APPDATA%\\FolderSync\\versions, hors OneDrive)",
                      foreground="#666").grid(row=3, column=1, sticky="w", padx=6)

            opt = ttk.LabelFrame(self.root, text="Stratégie", padding=8)
            opt.pack(fill="x", padx=8)
            ttk.Label(opt, text="Mode :").grid(row=0, column=0, sticky="w")
            ttk.Combobox(opt, textvariable=self.v_mode, state="readonly", width=32,
                         values=["Bidirectionnel (recommandé)", "Miroir : source → destination"]
                         ).grid(row=0, column=1, sticky="w", padx=6)
            ttk.Label(opt, text="Garder les versions (jours) :").grid(row=0, column=2, sticky="e", padx=(20, 0))
            ttk.Spinbox(opt, from_=0, to=3650, width=6, textvariable=self.v_days).grid(row=0, column=3, padx=6)
            ttk.Label(opt, text="(0 = pour toujours)", foreground="#666").grid(row=0, column=4, sticky="w")
            ttk.Checkbutton(opt, text="Propager les suppressions (toujours archivées)", variable=self.v_del
                            ).grid(row=1, column=0, columnspan=2, sticky="w")
            ttk.Checkbutton(opt, text="Comparer le contenu (SHA-256) — plus lent", variable=self.v_hash
                            ).grid(row=1, column=2, columnspan=3, sticky="w", padx=(20, 0))
            ttk.Checkbutton(opt, text="Ne pas télécharger les fichiers OneDrive « en ligne seulement »",
                            variable=self.v_online).grid(row=2, column=0, columnspan=2, sticky="w")
            ttk.Checkbutton(opt, text="Vérifier chaque copie (SHA-256)", variable=self.v_verify
                            ).grid(row=2, column=2, columnspan=3, sticky="w", padx=(20, 0))
            ttk.Label(opt, text="Exclusions :").grid(row=3, column=0, sticky="w", pady=(4, 0))
            ttk.Entry(opt, textvariable=self.v_excl).grid(row=3, column=1, columnspan=4, sticky="ew", padx=6, pady=(4, 0))
            opt.columnconfigure(4, weight=1)

            self.v_auto = tk.BooleanVar(value=c.auto_enabled)
            self.v_interval = tk.IntVar(value=c.interval_min)
            self.v_onstart = tk.BooleanVar(value=c.auto_on_start)
            self.v_winstart = tk.BooleanVar(value=c.start_with_windows)
            self.v_next = tk.StringVar()
            au = ttk.LabelFrame(self.root, text="Synchronisation automatique", padding=8)
            au.pack(fill="x", padx=8, pady=(6, 0))
            ttk.Checkbutton(au, text="Synchroniser automatiquement toutes les", variable=self.v_auto
                            ).grid(row=0, column=0, sticky="w")
            ttk.Spinbox(au, from_=1, to=1440, width=6, textvariable=self.v_interval).grid(row=0, column=1, padx=6)
            ttk.Label(au, text="minutes").grid(row=0, column=2, sticky="w")
            ttk.Label(au, textvariable=self.v_next, foreground="#0a5").grid(row=0, column=3, sticky="e", padx=(20, 0))
            ttk.Checkbutton(au, text="Synchroniser dès l'ouverture de l'application", variable=self.v_onstart
                            ).grid(row=1, column=0, columnspan=3, sticky="w")
            wcb = ttk.Checkbutton(au, text="Ouvrir l'application (réduite) au démarrage de Windows",
                                  variable=self.v_winstart)
            wcb.grid(row=2, column=0, columnspan=3, sticky="w")
            if not IS_WIN:
                wcb.state(["disabled"])
            ttk.Label(au, text="Les suppressions massives et les erreurs sont notées dans le journal ; les conflits "
                               "gardent les deux versions.", foreground="#666").grid(row=3, column=0, columnspan=4, sticky="w")
            au.columnconfigure(3, weight=1)
            self._auto_sig = (c.auto_enabled, c.interval_min)
            for v in (self.v_src, self.v_dst, self.v_mode, self.v_del, self.v_hash, self.v_online, self.v_verify,
                      self.v_ver, self.v_days, self.v_excl, self.v_auto, self.v_interval, self.v_onstart,
                      self.v_winstart):
                v.trace_add("write", lambda *a: self._changed())

            bar = ttk.Frame(self.root, padding=8)
            bar.pack(fill="x")
            self.b_scan = ttk.Button(bar, text="1. Analyser les différences", command=self.analyze)
            self.b_scan.pack(side="left")
            self.b_sync = ttk.Button(bar, text="2. Synchroniser la sélection", command=self.synchronize, state="disabled")
            self.b_sync.pack(side="left", padx=8)
            ttk.Button(bar, text="Dossier des versions", command=lambda: open_in_explorer(self.cfg.versions_root())
                       ).pack(side="right")
            ttk.Button(bar, text="Journal", command=self._open_log).pack(side="right", padx=6)

            frame = ttk.Frame(self.root, padding=(8, 0))
            frame.pack(fill="both", expand=True)
            cols = ("sel", "file", "action", "src", "dst", "note")
            self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="extended")
            for col, title, w in (("sel", "✓", 36), ("file", "Fichier", 330), ("action", "Action", 250),
                                  ("src", "Source", 150), ("dst", "Destination", 150), ("note", "Remarque", 250)):
                self.tree.heading(col, text=title)
                self.tree.column(col, width=w, anchor="center" if col == "sel" else "w", stretch=col != "sel")
            for op, color in self.TAGS.items():
                self.tree.tag_configure(op, background=color)
            sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=sb.set)
            self.tree.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            self.tree.bind("<Button-1>", self._click)
            self.tree.bind("<Double-1>", lambda e: self._diff())
            self.tree.bind("<Button-3>", self._menu)
            self.tree.bind("<Button-2>", self._menu)

            self.menu = tk.Menu(self.root, tearoff=0)
            self.menu.add_command(label="Forcer →  source vers destination (ancienne version archivée)",
                                  command=lambda: self._force("push"))
            self.menu.add_command(label="Forcer ←  destination vers source (ancienne version archivée)",
                                  command=lambda: self._force("pull"))
            self.menu.add_command(label="Garder les deux versions", command=lambda: self._force("conflict"))
            self.menu.add_command(label="Ignorer", command=lambda: self._force("skip"))
            self.menu.add_separator()
            self.menu.add_command(label="Voir les différences (fichiers texte)", command=self._diff)

            st = ttk.Frame(self.root, padding=8)
            st.pack(fill="x")
            self.status = ttk.Label(st, text="Prêt. Clic sur la colonne ✓ pour cocher/décocher, clic droit pour changer l'action.")
            self.status.pack(side="left")
            self.prog = ttk.Progressbar(st, mode="indeterminate", length=180)
            self.prog.pack(side="right")

        # -- helpers ---------------------------------------------------------
        def _browse(self, var):
            d = filedialog.askdirectory(initialdir=var.get() or None)
            if d:
                var.set(os.path.normpath(d))

        def _read_cfg(self):
            c = self.cfg
            c.source, c.dest = self.v_src.get().strip(), self.v_dst.get().strip()
            c.mode = "two_way" if self.v_mode.get().startswith("Bidir") else "mirror"
            c.propagate_deletes, c.always_hash = self.v_del.get(), self.v_hash.get()
            c.skip_online_only, c.verify = self.v_online.get(), self.v_verify.get()
            c.versions_dir = self.v_ver.get().strip()
            try:
                c.keep_days = int(self.v_days.get())
            except (ValueError, tk.TclError):
                c.keep_days = 90
            c.excludes = [p.strip() for p in self.v_excl.get().split(";") if p.strip()]
            c.auto_enabled, c.auto_on_start = self.v_auto.get(), self.v_onstart.get()
            c.start_with_windows = self.v_winstart.get()
            try:
                c.interval_min = max(1, int(self.v_interval.get()))
            except (ValueError, tk.TclError):
                c.interval_min = 15
            c.save()

        def _changed(self):
            """Tout réglage est enregistré immédiatement (répertoires, options, délai)."""
            prev_win = self.cfg.start_with_windows
            try:
                self._read_cfg()
            except Exception:  # noqa
                return
            if (self.cfg.auto_enabled, self.cfg.interval_min) != self._auto_sig:
                self._auto_sig = (self.cfg.auto_enabled, self.cfg.interval_min)
                self._schedule()
            if self.cfg.start_with_windows != prev_win:
                try:
                    set_startup(self.cfg.start_with_windows)
                except Exception as e:  # noqa
                    messagebox.showerror("Démarrage de Windows", str(e))

        # -- synchronisation automatique ---------------------------------------
        def _schedule(self):
            self.next_run = (datetime.now() + timedelta(minutes=max(1, self.cfg.interval_min))
                             if self.cfg.auto_enabled else None)

        def _auto_tick(self):
            if self.cfg.auto_enabled and self.next_run:
                left = (self.next_run - datetime.now()).total_seconds()
                if left <= 0:
                    if self.busy:
                        self.next_run = datetime.now() + timedelta(seconds=30)
                    else:
                        self._auto_run()
                else:
                    self.v_next.set(f"Prochaine synchro dans {int(left // 60)} min {int(left % 60):02d} s")
            else:
                self.v_next.set("Synchronisation automatique désactivée")
            self.root.after(1000, self._auto_tick)

        def _auto_run(self):
            if self.busy or not self.cfg.auto_enabled:
                return
            try:
                self._read_cfg()
                self.cfg.validate()
            except ValueError as e:
                self.status.config(text=f"Synchro auto impossible : {e}")
                self._schedule()
                return
            self._set_busy(True, "Synchronisation automatique en cours…")

            def work():
                try:
                    plan, state = run_analysis(self.cfg)
                    todo = [a for a in plan.actions if a.op != "skip"]
                    warn = deletion_warning(plan, todo)
                    if warn:
                        msg = f"cycle sauté — {warn}"
                    elif todo:
                        res = run_sync(self.cfg, plan, todo, state)
                        msg = f"{res['done']} actions, {len(res['errors'])} erreur(s)"
                    else:
                        msg = "rien à faire"
                except Exception as e:  # noqa
                    msg = f"erreur — {e}"
                with open(self.cfg.log_path(), "a", encoding="utf-8") as f:
                    f.write(f"{datetime.now():%H:%M:%S} [auto] {msg}\n")
                self.q.put(("auto", f"Synchro auto {datetime.now():%H:%M} : {msg}"))
            threading.Thread(target=work, daemon=True).start()

        def _set_busy(self, busy, msg=None):
            self.busy = busy
            self.b_scan.config(state="disabled" if busy else "normal")
            self.b_sync.config(state="disabled" if busy or not self.plan else "normal")
            (self.prog.start if busy else self.prog.stop)(*([] if busy else []))
            if msg:
                self.status.config(text=msg)

        def _open_log(self):
            p = self.cfg.log_path()
            if not os.path.exists(p):
                open(p, "a").close()
            os.startfile(p) if IS_WIN else open_in_explorer(os.path.dirname(p))

        # -- analyse ---------------------------------------------------------
        def analyze(self):
            try:
                self._read_cfg()
                self.cfg.validate()
            except ValueError as e:
                messagebox.showwarning("Configuration", str(e))
                return
            self.plan = None
            self.tree.delete(*self.tree.get_children())
            self._set_busy(True, "Analyse en cours…")

            def work():
                try:
                    plan, state = run_analysis(self.cfg, lambda m: self.q.put(("msg", m)))
                    self.q.put(("plan", (plan, state)))
                except Exception as e:  # noqa
                    self.q.put(("error", str(e)))
            threading.Thread(target=work, daemon=True).start()

        def _fill(self):
            self.tree.delete(*self.tree.get_children())
            order = {"conflict": 0, "del_s": 1, "del_d": 1, "push": 2, "pull": 2, "skip": 3}
            self.plan.actions.sort(key=lambda a: (order[a.op], a.rel.lower()))
            for i, a in enumerate(self.plan.actions):
                self._row(i, a)
            counts = {}
            for a in self.plan.actions:
                counts[a.op] = counts.get(a.op, 0) + 1
            self.status.config(text=(
                f"{len(self.plan.same)} identiques | →{counts.get('push', 0)} ←{counts.get('pull', 0)} "
                f"✕{counts.get('del_s', 0) + counts.get('del_d', 0)} ⚠{counts.get('conflict', 0)} "
                f"⏸{counts.get('skip', 0)}"))

        def _row(self, i, a, update=False):
            s = f"{fmt_size(a.s.size)}  {fmt_time(a.s.mtime)}" if a.s else "—"
            d = f"{fmt_size(a.d.size)}  {fmt_time(a.d.mtime)}" if a.d else "—"
            vals = ("☑" if a.selected and a.op != "skip" else "☐", a.rel, a.label(), s, d, a.note)
            if update:
                self.tree.item(str(i), values=vals, tags=(a.op,))
            else:
                self.tree.insert("", "end", iid=str(i), values=vals, tags=(a.op,))

        def _click(self, e):
            if self.tree.identify_region(e.x, e.y) != "cell" or self.tree.identify_column(e.x) != "#1":
                return
            iid = self.tree.identify_row(e.y)
            if iid:
                a = self.plan.actions[int(iid)]
                if a.op != "skip":
                    a.selected = not a.selected
                    self._row(int(iid), a, True)
                return "break"

        def _menu(self, e):
            iid = self.tree.identify_row(e.y)
            if iid:
                if iid not in self.tree.selection():
                    self.tree.selection_set(iid)
                self.menu.tk_popup(e.x_root, e.y_root)

        def _force(self, op):
            for iid in self.tree.selection():
                a = self.plan.actions[int(iid)]
                if force_action(a, op):
                    a.selected = op != "skip"
                    self._row(int(iid), a, True)

        def _diff(self):
            sel = self.tree.selection()
            if not sel or not self.plan:
                return
            a = self.plan.actions[int(sel[0])]
            if not (a.s and a.d):
                messagebox.showinfo("Différences", "Ce fichier n'existe que d'un côté.")
                return
            if a.s.size > 2_000_000 or a.d.size > 2_000_000 or a.s.online or a.d.online:
                messagebox.showinfo("Différences", "Fichier trop gros (> 2 Mo) ou en ligne seulement.")
                return

            def lines(p):
                with open(lp(p), "rb") as f:
                    raw = f.read()
                if b"\0" in raw:
                    return None
                return raw.decode("utf-8", "replace").splitlines(keepends=True)
            ls, ld = lines(a.s.path), lines(a.d.path)
            if ls is None or ld is None:
                messagebox.showinfo("Différences", "Fichier binaire : comparaison texte impossible.")
                return
            win = tk.Toplevel(self.root)
            win.title(f"Différences — {a.rel}")
            win.geometry("900x600")
            txt = tk.Text(win, wrap="none", font=("Consolas", 10))
            txt.pack(fill="both", expand=True)
            txt.tag_configure("add", background="#dff5df")
            txt.tag_configure("del", background="#fbe0e0")
            txt.tag_configure("hdr", foreground="#555")
            for line in difflib.unified_diff(ls, ld, "source", "destination"):
                tag = "add" if line.startswith("+") else "del" if line.startswith("-") else \
                    "hdr" if line.startswith("@") else ""
                txt.insert("end", line, tag)
            txt.config(state="disabled")

        # -- synchronisation -------------------------------------------------
        def synchronize(self):
            acts = [a for a in self.plan.actions if a.selected and a.op != "skip"]
            if not acts:
                messagebox.showinfo("Synchronisation", "Aucune action sélectionnée.")
                return
            n = {k: sum(1 for a in acts if a.op == k) for k in ("push", "pull", "del_s", "del_d", "conflict")}
            msg = (f"→ vers destination : {n['push']}\n← vers source : {n['pull']}\n"
                   f"✕ suppressions (archivées) : {n['del_s'] + n['del_d']}\n⚠ conflits : {n['conflict']}\n\n"
                   f"Les anciennes versions seront archivées dans :\n{self.cfg.versions_root()}")
            warn = deletion_warning(self.plan, acts)
            if warn:
                msg = "ATTENTION — " + warn + "\n\n" + msg
            if not messagebox.askokcancel("Confirmer la synchronisation", msg,
                                          icon="warning" if warn else "question"):
                return
            self._set_busy(True, "Synchronisation en cours…")

            def work():
                try:
                    res = run_sync(self.cfg, self.plan, acts, self.state,
                                   lambda i, n, r: self.q.put(("msg", f"{i}/{n}  {r}")))
                    self.q.put(("done", res))
                except Exception as e:  # noqa
                    self.q.put(("error", str(e)))
            threading.Thread(target=work, daemon=True).start()

        # -- file de messages du thread de travail ----------------------------
        def _poll(self):
            try:
                while True:
                    kind, data = self.q.get_nowait()
                    if kind == "msg":
                        self.status.config(text=data)
                    elif kind == "plan":
                        self.plan, self.state = data
                        self._set_busy(False)
                        self._fill()
                    elif kind == "auto":
                        self.plan = None
                        self.tree.delete(*self.tree.get_children())
                        self._set_busy(False, data)
                        self._schedule()
                    elif kind == "error":
                        self._set_busy(False, "Erreur.")
                        messagebox.showerror("Erreur", data)
                    elif kind == "done":
                        errs = data["errors"]
                        text = f"{data['done']} actions effectuées, {len(errs)} erreur(s).\nVersions : {data['archive']}"
                        if errs:
                            text += "\n\n" + "\n".join(f"{r}: {e}" for r, e in errs[:15])
                        (messagebox.showwarning if errs else messagebox.showinfo)("Terminé", text)
                        self.plan = None
                        self.analyze()
            except queue.Empty:
                pass
            self.root.after(100, self._poll)

        def _close(self):
            if self.busy:
                return
            try:
                self._read_cfg()
            except Exception:  # noqa
                pass
            if self.cfg.auto_enabled:
                r = messagebox.askyesnocancel(
                    "FolderSync", "La synchronisation automatique est active.\n\n"
                                  "Oui = réduire la fenêtre (la synchro continue)\nNon = quitter l'application")
                if r is None:
                    return
                if r:
                    self.root.iconify()
                    return
            self.root.destroy()

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Synchronisation sûre dossier local <-> OneDrive")
    ap.add_argument("--cli", action="store_true", help="mode console (utilise la configuration enregistrée)")
    ap.add_argument("--apply", action="store_true", help="avec --cli : exécuter la synchronisation")
    ap.add_argument("--force", action="store_true", help="avec --apply : ignorer le garde-fou de suppressions")
    ap.add_argument("--background", action="store_true", help="boucle de synchronisation automatique, sans fenêtre")
    ap.add_argument("--interval", type=int, default=15, help="minutes entre deux cycles (défaut 15)")
    ap.add_argument("--install-task", action="store_true", help="Windows : lancer en arrière-plan à chaque ouverture de session")
    ap.add_argument("--remove-task", action="store_true", help="Windows : retirer la tâche d'arrière-plan")
    ap.add_argument("--minimized", action="store_true", help="ouvrir l'interface réduite")
    args = ap.parse_args()
    if args.install_task:
        sys.exit(install_task(args.interval))
    if args.remove_task:
        sys.exit(remove_task())
    if args.background:
        sys.exit(background(args.interval))
    sys.exit(cli(args.apply, args.force) if args.cli else gui(args.minimized))
