import json
import re
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk
from typing import Dict, Optional

from pywinauto import Desktop

CONFIG_FILENAME = ".win_auto_program.json"


def _config_path() -> Path:
    return Path.cwd() / CONFIG_FILENAME


def _derive_keyword(title: str) -> str:
    parts = [p.strip() for p in re.split(r"\s*[-–—]\s*", title) if p.strip()]
    if parts:
        return parts[-1]
    return title.strip()


def _load_config() -> Optional[Dict[str, str]]:
    path = _config_path()
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("config is not a dict")
        if "class_name" in data and "keyword" in data:
            return {"class_name": str(data["class_name"]), "keyword": str(data["keyword"])}
    except Exception as exc:
        print(f"Konfiguration konnte nicht gelesen werden ({exc}). Bitte neu auswählen.")
    return None


def _save_config(config: Dict[str, str]) -> None:
    path = _config_path()
    path.write_text(json.dumps(config, ensure_ascii=True, indent=2), encoding="utf-8")


def _show_window_selection_dialog(windows):
    """Zeigt einen grafischen Dialog zur Auswahl eines Fensters."""
    result = [None]  # Mutable container für das Ergebnis

    dialog = tk.Tk()
    dialog.title("OTBioLab Fenster auswählen")
    dialog.geometry("700x500")
    dialog.resizable(True, True)

    # Info-Label
    info_label = ttk.Label(
        dialog,
        text="Bitte wählen Sie das Programm für den OTBioLab Speichern-Dialog aus:",
        font=("Segoe UI", 10)
    )
    info_label.pack(padx=20, pady=(20, 10), anchor="w")

    # Frame für Listbox mit Scrollbar
    list_frame = ttk.Frame(dialog)
    list_frame.pack(padx=20, pady=10, fill="both", expand=True)

    # Scrollbar
    scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
    scrollbar.pack(side="right", fill="y")

    # Listbox
    listbox = tk.Listbox(
        list_frame,
        yscrollcommand=scrollbar.set,
        font=("Segoe UI", 9),
        height=20
    )
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    # Fenster zur Listbox hinzufügen
    for idx, info in enumerate(windows):
        display_text = f"{info['title']} (Klasse: {info['class_name']})"
        listbox.insert(tk.END, display_text)

    # Button-Frame
    button_frame = ttk.Frame(dialog)
    button_frame.pack(padx=20, pady=(10, 20), fill="x")

    def on_select():
        selection = listbox.curselection()
        if selection:
            idx = selection[0]
            selected = windows[idx]
            keyword = _derive_keyword(selected["title"])
            result[0] = {"class_name": selected["class_name"], "keyword": keyword}
            dialog.destroy()

    def on_cancel():
        dialog.destroy()

    # Buttons
    ttk.Button(button_frame, text="Abbrechen", command=on_cancel).pack(side="right", padx=(10, 0))
    ttk.Button(button_frame, text="Auswählen", command=on_select).pack(side="right")

    # Doppelklick auf Eintrag = Auswahl
    listbox.bind("<Double-Button-1>", lambda e: on_select())

    # Dialog modal machen
    dialog.transient()
    dialog.grab_set()
    dialog.mainloop()

    return result[0]


def _prompt_for_config() -> Optional[Dict[str, str]]:
    windows = []
    seen = set()
    try:
        desktop = Desktop(backend="win32")
    except Exception as exc:
        print(f"Desktop konnte nicht initialisiert werden ({exc}).")
        return None

    for win in desktop.windows():
        try:
            if not win.is_visible():
                continue
            title = win.window_text().strip()
            class_name = (win.element_info.class_name or "").strip()
            if not title or not class_name:
                continue
            key = (class_name.lower(), title.lower())
            if key in seen:
                continue
            seen.add(key)
            windows.append({"title": title, "class_name": class_name})
        except Exception:
            continue

    if not windows:
        print("Keine sichtbaren Fenster gefunden.")
        return None

    # Verwende GUI-Dialog statt Konsole
    config = _show_window_selection_dialog(windows)
    if config:
        print(
            f"Auswahl gespeichert: Klasse '{config['class_name']}', Schlüsselwort '{config['keyword']}'."
        )
    return config


def _ensure_config() -> Optional[Dict[str, str]]:
    config = _load_config()
    if config:
        print(
            f"Verwende gespeicherte Auswahl: Klasse '{config['class_name']}', "
            f"Schlüsselwort '{config['keyword']}'."
        )
        return config
    config = _prompt_for_config()
    if config:
        _save_config(config)
    return config


def _get_pid(config: Dict[str, str]) -> Optional[int]:
    class_name = config.get("class_name", "")
    keyword = config.get("keyword", "").lower()
    for be in ("uia", "win32"):
        try:
            for w in Desktop(backend=be).windows():
                try:
                    if class_name and w.element_info.class_name != class_name:
                        continue
                    if keyword and keyword not in w.window_text().lower():
                        continue
                    return w.element_info.process_id
                except Exception:
                    continue
        except Exception:
            continue
    return None

def save_in_word_dialog(path, timeout=20):
    config = _ensure_config()
    if not config:
        print("Kein Programm ausgewählt. Vorgang abgebrochen.")
        return False

    pid = _get_pid(config)
    if not pid:
        print(
            f"Kein passender Prozess gefunden (Klasse '{config.get('class_name')}', "
            f"Schlüsselwort '{config.get('keyword')}')."
        )
        return False

    t0 = time.time()
    while time.time() - t0 < timeout:
        # 1) Versuche klassischen Dialog (#32770) dieses Prozesses
        for be in ("uia","win32"):
            d = Desktop(backend=be)
            try:
                dlg = d.window(class_name="#32770", process=pid)
                if dlg.exists() and dlg.is_visible():
                    print(f"[hit {be}] #32770")
                    return _fill_and_save_win32(dlg, path)
            except Exception:
                pass

        # 2) UIA-Variante: Titel enthält Speichern/Save & gehört zum Zielprozess
        try:
            d = Desktop(backend="uia")
            for w in d.windows():
                ei = w.element_info
                if (ei.process_id == pid and
                    ei.control_type in ("Window","Pane") and
                    re.search(r"(Speichern|Save)", w.window_text(), re.I)):
                    print("[hit uia] title contains Speichern/Save")
                    return _fill_and_save(w, path)
        except Exception:
            pass

        # 3) Fallback: CFD-Hilfsfenster → gehe einen Schritt nach oben
        try:
            d = Desktop(backend="win32")
            cfd = d.window(title_re="CFD File .* Window", process=pid)
            if cfd.exists():
                dlg = cfd.top_level_parent()
                if dlg.exists() and dlg.is_visible():
                    print("[hit win32] via CFD parent")
                    return _fill_and_save(dlg, path)
        except Exception:
            pass

        time.sleep(0.2)
    return False

def _fill_and_save_win32(dlg, path):
    # 1) Dateiname-Edit: hat i.d.R. einen nicht-leeren Titel wie "tmp.docx"
    try:
        edit = dlg.child_window(class_name="Edit", title_re=r".+\.\w+$").wrapper_object()
    except Exception:
        # Fallback: nimm das Edit innerhalb der ComboBox mit nicht-leerem Titel
        combo = dlg.child_window(class_name="ComboBox", title_re=r".+\.\w+$").wrapper_object()
        edit = combo.child_window(class_name="Edit").wrapper_object()

    # Eingabe
    edit.set_focus()
    try:
        edit.set_edit_text("")  # schneller als ^a{BACKSPACE}
    except Exception:
        pass
    edit.type_keys(path, with_spaces=True, set_foreground=True)

    # 2) Speichern klicken (deutsch/englisch)
    try:
        btn = (dlg.child_window(title_re=r"^S&?peichern$", class_name="Button")
                  .wrapper_object())
    except Exception:
        try:
            btn = dlg.child_window(title_re=r"^Save$", class_name="Button").wrapper_object()
        except Exception:
            btn = None

    if btn:
        btn.click_input()
    else:
        dlg.type_keys("{ENTER}")

    return True


if __name__ == "__main__":
    ok = save_in_word_dialog(r"C:\Temp\auto_saved.docx", timeout=25)
    print("Done:", ok)
