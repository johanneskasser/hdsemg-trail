from __future__ import annotations

import json
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional, Tuple

ROOT_DIR = Path(__file__).resolve().parent.parent
RESOURCES_DIR = ROOT_DIR / "resources"
if str(RESOURCES_DIR) not in sys.path:
    sys.path.insert(0, str(RESOURCES_DIR))

try:
    from get_save_dialog import save_in_word_dialog
except Exception:
    print("Warnung: OTBioLab Interceptor konnte nicht geladen werden.")
    save_in_word_dialog = None


@dataclass
class FieldConfig:
    field_id: str
    label: str
    kind: str = "string"
    required: bool = False
    placeholder: Optional[str] = None
    options: List[str] = field(default_factory=list)
    use_from_ref: bool = False  # Wert aus Referenz-File übernehmen?
    otbiolab_template: Optional[str] = None  # Dateinamen-Template für dieses Feld
    repeated_measurement: bool = False  # Kann diese Messung wiederholt werden?
    repeated_fields: List['FieldConfig'] = field(default_factory=list)  # Sub-Felder für jede Wiederholung


@dataclass
class StepConfig:
    step_id: str
    title: str
    description: str = ""
    expected_duration_seconds: Optional[int] = None
    fields: List[FieldConfig] = field(default_factory=list)
    otbiolab_template: Optional[str] = None
    notes_placeholder: Optional[str] = None


@dataclass
class Declaration:
    title: str
    description: str
    metadata_fields: List[FieldConfig]
    steps: List[StepConfig]


@dataclass
class StepResult:
    config: StepConfig
    started_at: Optional[datetime]
    completed_at: Optional[datetime]
    duration: Optional[timedelta]
    values: Dict[str, Any]
    notes: str
    otbiolab_paths: List[str]  # Liste aller übergebenen OTBioLab-Dateien (Schritt-Ebene, legacy)
    field_otbiolab_files: Dict[str, List[str]] = field(default_factory=dict)  # Feld-ID → Liste von Dateien
    repeated_measurements: Dict[str, List[Dict[str, Any]]] = field(default_factory=dict)  # Feld-ID → Liste von Versuchen (jeder Versuch = Dict von Sub-Feld-Werten)


def _parse_field_config(field_data: dict) -> FieldConfig:
    """Parse a field config from JSON, including nested repeated_fields."""
    repeated_fields = []
    if "repeated_fields" in field_data:
        repeated_fields = [_parse_field_config(rf) for rf in field_data["repeated_fields"]]

    return FieldConfig(
        field_id=str(field_data["id"]),
        label=str(field_data.get("label", field_data["id"])),
        kind=str(field_data.get("type", "string")),
        required=bool(field_data.get("required", False)),
        placeholder=field_data.get("placeholder"),
        options=list(field_data.get("options", [])),
        use_from_ref=bool(field_data.get("use_from_ref", False)),
        otbiolab_template=field_data.get("otbiolab_filename_template"),
        repeated_measurement=bool(field_data.get("type") == "repeated_measurement"),
        repeated_fields=repeated_fields,
    )


def load_declaration(path: Path) -> Declaration:
    data = json.loads(path.read_text(encoding="utf-8"))
    metadata_fields = [_parse_field_config(item) for item in data.get("metadata_fields", [])]
    steps = []
    for raw_step in data.get("steps", []):
        step_fields = [_parse_field_config(field) for field in raw_step.get("fields", [])]
        steps.append(
            StepConfig(
                step_id=str(raw_step["id"]),
                title=str(raw_step.get("title", raw_step["id"])),
                description=str(raw_step.get("description", "")),
                expected_duration_seconds=raw_step.get("expected_duration_seconds"),
                fields=step_fields,
                otbiolab_template=raw_step.get("otbiolab_filename_template"),
                notes_placeholder=raw_step.get("notes_placeholder"),
            )
        )
    if not steps:
        raise ValueError("Deklarationsdatei enthält keine Schritte.")
    return Declaration(
        title=str(data.get("title", "Versuchsreihe")),
        description=str(data.get("description", "")),
        metadata_fields=metadata_fields,
        steps=steps,
    )


def seconds_to_clock(value: Optional[float]) -> str:
    """Konvertiert Sekunden zu HH:MM:SS Format."""
    if value is None or value < 0:
        return "00:00:00"
    seconds = int(value)
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def seconds_to_minutes_clock(value: Optional[float]) -> str:
    """Konvertiert Sekunden zu MM:SS Format (Minuten können über 60 gehen)."""
    if value is None or value < 0:
        return "00:00"
    seconds = int(value)
    minutes, secs = divmod(seconds, 60)
    return f"{minutes:02d}:{secs:02d}"


def windows_pick_directory(parent: tk.Tk, title: str, initial_dir: Path) -> Optional[Path]:
    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return None

    shell32 = ctypes.windll.shell32
    user32 = ctypes.windll.user32

    BIF_NEWDIALOGSTYLE = 0x0040
    BIF_RETURNONLYFSDIRS = 0x0001
    BIF_EDITBOX = 0x0010

    BFFM_INITIALIZED = 1
    BFFM_SETSELECTIONW = 0x0467

    initial_path = initial_dir.resolve()
    initial_str = str(initial_path)
    initial_buf = ctypes.create_unicode_buffer(initial_str)

    BFFCALLBACK = ctypes.WINFUNCTYPE(ctypes.c_int, wintypes.HWND, ctypes.c_uint, wintypes.LPARAM, wintypes.LPARAM)

    @BFFCALLBACK
    def browse_callback(hwnd, msg, lParam, lpData):
        if msg == BFFM_INITIALIZED and initial_buf.value:
            user32.SendMessageW(hwnd, BFFM_SETSELECTIONW, 1, ctypes.cast(initial_buf, ctypes.c_void_p).value)
        return 0

    class BROWSEINFO(ctypes.Structure):
        _fields_ = [
            ("hwndOwner", wintypes.HWND),
            ("pidlRoot", ctypes.c_void_p),
            ("pszDisplayName", wintypes.LPWSTR),
            ("lpszTitle", wintypes.LPWSTR),
            ("ulFlags", ctypes.c_uint),
            ("lpfn", ctypes.c_void_p),
            ("lParam", ctypes.c_void_p),
            ("iImage", ctypes.c_int),
        ]

    browse_info = BROWSEINFO()
    browse_info.hwndOwner = parent.winfo_id() if parent else None
    browse_info.pidlRoot = None
    display_buf = ctypes.create_unicode_buffer(260)
    browse_info.pszDisplayName = ctypes.cast(display_buf, wintypes.LPWSTR)
    browse_info.lpszTitle = title
    browse_info.ulFlags = BIF_NEWDIALOGSTYLE | BIF_RETURNONLYFSDIRS | BIF_EDITBOX
    browse_info.lpfn = ctypes.cast(browse_callback, ctypes.c_void_p)
    browse_info.lParam = None
    browse_info.iImage = 0

    pidl = shell32.SHBrowseForFolderW(ctypes.byref(browse_info))
    if not pidl:
        return None

    path_buf = ctypes.create_unicode_buffer(260)
    success = shell32.SHGetPathFromIDListW(pidl, path_buf)
    ctypes.windll.ole32.CoTaskMemFree(pidl)
    if success:
        result = path_buf.value
        return Path(result)
    return None


class FieldControl:
    def __init__(self, config: FieldConfig, widget: tk.Widget, variable: Optional[tk.Variable] = None):
        self.config = config
        self.widget = widget
        self.variable = variable
        self._change_callbacks: List[Any] = []

    def get_value(self) -> str:
        if isinstance(self.widget, tk.Text):
            return self.widget.get("1.0", "end-1c").strip()
        if self.variable is not None:
            return str(self.variable.get()).strip()
        if hasattr(self.widget, "get"):
            return str(self.widget.get()).strip()
        return ""

    def set_value(self, value: Any) -> None:
        if isinstance(self.widget, tk.Text):
            self.widget.delete("1.0", "end")
            if value:
                self.widget.insert("1.0", str(value))
            return
        if self.variable is not None:
            self.variable.set("" if value is None else str(value))
            return
        if hasattr(self.widget, "delete") and hasattr(self.widget, "insert"):
            self.widget.delete(0, "end")
            if value:
                self.widget.insert(0, str(value))

    def bind_on_change(self, callback: Any) -> None:
        if isinstance(self.widget, tk.Text):
            def _on_change(event: tk.Event) -> None:
                if self.widget.edit_modified():
                    self.widget.edit_modified(False)
                    callback()

            self.widget.bind("<<Modified>>", _on_change)
            return
        if self.variable is not None:
            self.variable.trace_add("write", lambda *args: callback())
        elif hasattr(self.widget, "bind"):
            self.widget.bind("<<ComboboxSelected>>", lambda *_: callback())


class SessionApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("HDsEMG Versuchsreihe Assistent")
        self.geometry("1200x800")
        self.minsize(960, 640)

        self.declaration: Optional[Declaration] = None
        self.declaration_path: Optional[Path] = None
        self.output_dir: Optional[Path] = None
        self.metadata_controls: Dict[str, FieldControl] = {}
        self.metadata_values: Dict[str, Any] = {}
        self.step_controls: Dict[str, FieldControl] = {}
        self.step_results: List[StepResult] = []
        self.current_step_index: int = -1
        self.session_started_at: Optional[datetime] = None
        self.current_step_started_at: Optional[datetime] = None
        self.current_step_otbiolab_paths: List[str] = []  # Liste aller OTBioLab-Dateien für aktuellen Schritt (legacy)
        self.current_step_field_otbiolab_files: Dict[str, List[str]] = {}  # Feld-ID → Liste von OTBioLab-Dateien
        self.current_step_repeated_measurements: Dict[str, List[Dict[str, Any]]] = {}  # Feld-ID → Liste von Versuchen für wiederholbare Messungen
        self.session_timestamp: Optional[str] = None
        self._timer_after_id: Optional[str] = None
        self.session_finished: bool = False
        self.reference_data: Optional[Dict[str, Any]] = None  # Geladene Referenz-Daten

        # Bind close event
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        self._build_ui()
        self._show_frame("start")

    # UI Aufbau ------------------------------------------------------------------
    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.container = ttk.Frame(self)
        self.container.grid(row=0, column=0, sticky="nsew")
        self.container.columnconfigure(0, weight=1)
        self.container.rowconfigure(0, weight=1)

        self.frames: Dict[str, ttk.Frame] = {}
        for name in ("start", "step", "summary"):
            frame = ttk.Frame(self.container, padding=24)
            frame.grid(row=0, column=0, sticky="nsew")
            frame.columnconfigure(0, weight=1)
            self.frames[name] = frame

        self._build_start_frame()
        self._build_step_frame()
        self._build_summary_frame()

    def _show_frame(self, name: str) -> None:
        frame = self.frames.get(name)
        if frame:
            frame.tkraise()

    # Startseite -----------------------------------------------------------------
    def _build_start_frame(self) -> None:
        frame = self.frames["start"]

        title = ttk.Label(frame, text="HDsEMG Versuchsreihe starten", font=("Segoe UI", 24, "bold"))
        title.grid(row=0, column=0, sticky="w")

        self.declaration_info_var = tk.StringVar(value="Bitte Deklarationsdatei auswählen.")
        info_label = ttk.Label(frame, textvariable=self.declaration_info_var, wraplength=900, justify="left")
        info_label.grid(row=1, column=0, sticky="w", pady=(12, 18))

        decl_row = ttk.Frame(frame)
        decl_row.grid(row=2, column=0, sticky="ew", pady=6)
        decl_row.columnconfigure(0, weight=1)
        self.declaration_path_var = tk.StringVar()
        decl_entry = ttk.Entry(decl_row, textvariable=self.declaration_path_var, state="readonly")
        decl_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(decl_row, text="Deklaration laden…", command=self._choose_declaration).grid(row=0, column=1)

        out_row = ttk.Frame(frame)
        out_row.grid(row=3, column=0, sticky="ew", pady=6)
        out_row.columnconfigure(0, weight=1)
        self.output_dir_var = tk.StringVar()
        out_entry = ttk.Entry(out_row, textvariable=self.output_dir_var, state="readonly")
        out_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(out_row, text="OTBioLab Zielordner…", command=self._choose_output_dir).grid(row=0, column=1)

        # Referenz-Datei Zeile
        ref_row = ttk.Frame(frame)
        ref_row.grid(row=4, column=0, sticky="ew", pady=6)
        ref_row.columnconfigure(0, weight=1)
        self.reference_file_var = tk.StringVar()
        ref_entry = ttk.Entry(ref_row, textvariable=self.reference_file_var, state="readonly")
        ref_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(ref_row, text="Referenz-Datei laden… (optional)", command=self._load_reference_file).grid(row=0, column=1)

        self.metadata_group = ttk.LabelFrame(frame, text="Session-Informationen")
        self.metadata_group.grid(row=5, column=0, sticky="nsew", pady=(18, 18))
        self.metadata_group.columnconfigure(0, weight=1)
        self.metadata_fields_frame = ttk.Frame(self.metadata_group)
        self.metadata_fields_frame.grid(row=0, column=0, sticky="nsew")
        self.metadata_fields_frame.columnconfigure(1, weight=1)

        self.start_button = ttk.Button(frame, text="Messung starten", command=self._start_session, state="disabled")
        self.start_button.grid(row=6, column=0, sticky="e")

        frame.rowconfigure(7, weight=1)

    def _choose_declaration(self) -> None:
        initial_dir = str((ROOT_DIR / "config").resolve())
        path_str = filedialog.askopenfilename(
            parent=self,
            title="Deklarationsdatei auswählen",
            initialdir=initial_dir,
            filetypes=[("JSON Dateien", "*.json"), ("Alle Dateien", "*.*")],
        )
        if not path_str:
            return
        path = Path(path_str)
        try:
            declaration = load_declaration(path)
        except Exception as exc:
            messagebox.showerror("Fehler", f"Deklarationsdatei konnte nicht geladen werden:\n{exc}")
            return
        self.declaration = declaration
        self.declaration_path = path
        self.declaration_path_var.set(str(path))
        desc = f"{declaration.title}: {declaration.description}".strip(": ")
        self.declaration_info_var.set(desc or declaration.title)
        self._rebuild_metadata_form()
        self._update_start_button_state()

    def _choose_output_dir(self) -> None:
        print("DEBUG: _choose_output_dir gestartet")
        try:
            # Verwende PowerShell für zuverlässigen Ordner-Dialog auf Windows
            if sys.platform.startswith("win"):
                print("DEBUG: Öffne PowerShell Ordner-Dialog...")
                powershell_script = """
                Add-Type -AssemblyName System.Windows.Forms
                $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
                $dialog.Description = "OTBioLab Zielordner auswählen"
                $dialog.ShowNewFolderButton = $true
                $result = $dialog.ShowDialog()
                if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                    Write-Output $dialog.SelectedPath
                }
                """
                result = subprocess.run(
                    ["powershell", "-NoProfile", "-Command", powershell_script],
                    capture_output=True,
                    text=True,
                    timeout=120
                )
                directory_str = result.stdout.strip()
                print(f"DEBUG: Dialog geschlossen, Ergebnis: {directory_str}")
            else:
                # Fallback für andere Systeme
                print("DEBUG: Öffne tkinter filedialog...")
                directory_str = filedialog.askdirectory(title="OTBioLab Zielordner auswählen")
                print(f"DEBUG: Dialog geschlossen, Ergebnis: {directory_str}")

            if directory_str:
                self.output_dir = Path(directory_str)
                self.output_dir_var.set(str(self.output_dir))
                self._update_start_button_state()
                print("DEBUG: Ordner erfolgreich gesetzt")
        except Exception as e:
            print(f"DEBUG: Exception aufgetreten: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Fehler", f"Fehler beim Öffnen des Ordner-Dialogs: {e}")

    def _rebuild_metadata_form(self) -> None:
        for child in self.metadata_fields_frame.winfo_children():
            child.destroy()
        self.metadata_controls.clear()
        if not self.declaration:
            return
        for row, field_cfg in enumerate(self.declaration.metadata_fields):
            label = ttk.Label(self.metadata_fields_frame, text=field_cfg.label + (":" if not field_cfg.label.endswith(":") else ""))
            label.grid(row=row, column=0, sticky="e", padx=(0, 12), pady=4)
            control = self._create_field_control(self.metadata_fields_frame, field_cfg)
            control.widget.grid(row=row, column=1, sticky="ew", pady=4)
            control.bind_on_change(self._update_start_button_state)
            self.metadata_controls[field_cfg.field_id] = control

            # Setze "Mess-Tag" automatisch auf heute (prüfe field_id UND label)
            field_lower = field_cfg.field_id.lower()
            label_lower = field_cfg.label.lower()
            if ("mess" in field_lower and "tag" in field_lower) or ("mess" in label_lower and "tag" in label_lower):
                today = datetime.now().strftime("%d.%m.%Y")
                control.set_value(today)
                print(f"DEBUG: Mess-Tag auf {today} gesetzt")

        # Wende Referenz-Daten an, falls bereits geladen
        if self.reference_data:
            self._apply_reference_data()

    # Schritt-Ansicht ------------------------------------------------------------
    def _build_step_frame(self) -> None:
        frame = self.frames["step"]
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(4, weight=1)

        header = ttk.Frame(frame)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        self.session_title_var = tk.StringVar()
        ttk.Label(header, textvariable=self.session_title_var, font=("Segoe UI", 22, "bold")).grid(row=0, column=0, sticky="w")

        self.total_timer_var = tk.StringVar(value="00:00:00")
        total_label = ttk.Label(header, textvariable=self.total_timer_var, font=("Segoe UI", 32, "bold"))
        total_label.grid(row=0, column=1, sticky="e", padx=(24, 0))

        step_header = ttk.Frame(frame)
        step_header.grid(row=1, column=0, sticky="ew", pady=(18, 6))
        step_header.columnconfigure(0, weight=1)

        self.current_step_var = tk.StringVar()
        ttk.Label(step_header, textvariable=self.current_step_var, font=("Segoe UI", 18)).grid(row=0, column=0, sticky="w")

        self.step_timer_var = tk.StringVar(value="00:00:00")
        # Verwende tk.Label statt ttk.Label, um Farbe ändern zu können
        self.step_timer_label = tk.Label(
            step_header,
            textvariable=self.step_timer_var,
            font=("Segoe UI", 28, "bold"),
            fg="#000000"  # Schwarz als Standard
        )
        self.step_timer_label.grid(row=0, column=1, sticky="e")

        self.step_description_label = ttk.Label(frame, wraplength=900, justify="left")
        self.step_description_label.grid(row=2, column=0, sticky="ew", pady=(6, 6))

        self.expected_duration_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.expected_duration_var).grid(row=3, column=0, sticky="w")

        self.fields_group = ttk.LabelFrame(frame, text="Messwerte")
        self.fields_group.grid(row=4, column=0, sticky="nsew", pady=(18, 12))
        self.fields_group.columnconfigure(0, weight=1)
        self.step_fields_frame = ttk.Frame(self.fields_group)
        self.step_fields_frame.grid(row=0, column=0, sticky="nsew")
        self.step_fields_frame.columnconfigure(1, weight=1)

        notes_container = ttk.Frame(frame)
        notes_container.grid(row=5, column=0, sticky="nsew")
        notes_container.columnconfigure(0, weight=1)
        notes_container.rowconfigure(1, weight=1)

        self.notes_placeholder_var = tk.StringVar(value="Notizen zur Messung")
        ttk.Label(notes_container, textvariable=self.notes_placeholder_var).grid(row=0, column=0, sticky="w")
        self.notes_text = tk.Text(notes_container, height=8, wrap="word")
        self.notes_text.grid(row=1, column=0, sticky="nsew")
        notes_scroll = ttk.Scrollbar(notes_container, orient="vertical", command=self.notes_text.yview)
        notes_scroll.grid(row=1, column=1, sticky="ns")
        self.notes_text.configure(yscrollcommand=notes_scroll.set)

        button_row = ttk.Frame(frame)
        button_row.grid(row=6, column=0, sticky="ew", pady=(18, 6))
        button_row.columnconfigure(2, weight=1)

        self.back_button = ttk.Button(button_row, text="Zurück", command=self._back_to_previous_step)
        self.back_button.grid(row=0, column=0, padx=(0, 8))

        self.trigger_button = ttk.Button(button_row, text="Dateinamen an OTBioLab übergeben", command=self._trigger_otbiolab_save)
        self.trigger_button.grid(row=0, column=1, padx=(0, 8))

        self.next_button = ttk.Button(button_row, text="Weiter", command=self._complete_step)
        self.next_button.grid(row=0, column=3)

        self.status_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.status_var, foreground="#1f6aa5").grid(row=7, column=0, sticky="w", pady=(6, 0))

    # Zusammenfassung ------------------------------------------------------------
    def _build_summary_frame(self) -> None:
        frame = self.frames["summary"]
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        ttk.Label(frame, text="Messung abgeschlossen", font=("Segoe UI", 22, "bold")).grid(row=0, column=0, sticky="w")

        self.summary_info_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.summary_info_var, justify="left").grid(row=1, column=0, sticky="w", pady=(12, 12))

        self.summary_text = tk.Text(frame, wrap="word")
        self.summary_text.grid(row=3, column=0, sticky="nsew")
        summary_scroll = ttk.Scrollbar(frame, orient="vertical", command=self.summary_text.yview)
        summary_scroll.grid(row=3, column=1, sticky="ns")
        self.summary_text.configure(yscrollcommand=summary_scroll.set)
        self.summary_text.configure(state="disabled")

        button_row = ttk.Frame(frame)
        button_row.grid(row=4, column=0, sticky="ew", pady=(18, 0))
        button_row.columnconfigure(1, weight=1)

        self.export_button = ttk.Button(button_row, text="Protokoll exportieren", command=self._export_protocol)
        self.export_button.grid(row=0, column=0, padx=(0, 8))

        ttk.Button(button_row, text="Neue Session", command=self._reset_to_start).grid(row=0, column=2)

    # Formular-Utilities ---------------------------------------------------------
    def _create_field_control(self, parent: ttk.Frame, config: FieldConfig) -> FieldControl:
        kind = config.kind.lower()
        if kind == "multiline":
            widget = tk.Text(parent, height=4, wrap="word")
            return FieldControl(config, widget)
        if kind == "choice":
            variable = tk.StringVar(value=config.options[0] if config.options else "")
            widget = ttk.Combobox(parent, textvariable=variable, values=config.options, state="readonly")
            return FieldControl(config, widget, variable)
        variable = tk.StringVar()
        widget = ttk.Entry(parent, textvariable=variable)
        if config.placeholder:
            widget.insert(0, "")
        return FieldControl(config, widget, variable)

    def _add_repeated_measurement_ui(self, start_row: int, field_cfg: FieldConfig) -> None:
        """Erstellt UI für wiederholbare Messungen mit dynamischen Versuchen."""
        # LabelFrame für diese wiederholbare Messung
        frame = ttk.LabelFrame(self.step_fields_frame, text=field_cfg.label, padding=10)
        frame.grid(row=start_row, column=0, columnspan=4, sticky="ew", pady=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        # Hole existierende Versuche
        attempts = self.current_step_repeated_measurements.get(field_cfg.field_id, [])

        # Container mit Canvas und Scrollbar für scrollbare Versuche
        canvas_container = ttk.Frame(frame)
        canvas_container.grid(row=0, column=0, columnspan=2, sticky="nsew")
        canvas_container.columnconfigure(0, weight=1)
        canvas_container.rowconfigure(0, weight=1)

        # Canvas mit maximaler Höhe
        canvas = tk.Canvas(canvas_container, height=300, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")

        # Scrollbar
        scrollbar = ttk.Scrollbar(canvas_container, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Frame für alle Versuche im Canvas
        attempts_frame = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=attempts_frame, anchor="nw")
        attempts_frame.columnconfigure(1, weight=1)

        # Zeige jeden Versuch
        for attempt_idx, attempt_data in enumerate(attempts):
            self._render_single_attempt(attempts_frame, field_cfg, attempt_idx, attempt_data)

        # Update scroll region und Canvas-Breite
        def _configure_canvas_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Canvas-Breite an Frame-Breite anpassen
            canvas.itemconfig(canvas_window, width=canvas.winfo_width())

        attempts_frame.bind("<Configure>", _configure_canvas_scroll)
        canvas.bind("<Configure>", _configure_canvas_scroll)

        # Mausrad-Scrolling (nur wenn Maus über Canvas)
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _on_enter(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _on_leave(event):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _on_enter)
        canvas.bind("<Leave>", _on_leave)

        # "Versuch hinzufügen" Button
        add_btn = ttk.Button(
            frame,
            text="+ Versuch hinzufügen",
            command=lambda: self._add_new_attempt(field_cfg)
        )
        add_btn.grid(row=1, column=0, columnspan=2, pady=(8, 0), sticky="w")

    def _render_single_attempt(self, parent: ttk.Frame, field_cfg: FieldConfig, attempt_idx: int, attempt_data: Dict[str, Any]) -> None:
        """Rendert einen einzelnen Versuch mit allen Sub-Feldern."""
        # Separator
        if attempt_idx > 0:
            sep = ttk.Separator(parent, orient="horizontal")
            sep.grid(row=attempt_idx * 100, column=0, columnspan=4, sticky="ew", pady=(8, 8))

        # Header: Versuch #N
        header = ttk.Label(parent, text=f"Versuch {attempt_idx + 1}", font=("Segoe UI", 10, "bold"))
        header.grid(row=attempt_idx * 100 + 1, column=0, columnspan=4, sticky="w", pady=(0, 4))

        # Render Sub-Felder
        for sub_row, sub_field_cfg in enumerate(field_cfg.repeated_fields):
            actual_row = attempt_idx * 100 + 2 + sub_row

            # Label
            label = ttk.Label(parent, text=sub_field_cfg.label + ":")
            label.grid(row=actual_row, column=0, sticky="e", padx=(0, 12), pady=2)

            # Control
            control_key = f"{field_cfg.field_id}__attempt{attempt_idx}__{sub_field_cfg.field_id}"
            control = self._create_field_control(parent, sub_field_cfg)
            control.widget.grid(row=actual_row, column=1, sticky="ew", pady=2)

            # Setze gespeicherten Wert
            if sub_field_cfg.field_id in attempt_data:
                control.set_value(attempt_data[sub_field_cfg.field_id])

            self.step_controls[control_key] = control

            # OTBioLab Button (wenn Template vorhanden)
            if field_cfg.otbiolab_template and save_in_word_dialog is not None:
                btn = ttk.Button(
                    parent,
                    text="📁",
                    width=3,
                    command=lambda fid=field_cfg.field_id, idx=attempt_idx: self._trigger_repeated_measurement_otbiolab_save(fid, idx)
                )
                btn.grid(row=actual_row, column=2, sticky="w", padx=(8, 0), pady=2)

                # Zeige Dateiname, wenn vorhanden
                if "otbiolab_file" in attempt_data:
                    file_label = ttk.Label(
                        parent,
                        text=f"✓ {Path(attempt_data['otbiolab_file']).name}",
                        foreground="#1f6aa5"
                    )
                    file_label.grid(row=actual_row, column=3, sticky="w", padx=(4, 0), pady=2)

    def _add_new_attempt(self, field_cfg: FieldConfig) -> None:
        """Fügt einen neuen Versuch hinzu und aktualisiert die UI."""
        if field_cfg.field_id not in self.current_step_repeated_measurements:
            self.current_step_repeated_measurements[field_cfg.field_id] = []

        # Neuen leeren Versuch hinzufügen
        new_attempt = {}
        self.current_step_repeated_measurements[field_cfg.field_id].append(new_attempt)

        # UI neu aufbauen um den neuen Versuch anzuzeigen
        self._show_current_step()

    def _trigger_repeated_measurement_otbiolab_save(self, field_id: str, attempt_idx: int) -> None:
        """Triggered OTBioLab Save für einen spezifischen Versuch."""
        if not self.declaration or save_in_word_dialog is None:
            return

        step = self.declaration.steps[self.current_step_index]
        field_cfg = None
        for f in step.fields:
            if f.field_id == field_id:
                field_cfg = f
                break

        if not field_cfg or not field_cfg.otbiolab_template:
            return

        # Baue Template-Context
        context = self._build_template_context(step)
        context["field_id"] = field_id
        context["attempt_number"] = attempt_idx + 1  # 1-based für Dateinamen

        filename = field_cfg.otbiolab_template.format(**context)
        full_path = self.output_dir / filename

        self.status_var.set(f"Warte auf OTBioLab Speichern-Dialog für Versuch {attempt_idx + 1}...")
        self.update_idletasks()

        def callback():
            success = False
            message = "Timeout oder Fehler beim Speichern."
            try:
                success = save_in_word_dialog(str(full_path), timeout=30)
                if success:
                    message = f"Datei gespeichert: {filename}"
            except Exception as exc:
                message = f"Fehler: {exc}"

            self.after(0, lambda: self._on_repeated_measurement_interceptor_finished(success, message, full_path, field_id, attempt_idx))

        threading.Thread(target=callback, daemon=True).start()

    def _on_repeated_measurement_interceptor_finished(self, success: bool, message: str, path: Path, field_id: str, attempt_idx: int) -> None:
        """Callback nach OTBioLab Save für wiederholbare Messung."""
        if success:
            # Speichere Dateiname in attempt_data
            attempts = self.current_step_repeated_measurements.get(field_id, [])
            if attempt_idx < len(attempts):
                attempts[attempt_idx]["otbiolab_file"] = str(path)

            # Speichere alle aktuellen Werte vor UI-Rebuild
            step = self.declaration.steps[self.current_step_index]
            self._save_repeated_measurement_values(step)

            # UI neu aufbauen
            self._show_current_step()

            self.status_var.set(f"✓ {message}")
        else:
            self.status_var.set(f"✗ {message}")

    def _save_repeated_measurement_values(self, step: StepConfig) -> None:
        """Speichert alle Werte aus wiederholbaren Messungen vor UI-Rebuild."""
        for field_cfg in step.fields:
            if not field_cfg.repeated_measurement:
                continue

            attempts = self.current_step_repeated_measurements.get(field_cfg.field_id, [])
            for attempt_idx in range(len(attempts)):
                for sub_field_cfg in field_cfg.repeated_fields:
                    control_key = f"{field_cfg.field_id}__attempt{attempt_idx}__{sub_field_cfg.field_id}"
                    control = self.step_controls.get(control_key)
                    if control:
                        attempts[attempt_idx][sub_field_cfg.field_id] = control.get_value()

    def _update_start_button_state(self) -> None:
        ready = self.declaration is not None and self.output_dir is not None
        if ready and self.declaration:
            for field_cfg in self.declaration.metadata_fields:
                control = self.metadata_controls.get(field_cfg.field_id)
                value = control.get_value() if control else ""
                if field_cfg.required and not value:
                    ready = False
                    break
        self.start_button.configure(state=("normal" if ready else "disabled"))

    def _coerce_value(self, field_cfg: FieldConfig, value: str) -> Tuple[Any, bool]:
        if value == "":
            return ("", not field_cfg.required)
        if field_cfg.kind == "integer":
            try:
                return int(value), True
            except ValueError:
                return value, False
        if field_cfg.kind == "float":
            try:
                safe_value = value.replace(",", ".")
                return float(safe_value), True
            except ValueError:
                return value, False
        return value, True

    # Session-Steuerung ----------------------------------------------------------
    def _start_session(self) -> None:
        if not self.declaration:
            return
        metadata_values: Dict[str, Any] = {}
        for field_cfg in self.declaration.metadata_fields:
            control = self.metadata_controls.get(field_cfg.field_id)
            raw_value = control.get_value() if control else ""
            typed_value, ok = self._coerce_value(field_cfg, raw_value)
            if not ok:
                messagebox.showwarning("Eingabe prüfen", f"Bitte gültigen Wert für '{field_cfg.label}' eintragen.")
                return
            if field_cfg.required and raw_value == "":
                messagebox.showwarning("Eingabe fehlt", f"Bitte Feld '{field_cfg.label}' ausfüllen.")
                return
            metadata_values[field_cfg.field_id] = typed_value if raw_value != "" else ""
        self.metadata_values = metadata_values
        self.session_started_at = datetime.now()
        self.session_timestamp = self.session_started_at.strftime("%Y%m%d_%H%M%S")
        self.current_step_index = 0
        self.current_step_started_at = datetime.now()
        self.step_results = []
        self.current_step_otbiolab_paths = []
        self.current_step_field_otbiolab_files = {}
        self.current_step_repeated_measurements = {}
        self.session_title_var.set(self.declaration.title)
        self._schedule_timer(reset=True)
        self._show_frame("step")
        self._show_current_step()

    def _schedule_timer(self, reset: bool = False) -> None:
        if reset and self._timer_after_id:
            self.after_cancel(self._timer_after_id)
            self._timer_after_id = None
        self._update_timer_labels()
        self._timer_after_id = self.after(500, self._schedule_timer)

    def _update_timer_labels(self) -> None:
        now = datetime.now()

        # Gesamt-Timer: Aktuelle Zeit / Erwartete Gesamtdauer
        if self.session_started_at and self.declaration:
            total_seconds = (now - self.session_started_at).total_seconds()

            # Berechne erwartete Gesamtdauer
            expected_total = sum(
                step.expected_duration_seconds or 0
                for step in self.declaration.steps
            )

            current_time = seconds_to_clock(total_seconds)
            expected_time = seconds_to_clock(expected_total)
            self.total_timer_var.set(f"{current_time}/{expected_time}")
        else:
            self.total_timer_var.set("00:00:00")

        # Schritt-Timer: MM:SS Format
        if self.current_step_started_at and self.declaration:
            step_seconds = (now - self.current_step_started_at).total_seconds()

            # Prüfe ob Zeit überschritten
            step = self.declaration.steps[self.current_step_index]
            expected = step.expected_duration_seconds

            if expected and step_seconds > expected:
                # Zeit überschritten - zeige Überschreitung und färbe rot
                overtime = step_seconds - expected
                self.step_timer_var.set(f"-{seconds_to_minutes_clock(overtime)}")
                self.step_timer_label.config(fg="#DC143C")  # Crimson Red
            else:
                # Normal - zeige Zeit und färbe schwarz
                self.step_timer_var.set(seconds_to_minutes_clock(step_seconds))
                self.step_timer_label.config(fg="#000000")  # Schwarz
        else:
            self.step_timer_var.set("00:00")
            if hasattr(self, 'step_timer_label'):
                self.step_timer_label.config(fg="#000000")

    def _show_current_step(self) -> None:
        if not self.declaration:
            return
        step = self.declaration.steps[self.current_step_index]
        total_steps = len(self.declaration.steps)
        self.current_step_var.set(f"Schritt {self.current_step_index + 1}/{total_steps} – {step.title}")
        self.step_description_label.configure(text=step.description)
        if step.expected_duration_seconds:
            expected = seconds_to_minutes_clock(step.expected_duration_seconds)
            self.expected_duration_var.set(f"Erwartete Dauer: {expected}")
        else:
            self.expected_duration_var.set("")

        for child in self.step_fields_frame.winfo_children():
            child.destroy()
        self.step_controls.clear()

        # Konfiguriere Spalten: Label (0), Control (1), OTBioLab-Button (2)
        self.step_fields_frame.columnconfigure(1, weight=1)

        row = 0
        for field_cfg in step.fields:
            if field_cfg.repeated_measurement:
                # Wiederholbare Messung - spezielle UI
                self._add_repeated_measurement_ui(row, field_cfg)
                row += 1  # Repeated measurement nimmt eine Zeile (wird intern erweitert)
            else:
                # Normales Feld
                # Label
                label = ttk.Label(self.step_fields_frame, text=field_cfg.label + (":" if not field_cfg.label.endswith(":") else ""))
                label.grid(row=row, column=0, sticky="e", padx=(0, 12), pady=4)

                # Control (Input-Feld)
                control = self._create_field_control(self.step_fields_frame, field_cfg)
                control.widget.grid(row=row, column=1, sticky="ew", pady=4)
                self.step_controls[field_cfg.field_id] = control

                # OTBioLab-Button (wenn Feld ein Template hat)
                if field_cfg.otbiolab_template and save_in_word_dialog is not None:
                    btn = ttk.Button(
                        self.step_fields_frame,
                        text="📁",  # Datei-Symbol
                        width=3,
                        command=lambda fid=field_cfg.field_id: self._trigger_field_otbiolab_save(fid)
                    )
                    btn.grid(row=row, column=2, sticky="w", padx=(8, 0), pady=4)

                    # Zeige Anzahl der bereits übergebenen Dateien
                    if field_cfg.field_id in self.current_step_field_otbiolab_files:
                        count = len(self.current_step_field_otbiolab_files[field_cfg.field_id])
                        if count > 0:
                            count_label = ttk.Label(
                                self.step_fields_frame,
                                text=f"({count})",
                                foreground="#1f6aa5"
                            )
                            count_label.grid(row=row, column=3, sticky="w", padx=(4, 0), pady=4)
                row += 1

        # Wende Referenz-Daten an (falls vorhanden und Felder mit use_from_ref=true)
        if self.reference_data and "steps" in self.reference_data:
            ref_steps = self.reference_data.get("steps", {})
            if step.step_id in ref_steps:
                ref_step_values = ref_steps[step.step_id].get("values", {})
                for field_cfg in step.fields:
                    if field_cfg.use_from_ref and field_cfg.field_id in ref_step_values:
                        control = self.step_controls.get(field_cfg.field_id)
                        if control:
                            control.set_value(ref_step_values[field_cfg.field_id])
                            print(f"DEBUG: Referenz-Wert übernommen für Schritt '{step.step_id}', Feld '{field_cfg.field_id}': {ref_step_values[field_cfg.field_id]}")

        existing = self.step_results[self.current_step_index] if len(self.step_results) > self.current_step_index else None
        if existing:
            for field_cfg in step.fields:
                if not field_cfg.repeated_measurement:
                    control = self.step_controls.get(field_cfg.field_id)
                    if control:
                        control.set_value(existing.values.get(field_cfg.field_id, ""))
            self.notes_text.delete("1.0", "end")
            self.notes_text.insert("1.0", existing.notes or "")
            self.current_step_otbiolab_paths = list(existing.otbiolab_paths)  # Liste kopieren (legacy)
            self.current_step_field_otbiolab_files = {k: list(v) for k, v in existing.field_otbiolab_files.items()}  # Deep copy
            # Restore repeated measurements
            self.current_step_repeated_measurements = {k: [dict(attempt) for attempt in v] for k, v in existing.repeated_measurements.items()}  # Deep copy
        else:
            # Wenn wir zu einem neuen Schritt gewechselt sind (erkennbar an leeren Dictionaries), lösche Notizen
            # Beim UI-Rebuild während der Bearbeitung bleiben die Werte erhalten
            if not self.current_step_repeated_measurements and not self.current_step_field_otbiolab_files and not self.current_step_otbiolab_paths:
                self.notes_text.delete("1.0", "end")

        placeholder = step.notes_placeholder or "Notizen zur Messung"
        self.notes_placeholder_var.set(placeholder)
        self.notes_text.edit_modified(False)
        self.back_button.configure(state=("normal" if self.current_step_index > 0 else "disabled"))
        trigger_allowed = bool(step.otbiolab_template) and save_in_word_dialog is not None
        self.trigger_button.configure(state=("normal" if trigger_allowed else "disabled"))
        self.status_var.set("")
        self.current_step_started_at = datetime.now()

    def _complete_step(self) -> None:
        if not self.declaration:
            return
        step = self.declaration.steps[self.current_step_index]

        # Validiere und sammle normale Feld-Werte
        values: Dict[str, Any] = {}
        for field_cfg in step.fields:
            if field_cfg.repeated_measurement:
                # Überspringe wiederholbare Messungen - die werden separat gespeichert
                continue

            control = self.step_controls.get(field_cfg.field_id)
            raw_value = control.get_value() if control else ""
            typed_value, ok = self._coerce_value(field_cfg, raw_value)
            if field_cfg.required and raw_value == "":
                messagebox.showwarning("Eingabe fehlt", f"Bitte Feld '{field_cfg.label}' ausfüllen.")
                return
            if not ok:
                messagebox.showwarning("Eingabe prüfen", f"Bitte gültigen Wert für '{field_cfg.label}' eintragen.")
                return
            values[field_cfg.field_id] = typed_value if raw_value != "" else ""

        # Speichere aktuelle Werte aus wiederholbaren Messungen
        self._save_repeated_measurement_values(step)

        notes_text = self.notes_text.get("1.0", "end-1c").strip()
        completed_at = datetime.now()
        duration = completed_at - self.current_step_started_at if self.current_step_started_at else None
        result = StepResult(
            config=step,
            started_at=self.current_step_started_at,
            completed_at=completed_at,
            duration=duration,
            values=values,
            notes=notes_text,
            otbiolab_paths=list(self.current_step_otbiolab_paths),  # Liste kopieren (legacy)
            field_otbiolab_files={k: list(v) for k, v in self.current_step_field_otbiolab_files.items()},  # Deep copy
            repeated_measurements={k: [dict(attempt) for attempt in v] for k, v in self.current_step_repeated_measurements.items()},  # Deep copy
        )

        if len(self.step_results) > self.current_step_index:
            self.step_results[self.current_step_index] = result
        else:
            self.step_results.append(result)

        if self.current_step_index + 1 < len(self.declaration.steps):
            self.current_step_index += 1
            self.current_step_started_at = datetime.now()
            self.current_step_otbiolab_paths = []  # Liste für nächsten Schritt zurücksetzen (legacy)
            self.current_step_field_otbiolab_files = {}  # Dictionary für nächsten Schritt zurücksetzen
            self.current_step_repeated_measurements = {}  # Dictionary für nächsten Schritt zurücksetzen
            self._show_current_step()
        else:
            self._finish_session()

    def _back_to_previous_step(self) -> None:
        if self.current_step_index <= 0:
            return
        self.current_step_index -= 1
        self.current_step_started_at = datetime.now()
        self._show_current_step()

    def _finish_session(self) -> None:
        if self._timer_after_id:
            self.after_cancel(self._timer_after_id)
            self._timer_after_id = None
        total_duration = datetime.now() - self.session_started_at if self.session_started_at else None
        protocol_text = self._build_protocol_text(total_duration)
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", protocol_text)
        self.summary_text.configure(state="disabled")
        pid = self.metadata_values.get("pid", "unbekannt")
        info_lines = [
            f"PID: {pid}",
            f"Dauer gesamt: {seconds_to_clock(total_duration.total_seconds()) if total_duration else 'n/a'}",
            f"Schritte dokumentiert: {len(self.step_results)}",
        ]
        self.summary_info_var.set("\n".join(info_lines))
        self.session_finished = True

        # Automatisches Speichern des Protokolls und der Referenz-Datei
        self._auto_save_protocol(protocol_text)
        self._save_reference_file()

        self._show_frame("summary")

    def _build_protocol_text(self, total_duration: Optional[timedelta]) -> str:
        lines: List[str] = []
        session_time = self.session_started_at.isoformat(sep=" ", timespec="seconds") if self.session_started_at else "-"
        lines.append("=" * 70)
        lines.append("HDsEMG VERSUCHSREIHE PROTOKOLL")
        lines.append("=" * 70)
        lines.append(f"Session gestartet: {session_time}")
        session_end_time = datetime.now().isoformat(sep=" ", timespec="seconds")
        lines.append(f"Session beendet:   {session_end_time}")
        if total_duration:
            lines.append(f"Gesamtdauer:       {seconds_to_clock(total_duration.total_seconds())}")
        lines.append("")
        lines.append("Metadaten:")
        lines.append("-" * 70)
        for key, value in self.metadata_values.items():
            lines.append(f"  {key}: {value}")
        lines.append("")
        lines.append("Schritte:")
        lines.append("=" * 70)
        for idx, result in enumerate(self.step_results, start=1):
            lines.append("")
            lines.append(f"[Schritt {idx}] {result.config.title} ({result.config.step_id})")
            lines.append("-" * 70)
            if result.started_at:
                lines.append(f"    ⏱ Start:      {result.started_at.strftime('%Y-%m-%d %H:%M:%S')}")
            if result.completed_at:
                lines.append(f"    ⏱ Ende:       {result.completed_at.strftime('%Y-%m-%d %H:%M:%S')}")
                lines.append(f"    ✓ Weiter gedrückt um: {result.completed_at.strftime('%H:%M:%S')}")
            if result.duration:
                lines.append(f"    ⌛ Dauer:      {seconds_to_clock(result.duration.total_seconds())}")
            if result.config.expected_duration_seconds:
                lines.append(f"    📋 Erwartet:   {seconds_to_clock(result.config.expected_duration_seconds)}")
            if result.values:
                lines.append("")
                lines.append("    Eingaben:")
                for field_cfg in result.config.fields:
                    if field_cfg.repeated_measurement:
                        # Überspringe - wiederholbare Messungen werden separat ausgegeben
                        continue

                    value = result.values.get(field_cfg.field_id, "")
                    lines.append(f"      • {field_cfg.label}: {value}")

                    # Zeige OTBioLab-Dateien für dieses Feld (falls vorhanden)
                    if field_cfg.field_id in result.field_otbiolab_files:
                        field_files = result.field_otbiolab_files[field_cfg.field_id]
                        if field_files:
                            if len(field_files) == 1:
                                lines.append(f"        💾 OTBioLab: {field_files[0]}")
                            else:
                                lines.append(f"        💾 OTBioLab-Dateien ({len(field_files)}):")
                                for i, otb_path in enumerate(field_files, start=1):
                                    lines.append(f"           {i}. {otb_path}")

            # Zeige wiederholbare Messungen
            if result.repeated_measurements:
                lines.append("")
                lines.append("    Wiederholbare Messungen:")
                for field_cfg in result.config.fields:
                    if not field_cfg.repeated_measurement:
                        continue

                    if field_cfg.field_id not in result.repeated_measurements:
                        continue

                    attempts = result.repeated_measurements[field_cfg.field_id]
                    if not attempts:
                        continue

                    lines.append(f"      ▸ {field_cfg.label}:")
                    for attempt_idx, attempt_data in enumerate(attempts, start=1):
                        lines.append(f"        Versuch {attempt_idx}:")
                        for sub_field_cfg in field_cfg.repeated_fields:
                            sub_value = attempt_data.get(sub_field_cfg.field_id, "")
                            lines.append(f"          • {sub_field_cfg.label}: {sub_value}")

                        # Zeige OTBioLab-Datei für diesen Versuch
                        if "otbiolab_file" in attempt_data:
                            lines.append(f"          💾 OTBioLab: {attempt_data['otbiolab_file']}")

            if result.notes:
                lines.append("")
                lines.append(f"    📝 Notizen:")
                for line in result.notes.split("\n"):
                    lines.append(f"       {line}")

            # Schritt-basierte OTBioLab-Dateien (legacy)
            if result.otbiolab_paths:
                lines.append("")
                if len(result.otbiolab_paths) == 1:
                    lines.append(f"    💾 OTBioLab-Datei (Schritt): {result.otbiolab_paths[0]}")
                else:
                    lines.append(f"    💾 OTBioLab-Dateien (Schritt, {len(result.otbiolab_paths)}):")
                    for i, otb_path in enumerate(result.otbiolab_paths, start=1):
                        lines.append(f"       {i}. {otb_path}")

            lines.append("")
        lines.append("=" * 70)
        lines.append("Ende des Protokolls")
        lines.append("=" * 70)
        return "\n".join(lines)

    # OTBioLab Integration -------------------------------------------------------
    def _trigger_otbiolab_save(self) -> None:
        if not self.declaration or not self.output_dir:
            return
        step = self.declaration.steps[self.current_step_index]
        if not step.otbiolab_template:
            messagebox.showinfo("Kein Dateiname", "Dieser Schritt erwartet keine OTBioLab Datei.")
            return
        if save_in_word_dialog is None:
            messagebox.showwarning("Interceptor fehlt", "Das OTBioLab Skript konnte nicht geladen werden.")
            return

        context = self._build_template_context(step)
        try:
            filename = step.otbiolab_template.format(**context)
        except KeyError as exc:
            messagebox.showerror("Template Fehler", f"Platzhalter {exc} konnte nicht gefüllt werden.")
            return
        if not filename.lower().endswith(".otb4"):
            filename += ".otb4"
        target_path = (self.output_dir / filename).resolve()
        target_path.parent.mkdir(parents=True, exist_ok=True)

        self.status_var.set(f"Übergebe Dateiname an OTBioLab… ({target_path.name})")
        self.trigger_button.configure(state="disabled")
        self.next_button.configure(state="disabled")

        # Hintergrundthread, damit die UI flüssig bleibt.
        def worker() -> None:
            try:
                success = save_in_word_dialog(str(target_path), timeout=25)
                message = "Dateiname übergeben." if success else "Speichern-Dialog wurde nicht gefunden."
            except Exception as exc:
                success = False
                message = f"Fehler beim Zugriff auf den Speichern-Dialog: {exc}"
            self.after(0, lambda: self._on_interceptor_finished(success, message, target_path))

        threading.Thread(target=worker, daemon=True).start()

    def _trigger_field_otbiolab_save(self, field_id: str) -> None:
        """Übergebe OTBioLab-Dateiname für ein spezifisches Feld."""
        if not self.declaration or not self.output_dir:
            return
        step = self.declaration.steps[self.current_step_index]

        # Finde das Feld
        field_cfg = None
        for f in step.fields:
            if f.field_id == field_id:
                field_cfg = f
                break

        if not field_cfg or not field_cfg.otbiolab_template:
            messagebox.showinfo("Kein Dateiname", "Dieses Feld hat keine OTBioLab-Datei-Vorlage.")
            return

        if save_in_word_dialog is None:
            messagebox.showwarning("Interceptor fehlt", "Das OTBioLab Skript konnte nicht geladen werden.")
            return

        # Baue Kontext mit Feld-spezifischen Informationen
        context = self._build_template_context(step)
        context["field_id"] = field_id
        context["field_label"] = field_cfg.label

        # Zähle bereits vorhandene Dateien für dieses Feld
        existing_count = len(self.current_step_field_otbiolab_files.get(field_id, []))
        context["file_number"] = existing_count + 1

        try:
            filename = field_cfg.otbiolab_template.format(**context)
        except KeyError as exc:
            messagebox.showerror("Template Fehler", f"Platzhalter {exc} konnte nicht gefüllt werden.")
            return

        if not filename.lower().endswith(".otb4"):
            filename += ".otb4"
        target_path = (self.output_dir / filename).resolve()
        target_path.parent.mkdir(parents=True, exist_ok=True)

        self.status_var.set(f"Übergebe Dateiname für '{field_cfg.label}' an OTBioLab… ({target_path.name})")

        # Hintergrundthread
        def worker() -> None:
            try:
                success = save_in_word_dialog(str(target_path), timeout=25)
                message = f"Dateiname für '{field_cfg.label}' übergeben." if success else "Speichern-Dialog wurde nicht gefunden."
            except Exception as exc:
                success = False
                message = f"Fehler beim Zugriff auf den Speichern-Dialog: {exc}"
            self.after(0, lambda: self._on_field_interceptor_finished(success, message, target_path, field_id))

        threading.Thread(target=worker, daemon=True).start()

    def _build_template_context(self, step: StepConfig) -> Dict[str, Any]:
        context = dict(self.metadata_values)
        context["pid"] = context.get("pid", "PID")
        context["step_id"] = step.step_id
        context["step_title"] = step.title
        context["step_index"] = self.current_step_index + 1
        context["timestamp"] = datetime.now().strftime("%Y%m%d_%H%M%S")
        context["session_timestamp"] = self.session_timestamp or context["timestamp"]
        return context

    def _on_interceptor_finished(self, success: bool, message: str, path: Path) -> None:
        self.trigger_button.configure(state="normal")
        self.next_button.configure(state="normal")
        self.status_var.set(message)
        if success:
            # Füge Pfad zur Liste hinzu
            self.current_step_otbiolab_paths.append(str(path))
            # Zeige Anzahl der übergebenen Dateien an
            count = len(self.current_step_otbiolab_paths)
            self.status_var.set(f"{message} ({count} Datei{'en' if count != 1 else ''} für diesen Schritt)")

    def _on_field_interceptor_finished(self, success: bool, message: str, path: Path, field_id: str) -> None:
        """Callback nach Feld-spezifischer OTBioLab-Dateiübergabe."""
        self.status_var.set(message)
        if success:
            # Füge Pfad zur Feld-spezifischen Liste hinzu
            if field_id not in self.current_step_field_otbiolab_files:
                self.current_step_field_otbiolab_files[field_id] = []
            self.current_step_field_otbiolab_files[field_id].append(str(path))

            # Zeige Anzahl der übergebenen Dateien für dieses Feld
            count = len(self.current_step_field_otbiolab_files[field_id])
            self.status_var.set(f"{message} ({count} Datei{'en' if count != 1 else ''} für dieses Feld)")

            # Speichere aktuelle Werte vor UI-Update
            if not self.declaration:
                return
            step = self.declaration.steps[self.current_step_index]
            current_values = {}
            for field_cfg in step.fields:
                control = self.step_controls.get(field_cfg.field_id)
                if control:
                    current_values[field_cfg.field_id] = control.get_value()

            # Aktualisiere die UI, um die Anzahl neben dem Button anzuzeigen
            self._show_current_step()

            # Stelle gespeicherte Werte wieder her
            for field_cfg in step.fields:
                control = self.step_controls.get(field_cfg.field_id)
                if control and field_cfg.field_id in current_values:
                    control.set_value(current_values[field_cfg.field_id])

    # Zusammenfassung/Export ------------------------------------------------------
    def _export_protocol(self) -> None:
        if not self.output_dir:
            messagebox.showwarning("Kein Ordner", "Es wurde kein Zielordner ausgewählt.")
            return
        protocol_dir = self.output_dir / "protokolle"
        protocol_dir.mkdir(parents=True, exist_ok=True)
        pid = self.metadata_values.get("pid", "PID")
        filename = f"{pid}_{self.session_timestamp}_protokoll.txt"
        target_path = protocol_dir / filename
        content = self.summary_text.get("1.0", "end-1c")
        target_path.write_text(content, encoding="utf-8")
        messagebox.showinfo("Protokoll gespeichert", f"Protokoll gespeichert unter:\n{target_path}")

    def _reset_to_start(self) -> None:
        if self._timer_after_id:
            self.after_cancel(self._timer_after_id)
            self._timer_after_id = None
        self.session_started_at = None
        self.current_step_started_at = None
        self.current_step_index = -1
        self.step_results = []
        self.current_step_otbiolab_paths = []
        self.current_step_field_otbiolab_files = {}
        self.metadata_values = {}
        self.session_finished = False
        self.total_timer_var.set("00:00:00")
        self.step_timer_var.set("00:00:00")
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.configure(state="disabled")
        self.summary_info_var.set("")
        for control in self.metadata_controls.values():
            control.set_value("")
        self._update_start_button_state()
        self._show_frame("start")

    # Referenz-File Methods -------------------------------------------------------
    def _load_reference_file(self) -> None:
        """Lädt eine Referenz-Datei und füllt Felder mit use_from_ref=true aus."""
        initial_dir = str(self.output_dir.resolve()) if self.output_dir else str(ROOT_DIR)
        path_str = filedialog.askopenfilename(
            parent=self,
            title="Referenz-Datei auswählen",
            initialdir=initial_dir,
            filetypes=[("JSON Dateien", "*.json"), ("Alle Dateien", "*.*")],
        )
        if not path_str:
            return

        try:
            path = Path(path_str)
            data = json.loads(path.read_text(encoding="utf-8"))

            # Validierung
            if "metadata" not in data:
                messagebox.showerror("Fehler", "Ungültige Referenz-Datei: 'metadata' fehlt.")
                return

            self.reference_data = data
            self.reference_file_var.set(str(path))

            # Felder mit use_from_ref=true automatisch ausfüllen
            self._apply_reference_data()

            messagebox.showinfo("Erfolg", f"Referenz-Datei geladen:\n{path.name}")
        except Exception as exc:
            messagebox.showerror("Fehler", f"Referenz-Datei konnte nicht geladen werden:\n{exc}")

    def _apply_reference_data(self) -> None:
        """Wendet Referenz-Daten auf Felder mit use_from_ref=true an."""
        if not self.reference_data or not self.declaration:
            return

        ref_metadata = self.reference_data.get("metadata", {})

        # Wende auf Metadata-Felder an
        for field_cfg in self.declaration.metadata_fields:
            if field_cfg.use_from_ref and field_cfg.field_id in ref_metadata:
                control = self.metadata_controls.get(field_cfg.field_id)
                if control:
                    control.set_value(ref_metadata[field_cfg.field_id])
                    print(f"DEBUG: Referenz-Wert übernommen für '{field_cfg.field_id}': {ref_metadata[field_cfg.field_id]}")

    def _save_reference_file(self) -> None:
        """Speichert Session-Daten als Referenz-File (wird beim Session-Ende aufgerufen)."""
        if not self.output_dir:
            return

        try:
            ref_dir = self.output_dir / "referenzen"
            ref_dir.mkdir(parents=True, exist_ok=True)

            pid = self.metadata_values.get("pid", "PID")
            filename = f"{pid}_{self.session_timestamp}_referenz.json"
            target_path = ref_dir / filename

            # Sammle Schritt-Daten
            steps_data = {}
            for result in self.step_results:
                steps_data[result.config.step_id] = {
                    "values": dict(result.values),
                    "notes": result.notes,
                }

            # Erstelle Referenz-Daten-Struktur
            ref_data = {
                "session_timestamp": self.session_timestamp,
                "session_started_at": self.session_started_at.isoformat() if self.session_started_at else None,
                "metadata": dict(self.metadata_values),
                "steps": steps_data,
                "declaration_title": self.declaration.title if self.declaration else None,
            }

            target_path.write_text(json.dumps(ref_data, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"INFO: Referenz-Datei gespeichert: {target_path}")
        except Exception as exc:
            print(f"FEHLER: Speichern der Referenz-Datei fehlgeschlagen: {exc}")

    # Utility Methods -------------------------------------------------------------
    def _auto_save_protocol(self, protocol_text: str) -> None:
        """Automatisches Speichern des Protokolls nach Messungsabschluss."""
        if not self.output_dir:
            print("WARNUNG: Kein Ausgabeordner gesetzt - Protokoll wird nicht automatisch gespeichert.")
            return

        try:
            protocol_dir = self.output_dir / "protokolle"
            protocol_dir.mkdir(parents=True, exist_ok=True)
            pid = self.metadata_values.get("pid", "PID")
            filename = f"{pid}_{self.session_timestamp}_protokoll.txt"
            target_path = protocol_dir / filename
            target_path.write_text(protocol_text, encoding="utf-8")
            print(f"INFO: Protokoll automatisch gespeichert: {target_path}")
        except Exception as exc:
            print(f"FEHLER: Automatisches Speichern des Protokolls fehlgeschlagen: {exc}")

    def _on_closing(self) -> None:
        """Handle window close event - zeige Warnung wenn Messung läuft."""
        # Prüfe ob eine Messung läuft
        if self.session_started_at and not self.session_finished:
            # Messung läuft - zeige Warnung
            response = messagebox.askyesno(
                "Messung läuft",
                "Eine Messung ist noch nicht abgeschlossen!\n\n"
                "Möchten Sie die Anwendung wirklich beenden?\n"
                "Alle nicht gespeicherten Daten gehen verloren.",
                icon="warning"
            )
            if not response:
                # Benutzer hat "Nein" gewählt - Schließen abbrechen
                return

        # Beenden
        if self._timer_after_id:
            self.after_cancel(self._timer_after_id)
        self.destroy()


def main() -> None:
    app = SessionApp()
    app.mainloop()


if __name__ == "__main__":
    main()
