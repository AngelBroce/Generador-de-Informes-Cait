"""
Aplicación principal con interfaz gráfica
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk
from pathlib import Path
from PIL import Image, ImageTk
from tkcalendar import DateEntry
import os
import re
import sys
import json
import shutil
import subprocess
import platform
from datetime import datetime

# Asegurar que el proyecto raíz esté en sys.path para permitir "import src.*"
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.core.report_outline import get_content_outline
from src.core.result_schemes import RESULT_SCHEMES
from src.services.evaluators_repository import EvaluatorRepository, build_technical_details
from src.services.counterparts_repository import CounterpartRepository


class MainApplication:
    """Aplicación principal de interfaz gráfica"""
    
    def __init__(self):
        """Inicializa la aplicación"""
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("green")
        ctk.set_widget_scaling(1.25)
        ctk.set_window_scaling(1.1)

        self.colors = {
            "bg": "#F7F8FA",
            "surface": "#FFFFFF",
            "border": "#EEF0F2",
            "primary": "#1B7A3D",
            "primary_dark": "#15632F",
            "primary_muted": "#E8F5EC",
            "field_bg": "#F3F6F4",
            "field_border": "#D8E2DD",
            "chip_bg": "#E6F3EA",
            "chip_text": "#1B7A3D",
            "error": "#D92D20",
            "text": "#1A1D21",
            "text_muted": "#5F6B7A",
            "text_light": "#C8E6C9",
            "info_bg": "#EBF5FB",
            "info_border": "#BEE0F5",
            "info_text": "#1565C0",
        }

        self.root = ctk.CTk()
        self.root.title("Generador de Informes Clínicos - CAIT Panamá")
        self.root.geometry("1500x980")
        self.root.minsize(1150, 780)
        self.root.configure(fg_color=self.colors["bg"])
        self.date_validation_cmd = self.root.register(self._validate_date_input)
        self.numeric_validation_cmd = self.root.register(self._validate_numeric_input)
        
        # Datos del informe actual
        self.current_report = None
        self.logo_path = None
        self.report_type = None
        self.active_section = None
        self.evaluated_entries = {key: [] for key in RESULT_SCHEMES}
        self.result_blocks = {}
        self.results_section_name = "Resultados de las pruebas"
        self.results_canvas = None
        self.result_palettes = {
            scheme_key: {
                option["label"]: {
                    "code": option["key"],
                    "label": option["label"],
                    "bg": option.get("bg", "#FFFFFF"),
                    "fg": option.get("fg", "#000000"),
                }
                for option in scheme.get("options", [])
            }
            for scheme_key, scheme in RESULT_SCHEMES.items()
        }
        self.conclusion_text_widget = None
        self.conclusion_section_name = "Conclusión"
        self.recommendations_text_widget = None
        self.recommendations_section_name = "Recomendaciones"
        self.calibration_section_name = "Certificados de calibración"
        self.calibration_files = []
        self.calibration_tree = None
        self.report_attachments_section_name = "Adjuntos de resultados"
        self.test_attachment_files = {"audiometria": [], "espirometria": []}
        self.test_attachment_trees = {}
        self.test_attachment_buttons = {}
        self.test_attachment_status = {}
        self.test_dataset_labels = {
            "audiometria": "Audiometría",
            "espirometria": "Espirometría",
        }
        self.attendance_section_name = "Listados de asistencia"
        self.attendance_files = []
        self.attendance_tree = None
        self.attendance_status_label = None

        # Catálogo de evaluadores cargado desde JSON
        self.evaluators_repo = EvaluatorRepository()
        self.evaluator_profiles = {}
        self.selected_evaluator_id = tk.StringVar()
        self.evaluator_var = tk.StringVar()
        self.evaluator_combo = None
        self.evaluator_detail_label = None

        # Catálogo de contrapartes técnicas
        self.counterparts_repo = CounterpartRepository()
        self.counterpart_profiles = {}
        self.selected_counterpart_id = tk.StringVar()
        self.counterpart_var = tk.StringVar(value="Sin contraparte")
        self.counterpart_combo = None
        self.counterpart_role_entry = None
        
        # Ruta base del proyecto
        self.project_root = Path(__file__).parent.parent.parent

        # Rutas de icono/logo de la aplicación
        runtime_base = Path(getattr(sys, "_MEIPASS", self.project_root))
        self.app_icon_path = runtime_base / "logo-apli-removebg-preview.ico"
        if not self.app_icon_path.exists():
            self.app_icon_path = self.project_root / "logo-apli-removebg-preview.ico"
        self.app_logo_path = Path(__file__).parent.parent / "assets" / "logo_cait.png"
        self.header_logo_path = self.app_icon_path if self.app_icon_path.exists() else self.app_logo_path
        if self.app_icon_path.exists():
            try:
                self.root.iconbitmap(default=str(self.app_icon_path))
            except Exception:
                pass
        
        self.setup_styles()
        self.setup_ui()
        self._purge_old_drafts()
    
    def setup_styles(self):
        """Configura los estilos de la aplicación."""
        style = ttk.Style(self.root)
        style.theme_use("clam")

        bg = self.colors["bg"]
        surface = self.colors["surface"]
        border = self.colors["border"]
        primary = self.colors["primary"]
        primary_dark = self.colors["primary_dark"]
        primary_muted = self.colors["primary_muted"]
        text = self.colors["text"]
        muted = self.colors["text_muted"]
        field_bg = self.colors["field_bg"]

        style.configure("TFrame", background=surface)
        style.configure("Header.TFrame", background=primary)
        style.configure(
            "Header.TLabel",
            background=primary,
            foreground=surface,
            font=("Segoe UI", 20, "bold"),
        )
        style.configure("Title.TLabel", font=("Segoe UI", 17, "bold"), foreground=primary)
        style.configure("Subtitle.TLabel", font=("Segoe UI", 13), foreground=muted)
        style.configure("TLabel", background=surface, foreground=text)

        style.configure("Action.TButton", font=("Segoe UI", 11, "bold"), padding=7)
        style.map(
            "Action.TButton",
            background=[("pressed", primary_dark), ("active", primary)],
            foreground=[("active", surface)],
        )

        style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), padding=7)
        style.map(
            "Primary.TButton",
            background=[("pressed", primary_dark), ("active", primary)],
            foreground=[("active", surface)],
        )

        style.configure(
            "Menu.TButton",
            font=("Segoe UI", 11),
            padding=9,
            background=surface,
            foreground=primary,
        )
        style.map("Menu.TButton", background=[("active", primary_muted)])

        style.configure(
            "MenuActive.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=9,
            background=primary,
            foreground=surface,
        )
        style.map("MenuActive.TButton", background=[("active", primary)])

        style.configure("TSeparator", background=border)

        style.configure(
            "Treeview",
            background=surface,
            fieldbackground=surface,
            foreground=text,
            bordercolor=border,
            lightcolor=border,
            darkcolor=border,
            rowheight=28,
        )
        style.configure(
            "Treeview.Heading",
            background=primary_muted,
            foreground=primary,
            font=("Segoe UI", 11, "bold"),
        )
        style.map(
            "Treeview",
            background=[("selected", primary)],
            foreground=[("selected", surface)],
        )
        style.configure("TEntry", fieldbackground=field_bg, padding=6)
        style.configure("TCombobox", padding=6)

    def _create_card(self, parent, padding: int = 16):
        """Crea una tarjeta con bordes suaves y contenido interno."""

        card = ctk.CTkFrame(
            parent,
            fg_color=self.colors["surface"],
            corner_radius=16,
            border_width=1,
            border_color=self.colors["border"],
        )
        inner = ctk.CTkFrame(card, fg_color="transparent", corner_radius=0)
        inner.pack(fill=tk.BOTH, expand=True, padx=padding, pady=padding)
        return card, inner

    def _create_section_header(self, parent, title: str, subtitle: str | None = None):
        """Encabezado de seccion con tipografia consistente."""

        wrapper = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        title_label = ctk.CTkLabel(
            wrapper,
            text=title,
            font=ctk.CTkFont("Segoe UI", 16, "bold"),
            text_color=self.colors["primary"],
            anchor="w",
        )
        title_label.pack(anchor=tk.W)
        if subtitle:
            subtitle_label = ctk.CTkLabel(
                wrapper,
                text=subtitle,
                font=ctk.CTkFont("Segoe UI", 11),
                text_color=self.colors["text_muted"],
                anchor="w",
            )
            subtitle_label.pack(anchor=tk.W, pady=(2, 0))
        return wrapper

    def _create_pill_label(self, parent, text: str):
        """Etiqueta ovalada para indicar campos o bloques."""

        return ctk.CTkLabel(
            parent,
            text=text,
            fg_color=self.colors["chip_bg"],
            text_color=self.colors["chip_text"],
            corner_radius=999,
            padx=12,
            pady=4,
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
        )

    def _create_text_entry(self, parent, text_var: tk.StringVar, width: int = 260):
        """Campo de entrada con estilo moderno."""

        return ctk.CTkEntry(
            parent,
            textvariable=text_var,
            width=width,
            height=36,
            corner_radius=10,
            border_width=1,
            border_color=self.colors["field_border"],
            fg_color=self.colors["field_bg"],
            text_color=self.colors["text"],
            font=ctk.CTkFont("Segoe UI", 11),
        )

    def _create_combo(self, parent, text_var: tk.StringVar, values: list[str], command=None, width: int = 260):
        """Combo con estilo alineado a la interfaz."""
        combo = ctk.CTkComboBox(
            parent,
            variable=text_var,
            values=values,
            width=width,
            height=36,
            state="readonly",
            command=command,
            corner_radius=10,
            border_width=1,
            border_color=self.colors["field_border"],
            fg_color=self.colors["field_bg"],
            text_color=self.colors["text"],
            button_color=self.colors["primary"],
            button_hover_color=self.colors["primary_dark"],
            font=ctk.CTkFont("Segoe UI", 11),
        )
        self._bind_combo_click_anywhere(combo)
        return combo

    def _create_validated_entry(self, parent, text_var: tk.StringVar, width: int = 260):
        """Crea un entry con etiqueta de error debajo."""

        container = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        entry = self._create_text_entry(container, text_var, width=width)
        entry.pack(fill=tk.X)
        error_label = ctk.CTkLabel(
            container,
            text="",
            font=ctk.CTkFont("Segoe UI", 9),
            text_color=self.colors["error"],
            anchor="w",
        )
        error_label.pack(anchor=tk.W, pady=(2, 0))
        return container, entry, error_label

    def _create_validated_combo(self, parent, text_var: tk.StringVar, values: list[str], command=None, width: int = 260):
        """Crea un combo con etiqueta de error debajo."""

        container = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        combo = self._create_combo(container, text_var, values, command=command, width=width)
        combo.pack(fill=tk.X)
        error_label = ctk.CTkLabel(
            container,
            text="",
            font=ctk.CTkFont("Segoe UI", 9),
            text_color=self.colors["error"],
            anchor="w",
        )
        error_label.pack(anchor=tk.W, pady=(2, 0))
        return container, combo, error_label

    def _set_field_error(self, widget, error_label: ctk.CTkLabel, message: str | None) -> None:
        """Marca un campo con error y muestra el mensaje."""

        border_color = self.colors["field_border"]
        if message:
            border_color = self.colors["error"]
            error_label.configure(text=message)
        else:
            error_label.configure(text="")

        if widget is None:
            return
        try:
            widget.configure(border_color=border_color)
        except (tk.TclError, AttributeError):
            pass

    def _attach_required_validation(
        self,
        text_var: tk.StringVar,
        widget,
        error_label: ctk.CTkLabel,
        message: str,
        show_on_init: bool = False,
    ):
        """Valida que el campo no este vacio en cada cambio."""

        def _validate(_evt=None):
            value = (text_var.get() or "").strip()
            self._set_field_error(widget, error_label, message if not value else None)
            return bool(value)

        widget.bind("<KeyRelease>", _validate)
        widget.bind("<FocusOut>", _validate)
        if show_on_init:
            _validate()
        return _validate

    def _attach_date_validation(
        self,
        text_var: tk.StringVar,
        wrapper,
        error_label: ctk.CTkLabel,
        show_on_init: bool = False,
    ):
        """Valida formato de fecha dd/MM/yyyy."""

        def _validate(_evt=None):
            value = (text_var.get() or "").strip()
            if not value:
                self._set_field_error(wrapper, error_label, "Requerido")
                return False
            try:
                datetime.strptime(value, "%d/%m/%Y")
            except ValueError:
                self._set_field_error(wrapper, error_label, "Fecha invalida")
                return False
            self._set_field_error(wrapper, error_label, None)
            return True

        if show_on_init:
            _validate()
        return _validate

    def _attach_numeric_validation(
        self,
        text_var: tk.StringVar,
        widget,
        error_label: ctk.CTkLabel,
        message: str,
        show_on_init: bool = False,
    ):
        """Valida que el valor sea numerico cuando no esta vacio."""

        def _validate(_evt=None):
            value = (text_var.get() or "").strip()
            if value and not value.isdigit():
                self._set_field_error(widget, error_label, message)
                return False
            self._set_field_error(widget, error_label, None)
            return True

        widget.bind("<KeyRelease>", _validate)
        widget.bind("<FocusOut>", _validate)
        if show_on_init:
            _validate()
        return _validate

    def _format_panama_id(self, raw: str) -> str:
        """Formatea la cedula panamena con guiones automaticamente."""

        cleaned = (raw or "").upper().replace(" ", "")
        if not cleaned:
            return ""

        prefix = ""
        remainder = ""
        if cleaned.startswith("PE"):
            prefix = "PE"
            remainder = cleaned[2:]
        elif cleaned.startswith("E"):
            prefix = "E"
            remainder = cleaned[1:]
        elif cleaned.startswith("N"):
            prefix = "N"
            remainder = cleaned[1:]
        else:
            match = re.match(r"^(\d{1,2})", cleaned)
            if match:
                candidate = match.group(0)
                if candidate.isdigit() and 1 <= int(candidate) <= 13:
                    prefix = candidate
                    remainder = cleaned[len(prefix):]
                else:
                    remainder = cleaned
            else:
                remainder = cleaned

        if "-" in remainder:
            parts = [part for part in remainder.split("-") if part]
            tomo = re.sub(r"\D", "", parts[0]) if len(parts) > 0 else ""
            asiento = re.sub(r"\D", "", parts[1]) if len(parts) > 1 else ""
        else:
            digits = re.sub(r"\D", "", remainder)
            tomo = digits[:4]
            asiento = digits[4:9]

        parts = [prefix] if prefix else []
        if tomo:
            parts.append(tomo)
        if asiento:
            parts.append(asiento)
        return "-".join(parts)

    def _sanitize_panama_id_input(self, raw: str) -> str:
        """Permite solo prefijos E, N o PE y elimina letras no permitidas."""

        value = (raw or "").upper().replace(" ", "")
        if not value:
            return ""

        prefix = ""
        if value.startswith("PE"):
            prefix = "PE"
            remainder = value[2:]
        elif value.startswith("E"):
            prefix = "E"
            remainder = value[1:]
        elif value.startswith("N"):
            prefix = "N"
            remainder = value[1:]
        else:
            remainder = value

        remainder = re.sub(r"[^0-9-]", "", remainder)
        return f"{prefix}{remainder}"

    def _is_valid_panama_id(self, value: str) -> bool:
        """Valida cedulas panamenas (nacionales y casos especiales)."""

        if not value:
            return False
        cleaned = value.upper().strip()
        patterns = [
            r"^(1[0-3]|[1-9])-\d{1,4}-\d{1,5}$",
            r"^E-\d{1,4}-\d{1,7}$",
            r"^N-\d{1,4}-\d{1,5}$",
            r"^PE-\d{1,4}-\d{1,5}$",
        ]
        return any(re.match(pattern, cleaned) for pattern in patterns)

    def _attach_id_validation(
        self,
        text_var: tk.StringVar,
        widget,
        error_label: ctk.CTkLabel,
        show_on_init: bool = False,
    ):
        """Valida y formatea la cedula panamena en tiempo real."""

        def _validate(_evt=None):
            sanitized = self._sanitize_panama_id_input(text_var.get())
            if sanitized != text_var.get():
                text_var.set(sanitized)
            value = (text_var.get() or "").strip()
            if not value:
                self._set_field_error(widget, error_label, "Requerido")
                return False
            if not self._is_valid_panama_id(value):
                self._set_field_error(widget, error_label, "Cedula invalida")
                return False
            self._set_field_error(widget, error_label, None)
            return True

        def _format_on_blur(_evt=None):
            formatted = self._format_panama_id(text_var.get())
            if formatted:
                text_var.set(formatted)
            _validate()

        widget.bind("<KeyRelease>", _validate)
        widget.bind("<FocusOut>", _format_on_blur)
        if show_on_init:
            _validate()
        return _validate

    def _bind_combo_click_anywhere(self, combo: ctk.CTkComboBox) -> None:
        """Abre el desplegable al hacer clic en cualquier parte del combo."""

        click_handler = getattr(combo, "_clicked", None)
        canvas = getattr(combo, "_canvas", None)
        entry = getattr(combo, "_entry", None)
        if not callable(click_handler) or canvas is None:
            return

        for tag in ("inner_parts_left", "border_parts_left", "inner_parts_right", "border_parts_right"):
            try:
                canvas.tag_bind(tag, "<Button-1>", click_handler)
            except tk.TclError:
                continue

        if entry is not None:
            entry.bind("<Button-1>", click_handler)
    
    def setup_ui(self):
        """Configura la interfaz de usuario"""
        primary = self.colors["primary"]
        primary_dark = self.colors["primary_dark"]
        primary_muted = self.colors["primary_muted"]
        surface = self.colors["surface"]
        border = self.colors["border"]
        text = self.colors["text"]
        text_muted = self.colors["text_muted"]

        main_container = ctk.CTkFrame(self.root, fg_color=self.colors["bg"], corner_radius=0)
        main_container.pack(fill=tk.BOTH, expand=True)

        header_frame = ctk.CTkFrame(main_container, fg_color=primary, corner_radius=0)
        header_frame.pack(fill=tk.X)

        header_content = ctk.CTkFrame(header_frame, fg_color=primary, corner_radius=0)
        header_content.pack(pady=16, padx=20, fill=tk.X)

        try:
            if self.header_logo_path.exists():
                img = Image.open(self.header_logo_path).convert("RGBA")
                logo_size = (56, 56)
                self.logo_image = ctk.CTkImage(light_image=img, dark_image=img, size=logo_size)
                logo_label = ctk.CTkLabel(
                    header_content,
                    image=self.logo_image,
                    text="",
                    fg_color="transparent",
                )
                logo_label.pack(side=tk.LEFT, padx=(0, 12))
        except Exception as e:
            print(f"No se pudo cargar el logo: {e}")

        text_frame = ctk.CTkFrame(header_content, fg_color=primary, corner_radius=0)
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        title_label = ctk.CTkLabel(
            text_frame,
            text="CENTRO DE ATENCIÓN INTEGRAL TERAPÉUTICO PANAMÁ",
            font=ctk.CTkFont("Segoe UI", 16, "bold"),
            text_color=surface,
            anchor="w",
        )
        title_label.pack(anchor=tk.W)

        subtitle_label = ctk.CTkLabel(
            text_frame,
            text="Generador de Informes Clínicos",
            font=ctk.CTkFont("Segoe UI", 12),
            text_color=self.colors["text_light"],
            anchor="w",
        )
        subtitle_label.pack(anchor=tk.W)

        help_button = ctk.CTkButton(
            header_content,
            text="Ayuda",
            fg_color="transparent",
            hover_color="#1F8543",
            text_color=surface,
            border_width=1,
            border_color="#2B8B4F",
            font=ctk.CTkFont("Segoe UI", 11, "bold"),
            height=32,
            corner_radius=8,
        )
        help_button.pack(side=tk.RIGHT)
        self._attach_tooltip(
            help_button,
            "Crear informe genera solo el PDF. Exportar ZIP guarda el PDF y todos los adjuntos, "
            "ordenados en carpetas.",
        )

        ttk.Separator(main_container, orient="horizontal").pack(fill=tk.X)

        content_frame = ctk.CTkFrame(main_container, fg_color=self.colors["bg"], corner_radius=0)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)

        actions_header = ctk.CTkFrame(content_frame, fg_color=self.colors["bg"], corner_radius=0)
        actions_header.pack(fill=tk.X, pady=(0, 6))

        ctk.CTkLabel(
            actions_header,
            text="Acciones Principales",
            font=ctk.CTkFont("Segoe UI", 14, "bold"),
            text_color=text,
        ).pack(side=tk.LEFT)

        right_buttons = ctk.CTkFrame(actions_header, fg_color=self.colors["bg"], corner_radius=0)
        right_buttons.pack(side=tk.RIGHT)

        ctk.CTkButton(
            right_buttons,
            text="+ Nuevo informe",
            fg_color=primary_muted,
            text_color=primary,
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 11, "bold"),
            corner_radius=10,
            height=36,
            command=self._reset_form_state,
        ).pack(side=tk.LEFT, padx=6)

        ctk.CTkButton(
            right_buttons,
            text="Guardar borrador",
            fg_color=primary_muted,
            text_color=primary,
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 11, "bold"),
            corner_radius=10,
            height=36,
            command=self.save_report_draft,
        ).pack(side=tk.LEFT, padx=6)

        ctk.CTkButton(
            right_buttons,
            text="Cargar borrador",
            fg_color="transparent",
            text_color=primary,
            border_width=2,
            border_color=primary,
            hover_color=primary_muted,
            font=ctk.CTkFont("Segoe UI", 11, "bold"),
            corner_radius=10,
            height=36,
            command=self.load_report_draft,
        ).pack(side=tk.LEFT, padx=6)

        ctk.CTkButton(
            right_buttons,
            text="Exportar ZIP",
            fg_color=primary,
            text_color=surface,
            hover_color=primary_dark,
            font=ctk.CTkFont("Segoe UI", 11, "bold"),
            corner_radius=10,
            height=36,
            command=self.export_zip,
        ).pack(side=tk.LEFT, padx=6)

        ttk.Separator(content_frame, orient="horizontal").pack(fill=tk.X, pady=8)

        body_frame = ctk.CTkFrame(content_frame, fg_color=self.colors["bg"], corner_radius=0)
        body_frame.pack(fill=tk.BOTH, expand=True)

        menu_frame = ctk.CTkFrame(
            body_frame,
            fg_color=surface,
            corner_radius=12,
            border_width=1,
            border_color=border,
            width=260,
        )
        menu_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 16))

        menu_header = ctk.CTkFrame(menu_frame, fg_color=surface, corner_radius=0)
        menu_header.pack(fill=tk.X, padx=14, pady=(14, 6))
        ctk.CTkLabel(
            menu_header,
            text="Secciones del Informe",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            text_color=text_muted,
            anchor="w",
        ).pack(fill=tk.X)
        ttk.Separator(menu_frame, orient="horizontal").pack(fill=tk.X, padx=12, pady=(0, 8))

        sections_container = ctk.CTkFrame(
            body_frame,
            fg_color=self.colors["bg"],
            corner_radius=0,
        )
        sections_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sections_container.grid_rowconfigure(0, weight=1)
        sections_container.grid_columnconfigure(0, weight=1)

        self.company_section_name = "Datos de la Empresa"
        self.drafts_section_name = "Borradores"
        additional_sections = [
            self.company_section_name,
            self.results_section_name,
            self.conclusion_section_name,
            self.recommendations_section_name,
            self.calibration_section_name,
            self.report_attachments_section_name,
            self.attendance_section_name,
        ]
        self.section_names = ["Presentación del informe", "Contenido del informe"] + additional_sections
        self.section_buttons = {}
        self.section_frames = {}
        self.section_bodies = {}
        self.sections_container = sections_container
        self.content_outline_labels = []
        self.drafts_tree = None
        self.drafts_window = None

        # Variables del formulario principal
        self.report_type_var = tk.StringVar(value="audiometría")
        self.company_var = tk.StringVar()
        self.location_var = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.plant_var = tk.StringVar()
        self.activity_var = tk.StringVar()
        self.company_counterpart_var = tk.StringVar()
        self.counterpart_role_var = tk.StringVar()
        self.country_var = tk.StringVar()
        self.study_dates_var = tk.StringVar(value=self.date_var.get())

        for name in self.section_names:
            btn = ctk.CTkButton(
                menu_frame,
                text=name,
                fg_color="transparent",
                text_color=primary,
                hover_color=primary_muted,
                font=ctk.CTkFont("Segoe UI", 11),
                corner_radius=8,
                height=36,
                anchor="w",
                command=lambda n=name: self.show_section(n),
            )
            btn.pack(fill=tk.X, padx=10, pady=4)
            self.section_buttons[name] = btn

            frame = ctk.CTkFrame(
                sections_container,
                fg_color=surface,
                corner_radius=12,
                border_width=1,
                border_color=border,
            )
            frame.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
            self.section_frames[name] = frame
            if name == self.results_section_name:
                self.section_bodies[name] = frame
            else:
                self.section_bodies[name] = self._create_scrollable_section(frame)

        self._build_presentation_section()
        self._build_content_section()
        self._build_additional_sections()
        self.show_section(self.section_names[0])
        
        progress_frame = ctk.CTkFrame(menu_frame, fg_color=surface, corner_radius=0)
        progress_frame.pack(fill=tk.X, padx=12, pady=(8, 12), side=tk.BOTTOM)
        self.progress_caption = ctk.CTkLabel(
            progress_frame,
            text="Progreso del informe",
            font=ctk.CTkFont("Segoe UI", 9),
            text_color=text_muted,
            anchor="w",
        )
        self.progress_caption.pack(anchor=tk.W)
        self.progress_bar = ctk.CTkProgressBar(
            progress_frame,
            height=6,
            fg_color=border,
            progress_color=primary,
            corner_radius=4,
        )
        self.progress_bar.pack(fill=tk.X, pady=(6, 6))
        self.progress_label = ctk.CTkLabel(
            progress_frame,
            text="",
            font=ctk.CTkFont("Segoe UI", 9),
            text_color=text_muted,
            anchor="w",
        )
        self.progress_label.pack(anchor=tk.W)
        self._update_section_progress()

        ttk.Separator(content_frame, orient="horizontal").pack(fill=tk.X, pady=10)

        self.status_label = ctk.CTkLabel(
            content_frame,
            text="Listo para crear un nuevo informe",
            font=ctk.CTkFont("Segoe UI", 10),
            text_color=primary,
            anchor="w",
        )
        self.status_label.pack(pady=10, anchor=tk.W)
    
    def _build_presentation_section(self):
        """Construye la sección de presentación del informe."""

        frame = self.section_bodies.get("Presentación del informe")
        if frame is None:
            return

        for child in frame.winfo_children():
            child.destroy()
        header = self._create_section_header(
            frame,
            "Información del Informe",
            "Complete los datos generales de presentación del informe clínico.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        card, card_body = self._create_card(frame)
        card.pack(fill=tk.X, pady=4)

        form = ctk.CTkFrame(card_body, fg_color="transparent", corner_radius=0)
        form.pack(fill=tk.X)
        form.grid_columnconfigure(1, weight=1)

        self._create_pill_label(form, "Tipo de informe *").grid(row=0, column=0, sticky=tk.NW, pady=8)
        type_combo = self._create_combo(
            form,
            self.report_type_var,
            ["audiometría", "espirometría", "audiometría + espirometría"],
            command=lambda _value=None: self._handle_report_type_change(),
            width=280,
        )
        type_combo.grid(row=0, column=1, sticky="nsew", padx=12, pady=8)

        self._create_pill_label(form, "Empresa *").grid(row=1, column=0, sticky=tk.NW, pady=8)
        company_container, company_entry, company_error = self._create_validated_entry(
            form, self.company_var, width=280
        )
        company_container.grid(row=1, column=1, sticky="nsew", padx=12, pady=8)
        self._attach_required_validation(self.company_var, company_entry, company_error, "Requerido")

        self._create_pill_label(form, "Ubicación / Sede *").grid(row=2, column=0, sticky=tk.NW, pady=8)
        location_container, location_entry, location_error = self._create_validated_entry(
            form, self.location_var, width=280
        )
        location_container.grid(row=2, column=1, sticky="nsew", padx=12, pady=8)
        self._attach_required_validation(self.location_var, location_entry, location_error, "Requerido")

        self._create_pill_label(form, "Evaluador / Responsable *").grid(row=3, column=0, sticky=tk.W, pady=8)
        self.evaluator_combo = self._create_combo(
            form,
            self.evaluator_var,
            [],
            command=lambda _value=None: self._handle_evaluator_selection(),
            width=280,
        )
        self.evaluator_combo.grid(row=3, column=1, sticky="nsew", padx=12, pady=8)

        evaluator_actions = ctk.CTkFrame(form, fg_color="transparent", corner_radius=0)
        evaluator_actions.grid(row=3, column=2, sticky=tk.W, padx=6, pady=8)
        ctk.CTkButton(
            evaluator_actions,
            text="Agregar",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._open_add_evaluator_dialog,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkButton(
            evaluator_actions,
            text="Eliminar",
            fg_color="transparent",
            text_color=self.colors["primary"],
            border_width=1,
            border_color=self.colors["primary"],
            hover_color=self.colors["primary_muted"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._remove_selected_evaluator,
        ).pack(side=tk.LEFT)

        self.evaluator_detail_label = ctk.CTkLabel(
            form,
            text="",
            font=ctk.CTkFont("Segoe UI", 10),
            text_color=self.colors["text_muted"],
            anchor="w",
        )
        self.evaluator_detail_label.grid(row=4, column=1, columnspan=2, sticky=tk.W, padx=12, pady=(0, 6))

        self._create_pill_label(form, "Contraparte tecnica").grid(row=5, column=0, sticky=tk.W, pady=8)
        self.counterpart_combo = self._create_combo(
            form,
            self.counterpart_var,
            ["Sin contraparte"],
            command=lambda _value=None: self._handle_counterpart_selection(),
            width=280,
        )
        self.counterpart_combo.grid(row=5, column=1, sticky="nsew", padx=12, pady=8)

        counterpart_actions = ctk.CTkFrame(form, fg_color="transparent", corner_radius=0)
        counterpart_actions.grid(row=5, column=2, sticky=tk.W, padx=6, pady=8)
        ctk.CTkButton(
            counterpart_actions,
            text="Agregar",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._open_add_counterpart_dialog,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkButton(
            counterpart_actions,
            text="Eliminar",
            fg_color="transparent",
            text_color=self.colors["primary"],
            border_width=1,
            border_color=self.colors["primary"],
            hover_color=self.colors["primary_muted"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._remove_selected_counterpart,
        ).pack(side=tk.LEFT)

        self._create_pill_label(form, "Cargo contraparte").grid(row=6, column=0, sticky=tk.W, pady=8)
        self.counterpart_role_entry = self._create_text_entry(form, self.counterpart_role_var, width=280)
        self.counterpart_role_entry.grid(row=6, column=1, sticky="nsew", padx=12, pady=8)

        self._create_pill_label(form, "Fecha de evaluación *").grid(row=7, column=0, sticky=tk.W, pady=8)
        date_container = ctk.CTkFrame(form, fg_color="transparent", corner_radius=0)
        date_container.grid(row=7, column=1, sticky="nsew", padx=12, pady=8)
        date_wrapper = ctk.CTkFrame(
            date_container,
            fg_color=self.colors["field_bg"],
            corner_radius=10,
            border_width=1,
            border_color=self.colors["field_border"],
        )
        date_wrapper.configure(width=280, height=36)
        date_wrapper.pack_propagate(False)
        date_wrapper.pack(anchor=tk.W, fill=tk.X)
        date_entry = self._create_date_entry(date_wrapper, self.date_var, width=32)
        date_entry.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        date_error = ctk.CTkLabel(
            date_container,
            text="",
            font=ctk.CTkFont("Segoe UI", 9),
            text_color=self.colors["error"],
            anchor="w",
        )
        date_error.pack(anchor=tk.W, pady=(2, 0))
        self._attach_date_validation(self.date_var, date_wrapper, date_error)

        self._reload_evaluators()
        self._reload_counterparts()

        helper_text = ctk.CTkLabel(
            frame,
            text="Selecciona otras partes del informe desde el menú lateral para continuar.",
            font=ctk.CTkFont("Segoe UI", 10),
            text_color=self.colors["text_muted"],
            anchor="w",
        )
        helper_text.pack(anchor=tk.W, pady=(10, 0))

    def _reload_evaluators(self, prefer_id=None) -> None:
        """Carga o actualiza el listado de evaluadores desde el repositorio."""

        profiles = self.evaluators_repo.list_all()
        self.evaluator_profiles = {profile.get("id"): profile for profile in profiles if profile.get("id")}
        names = [profile.get("name", "") for profile in profiles]

        if self.evaluator_combo is not None:
            self.evaluator_combo.configure(values=names)

        if not profiles:
            self.selected_evaluator_id.set("")
            self.evaluator_var.set("")
            self._update_evaluator_details_preview()
            return

        target_id = prefer_id or self.selected_evaluator_id.get()
        if target_id not in self.evaluator_profiles:
            target_id = profiles[0].get("id")
        if target_id:
            self._select_evaluator(target_id, update_combo=True)

    def _reload_counterparts(self, prefer_id: str | None = None) -> None:
        """Carga o actualiza el listado de contrapartes desde el repositorio."""

        profiles = self.counterparts_repo.list_all()
        self.counterpart_profiles = {
            profile.get("id"): profile
            for profile in profiles
            if profile.get("id")
        }
        labels = [self._format_counterpart_label(profile) for profile in profiles]
        options = ["Sin contraparte"] + labels

        if self.counterpart_combo is not None:
            self.counterpart_combo.configure(values=options)

        target_id = prefer_id or self.selected_counterpart_id.get()
        if target_id and target_id in self.counterpart_profiles:
            self._select_counterpart(target_id, update_combo=True)
            return

        self._select_counterpart(None, update_combo=True)

    def _format_counterpart_label(self, profile: dict) -> str:
        """Formatea la etiqueta visible de una contraparte."""

        name = (profile.get("name") or "").strip()
        role = (profile.get("role") or "").strip()
        return f"{name} ({role})" if role else name

    def _select_counterpart(self, counterpart_id: str | None, update_combo: bool = False) -> None:
        """Aplica la selección de contraparte y actualiza los campos."""

        if not counterpart_id or counterpart_id not in self.counterpart_profiles:
            self.selected_counterpart_id.set("")
            self.company_counterpart_var.set("")
            self.counterpart_role_var.set("")
            self._set_counterpart_role_state(False)
            if update_combo and self.counterpart_combo is not None:
                self.counterpart_var.set("Sin contraparte")
                self.counterpart_combo.set("Sin contraparte")
            return

        profile = self.counterpart_profiles[counterpart_id]
        self.selected_counterpart_id.set(counterpart_id)
        self.company_counterpart_var.set(profile.get("name", ""))
        self.counterpart_role_var.set(profile.get("role", ""))
        self._set_counterpart_role_state(True)
        if update_combo and self.counterpart_combo is not None:
            label = self._format_counterpart_label(profile)
            self.counterpart_var.set(label)
            self.counterpart_combo.set(label)

    def _handle_counterpart_selection(self, *_args) -> None:
        """Sincroniza la contraparte seleccionada desde el combo."""

        selected_label = (self.counterpart_var.get() or "").strip()
        if not selected_label or selected_label == "Sin contraparte":
            self._select_counterpart(None, update_combo=True)
            return

        for counter_id, profile in self.counterpart_profiles.items():
            if self._format_counterpart_label(profile) == selected_label:
                self._select_counterpart(counter_id)
                return

        self._select_counterpart(None, update_combo=True)

    def _set_counterpart_role_state(self, enabled: bool) -> None:
        """Activa o desactiva el campo de cargo de la contraparte."""

        if self.counterpart_role_entry is None:
            return
        state = "normal" if enabled else "disabled"
        self.counterpart_role_entry.configure(state=state)

    def _open_add_counterpart_dialog(self) -> None:
        """Muestra el cuadro para agregar una nueva contraparte tecnica."""

        dialog = tk.Toplevel(self.root)
        dialog.title("Nueva contraparte tecnica")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        name_var = tk.StringVar()
        role_var = tk.StringVar()

        container = ttk.Frame(dialog, padding=15)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="Nombre completo:", style="Subtitle.TLabel").grid(
            row=0, column=0, sticky=tk.W, pady=4
        )
        ttk.Entry(container, textvariable=name_var, width=40).grid(row=0, column=1, sticky=tk.EW, pady=4)

        ttk.Label(container, text="Cargo / rol:", style="Subtitle.TLabel").grid(
            row=1, column=0, sticky=tk.W, pady=4
        )
        ttk.Entry(container, textvariable=role_var, width=40).grid(row=1, column=1, sticky=tk.EW, pady=4)

        buttons = ttk.Frame(container)
        buttons.grid(row=2, column=0, columnspan=2, sticky=tk.E, pady=(12, 0))
        ttk.Button(buttons, text="Cancelar", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

        def _save_counterpart():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Datos incompletos", "El nombre de la contraparte es obligatorio.")
                return

            try:
                new_entry = self.counterparts_repo.add_counterpart(
                    {
                        "name": name,
                        "role": role_var.get().strip(),
                    }
                )
            except ValueError as exc:
                messagebox.showerror("No se pudo guardar", str(exc))
                return

            dialog.destroy()
            self._reload_counterparts(prefer_id=new_entry.get("id"))

        ttk.Button(buttons, text="Guardar", style="Primary.TButton", command=_save_counterpart).pack(
            side=tk.RIGHT
        )
        container.columnconfigure(1, weight=1)
        self._center_window(dialog)

    def _center_window(self, window: tk.Toplevel) -> None:
        """Centra un dialogo relativo a la ventana principal."""

        window.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()

        win_w = window.winfo_width()
        win_h = window.winfo_height()
        if win_w <= 1 or win_h <= 1:
            win_w = window.winfo_reqwidth()
            win_h = window.winfo_reqheight()

        pos_x = root_x + (root_w - win_w) // 2
        pos_y = root_y + (root_h - win_h) // 2
        window.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

    def _remove_selected_counterpart(self) -> None:
        """Elimina la contraparte seleccionada del catalogo."""

        counterpart_id = self.selected_counterpart_id.get()
        if not counterpart_id:
            messagebox.showinfo("Sin seleccion", "Selecciona una contraparte para eliminarla.")
            return

        profile = self.counterpart_profiles.get(counterpart_id, {})
        name = profile.get("name", "contraparte")
        confirm = messagebox.askyesno(
            "Confirmar eliminacion",
            f"Se eliminara la contraparte '{name}'. ¿Deseas continuar?",
        )
        if not confirm:
            return

        if self.counterparts_repo.remove_counterpart(counterpart_id):
            self._reload_counterparts()

    def _remove_selected_evaluator(self) -> None:
        """Elimina el evaluador seleccionado del catalogo."""

        evaluator_id = self.selected_evaluator_id.get()
        if not evaluator_id:
            messagebox.showinfo("Sin seleccion", "Selecciona un evaluador para eliminarlo.")
            return

        profile = self.evaluator_profiles.get(evaluator_id, {})
        name = profile.get("name", "evaluador")
        confirm = messagebox.askyesno(
            "Confirmar eliminacion",
            f"Se eliminara el evaluador '{name}'. ¿Deseas continuar?",
        )
        if not confirm:
            return

        if self.evaluators_repo.remove_evaluator(evaluator_id):
            self._reload_evaluators()

    def _select_evaluator(self, evaluator_id: str, update_combo: bool = False) -> None:
        """Aplica la selección actualizando campos dependientes."""

        profile = self.evaluator_profiles.get(evaluator_id)
        if not profile:
            return

        self.selected_evaluator_id.set(evaluator_id)
        self.evaluator_var.set(profile.get("name", ""))
        if update_combo and self.evaluator_combo is not None:
            self.evaluator_combo.set(profile.get("name", ""))
        self._update_evaluator_details_preview()

    def _handle_evaluator_selection(self, *_args) -> None:
        """Sincroniza el ID interno al cambiar la selección del combobox."""

        selected_name = (self.evaluator_var.get() or "").strip()
        for eval_id, profile in self.evaluator_profiles.items():
            if profile.get("name") == selected_name:
                self._select_evaluator(eval_id)
                return
        self._update_evaluator_details_preview()

    def _update_evaluator_details_preview(self) -> None:
        """Refresca la etiqueta con la profesión y registro del evaluador."""

        if self.evaluator_detail_label is None:
            return

        profile = self._get_selected_evaluator_profile()
        if not profile:
            self.evaluator_detail_label.configure(text="Registra un evaluador para continuar.")
            return

        profession = (profile.get("profession") or "").strip()
        registry = (profile.get("registry") or "").strip()
        detail_parts = [value for value in (profession, registry) if value]
        if not detail_parts:
            detail_parts = [profile.get("title_label", "")] if profile.get("title_label") else []
        self.evaluator_detail_label.configure(text=" | ".join(detail_parts))

    def _select_evaluator_credential(self, target_var: tk.StringVar) -> None:
        """Permite escoger un PDF de idoneidad para un evaluador."""

        file_path = filedialog.askopenfilename(
            title="Selecciona la idoneidad (PDF)",
            filetypes=(("PDF", "*.pdf"), ("Todos los archivos", "*.*")),
        )
        if not file_path:
            return
        target_var.set(file_path)

    def _store_evaluator_credential(self, source_path: str, evaluator_name: str) -> str:
        """Copia la idoneidad a la carpeta interna y devuelve la ruta relativa."""

        if not source_path:
            return ""

        source = Path(source_path)
        if not source.exists():
            return ""

        target_dir = self.project_root / "data" / "attachments" / "idoneidad"
        target_dir.mkdir(parents=True, exist_ok=True)

        safe_name = self._sanitize_filename(evaluator_name or "evaluador")
        destination = target_dir / f"{safe_name}{source.suffix}"
        shutil.copy2(source, destination)

        return str(destination.relative_to(self.project_root))

    def _open_add_evaluator_dialog(self) -> None:
        """Muestra un cuadro de diálogo para registrar nuevos evaluadores."""

        dialog = tk.Toplevel(self.root)
        dialog.title("Nuevo evaluador")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        name_var = tk.StringVar()
        profession_var = tk.StringVar()
        registry_var = tk.StringVar()
        audio_var = tk.BooleanVar(value=True)
        espiro_var = tk.BooleanVar(value="espirom" in (self.report_type_var.get() or "").lower())
        credential_var = tk.StringVar()

        container = ttk.Frame(dialog, padding=15)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="Nombre completo:", style='Subtitle.TLabel').grid(row=0, column=0, sticky=tk.W, pady=4)
        ttk.Entry(container, textvariable=name_var, width=40).grid(row=0, column=1, sticky=tk.EW, pady=4)

        ttk.Label(container, text="Profesión o título:", style='Subtitle.TLabel').grid(row=1, column=0, sticky=tk.W, pady=4)
        ttk.Entry(container, textvariable=profession_var, width=40).grid(row=1, column=1, sticky=tk.EW, pady=4)

        ttk.Label(container, text="Registro/licencia:", style='Subtitle.TLabel').grid(row=2, column=0, sticky=tk.W, pady=4)
        ttk.Entry(container, textvariable=registry_var, width=40).grid(row=2, column=1, sticky=tk.EW, pady=4)

        ttk.Label(container, text="Idoneidad (PDF):", style='Subtitle.TLabel').grid(row=3, column=0, sticky=tk.W, pady=4)
        credential_entry = ttk.Entry(container, textvariable=credential_var, width=34, state="readonly")
        credential_entry.grid(row=3, column=1, sticky=tk.W, pady=4)
        ttk.Button(
            container,
            text="Seleccionar",
            style='Action.TButton',
            command=lambda: self._select_evaluator_credential(credential_var),
        ).grid(row=3, column=1, sticky=tk.E, pady=4)

        checks_frame = ttk.Frame(container)
        checks_frame.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(8, 4))
        ttk.Checkbutton(checks_frame, text="Disponible en audiometría", variable=audio_var).pack(anchor=tk.W)
        ttk.Checkbutton(checks_frame, text="Disponible en espirometría", variable=espiro_var).pack(anchor=tk.W)

        buttons = ttk.Frame(container)
        buttons.grid(row=5, column=0, columnspan=2, sticky=tk.E, pady=(12, 0))
        ttk.Button(buttons, text="Cancelar", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

        def _save_new_evaluator():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Datos incompletos", "El nombre del evaluador es obligatorio.")
                return

            applicable = []
            if audio_var.get():
                applicable.append("audiometria")
            if espiro_var.get():
                applicable.append("espirometria")
            if not applicable:
                messagebox.showwarning(
                    "Selecciona al menos un tipo",
                    "Debes indicar en qué tipo de informe participará el evaluador.",
                )
                return

            if not credential_var.get().strip():
                messagebox.showwarning(
                    "Idoneidad requerida",
                    "Debes adjuntar el archivo de idoneidad en PDF.",
                )
                return

            credential_path = self._store_evaluator_credential(
                credential_var.get().strip(),
                name,
            )
            if not credential_path:
                messagebox.showerror(
                    "No se pudo guardar",
                    "No fue posible copiar la idoneidad al directorio interno.",
                )
                return

            try:
                new_entry = self.evaluators_repo.add_evaluator(
                    {
                        "name": name,
                        "profession": profession_var.get().strip(),
                        "registry": registry_var.get().strip(),
                        "credential_file": credential_path,
                        "applicable_reports": applicable,
                    }
                )
            except ValueError as exc:
                messagebox.showerror("No se pudo guardar", str(exc))
                return

            dialog.destroy()
            self._reload_evaluators(prefer_id=new_entry.get("id"))

        ttk.Button(buttons, text="Guardar", style='Primary.TButton', command=_save_new_evaluator).pack(side=tk.RIGHT)
        container.columnconfigure(1, weight=1)
        self._center_window(dialog)

    def _get_selected_evaluator_profile(self):
        """Obtiene el diccionario del evaluador actualmente seleccionado."""

        evaluator_id = self.selected_evaluator_id.get()
        if evaluator_id and evaluator_id in self.evaluator_profiles:
            return self.evaluator_profiles[evaluator_id]

        selected_name = (self.evaluator_var.get() or "").strip()
        for profile in self.evaluator_profiles.values():
            if profile.get("name") == selected_name:
                return profile
        return None

    def _build_technical_team_payload(self, report_type: str) -> list:
        """Arma la lista enviada al generador de PDF para Equipo Técnico."""

        team_profiles = self.evaluators_repo.get_team_for_report(report_type)
        formatted = [self._format_technical_member(profile) for profile in team_profiles if profile]

        if not formatted:
            profile = self._get_selected_evaluator_profile()
            if profile:
                formatted.append(self._format_technical_member(profile))

        return formatted

    def _format_technical_member(self, profile: dict) -> dict:
        """Establece la estructura esperada por el generador de PDF."""

        return {
            "name": profile.get("name", ""),
            "details": build_technical_details(profile),
            "credential_file": (profile.get("credential_file") or "").strip(),
        }

    def _create_date_entry(self, parent, text_var: tk.StringVar, width: int = 32):
        """Crea un DateEntry con calendario y normalización consistente dd/MM/yyyy."""

        self._normalize_date_var(text_var)

        date_entry = DateEntry(
            parent,
            textvariable=text_var,
            selectmode="day",
            date_pattern="dd/MM/yyyy",
            showweeknumbers=False,
        )
        date_entry.configure(width=width)

        try:
            current_date = datetime.strptime(text_var.get(), "%d/%m/%Y")
        except ValueError:
            current_date = datetime.now()
            text_var.set(current_date.strftime("%d/%m/%Y"))

        date_entry.set_date(current_date)

        date_entry.bind(
            "<<DateEntrySelected>>",
            lambda _evt, var=text_var: self._normalize_date_var(var),
        )
        entry_widget = getattr(date_entry, "entry", date_entry)
        entry_widget.configure(state="readonly")
        return date_entry

    def _normalize_date_var(self, text_var: tk.StringVar) -> None:
        """Fuerza el valor del StringVar a dd/MM/yyyy usando hoy como respaldo."""

        value = (text_var.get() or "").strip()
        fallback = datetime.now()
        parsed_date = None

        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                parsed_date = datetime.strptime(value, fmt)
                break
            except ValueError:
                continue

        if parsed_date is None:
            parsed_date = fallback

        text_var.set(parsed_date.strftime("%d/%m/%Y"))

    def _validate_date_input(self, proposed: str) -> bool:
        """Valida la escritura del usuario para permitir solo números, '/' y longitud esperada."""

        if proposed == "":
            return True

        if len(proposed) > 10:
            return False

        if proposed.count("/") > 2:
            return False

        for ch in proposed:
            if not (ch.isdigit() or ch == "/"):
                return False

        segments = proposed.split("/")
        max_lengths = [2, 2, 4]
        for idx, segment in enumerate(segments):
            max_len = max_lengths[idx] if idx < len(max_lengths) else 0
            if max_len and len(segment) > max_len:
                return False

        return True

    def _validate_numeric_input(self, proposed: str) -> bool:
        """Permite únicamente dígitos en campos numéricos como la edad."""

        if proposed == "":
            return True
        return proposed.isdigit()

    def _enforce_date_mask(self, text_var: tk.StringVar) -> None:
        """Formatea el valor como dd/MM/yyyy mientras el usuario escribe."""

        digits = "".join(ch for ch in (text_var.get() or "") if ch.isdigit())[:8]

        if not digits:
            text_var.set("")
            return

        day = digits[:2]
        month = digits[2:4] if len(digits) > 2 else ""
        year = digits[4:] if len(digits) > 4 else ""

        formatted = day
        if len(day) == 2 and month:
            formatted += "/" + month
        elif 0 < len(month) < 2:
            formatted += "/" + month

        if len(month) == 2 and year:
            formatted += "/" + year
        elif len(month) < 2 and year:
            # esperar a que el mes tenga dos dígitos antes de añadir el año
            pass

        text_var.set(formatted)

    def _refresh_content_preview(self, *_args):
        """Actualiza los textos del índice visible en la sección de contenido."""

        if not self.content_outline_labels:
            return

        outline = get_content_outline(self.report_type_var.get())
        for label, title in zip(self.content_outline_labels, outline):
            label.configure(text=f"- {title}")

    def _handle_report_type_change(self, *_args):
        """Sincroniza UI dependientes del tipo de informe."""

        self._refresh_content_preview()
        self._render_result_blocks()
        self._update_test_attachment_state()

    def _build_content_section(self):
        """Construye la sección de vista previa del índice de contenido."""

        frame = self.section_bodies.get("Contenido del informe")
        if frame is None:
            return

        for child in frame.winfo_children():
            child.destroy()

        header = self._create_section_header(
            frame,
            "Contenido del informe",
            "Esta pagina siempre mostrara el indice del informe usando la plantilla oficial.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        card, card_body = self._create_card(frame)
        card.pack(fill=tk.BOTH, expand=True)

        list_container = ctk.CTkFrame(card_body, fg_color="transparent", corner_radius=0)
        list_container.pack(fill=tk.BOTH, expand=True)

        self.content_outline_labels = []
        outline = get_content_outline(self.report_type_var.get())
        for title in outline:
            label = ctk.CTkLabel(
                list_container,
                text=f"• {title}",
                font=ctk.CTkFont("Segoe UI", 11),
                text_color=self.colors["text_muted"],
                anchor="w",
            )
            label.pack(anchor=tk.W, pady=4)
            self.content_outline_labels.append(label)

        self._refresh_content_preview()

    def _build_additional_sections(self):
        """Crea marcos de contenido para las demás partes del informe."""

        for name in self.section_names[2:]:
            frame = self.section_bodies.get(name)
            if frame is None:
                continue

            if name == self.company_section_name:
                self._build_company_data_section(frame)
                continue
            if name == self.results_section_name:
                self._build_results_section(frame)
                continue
            if name == self.recommendations_section_name:
                self._build_recommendations_section(frame)
                continue
            if name == self.conclusion_section_name:
                self._build_conclusion_section(frame)
                continue
            if name == self.calibration_section_name:
                self._build_calibration_section(frame)
                continue
            if name == self.report_attachments_section_name:
                self._build_test_attachments_section(frame)
                continue
            if name == self.attendance_section_name:
                self._build_attendance_section(frame)
                continue

            ttk.Label(frame, text=name, style='Title.TLabel').pack(anchor=tk.W, pady=(0, 10))
            ttk.Label(
                frame,
                text="Contenido disponible próximamente.",
                style='Subtitle.TLabel'
            ).pack(anchor=tk.W)

    def _build_company_data_section(self, frame: ttk.Frame):
        """Sección Parte 3 con los campos adicionales para datos de la empresa."""
        for child in frame.winfo_children():
            child.destroy()

        header = self._create_section_header(
            frame,
            self.company_section_name,
            "Completa la informacion que se utilizara en la pagina 3 del PDF.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        card, card_body = self._create_card(frame)
        card.pack(fill=tk.X, pady=4)

        fields_frame = ctk.CTkFrame(card_body, fg_color="transparent", corner_radius=0)
        fields_frame.pack(fill=tk.X)
        fields_frame.grid_columnconfigure(1, weight=1)

        field_specs = [
            ("Planta evaluada:", self.plant_var, "entry"),
            ("Actividad principal:", self.activity_var, "entry"),
            ("País en que se realizó:", self.country_var, "entry"),
            ("Fechas del estudio:", self.study_dates_var, "date"),
        ]

        for idx, (label_text, var, field_type) in enumerate(field_specs):
            self._create_pill_label(fields_frame, label_text).grid(row=idx, column=0, sticky=tk.W, pady=8)
            if field_type == "date":
                date_wrapper = ctk.CTkFrame(
                    fields_frame,
                    fg_color=self.colors["field_bg"],
                    corner_radius=10,
                    border_width=1,
                    border_color=self.colors["field_border"],
                )
                date_wrapper.grid(row=idx, column=1, sticky=tk.W, padx=12, pady=8)
                widget = self._create_date_entry(date_wrapper, var, width=18)
                widget.pack(padx=8, pady=6)
            else:
                self._create_text_entry(fields_frame, var, width=340).grid(
                    row=idx, column=1, sticky=tk.EW, padx=12, pady=8
                )

    def _build_drafts_section(self, frame: ttk.Frame):
        """Seccion para listar y cargar borradores guardados."""

        for child in frame.winfo_children():
            child.destroy()

        header = self._create_section_header(
            frame,
            "Borradores guardados",
            "Administra los borradores guardados para continuar un informe.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        action_row = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=0)
        action_row.pack(fill=tk.X, pady=(0, 8))
        ctk.CTkButton(
            action_row,
            text="Actualizar lista",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._refresh_drafts_table,
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            action_row,
            text="Cargar seleccionado",
            fg_color=self.colors["primary"],
            text_color=self.colors["surface"],
            hover_color=self.colors["primary_dark"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._load_selected_draft,
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            action_row,
            text="Eliminar seleccionado",
            fg_color="transparent",
            text_color=self.colors["primary"],
            border_width=1,
            border_color=self.colors["primary"],
            hover_color=self.colors["primary_muted"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._delete_selected_draft,
        ).pack(side=tk.LEFT, padx=4)

        tree_frame = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=0)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        columns = ("name", "date")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=12)
        scrollbar = ctk.CTkScrollbar(
            tree_frame,
            orientation="vertical",
            command=tree.yview,
            width=12,
            corner_radius=999,
            fg_color=self.colors["border"],
            button_color=self.colors["primary"],
            button_hover_color=self.colors["primary_dark"],
        )
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(6, 0))

        tree.heading("name", text="Archivo")
        tree.heading("date", text="Modificado")
        tree.column("name", width=420, anchor=tk.W)
        tree.column("date", width=180, anchor=tk.CENTER)

        self.drafts_tree = tree
        self._refresh_drafts_table()

    def _build_results_section(self, frame: ttk.Frame):
        """Sección de resultados con bloques dinámicos por tipo de prueba."""
        for child in frame.winfo_children():
            child.destroy()

        header = self._create_section_header(
            frame,
            self.results_section_name,
            "Registra a los colaboradores evaluados para generar los cuadros oficiales por tipo de prueba.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        self._setup_scrollable_results_container(frame)
        self._render_result_blocks()

    def _setup_scrollable_results_container(self, parent: ttk.Frame):
        """Crea un contenedor desplazable para los bloques de resultados."""

        container = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, highlightthickness=0, bg=self.colors["bg"])
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        canvas.configure(yscrollcommand=None)

        inner_frame = ctk.CTkFrame(canvas, fg_color="transparent", corner_radius=0)
        window_id = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        def _update_scroll_region(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _resize_inner(event):
            canvas.itemconfig(window_id, width=event.width)

        inner_frame.bind("<Configure>", _update_scroll_region)
        canvas.bind("<Configure>", _resize_inner)
        self._bind_mousewheel(canvas)

        self.results_canvas = canvas
        self.results_section_container = inner_frame

    def _create_scrollable_section(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        """Crea un contenedor desplazable para una sección."""

        container = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        container.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        canvas = tk.Canvas(container, highlightthickness=0, bg=self.colors["surface"])
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ctk.CTkScrollbar(
            container,
            orientation="vertical",
            command=canvas.yview,
            width=12,
            corner_radius=999,
            fg_color=self.colors["border"],
            button_color=self.colors["primary"],
            button_hover_color=self.colors["primary_dark"],
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(6, 0))
        canvas.configure(yscrollcommand=scrollbar.set)

        inner_frame = ctk.CTkFrame(canvas, fg_color="transparent", corner_radius=0)
        window_id = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        def _update_scroll_region(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _resize_inner(event):
            canvas.itemconfig(window_id, width=event.width)

        inner_frame.bind("<Configure>", _update_scroll_region)
        canvas.bind("<Configure>", _resize_inner)
        self._bind_mousewheel(canvas)

        return inner_frame

    def _bind_mousewheel(self, canvas: tk.Canvas):
        """Activa el desplazamiento con la rueda del ratón dentro del canvas."""

        def _on_mousewheel(event):
            delta = 0
            if event.delta:
                delta = int(-1 * (event.delta / 120))
            elif event.num == 5:
                delta = 1
            elif event.num == 4:
                delta = -1
            if delta:
                canvas.yview_scroll(delta, "units")

        def _activate(_event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _deactivate(_event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _activate)
        canvas.bind("<Leave>", _deactivate)

    def _render_result_blocks(self):
        """Reconstruye los bloques de captura según el tipo de informe."""

        if not hasattr(self, "results_section_container"):
            return

        for child in self.results_section_container.winfo_children():
            child.destroy()

        self.result_blocks = {}
        dataset_keys = self._determine_result_dataset_keys()
        if not dataset_keys:
            ttk.Label(
                self.results_section_container,
                text="Selecciona un tipo de informe para habilitar los resultados.",
                style='Subtitle.TLabel'
            ).pack(anchor=tk.W, pady=10)
            return

        for dataset_key in dataset_keys:
            self._create_result_block(self.results_section_container, dataset_key)

    def _determine_result_dataset_keys(self):
        """Devuelve el listado de conjuntos a capturar basándose en la selección actual."""

        report_type = (self.report_type_var.get() or "").lower()
        has_audio = "audiometr" in report_type
        has_espiro = "espiro" in report_type

        if has_audio and has_espiro:
            return ["audiometria", "espirometria"]
        if has_espiro:
            return ["espirometria"]
        return ["audiometria"]

    def _create_result_block(self, parent: ttk.Frame, dataset_key: str):
        """Construye el formulario y la tabla para un tipo de prueba."""

        scheme = RESULT_SCHEMES.get(dataset_key, RESULT_SCHEMES["audiometria"])
        palette = self.result_palettes.get(dataset_key, {})
        result_values = list(palette.keys())
        default_result = result_values[0] if result_values else ""

        block_card, block_body = self._create_card(parent, padding=12)
        block_card.pack(fill=tk.BOTH, expand=True, pady=10)

        title_row = ctk.CTkFrame(block_body, fg_color="transparent", corner_radius=0)
        title_row.pack(fill=tk.X, pady=(0, 8))
        self._create_pill_label(title_row, scheme["title"]).pack(anchor=tk.W)

        form_vars = {
            "name": tk.StringVar(),
            "identification": tk.StringVar(),
            "age": tk.StringVar(),
            "position": tk.StringVar(),
            "result": tk.StringVar(value=default_result),
        }

        form = ctk.CTkFrame(block_body, fg_color="transparent", corner_radius=0)
        form.pack(fill=tk.X)
        form.grid_columnconfigure(1, weight=1)
        form.grid_columnconfigure(3, weight=1)

        self._create_pill_label(form, "Nombre completo").grid(row=0, column=0, sticky=tk.NW, pady=6)
        name_container, name_entry, name_error = self._create_validated_entry(
            form, form_vars["name"], width=280
        )
        name_container.grid(row=0, column=1, sticky="nsew", padx=12, pady=6)
        name_validator = self._attach_required_validation(form_vars["name"], name_entry, name_error, "Requerido")

        self._create_pill_label(form, "Cédula").grid(row=0, column=2, sticky=tk.NW, pady=6)
        id_container, id_entry, id_error = self._create_validated_entry(
            form, form_vars["identification"], width=220
        )
        id_container.grid(row=0, column=3, sticky="nsew", padx=12, pady=6)
        id_validator = self._attach_id_validation(form_vars["identification"], id_entry, id_error)

        self._create_pill_label(form, "Edad").grid(row=1, column=0, sticky=tk.NW, pady=6)
        age_container, age_entry, age_error = self._create_validated_entry(
            form, form_vars["age"], width=120
        )
        age_container.grid(row=1, column=1, sticky=tk.NW, padx=12, pady=6)
        age_entry.configure(validate="key", validatecommand=(self.numeric_validation_cmd, "%P"))
        age_validator = self._attach_numeric_validation(form_vars["age"], age_entry, age_error, "Solo numeros")

        self._create_pill_label(form, "Cargo").grid(row=1, column=2, sticky=tk.NW, pady=6)
        position_container, position_entry, position_error = self._create_validated_entry(
            form, form_vars["position"], width=220
        )
        position_container.grid(row=1, column=3, sticky="nsew", padx=12, pady=6)
        position_validator = self._attach_required_validation(
            form_vars["position"], position_entry, position_error, "Requerido"
        )

        self._create_pill_label(form, "Resultado").grid(row=2, column=0, sticky=tk.NW, pady=6)
        result_container, result_combo, result_error = self._create_validated_combo(
            form,
            form_vars["result"],
            result_values,
            command=lambda _value=None, key=dataset_key: self._update_result_preview(key),
            width=240,
        )
        result_container.grid(row=2, column=1, sticky=tk.NW, padx=12, pady=6)
        result_validator = self._attach_required_validation(
            form_vars["result"], result_combo, result_error, "Requerido"
        )

        preview_container = ctk.CTkFrame(form, fg_color="transparent", corner_radius=0)
        preview_container.grid(row=2, column=2, columnspan=2, sticky=tk.NW, padx=12, pady=6)
        ctk.CTkLabel(
            preview_container,
            text="Color del resultado",
            font=ctk.CTkFont("Segoe UI", 10),
            text_color=self.colors["text_muted"],
        ).pack(anchor=tk.W)
        preview_label = ctk.CTkLabel(
            preview_container,
            text="",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=999,
            height=28,
            fg_color=self.colors["field_bg"],
            text_color=self.colors["text"],
            padx=12,
        )
        preview_label.pack(anchor=tk.W, pady=(4, 0))

        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        button_frame = ctk.CTkFrame(block_body, fg_color="transparent", corner_radius=0)
        button_frame.pack(fill=tk.X, pady=(8, 4))
        ctk.CTkButton(
            button_frame,
            text="Agregar resultado",
            fg_color=self.colors["primary"],
            text_color=self.colors["surface"],
            hover_color=self.colors["primary_dark"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=lambda key=dataset_key: self._add_result_entry(key),
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            button_frame,
            text="Eliminar seleccionado",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=lambda key=dataset_key: self._remove_selected_entry(key),
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            button_frame,
            text="Limpiar lista",
            fg_color="transparent",
            text_color=self.colors["primary"],
            border_width=1,
            border_color=self.colors["primary"],
            hover_color=self.colors["primary_muted"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=lambda key=dataset_key: self._clear_results_entries(key),
        ).pack(side=tk.LEFT, padx=4)

        tree_frame = ctk.CTkFrame(block_body, fg_color="transparent", corner_radius=0)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        columns = ("index", "name", "identification", "age", "position", "result")
        tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            selectmode='extended',
            height=10,
        )
        scrollbar = ctk.CTkScrollbar(
            tree_frame,
            orientation="vertical",
            command=tree.yview,
            width=12,
            corner_radius=999,
            fg_color=self.colors["border"],
            button_color=self.colors["primary"],
            button_hover_color=self.colors["primary_dark"],
        )
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(6, 0))

        headings = {
            "index": "N°",
            "name": "Nombre",
            "identification": "Cédula",
            "age": "Edad",
            "position": "Cargo",
            "result": "Resultado",
        }
        for column, title in headings.items():
            tree.heading(column, text=title)

        tree.column("index", width=60, anchor=tk.CENTER, stretch=False)
        tree.column("name", width=220, anchor=tk.W)
        tree.column("identification", width=140, anchor=tk.CENTER)
        tree.column("age", width=80, anchor=tk.CENTER, stretch=False)
        tree.column("position", width=200, anchor=tk.W)
        tree.column("result", width=200, anchor=tk.CENTER)

        self.result_blocks[dataset_key] = {
            "form_vars": form_vars,
            "preview_label": preview_label,
            "tree": tree,
            "validators": {
                "name": name_validator,
                "identification": id_validator,
                "age": age_validator,
                "position": position_validator,
                "result": result_validator,
            },
        }
        self._update_result_preview(dataset_key)
        self._configure_result_tags(dataset_key)
        self._refresh_results_table(dataset_key)

    def _build_calibration_section(self, frame: ttk.Frame):
        """Permite adjuntar los certificados de calibración en PDF."""
        for child in frame.winfo_children():
            child.destroy()

        header = self._create_section_header(
            frame,
            self.calibration_section_name,
            "Carga los certificados en formato PDF para insertarlos directamente en el informe.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        upload_card, upload_body = self._create_card(frame)
        upload_card.pack(fill=tk.X, pady=(0, 12))

        ctk.CTkLabel(
            upload_body,
            text="Arrastra archivos aqui o selecciona desde tu equipo",
            font=ctk.CTkFont("Segoe UI", 11),
            text_color=self.colors["text_muted"],
        ).pack(anchor=tk.CENTER, pady=(8, 10))
        ctk.CTkButton(
            upload_body,
            text="Seleccionar archivos",
            fg_color=self.colors["primary"],
            text_color=self.colors["surface"],
            hover_color=self.colors["primary_dark"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=36,
            command=self._add_calibration_file,
        ).pack()

        action_row = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=0)
        action_row.pack(fill=tk.X, pady=(0, 8))
        ctk.CTkButton(
            action_row,
            text="Ver seleccionado",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._open_selected_calibration_file,
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            action_row,
            text="Eliminar seleccionado",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._remove_selected_calibration_file,
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            action_row,
            text="Limpiar lista",
            fg_color="transparent",
            text_color=self.colors["primary"],
            border_width=1,
            border_color=self.colors["primary"],
            hover_color=self.colors["primary_muted"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._clear_calibration_files,
        ).pack(side=tk.LEFT, padx=4)

        tree_frame = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=0)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        columns = ("file", "status")
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=10, selectmode='extended')
        scrollbar = ctk.CTkScrollbar(
            tree_frame,
            orientation="vertical",
            command=tree.yview,
            width=12,
            corner_radius=999,
            fg_color=self.colors["border"],
            button_color=self.colors["primary"],
            button_hover_color=self.colors["primary_dark"],
        )
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(6, 0))

        tree.heading("file", text="Archivo")
        tree.heading("status", text="Estado")
        tree.column("file", width=380, anchor=tk.W)
        tree.column("status", width=140, anchor=tk.CENTER)

        self.calibration_tree = tree
        self._refresh_calibration_table()

    def _add_calibration_file(self):
        """Selecciona y agrega un archivo PDF de certificado."""

        file_path = filedialog.askopenfilename(
            title="Selecciona el certificado en PDF",
            filetypes=(("PDF", "*.pdf"), ("Todos los archivos", "*.*")),
        )
        if not file_path:
            return

        if file_path not in self.calibration_files:
            self.calibration_files.append(file_path)
        self._refresh_calibration_table()

    def _remove_selected_calibration_file(self):
        """Elimina los archivos seleccionados de la lista."""

        if not self.calibration_tree:
            return

        selected = self.calibration_tree.selection()
        if not selected:
            messagebox.showinfo("Selecciona un archivo", "Primero selecciona un certificado para eliminarlo")
            return

        indexes = sorted({int(item) for item in selected if str(item).isdigit()}, reverse=True)
        for idx in indexes:
            if 0 <= idx < len(self.calibration_files):
                self.calibration_files.pop(idx)

        self._refresh_calibration_table()

    def _clear_calibration_files(self):
        """Limpia toda la lista de certificados cargados."""

        if not self.calibration_files:
            return

        confirm = messagebox.askyesno(
            "Confirmar limpieza",
            "Se eliminarán todos los certificados cargados. ¿Deseas continuar?",
        )
        if not confirm:
            return

        self.calibration_files.clear()
        self._refresh_calibration_table()

    def _open_selected_calibration_file(self):
        """Abre el certificado PDF seleccionado."""

        if not self.calibration_tree:
            return

        selected = self.calibration_tree.selection()
        if not selected:
            messagebox.showinfo("Selecciona un archivo", "Elige un certificado para abrirlo")
            return

        idx = selected[0]
        if not str(idx).isdigit():
            return
        index = int(idx)
        if 0 <= index < len(self.calibration_files):
            self.open_pdf(self.calibration_files[index])

    def _refresh_calibration_table(self):
        """Refresca la vista con los archivos seleccionados."""

        if not self.calibration_tree:
            return

        tree = self.calibration_tree
        for row in tree.get_children():
            tree.delete(row)

        for idx, file_path in enumerate(self.calibration_files):
            exists = os.path.exists(file_path)
            status = "Disponible" if exists else "No encontrado"
            tree.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(Path(file_path).name, status),
            )

    def _get_calibration_files(self):
        """Devuelve solo los archivos que existen actualmente."""

        return [path for path in self.calibration_files if os.path.exists(path)]

    def _build_test_attachments_section(self, frame: ttk.Frame):
        """Sección para cargar audiogramas y reportes de espirometría."""
        for child in frame.winfo_children():
            child.destroy()

        header = self._create_section_header(
            frame,
            self.report_attachments_section_name,
            "Adjunta los archivos PDF exportados por los equipos de medicion. Se incluiran como anexos del informe.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        dataset_specs = [
            ("audiometria", "Audiogramas (PDF)", "Disponible cuando el informe incluye audiometrías."),
            ("espirometria", "Reportes de espirometría (PDF)", "Disponible cuando el informe incluye espirometrías."),
        ]

        for dataset_key, title, helper in dataset_specs:
            block_card, block_body = self._create_card(frame)
            block_card.pack(fill=tk.BOTH, expand=True, pady=8)

            title_row = ctk.CTkFrame(block_body, fg_color="transparent", corner_radius=0)
            title_row.pack(fill=tk.X, pady=(0, 10))
            self._create_pill_label(title_row, title).pack(anchor=tk.W)

            btn_frame = ctk.CTkFrame(block_body, fg_color="transparent", corner_radius=0)
            btn_frame.pack(fill=tk.X, pady=(0, 6))

            add_btn = ctk.CTkButton(
                btn_frame,
                text="Agregar PDF",
                fg_color=self.colors["primary"],
                text_color=self.colors["surface"],
                hover_color=self.colors["primary_dark"],
                font=ctk.CTkFont("Segoe UI", 10, "bold"),
                corner_radius=10,
                height=34,
                command=lambda key=dataset_key: self._add_test_attachment_file(key),
            )
            add_btn.pack(side=tk.LEFT, padx=4)

            view_btn = ctk.CTkButton(
                btn_frame,
                text="Ver seleccionado",
                fg_color=self.colors["primary_muted"],
                text_color=self.colors["primary"],
                hover_color="#D4EDDA",
                font=ctk.CTkFont("Segoe UI", 10, "bold"),
                corner_radius=10,
                height=34,
                command=lambda key=dataset_key: self._open_selected_test_attachment_file(key),
            )
            view_btn.pack(side=tk.LEFT, padx=4)

            remove_btn = ctk.CTkButton(
                btn_frame,
                text="Eliminar seleccionado",
                fg_color=self.colors["primary_muted"],
                text_color=self.colors["primary"],
                hover_color="#D4EDDA",
                font=ctk.CTkFont("Segoe UI", 10, "bold"),
                corner_radius=10,
                height=34,
                command=lambda key=dataset_key: self._remove_selected_test_attachment_file(key),
            )
            remove_btn.pack(side=tk.LEFT, padx=4)

            clear_btn = ctk.CTkButton(
                btn_frame,
                text="Limpiar lista",
                fg_color="transparent",
                text_color=self.colors["primary"],
                border_width=1,
                border_color=self.colors["primary"],
                hover_color=self.colors["primary_muted"],
                font=ctk.CTkFont("Segoe UI", 10, "bold"),
                corner_radius=10,
                height=34,
                command=lambda key=dataset_key: self._clear_test_attachment_files(key),
            )
            clear_btn.pack(side=tk.LEFT, padx=4)

            tree_frame = ctk.CTkFrame(block_body, fg_color="transparent", corner_radius=0)
            tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
            columns = ("file", "status")
            tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=6, selectmode='extended')
            scrollbar = ctk.CTkScrollbar(
                tree_frame,
                orientation="vertical",
                command=tree.yview,
                width=12,
                corner_radius=999,
                fg_color=self.colors["border"],
                button_color=self.colors["primary"],
                button_hover_color=self.colors["primary_dark"],
            )
            tree.configure(yscrollcommand=scrollbar.set)
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(6, 0))

            tree.heading("file", text="Archivo")
            tree.heading("status", text="Estado")
            tree.column("file", width=320, anchor=tk.W)
            tree.column("status", width=140, anchor=tk.CENTER)

            status_label = ctk.CTkLabel(
                block_body,
                text=helper,
                font=ctk.CTkFont("Segoe UI", 10),
                text_color=self.colors["text_muted"],
                anchor="w",
            )
            status_label.pack(anchor=tk.W, padx=4, pady=(4, 0))

            self.test_attachment_trees[dataset_key] = tree
            self.test_attachment_buttons[dataset_key] = {
                "add": add_btn,
                "view": view_btn,
                "remove": remove_btn,
                "clear": clear_btn,
            }
            self.test_attachment_status[dataset_key] = status_label
            self._refresh_test_attachment_table(dataset_key)

        self._update_test_attachment_state()

        ttk.Separator(frame, orient='horizontal').pack(fill=tk.X, pady=10)

    def _add_test_attachment_file(self, dataset_key: str):
        """Agrega un PDF de audiometría o espirometría."""

        files = self.test_attachment_files.setdefault(dataset_key, [])
        file_path = filedialog.askopenfilename(
            title="Selecciona el archivo PDF",
            filetypes=(("PDF", "*.pdf"), ("Todos los archivos", "*.*")),
        )
        if not file_path:
            return

        if file_path not in files:
            files.append(file_path)
        self._refresh_test_attachment_table(dataset_key)

    def _open_selected_test_attachment_file(self, dataset_key: str):
        """Abre el PDF elegido dentro del bloque indicado."""

        tree = self.test_attachment_trees.get(dataset_key)
        files = self.test_attachment_files.get(dataset_key, [])
        if not tree or not files:
            return

        selected = tree.selection()
        if not selected:
            messagebox.showinfo("Selecciona un archivo", "Elige un PDF para visualizarlo")
            return

        idx = selected[0]
        if not str(idx).isdigit():
            return
        index = int(idx)
        if 0 <= index < len(files):
            self.open_pdf(files[index])

    def _remove_selected_test_attachment_file(self, dataset_key: str):
        """Elimina los PDFs seleccionados en el bloque indicado."""

        tree = self.test_attachment_trees.get(dataset_key)
        files = self.test_attachment_files.get(dataset_key, [])
        if not tree or not files:
            return

        selected = tree.selection()
        if not selected:
            messagebox.showinfo("Selecciona un archivo", "Primero selecciona al menos un PDF para quitarlo")
            return

        indexes = sorted({int(item) for item in selected if str(item).isdigit()}, reverse=True)
        for idx in indexes:
            if 0 <= idx < len(files):
                files.pop(idx)

        self._refresh_test_attachment_table(dataset_key)

    def _clear_test_attachment_files(self, dataset_key: str):
        """Vacía la lista del bloque de adjuntos indicado."""

        files = self.test_attachment_files.get(dataset_key)
        if not files:
            return

        confirm = messagebox.askyesno(
            "Confirmar limpieza",
            "Se eliminarán todos los archivos cargados en este bloque. ¿Deseas continuar?",
        )
        if not confirm:
            return

        files.clear()
        self._refresh_test_attachment_table(dataset_key)

    def _refresh_test_attachment_table(self, dataset_key: str):
        """Refresca la tabla de adjuntos por tipo de prueba."""

        tree = self.test_attachment_trees.get(dataset_key)
        if not tree:
            return

        for row in tree.get_children():
            tree.delete(row)

        files = self.test_attachment_files.get(dataset_key, [])
        for idx, file_path in enumerate(files):
            exists = os.path.exists(file_path)
            status = "Disponible" if exists else "No encontrado"
            tree.insert("", tk.END, iid=str(idx), values=(Path(file_path).name, status))

    def _get_test_attachment_files(self, dataset_key: str):
        """Devuelve los adjuntos existentes del tipo solicitado."""

        files = self.test_attachment_files.get(dataset_key, [])
        return [path for path in files if os.path.exists(path)]

    def _build_attendance_section(self, frame: ttk.Frame):
        """Permite cargar los listados de asistencia firmados en PDF."""
        for child in frame.winfo_children():
            child.destroy()

        header = self._create_section_header(
            frame,
            self.attendance_section_name,
            "Adjunta los listados oficiales de asistencia para integrarlos en el informe.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        upload_card, upload_body = self._create_card(frame)
        upload_card.pack(fill=tk.X, pady=(0, 12))

        ctk.CTkLabel(
            upload_body,
            text="Arrastra archivos aqui o selecciona desde tu equipo",
            font=ctk.CTkFont("Segoe UI", 11),
            text_color=self.colors["text_muted"],
        ).pack(anchor=tk.CENTER, pady=(8, 10))
        ctk.CTkButton(
            upload_body,
            text="Seleccionar archivos",
            fg_color=self.colors["primary"],
            text_color=self.colors["surface"],
            hover_color=self.colors["primary_dark"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=36,
            command=self._add_attendance_file,
        ).pack()

        button_frame = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=0)
        button_frame.pack(fill=tk.X, pady=6)
        ctk.CTkButton(
            button_frame,
            text="Ver seleccionado",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._open_selected_attendance_file,
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            button_frame,
            text="Eliminar seleccionado",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._remove_selected_attendance_file,
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            button_frame,
            text="Limpiar lista",
            fg_color="transparent",
            text_color=self.colors["primary"],
            border_width=1,
            border_color=self.colors["primary"],
            hover_color=self.colors["primary_muted"],
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._clear_attendance_files,
        ).pack(side=tk.LEFT, padx=4)

        tree_frame = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=0)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        columns = ("file", "status")
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=6, selectmode='extended')
        scrollbar = ctk.CTkScrollbar(
            tree_frame,
            orientation="vertical",
            command=tree.yview,
            width=12,
            corner_radius=999,
            fg_color=self.colors["border"],
            button_color=self.colors["primary"],
            button_hover_color=self.colors["primary_dark"],
        )
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(6, 0))

        tree.heading("file", text="Archivo")
        tree.heading("status", text="Estado")
        tree.column("file", width=360, anchor=tk.W)
        tree.column("status", width=140, anchor=tk.CENTER)

        self.attendance_status_label = ctk.CTkLabel(
            frame,
            text="Los listados se insertaran inmediatamente despues de los resultados.",
            font=ctk.CTkFont("Segoe UI", 10),
            text_color=self.colors["text_muted"],
            anchor="w",
        )
        self.attendance_status_label.pack(anchor=tk.W, pady=(6, 0))

        self.attendance_tree = tree
        self._refresh_attendance_table()

    def _add_attendance_file(self):
        """Selecciona un PDF para el listado de asistencia."""

        file_path = filedialog.askopenfilename(
            title="Selecciona el listado en PDF",
            filetypes=(("PDF", "*.pdf"), ("Todos los archivos", "*.*")),
        )
        if not file_path:
            return

        if file_path not in self.attendance_files:
            self.attendance_files.append(file_path)
        self._refresh_attendance_table()

    def _open_selected_attendance_file(self):
        """Abre el PDF seleccionado en la tabla de asistencia."""

        if not self.attendance_tree:
            return
        selected = self.attendance_tree.selection()
        if not selected:
            messagebox.showinfo("Selecciona un archivo", "Elige un listado para abrirlo")
            return

        idx = selected[0]
        if not str(idx).isdigit():
            return
        index = int(idx)
        if 0 <= index < len(self.attendance_files):
            self.open_pdf(self.attendance_files[index])

    def _remove_selected_attendance_file(self):
        """Elimina los PDF seleccionados en el bloque de asistencia."""

        if not self.attendance_tree:
            return

        selected = self.attendance_tree.selection()
        if not selected:
            messagebox.showinfo("Selecciona un archivo", "Primero selecciona un listado para eliminarlo")
            return

        indexes = sorted({int(item) for item in selected if str(item).isdigit()}, reverse=True)
        for idx in indexes:
            if 0 <= idx < len(self.attendance_files):
                self.attendance_files.pop(idx)

        self._refresh_attendance_table()

    def _clear_attendance_files(self):
        """Vacía la lista de listados de asistencia."""

        if not self.attendance_files:
            return

        confirm = messagebox.askyesno(
            "Confirmar limpieza",
            "Se eliminarán todos los listados cargados. ¿Deseas continuar?",
        )
        if not confirm:
            return

        self.attendance_files.clear()
        self._refresh_attendance_table()

    def _refresh_attendance_table(self):
        """Refresca la tabla con los listados cargados."""

        if not self.attendance_tree:
            return

        tree = self.attendance_tree
        for row in tree.get_children():
            tree.delete(row)

        for idx, file_path in enumerate(self.attendance_files):
            exists = os.path.exists(file_path)
            status = "Disponible" if exists else "No encontrado"
            tree.insert("", tk.END, iid=str(idx), values=(Path(file_path).name, status))

    def _get_attendance_files(self) -> list:
        """Devuelve únicamente los listados existentes."""

        return [path for path in self.attendance_files if os.path.exists(path)]

    def _update_test_attachment_state(self):
        """Habilita o deshabilita los bloques según el tipo de informe seleccionado."""

        active_dataset_keys = set(self._determine_result_dataset_keys())
        for dataset_key in self.test_attachment_files.keys():
            self._set_test_attachment_block_state(dataset_key, dataset_key in active_dataset_keys)

    def _set_test_attachment_block_state(self, dataset_key: str, enabled: bool):
        """Aplica el estado visual a un bloque de adjuntos."""

        buttons = self.test_attachment_buttons.get(dataset_key, {})
        state = tk.NORMAL if enabled else tk.DISABLED
        for button in buttons.values():
            button.configure(state=state)

        tree = self.test_attachment_trees.get(dataset_key)
        if tree:
            if enabled:
                tree.configure(selectmode='extended')
                try:
                    tree.state(("!disabled",))
                except tk.TclError:
                    pass
            else:
                tree.selection_remove(tree.selection())
                tree.configure(selectmode='none')
                try:
                    tree.state(("disabled",))
                except tk.TclError:
                    pass

        status_label = self.test_attachment_status.get(dataset_key)
        if status_label:
            if enabled:
                status_label.configure(
                    text="Carga los PDF proporcionados por el equipo para incluirlos como anexos.",
                    text_color="#2E7D32",
                )
            else:
                dataset_name = self.test_dataset_labels.get(dataset_key, "este estudio")
                status_label.configure(
                    text=f"Disponible solo cuando se selecciona {dataset_name} en la presentación.",
                    text_color="#757575",
                )

    def _build_recommendations_section(self, frame: ttk.Frame):
        """Sección editable para las recomendaciones del informe."""
        for child in frame.winfo_children():
            child.destroy()

        header = self._create_section_header(
            frame,
            self.recommendations_section_name,
            "Puedes mantener estas recomendaciones base o ajustarlas segun los hallazgos del informe.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        card, card_body = self._create_card(frame)
        card.pack(fill=tk.BOTH, expand=True)

        self.recommendations_text_widget = tk.Text(
            card_body,
            height=16,
            wrap=tk.WORD,
            font=("Segoe UI", 11),
            padx=12,
            pady=12,
        )
        self.recommendations_text_widget.pack(fill=tk.BOTH, expand=True)
        self._reset_recommendations_text_to_default(prompt=False)

        button_frame = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=0)
        button_frame.pack(anchor=tk.E, pady=10)
        ctk.CTkButton(
            button_frame,
            text="Regenerar texto sugerido",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._reset_recommendations_text_to_default,
        ).pack(side=tk.RIGHT)

    def _build_conclusion_section(self, frame: ttk.Frame):
        """Sección Conclusión con texto editable y plantilla dinámica."""
        for child in frame.winfo_children():
            child.destroy()

        header = self._create_section_header(
            frame,
            "Conclusión del informe",
            "Puedes mantener el texto sugerido o personalizarlo segun el informe.",
        )
        header.pack(anchor=tk.W, pady=(0, 16), fill=tk.X)

        card, card_body = self._create_card(frame)
        card.pack(fill=tk.BOTH, expand=True)

        self.conclusion_text_widget = tk.Text(
            card_body,
            height=18,
            wrap=tk.WORD,
            font=("Segoe UI", 11),
            padx=12,
            pady=12,
        )
        self.conclusion_text_widget.pack(fill=tk.BOTH, expand=True)
        self._reset_conclusion_text_to_default(prompt=False)

        button_frame = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=0)
        button_frame.pack(anchor=tk.E, pady=10)
        ctk.CTkButton(
            button_frame,
            text="Regenerar texto sugerido",
            fg_color=self.colors["primary_muted"],
            text_color=self.colors["primary"],
            hover_color="#D4EDDA",
            font=ctk.CTkFont("Segoe UI", 10, "bold"),
            corner_radius=10,
            height=34,
            command=self._reset_conclusion_text_to_default,
        ).pack(side=tk.RIGHT)

    def _reset_recommendations_text_to_default(self, prompt: bool = True):
        """Restablece el texto con la plantilla de recomendaciones adecuada."""

        if not self.recommendations_text_widget:
            return

        if prompt:
            confirm = messagebox.askyesno(
                "Confirmar",
                "Se reemplazará el texto actual de recomendaciones por la plantilla automática. ¿Deseas continuar?",
            )
            if not confirm:
                return

        template = self._generate_recommendations_template()
        self.recommendations_text_widget.delete("1.0", tk.END)
        self.recommendations_text_widget.insert("1.0", template)

    def _reset_conclusion_text_to_default(self, prompt: bool = True):
        """Reemplaza el texto con la plantilla generada a partir de los datos actuales."""

        if not self.conclusion_text_widget:
            return

        if prompt:
            confirm = messagebox.askyesno(
                "Confirmar",
                "Se reemplazará el texto actual por la plantilla automática. ¿Deseas continuar?",
            )
            if not confirm:
                return

        template = self._generate_conclusion_template()
        self.conclusion_text_widget.delete("1.0", tk.END)
        self.conclusion_text_widget.insert("1.0", template)

    def _generate_conclusion_template(self) -> str:
        """Genera un texto descriptivo usando los datos ingresados en el formulario."""

        context = self._build_conclusion_context()
        report_type = (self.report_type_var.get() or "").lower()
        has_audio = "audiometr" in report_type
        has_espiro = "espiro" in report_type

        if has_audio and has_espiro:
            return self._build_combined_conclusion(context)
        if has_espiro:
            return self._build_espirometry_conclusion()
        return self._build_audiometry_conclusion(context)

    def _build_conclusion_context(self) -> dict:
        """Reúne datos comunes usados en las plantillas dinámicas."""

        company = self.company_var.get().strip() or "la empresa evaluada"
        location = self.location_var.get().strip() or "sus instalaciones"
        plant = self.plant_var.get().strip() or location
        study_dates = self.study_dates_var.get().strip() or self.date_var.get().strip()
        evaluation_label = self._resolve_study_label()
        total_people = sum(len(entries) for entries in self.evaluated_entries.values())
        people_text = (
            f"a un total de {total_people} personas" if total_people else "al personal convocado"
        )

        return {
            "company": company,
            "location": location,
            "plant": plant,
            "study_dates": study_dates,
            "evaluation_label": evaluation_label,
            "people_text": people_text,
        }

    def _build_audiometry_conclusion(self, context: dict) -> str:
        """Plantilla base enfocada únicamente en audiometrías."""

        paragraphs = [
            (
                f"La empresa {context['company']}, realizó las {context['evaluation_label']} durante {context['study_dates']}, "
                f"a los colaboradores en la planta {context['plant']}, ubicada en {context['location']}."
            ),
            (
                f"Se aplicó la prueba de {context['evaluation_label']} {context['people_text']}, cada colaborador suministró "
                "su historia clínica y laboral de forma confidencial y recibió orientación sobre los cuidados de su audición, "
                "con énfasis en la higiene auditiva y la protección laboral que debe aplicar."
            ),
            (
                "El resultado de cada evaluación, junto con la historia clínica, se encuentra documentado dentro de este informe, "
                "acompañado de un resumen diagnóstico y un gráfico porcentual de los resultados."
            ),
            (
                "El audiómetro utilizado para las mediciones corresponde al equipo Otopod, modelo Otovation, "
                "con certificado de calibración anual ISO, lo que garantiza la confiabilidad de los resultados."
            ),
        ]

        return "\n\n".join(paragraphs)

    def _generate_recommendations_template(self) -> str:
        """Devuelve las recomendaciones sugeridas según el tipo de informe."""

        report_type = (self.report_type_var.get() or "").lower()
        has_audio = "audiometr" in report_type
        has_espiro = "espiro" in report_type

        if has_audio and has_espiro:
            return self._build_combined_recommendations_text()
        if has_espiro:
            return self._build_espirometry_recommendations_text()
        return self._build_audiometry_recommendations_text()

    def _build_espirometry_conclusion(self) -> str:
        """Plantilla fija solicitada para informes de espirometría."""

        paragraphs = [
            (
                "La empresa Productos Toledano S.A., realizó la toma de las pruebas de espirometrías, "
                "en el mes de febrero el día 21 del presente año, a los colaboradores que están expuestos a partículas, "
                "en la Planta Incubadora Chorrerana S.A., en el área de La Chorrera."
            ),
            (
                "Se aplicó la prueba de espirometrías laboral a un total de 45 colaboradores, cada uno de estos suministró "
                "a nosotros los especialistas idóneos, su historia clínica de salud y laborales confidencialmente, "
                "fueron orientados sobre sus resultados y la protección laboral que deben de aplicar, para crear conciencia en ellos."
            ),
            (
                "El resultado de las pruebas de cada colaborador, junto con su historia clínica, se encuentra documentados en el interior de este informe, "
                "además de un listado resumen diagnóstico y un gráfico porcentual de los resultados."
            ),
            (
                "El Espirómetro, utilizado en la toma de las pruebas respectivamente, poseen su certificado de calibración anual, "
                "lo que garantiza la confiabilidad de los resultados."
            ),
        ]

        return "\n\n".join(paragraphs)

    def _build_combined_conclusion(self, context: dict) -> str:
        """Plantilla híbrida para informes con audiometrías y espirometrías."""

        paragraphs = [
            (
                f"La empresa {context['company']}, coordinó jornadas de audiometrías y espirometrías durante {context['study_dates']}, "
                f"para los colaboradores en la planta {context['plant']}, ubicada en {context['location']}."
            ),
            (
                f"Se aplicaron ambas pruebas {context['people_text']}, recopilando de forma confidencial la historia clínica y laboral de cada participante. "
                "Todos recibieron orientación sobre la protección auditiva y respiratoria que deben aplicar, además de las medidas de control frente a la exposición a partículas."
            ),
            (
                "Los resultados individuales y la documentación clínica se incluyen dentro del informe, junto con resúmenes diagnósticos específicos y gráficos porcentuales comparativos "
                "que permiten visualizar la tendencia de cada tipo de estudio."
            ),
            (
                "El audiómetro Otopod modelo Otovation y los espirómetros utilizados cuentan con sus respectivos certificados de calibración anual, "
                "lo que respalda la confiabilidad de todas las mediciones presentadas." 
            ),
        ]

        return "\n\n".join(paragraphs)

    def _resolve_study_label(self) -> str:
        """Devuelve una descripción corta según el tipo de estudio seleccionado."""

        report_type = (self.report_type_var.get() or "").lower()
        if "audiometr" in report_type and "espirom" in report_type:
            return "audiometrías y espirometrías"
        if "espiro" in report_type:
            return "espirometrías"
        return "audiometrías"

    def _build_audiometry_recommendations_text(self) -> str:
        lines = self._get_audiometry_recommendation_lines()
        return "\n\n".join(lines)

    def _build_espirometry_recommendations_text(self) -> str:
        lines = self._get_espirometry_recommendation_lines()
        return "\n\n".join(lines)

    def _build_combined_recommendations_text(self) -> str:
        lines = self._get_audiometry_recommendation_lines() + self._get_espirometry_recommendation_lines()
        return "\n\n".join(lines)

    def _get_audiometry_recommendation_lines(self):
        return [
            "• Se recomienda la realización periódica (anual) de la evaluación auditiva (audiometría laboral) de los colaboradores, para mantener el registro y control auditivo correspondiente.",
            "• Continuar suministrándoles los equipos de seguridad a los colaboradores, ya que ayudan a disminuir el ruido laboral y las vibraciones presentes durante la jornada.",
            "• Los colaboradores fueron evaluados mediante la prueba de audiometría y solo se presentó un caso de vigilancia ocupacional; sin embargo, se recomendó a dos colaboradoras acudir a un especialista (Médico Otorrinolaringólogo) para determinar si requieren tratamiento por molestias en la articulación temporomandibular y la base del cuello, sin repercusión en su audición.",
        ]

    def _get_espirometry_recommendation_lines(self):
        return [
            "• Se recomienda la realización periódica (anual) de las pruebas de espirometría laboral de los colaboradores, para continuar llevando el registro y control pulmonar correspondiente.",
            "• Los colaboradores no deben exponerse a contaminantes (humo de cigarrillo, químicos, entre otros) fuera de horas laborales.",
            "• Se recomienda la utilización de protección respiratoria adecuada en las áreas de exposición a partículas o polvo, para evitar futuros procesos respiratorios que puedan convertirse en diagnósticos de restricción leve o moderada.",
            "• Preservar e incentivar el uso correcto de los implementos de seguridad laboral durante toda la jornada de trabajo y no parcialmente.",
        ]

    def _get_conclusion_text(self) -> str:
        """Obtiene el texto actual o genera la plantilla si aún no se ha editado."""

        if self.conclusion_text_widget:
            current_text = self.conclusion_text_widget.get("1.0", tk.END).strip()
            if current_text:
                return current_text
        return self._generate_conclusion_template()

    def _get_recommendations_text(self) -> str:
        """Obtiene el bloque actual de recomendaciones o genera uno nuevo."""

        if self.recommendations_text_widget:
            current_text = self.recommendations_text_widget.get("1.0", tk.END).strip()
            if current_text:
                return current_text
        return self._generate_recommendations_template()

    def _configure_result_tags(self, dataset_key: str):
        """Configura los colores de fila asociados a cada tipo de resultado."""

        block = self.result_blocks.get(dataset_key)
        if not block:
            return

        tree = block["tree"]
        palette = self.result_palettes.get(dataset_key, {})
        for config in palette.values():
            tree.tag_configure(
                config["code"],
                background=config["bg"],
                foreground=config["fg"],
            )

    def _get_result_palette_by_label(self, dataset_key: str, label: str) -> dict:
        """Devuelve la configuración de color según la etiqueta seleccionada."""

        palette = self.result_palettes.get(dataset_key, {})
        if label in palette:
            return palette[label]
        if palette:
            return next(iter(palette.values()))
        return {
            "code": "normal",
            "label": "Normal",
            "bg": "#E0E0E0",
            "fg": "#1B5E20",
        }

    def _update_result_preview(self, dataset_key: str):
        """Actualiza el recuadro de vista previa del bloque indicado."""

        block = self.result_blocks.get(dataset_key)
        if not block:
            return

        form_vars = block.get("form_vars") or {}
        preview_label = block.get("preview_label")
        if not preview_label or "result" not in form_vars:
            return

        selected_label = form_vars["result"].get()
        palette = self._get_result_palette_by_label(dataset_key, selected_label)
        preview_label.configure(
            text=palette["label"].upper(),
            fg_color=palette["bg"],
            text_color=palette["fg"],
        )

    def _add_result_entry(self, dataset_key: str):
        """Agrega un registro al listado del bloque indicado."""

        block = self.result_blocks.get(dataset_key)
        if not block:
            return

        form_vars = block.get("form_vars") or {}
        validators = block.get("validators") or {}
        for validator in validators.values():
            if callable(validator):
                validator()
        required_keys = ["name", "identification", "position", "result"]
        if any(key not in form_vars for key in required_keys):
            return

        name = form_vars["name"].get().strip()
        identification = form_vars["identification"].get().strip()
        age = form_vars["age"].get().strip()
        position = form_vars["position"].get().strip()
        result_label = form_vars["result"].get().strip()

        if not all([name, identification, position, result_label]):
            return
        if not self._is_valid_panama_id(identification):
            messagebox.showwarning("Cédula inválida", "Verifica el formato de la cédula.")
            return

        palette = self._get_result_palette_by_label(dataset_key, result_label)
        entry = {
            "name": name,
            "identification": identification,
            "age": age,
            "position": position,
            "result_label": palette["label"],
            "result_code": palette["code"],
            "test_type": dataset_key,
        }

        self.evaluated_entries.setdefault(dataset_key, []).append(entry)
        self._refresh_results_table(dataset_key)
        self._sync_report_evaluated()

        form_vars["name"].set("")
        form_vars["identification"].set("")
        form_vars["age"].set("")
        form_vars["position"].set("")
        form_vars["result"].set(palette["label"])
        self._update_result_preview(dataset_key)

    def _refresh_results_table(self, dataset_key: str):
        """Actualiza la tabla visual de un bloque específico."""

        block = self.result_blocks.get(dataset_key)
        if not block:
            return

        tree = block.get("tree")
        if not tree:
            return

        for row in tree.get_children():
            tree.delete(row)

        entries = self.evaluated_entries.get(dataset_key, [])
        for idx, entry in enumerate(entries, start=1):
            values = (
                idx,
                entry.get("name", "N/A"),
                entry.get("identification", "N/A"),
                entry.get("age", ""),
                entry.get("position", ""),
                entry.get("result_label", ""),
            )
            tag = entry.get("result_code", "normal")
            tree.insert("", tk.END, values=values, tags=(tag,))

        self._configure_result_tags(dataset_key)

    def _remove_selected_entry(self, dataset_key: str):
        """Elimina los registros seleccionados de un bloque."""

        block = self.result_blocks.get(dataset_key)
        if not block:
            return

        tree = block.get("tree")
        if not tree:
            return

        selected = tree.selection()
        if not selected:
            messagebox.showinfo("Selecciona un registro", "Primero selecciona una fila para eliminarla")
            return

        indexes = sorted(
            (int(tree.set(item, "index")) - 1) for item in selected if tree.set(item, "index").isdigit()
        )

        entries = self.evaluated_entries.get(dataset_key, [])
        for idx in reversed(indexes):
            if 0 <= idx < len(entries):
                entries.pop(idx)

        self._refresh_results_table(dataset_key)
        self._sync_report_evaluated()

    def _clear_results_entries(self, dataset_key: str):
        """Limpia todos los resultados registrados en el bloque indicado."""

        entries = self.evaluated_entries.get(dataset_key)
        if not entries:
            return

        confirm = messagebox.askyesno(
            "Confirmar limpieza",
            "Se eliminarán todos los registros de resultados. ¿Deseas continuar?",
        )
        if not confirm:
            return

        entries.clear()
        self._refresh_results_table(dataset_key)
        self._sync_report_evaluated()

    def _sync_report_evaluated(self):
        """Mantiene sincronizado el informe actual con las filas capturadas."""

        if not self.current_report:
            return

        self.current_report["evaluated"] = self._build_evaluated_payload()

    def _build_evaluated_payload(self):
        """Convierte la tabla interna a la estructura usada por el PDF."""

        payload = []
        for dataset_key, entries in self.evaluated_entries.items():
            for entry in entries:
                payload.append(
                    {
                        "name": entry.get("name"),
                        "identification": entry.get("identification"),
                        "age": entry.get("age"),
                        "position": entry.get("position"),
                        "result_label": entry.get("result_label"),
                        "result_code": entry.get("result_code"),
                        "test_type": dataset_key,
                    }
                )
        return payload

    def show_section(self, section_name: str):
        """Muestra la sección seleccionada cambiando el contenido sin refrescar la vista."""

        frame = self.section_frames.get(section_name)
        if frame is None:
            return

        frame.tkraise()
        for name, button in self.section_buttons.items():
            if name == section_name:
                button.configure(
                    fg_color=self.colors["primary"],
                    text_color=self.colors["surface"],
                    hover_color=self.colors["primary_dark"],
                )
            else:
                button.configure(
                    fg_color="transparent",
                    text_color=self.colors["primary"],
                    hover_color=self.colors["primary_muted"],
                )
        self.active_section = section_name
        self._update_section_progress()

    def _update_section_progress(self) -> None:
        """Actualiza la barra de progreso del menú lateral."""

        if not self.section_names:
            return

        try:
            index = self.section_names.index(self.active_section)
        except ValueError:
            index = 0

        total = len(self.section_names)
        progress = (index + 1) / total if total else 0

        if hasattr(self, "progress_bar") and self.progress_bar is not None:
            self.progress_bar.set(progress)

        if hasattr(self, "progress_label") and self.progress_label is not None:
            self.progress_label.configure(text=f"{index + 1} de {total} secciones")

    def new_report(self, silent: bool = False) -> bool:
        """Crea un nuevo informe con los datos capturados."""
        company = self.company_var.get().strip()
        location = self.location_var.get().strip()
        evaluator = self.evaluator_var.get().strip()
        evaluator_profile = self._get_selected_evaluator_profile()
        report_type = self.report_type_var.get()
        plant = self.plant_var.get().strip()
        activity = self.activity_var.get().strip()
        company_counterpart = self.company_counterpart_var.get().strip()
        counterpart_role = self.counterpart_role_var.get().strip()
        country = self.country_var.get().strip()
        study_dates = self.study_dates_var.get().strip()

        self._normalize_date_var(self.date_var)
        self._normalize_date_var(self.study_dates_var)
        study_dates = self.study_dates_var.get()
        
        if not all([company, location, evaluator]):
            messagebox.showwarning(
                "Campos incompletos",
                "Por favor completa: Empresa, Ubicación y Evaluador",
            )
            return False

        if evaluator and not evaluator_profile:
            messagebox.showwarning(
                "Evaluador inválido",
                "Debes seleccionar un evaluador registrado o agregar uno nuevo.",
            )
            return False

        if not plant:
            plant = location

        if not study_dates:
            study_dates = self.date_var.get().strip()

        conclusion_text = self._get_conclusion_text()
        recommendations_text = self._get_recommendations_text()
        
        technical_team = self._build_technical_team_payload(report_type)
        evaluator_credentials = []
        for member in technical_team:
            credential_file = (member.get("credential_file") or "").strip()
            if credential_file and not Path(credential_file).is_absolute():
                credential_file = str(self.project_root / credential_file)
            if credential_file:
                evaluator_credentials.append({
                    "name": member.get("name", ""),
                    "file": credential_file,
                })
        calibration_files = self._get_calibration_files()
        calibration_certificates = [{"file": path} for path in calibration_files]
        active_dataset_keys = set(self._determine_result_dataset_keys())
        audiogram_files = (
            self._get_test_attachment_files("audiometria") if "audiometria" in active_dataset_keys else []
        )
        spirometry_files = (
            self._get_test_attachment_files("espirometria") if "espirometria" in active_dataset_keys else []
        )
        attendance_files = self._get_attendance_files()
        combined_attachments = []
        for file_path in calibration_files + audiogram_files + spirometry_files + attendance_files:
            if file_path not in combined_attachments:
                combined_attachments.append(file_path)

        attachment_folder_links = {
            "calibration": str(Path(calibration_files[0]).parent) if calibration_files else "",
            "audiometria": str(Path(audiogram_files[0]).parent) if audiogram_files else "",
            "espirometria": str(Path(spirometry_files[0]).parent) if spirometry_files else "",
            "attendance": str(Path(attendance_files[0]).parent) if attendance_files else "",
        }

        self.current_report = {
            "id": f"REP_{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "type": report_type,
            "company": company,
            "location": location,
            "evaluator": evaluator_profile.get("name", evaluator) if evaluator_profile else evaluator,
            "date": self.date_var.get(),
            "plant": plant,
            "activity": activity,
            "company_counterpart": company_counterpart,
            "counterpart_role": counterpart_role,
            "country": country if country else location,
            "study_dates": study_dates,
            "evaluated": self._build_evaluated_payload(),
            "attachments": combined_attachments,
            "logo_path": self.logo_path,
            "conclusion": conclusion_text,
            "recommendations": recommendations_text,
            "evaluator_profile": evaluator_profile,
            "technical_team": technical_team,
            "evaluator_credentials": evaluator_credentials,
            "calibration_certificates": calibration_certificates,
            "calibration_files": list(calibration_files),
            "audiogram_files": list(audiogram_files),
            "spirometry_files": list(spirometry_files),
            "attendance_files": list(attendance_files),
        }
        
        self.status_label.configure(text=f"✅ Informe creado: {self.current_report['id']} ({report_type})")
        if not silent:
            messagebox.showinfo(
                "Éxito",
                f"Informe creado exitosamente:\n\nID: {self.current_report['id']}\n"
                f"Tipo: {report_type}\n"
                f"Empresa: {company}",
            )
        return True
    
    def _generate_pdf_file(self, pdf_path: Path) -> bool:
        """Genera el PDF actual en la ruta indicada."""

        if not self.current_report:
            return False

        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        from src.services.pdf_generator import PDFGenerator

        pdf_gen = PDFGenerator()
        logo_path = None
        if self.app_logo_path and self.app_logo_path.exists():
            logo_path = str(self.app_logo_path)
        return pdf_gen.generate(self.current_report, str(pdf_path), logo_path)

    def _build_export_base_name(self) -> str:
        """Construye el nombre base 'Informe <empresa> <año> <tipo>' para archivos exportados."""

        if not self.current_report:
            return "Informe"

        company = (self.current_report.get("company") or "Empresa").strip()
        report_type = (self.current_report.get("type") or "Estudio").strip()
        year = self._resolve_report_year()
        base_name = f"Informe {company} {year} {report_type}".strip()
        return self._sanitize_filename(base_name)

    def _resolve_report_year(self) -> str:
        """Determina el año del informe utilizando las fechas capturadas."""

        candidates = [
            self.current_report.get("study_dates") if self.current_report else "",
            self.current_report.get("date") if self.current_report else "",
            datetime.now().strftime("%Y"),
        ]

        for value in candidates:
            year = self._extract_year_from_text(value)
            if year:
                return year
        return datetime.now().strftime("%Y")

    def _extract_year_from_text(self, text: str) -> str:
        """Busca un año en cualquier cadena con formato libre."""

        if not text:
            return ""
        match = re.search(r"(19|20)\d{2}", text)
        return match.group(0) if match else ""

    def _sanitize_filename(self, value: str) -> str:
        """Remueve caracteres inválidos para nombres de archivos."""

        sanitized = re.sub(r'[<>:"/\\|?*]', "_", value)
        sanitized = re.sub(r"\s+", " ", sanitized).strip()
        return sanitized or "Informe"

    def _copy_report_attachments(self, package_dir: Path) -> int:
        """Copia los archivos adjuntos a la carpeta de exportación y devuelve la cantidad copiada."""

        if not self.current_report:
            return 0

        attachments_root = package_dir / "Adjuntos"
        report = self.current_report
        groups = [
            ("Calibracion", report.get("calibration_files") or []),
            ("Audiometrias", report.get("audiogram_files") or []),
            ("Espirometrias", report.get("spirometry_files") or []),
            ("Asistencia", report.get("attendance_files") or []),
        ]

        copied = 0
        seen_paths = set()
        created_root = False

        def ensure_root():
            nonlocal created_root
            if not created_root:
                attachments_root.mkdir(parents=True, exist_ok=True)
                created_root = True

        def normalize_path(path_obj: Path) -> str:
            try:
                return str(path_obj.resolve())
            except OSError:
                return str(path_obj)

        for folder_name, file_list in groups:
            existing_files = []
            for file_path in file_list:
                if not file_path:
                    continue
                candidate = Path(file_path)
                if not candidate.exists():
                    continue
                normalized = normalize_path(candidate)
                if normalized in seen_paths:
                    continue
                seen_paths.add(normalized)
                existing_files.append(candidate)

            if not existing_files:
                continue

            ensure_root()
            target_folder = attachments_root / folder_name
            target_folder.mkdir(parents=True, exist_ok=True)
            for source in existing_files:
                self._copy_file_with_unique_name(source, target_folder)
                copied += 1

        other_files = []
        for file_path in report.get("attachments") or []:
            if not file_path:
                continue
            candidate = Path(file_path)
            if not candidate.exists():
                continue
            normalized = normalize_path(candidate)
            if normalized in seen_paths:
                continue
            seen_paths.add(normalized)
            other_files.append(candidate)

        if other_files:
            ensure_root()
            target_folder = attachments_root / "Otros"
            target_folder.mkdir(parents=True, exist_ok=True)
            for source in other_files:
                self._copy_file_with_unique_name(source, target_folder)
                copied += 1

        return copied

    def _copy_evaluator_credentials(self, package_dir: Path) -> tuple[int, list]:
        """Copia las idoneidades por evaluador y devuelve cantidad y rutas destino."""

        if not self.current_report:
            return 0, []

        credentials = self.current_report.get("evaluator_credentials") or []
        if not credentials:
            return 0, []

        attachments_root = package_dir / "Adjuntos" / "Idoneidad"
        copied = 0
        destinations = []

        for entry in credentials:
            source_path = (entry.get("file") or "").strip()
            name = (entry.get("name") or "Evaluador").strip() or "Evaluador"
            if not source_path:
                continue
            source = Path(source_path)
            if not source.is_absolute():
                source = self.project_root / source
            if not source.exists():
                continue
            target_dir = attachments_root / self._sanitize_filename(name)
            target_dir.mkdir(parents=True, exist_ok=True)
            destination = target_dir / source.name
            shutil.copy2(source, destination)
            copied += 1
            destinations.append({
                "name": name,
                "file": destination,
                "folder": target_dir,
            })

        return copied, destinations

    def _copy_file_with_unique_name(self, source: Path, target_dir: Path) -> None:
        """Copia un archivo asegurando que no se sobrescriba otro con el mismo nombre."""

        destination = target_dir / source.name
        if not destination.exists():
            shutil.copy2(source, destination)
            return

        counter = 1
        stem = source.stem
        suffix = source.suffix
        while True:
            candidate = target_dir / f"{stem}_{counter}{suffix}"
            if not candidate.exists():
                shutil.copy2(source, candidate)
                break
            counter += 1

    def create_and_generate_pdf(self):
        """Crea el informe y genera el PDF en un solo paso"""
        created = self.new_report()
        if not created or not self.current_report:
            return

        if self.generate_pdf():
            self._reset_form_state()
    
    def generate_pdf(self) -> bool:
        """Genera el PDF del informe y devuelve True cuando se completa."""
        if not self.current_report:
            messagebox.showwarning(
                "Informe no creado",
                "Debes crear un informe primero",
            )
            return False
        
        self._sync_report_evaluated()

        try:
            exports_dir = self.project_root / "data" / "exports"
            exports_dir.mkdir(parents=True, exist_ok=True)

            default_name = f"{self.current_report['id']}.pdf"
            selected_path = filedialog.asksaveasfilename(
                title="Guardar PDF",
                defaultextension=".pdf",
                initialdir=str(exports_dir),
                initialfile=default_name,
                filetypes=(("PDF", "*.pdf"), ("Todos los archivos", "*.*")),
            )
            if not selected_path:
                return False

            pdf_path = Path(selected_path)
            
            success = self._generate_pdf_file(pdf_path)
            
            if success:
                # Mostrar mensaje de éxito con opción de abrir
                result = messagebox.askyesno("PDF Generado", 
                                   f"PDF generado exitosamente en:\n\n{pdf_path}\n\n"
                                   f"¿Deseas abrir el PDF ahora?")
                
                if result:
                    self.open_pdf(str(pdf_path))
                
                self.status_label.configure(text=f"✅ PDF generado: {pdf_path}")
                return True
            else:
                messagebox.showerror("Error", "No se pudo generar el PDF")
                return False
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar PDF: {e}")
            print(f"Error: {e}")
            return False

    def _collect_report_state(self) -> dict:
        """Construye el estado completo del formulario para guardarlo."""

        return {
            "version": 1,
            "report_type": self.report_type_var.get(),
            "company": self.company_var.get(),
            "location": self.location_var.get(),
            "date": self.date_var.get(),
            "plant": self.plant_var.get(),
            "activity": self.activity_var.get(),
            "company_counterpart": self.company_counterpart_var.get(),
            "counterpart_role": self.counterpart_role_var.get(),
            "country": self.country_var.get(),
            "study_dates": self.study_dates_var.get(),
            "selected_evaluator_id": self.selected_evaluator_id.get(),
            "evaluator_name": self.evaluator_var.get(),
            "selected_counterpart_id": self.selected_counterpart_id.get(),
            "counterpart_label": self.counterpart_var.get(),
            "evaluated_entries": self.evaluated_entries,
            "calibration_files": list(self.calibration_files),
            "test_attachment_files": self.test_attachment_files,
            "attendance_files": list(self.attendance_files),
            "conclusion_text": self._get_conclusion_text(),
            "recommendations_text": self._get_recommendations_text(),
        }

    def _apply_report_state(self, state: dict) -> None:
        """Restaura el estado completo del formulario desde un borrador."""

        self.report_type_var.set(state.get("report_type", "audiometría"))
        self.company_var.set(state.get("company", ""))
        self.location_var.set(state.get("location", ""))
        self.date_var.set(state.get("date", datetime.now().strftime("%d/%m/%Y")))
        self.plant_var.set(state.get("plant", ""))
        self.activity_var.set(state.get("activity", ""))
        self.company_counterpart_var.set(state.get("company_counterpart", ""))
        self.counterpart_role_var.set(state.get("counterpart_role", ""))
        self.country_var.set(state.get("country", ""))
        self.study_dates_var.set(state.get("study_dates", self.date_var.get()))

        self.evaluated_entries = {key: [] for key in RESULT_SCHEMES}
        for key, entries in (state.get("evaluated_entries") or {}).items():
            if key in self.evaluated_entries and isinstance(entries, list):
                self.evaluated_entries[key] = entries

        self.calibration_files = list(state.get("calibration_files") or [])
        self.test_attachment_files = state.get("test_attachment_files") or {"audiometria": [], "espirometria": []}
        self.attendance_files = list(state.get("attendance_files") or [])

        self._handle_report_type_change()

        evaluator_id = state.get("selected_evaluator_id")
        if evaluator_id and evaluator_id in self.evaluator_profiles:
            self._select_evaluator(evaluator_id, update_combo=True)
        else:
            selected_name = (state.get("evaluator_name") or "").strip()
            self.evaluator_var.set(selected_name)
            self._handle_evaluator_selection()

        counterpart_id = state.get("selected_counterpart_id")
        if counterpart_id and counterpart_id in self.counterpart_profiles:
            self._select_counterpart(counterpart_id, update_combo=True)
        else:
            selected_label = (state.get("counterpart_label") or "Sin contraparte").strip()
            self.counterpart_var.set(selected_label)
            if self.counterpart_combo is not None:
                self.counterpart_combo.set(selected_label)
            self._handle_counterpart_selection()

        if self.calibration_tree is not None:
            self._refresh_calibration_table()
        for dataset_key in list(self.test_attachment_files.keys()):
            self._refresh_test_attachment_table(dataset_key)
        if self.attendance_tree is not None:
            self._refresh_attendance_table()

        if self.conclusion_text_widget is not None:
            self.conclusion_text_widget.delete("1.0", tk.END)
            self.conclusion_text_widget.insert("1.0", state.get("conclusion_text", ""))
        if self.recommendations_text_widget is not None:
            self.recommendations_text_widget.delete("1.0", tk.END)
            self.recommendations_text_widget.insert("1.0", state.get("recommendations_text", ""))

        self.status_label.configure(text="Borrador cargado correctamente.")

    def save_report_draft(self) -> None:
        """Guarda un borrador en JSON para reabrirlo luego."""

        drafts_dir = self.project_root / "data" / "reports"
        drafts_dir.mkdir(parents=True, exist_ok=True)
        default_name = f"borrador_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        target = filedialog.asksaveasfilename(
            title="Guardar borrador",
            defaultextension=".json",
            initialdir=str(drafts_dir),
            initialfile=default_name,
            filetypes=(("JSON", "*.json"), ("Todos los archivos", "*.*")),
        )
        if not target:
            return

        state = self._collect_report_state()
        try:
            with open(target, "w", encoding="utf-8") as handler:
                json.dump(state, handler, ensure_ascii=False, indent=2)
            self.status_label.configure(text=f"Borrador guardado: {Path(target).name}")
        except OSError as exc:
            messagebox.showerror("Error", f"No se pudo guardar el borrador: {exc}")

    def load_report_draft(self) -> None:
        """Abre la pantalla interna de borradores para seleccionar cuál cargar."""

        draft_files = self._list_draft_files()
        if not draft_files:
            messagebox.showinfo(
                "Sin borradores",
                "No hay borradores guardados en la carpeta interna de la aplicación.",
            )
            return

        self._open_drafts_window()

    def _open_drafts_window(self) -> None:
        """Abre una ventana para administrar y cargar borradores internos."""

        if self.drafts_window is not None and self.drafts_window.winfo_exists():
            self.drafts_window.deiconify()
            self.drafts_window.lift()
            self.drafts_window.focus_force()
            self._refresh_drafts_table()
            return

        window = tk.Toplevel(self.root)
        window.title("Borradores")
        window.geometry("760x480")
        window.minsize(680, 420)
        window.transient(self.root)
        window.configure(bg=self.colors["bg"])

        self.drafts_window = window

        container = ctk.CTkFrame(window, fg_color=self.colors["bg"], corner_radius=0)
        container.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)
        self._build_drafts_section(container)

        self._center_window(window)

        def _on_close() -> None:
            self.drafts_tree = None
            self.drafts_window = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", _on_close)

        if self.drafts_tree is not None:
            first = self.drafts_tree.get_children()
            if first:
                self.drafts_tree.selection_set(first[0])
                self.drafts_tree.focus(first[0])
                self.drafts_tree.see(first[0])

        self.status_label.configure(text="Selecciona un borrador de la tabla y pulsa 'Cargar seleccionado'.")

    def _load_report_draft_from_path(self, source: str) -> None:
        """Carga un borrador desde la ruta indicada."""

        try:
            with open(source, "r", encoding="utf-8") as handler:
                state = json.load(handler)
        except (OSError, json.JSONDecodeError) as exc:
            messagebox.showerror("Error", f"No se pudo cargar el borrador: {exc}")
            return

        self._apply_report_state(state if isinstance(state, dict) else {})

    def _list_draft_files(self) -> list[Path]:
        """Devuelve los archivos JSON de borradores."""

        drafts_dir = self.project_root / "data" / "reports"
        drafts_dir.mkdir(parents=True, exist_ok=True)
        return sorted(drafts_dir.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)

    def _purge_old_drafts(self) -> None:
        """Elimina borradores con mas de 5 anos."""

        cutoff = datetime.now().timestamp() - (365 * 5 * 24 * 60 * 60)
        expired = []
        for path in self._list_draft_files():
            try:
                if path.stat().st_mtime < cutoff:
                    expired.append(path)
            except OSError:
                continue

        if not expired:
            return

        confirm = messagebox.askyesno(
            "Borradores antiguos",
            f"Se encontraron {len(expired)} borradores con mas de 5 anos. ¿Deseas eliminarlos?",
        )
        if not confirm:
            return

        for path in expired:
            try:
                path.unlink()
            except OSError:
                continue

    def _refresh_drafts_table(self) -> None:
        """Refresca la tabla de borradores."""

        if not self.drafts_tree:
            return

        tree = self.drafts_tree
        for row in tree.get_children():
            tree.delete(row)

        for path in self._list_draft_files():
            updated = datetime.fromtimestamp(path.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
            tree.insert("", tk.END, iid=str(path), values=(path.name, updated))

    def _get_selected_draft_path(self) -> Path | None:
        """Obtiene la ruta seleccionada en la tabla de borradores."""

        if not self.drafts_tree:
            return None

        selected = self.drafts_tree.selection()
        if not selected:
            return None

        try:
            return Path(selected[0])
        except OSError:
            return None

    def _load_selected_draft(self) -> None:
        """Carga el borrador seleccionado en la tabla."""

        draft_path = self._get_selected_draft_path()
        if not draft_path or not draft_path.exists():
            messagebox.showinfo("Sin seleccion", "Selecciona un borrador para cargarlo.")
            return

        self._load_report_draft_from_path(str(draft_path))

    def _delete_selected_draft(self) -> None:
        """Elimina el borrador seleccionado en la tabla."""

        draft_path = self._get_selected_draft_path()
        if not draft_path or not draft_path.exists():
            messagebox.showinfo("Sin seleccion", "Selecciona un borrador para eliminarlo.")
            return

        confirm = messagebox.askyesno(
            "Confirmar eliminacion",
            f"Se eliminara el borrador '{draft_path.name}'. ¿Deseas continuar?",
        )
        if not confirm:
            return

        try:
            draft_path.unlink()
        except OSError as exc:
            messagebox.showerror("Error", f"No se pudo eliminar el borrador: {exc}")
            return

        self._refresh_drafts_table()
    
    def _reset_form_state(self) -> None:
        """Limpia todos los campos para preparar un nuevo informe."""

        today = datetime.now().strftime("%d/%m/%Y")

        self.current_report = None
        self.logo_path = None

        for var in (
            self.company_var,
            self.location_var,
            self.plant_var,
            self.activity_var,
            self.company_counterpart_var,
            self.counterpart_role_var,
            self.country_var,
        ):
            var.set("")

        self.date_var.set(today)
        self.study_dates_var.set(today)

        self.selected_evaluator_id.set("")
        self.evaluator_var.set("")
        if self.evaluator_combo is not None:
            self.evaluator_combo.set("")
        self._update_evaluator_details_preview()

        self.selected_counterpart_id.set("")
        self.counterpart_var.set("Sin contraparte")
        if self.counterpart_combo is not None:
            self.counterpart_combo.set("Sin contraparte")
        self._set_counterpart_role_state(False)

        self.calibration_files = []
        self._refresh_calibration_table()

        for dataset_key in list(self.test_attachment_files.keys()):
            self.test_attachment_files[dataset_key] = []
            self._refresh_test_attachment_table(dataset_key)

        self.attendance_files = []
        self._refresh_attendance_table()

        self.evaluated_entries = {key: [] for key in RESULT_SCHEMES}
        self._handle_report_type_change()

        self._update_test_attachment_state()

        self._reset_recommendations_text_to_default(prompt=False)
        self._reset_conclusion_text_to_default(prompt=False)

        self.status_label.configure(text="Formulario limpio. Listo para crear un nuevo informe.")
    
    def open_pdf(self, pdf_path: str):
        """Abre el PDF en el lector predeterminado del sistema"""
        try:
            if not os.path.exists(pdf_path):
                messagebox.showerror("Error", f"El archivo no existe: {pdf_path}")
                return
            
            # Abrir con el programa predeterminado del sistema
            if platform.system() == 'Windows':
                os.startfile(pdf_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', pdf_path])
            else:  # Linux
                subprocess.Popen(['xdg-open', pdf_path])
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el PDF: {e}")
    
    def export_zip(self):
        """Exporta el informe en formato ZIP"""
        created = self.new_report(silent=True)
        if not created or not self.current_report:
            return

        self._sync_report_evaluated()

        exports_dir = self.project_root / "data" / "exports"
        exports_dir.mkdir(parents=True, exist_ok=True)

        base_name = self._build_export_base_name()
        target_dir = filedialog.askdirectory(
            title="Selecciona la carpeta para guardar el ZIP",
            initialdir=str(exports_dir),
        )
        if not target_dir:
            return

        target_root = Path(target_dir)
        package_dir = target_root / base_name
        if package_dir.exists():
            shutil.rmtree(package_dir)
        package_dir.mkdir(parents=True, exist_ok=True)

        attachments_copied = self._copy_report_attachments(package_dir)
        credentials_copied, credential_destinations = self._copy_evaluator_credentials(package_dir)

        attachment_links = {
            "calibration": "Adjuntos/Calibracion",
            "audiometria": "Adjuntos/Audiometrias",
            "espirometria": "Adjuntos/Espirometrias",
            "attendance": "Adjuntos/Asistencia",
        }

        self.current_report["link_mode"] = "relative"
        self.current_report["attachment_folder_links"] = attachment_links
        self.current_report["evaluator_credentials_links"] = [
            {
                "name": item["name"],
                "file": str(Path(item["file"]).relative_to(package_dir)),
                "folder": str(Path(item["folder"]).relative_to(package_dir)),
            }
            for item in credential_destinations
        ]

        pdf_path = package_dir / f"{base_name}.pdf"
        if not self._generate_pdf_file(pdf_path):
            shutil.rmtree(package_dir, ignore_errors=True)
            messagebox.showerror("Error", "No se pudo generar el PDF para el paquete ZIP")
            return

        zip_base = target_root / base_name
        zip_path = zip_base.with_suffix(".zip")
        if zip_path.exists():
            zip_path.unlink()

        shutil.make_archive(str(zip_base), "zip", root_dir=target_root, base_dir=base_name)

        summary = [
            f"Paquete ZIP creado en:\n{zip_path}",
            f"✔ PDF: {pdf_path.name}",
        ]
        total_copied = attachments_copied + credentials_copied
        if total_copied:
            summary.append(f"✔ Adjuntos copiados: {total_copied}")
        else:
            summary.append("✔ Sin adjuntos adicionales disponibles")

        result = messagebox.askyesno(
            "ZIP Creado",
            "\n\n".join(summary) + "\n\n¿Deseas abrir la carpeta de exportaciones?",
        )

        if result:
            self.open_folder(str(target_root))

        self.status_label.configure(text=f"✅ ZIP exportado: {zip_path}")
        self._reset_form_state()
    
    def open_folder(self, folder_path: str):
        """Abre la carpeta en el explorador del sistema"""
        try:
            if not os.path.exists(folder_path):
                messagebox.showerror("Error", f"La carpeta no existe: {folder_path}")
                return
            
            # Abrir carpeta con el explorador predeterminado
            if platform.system() == 'Windows':
                os.startfile(folder_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', folder_path])
            else:  # Linux
                subprocess.Popen(['xdg-open', folder_path])
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta: {e}")

    def _attach_tooltip(self, widget: tk.Widget, text: str) -> None:
        """Asocia un tooltip sencillo a un widget."""

        tooltip = {"window": None}

        def _show(_event):
            if tooltip["window"] is not None:
                return
            win = tk.Toplevel(self.root)
            win.wm_overrideredirect(True)
            win.attributes("-topmost", True)
            label = tk.Label(
                win,
                text=text,
                background="#FFFFE0",
                foreground="#000000",
                relief=tk.SOLID,
                borderwidth=1,
                font=("Arial", 9),
            )
            label.pack(ipadx=6, ipady=4)
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + widget.winfo_height() + 8
            win.wm_geometry(f"+{x}+{y}")
            tooltip["window"] = win

        def _hide(_event):
            win = tooltip["window"]
            if win is not None:
                win.destroy()
                tooltip["window"] = None

        widget.bind("<Enter>", _show)
        widget.bind("<Leave>", _hide)
    
    def run(self):
        """Inicia la aplicación"""
        self.root.mainloop()


if __name__ == "__main__":
    # Punto de entrada cuando se ejecuta como script
    app = MainApplication()
    app.run()
