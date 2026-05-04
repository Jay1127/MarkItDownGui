import codecs
import mimetypes
import os
import shutil
import subprocess
import threading
import tkinter as tk
from importlib.metadata import entry_points
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from markitdown import MarkItDown, StreamInfo
except ImportError:
    MarkItDown = None
    StreamInfo = None


class MarkItDownGui(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("MarkItDown GUI")
        self.geometry("1100x860")
        self.minsize(900, 640)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.extension_hint = tk.StringVar()
        self.mime_type_hint = tk.StringVar()
        self.charset_hint = tk.StringVar()
        self.use_plugins = tk.BooleanVar(value=False)
        self.keep_data_uris = tk.BooleanVar(value=False)
        self.use_llm = tk.BooleanVar(value=False)
        self.llm_model = tk.StringVar(value="gpt-4o")
        self.llm_prompt = tk.StringVar()
        self.exiftool_path = tk.StringVar()
        self.style_map = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready")
        self.plugin_vars = {}
        self.extension_values = [
            "",
            ".pdf",
            ".docx",
            ".pptx",
            ".xlsx",
            ".html",
            ".csv",
            ".json",
            ".xml",
            ".zip",
            ".epub",
            ".jpg",
            ".png",
            ".mp3",
            ".wav",
        ]
        self.mime_type_values = [
            "",
            "application/pdf",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "text/html",
            "text/csv",
            "application/json",
            "application/xml",
            "application/zip",
            "application/epub+zip",
            "image/jpeg",
            "image/png",
            "audio/mpeg",
            "audio/wav",
        ]
        self.charset_values = ["", "utf-8", "cp949", "euc-kr", "utf-16", "latin-1"]

        self._build_ui()
        self._update_dependency_status()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=12)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Input file / URL").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(top, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(top, text="Browse", command=self.choose_input).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(top, text="Output .md").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(top, textvariable=self.output_path).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(top, text="Browse", command=self.choose_output).grid(row=1, column=2, padx=(8, 0), pady=4)

        options = ttk.Frame(top)
        options.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        options.columnconfigure(0, weight=1)
        options.columnconfigure(1, weight=1)
        options.columnconfigure(2, weight=1)

        self._build_basic_options(options).grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self._build_hint_options(options).grid(row=0, column=1, sticky="nsew", padx=6)
        self._build_plugin_options(options).grid(row=0, column=2, sticky="nsew", padx=(6, 0))
        self._build_llm_options(options).grid(row=1, column=0, columnspan=2, sticky="nsew", padx=(0, 6), pady=(8, 0))
        self._build_metadata_docx_options(options).grid(
            row=1, column=2, sticky="nsew", padx=(6, 0), pady=(8, 0)
        )

        actions = ttk.Frame(top)
        actions.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        actions.columnconfigure(4, weight=1)
        ttk.Button(actions, text="Refresh plugins", command=self.refresh_plugins).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Convert", command=self.start_convert).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="Open output", command=self.open_output).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(actions, text="Clear preview", command=self.clear_preview).grid(row=0, column=3, padx=(0, 8))

        body = ttk.PanedWindow(self, orient=tk.VERTICAL)
        body.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        preview_frame = ttk.LabelFrame(body, text="Markdown preview", padding=8)
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)

        self.preview = tk.Text(preview_frame, wrap="word", undo=False)
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=preview_scroll.set)
        self.preview.grid(row=0, column=0, sticky="nsew")
        preview_scroll.grid(row=0, column=1, sticky="ns")

        log_frame = ttk.LabelFrame(body, text="Log", padding=8)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log = tk.Text(log_frame, height=8, wrap="word", state="disabled")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        log_scroll.grid(row=0, column=1, sticky="ns")

        body.add(preview_frame, weight=4)
        body.add(log_frame, weight=1)

        status = ttk.Label(self, textvariable=self.status_text, anchor="w", padding=(12, 0, 12, 8))
        status.grid(row=2, column=0, sticky="ew")

    def _build_basic_options(self, parent):
        frame = ttk.LabelFrame(parent, text="Basic options", padding=8)
        frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(frame, text="Keep data URIs", variable=self.keep_data_uris).grid(row=0, column=0, sticky="w")
        return frame

    def _build_plugin_options(self, parent):
        frame = ttk.LabelFrame(parent, text="Plugin options", padding=8)
        frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(frame, text="Use plugins", variable=self.use_plugins).grid(row=0, column=0, sticky="w")
        self.plugin_frame = ttk.Frame(frame)
        self.plugin_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.plugin_frame.columnconfigure(0, weight=1)
        return frame

    def _build_hint_options(self, parent):
        frame = ttk.LabelFrame(parent, text="Type hints", padding=8)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Extension").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Combobox(frame, textvariable=self.extension_hint, values=self.extension_values).grid(
            row=0, column=1, sticky="ew", pady=3
        )

        ttk.Label(frame, text="MIME type").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Combobox(frame, textvariable=self.mime_type_hint, values=self.mime_type_values).grid(
            row=1, column=1, sticky="ew", pady=3
        )

        ttk.Label(frame, text="Charset").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Combobox(frame, textvariable=self.charset_hint, values=self.charset_values).grid(
            row=2, column=1, sticky="ew", pady=3
        )

        ttk.Label(frame, text="Examples: .pdf, application/pdf, utf-8").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )
        return frame

    def _build_llm_options(self, parent):
        frame = ttk.LabelFrame(parent, text="LLM image options", padding=8)
        frame.columnconfigure(1, weight=1)

        ttk.Checkbutton(frame, text="Use LLM for image descriptions", variable=self.use_llm).grid(
            row=0, column=0, columnspan=2, sticky="w"
        )

        ttk.Label(frame, text="LLM model").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(frame, textvariable=self.llm_model).grid(row=1, column=1, sticky="ew", pady=3)

        ttk.Label(frame, text="LLM prompt").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(frame, textvariable=self.llm_prompt).grid(row=2, column=1, sticky="ew", pady=3)
        return frame

    def _build_metadata_docx_options(self, parent):
        frame = ttk.LabelFrame(parent, text="Metadata / DOCX options", padding=8)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="exiftool path").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        exif_row = ttk.Frame(frame)
        exif_row.grid(row=0, column=1, sticky="ew", pady=3)
        exif_row.columnconfigure(0, weight=1)
        ttk.Entry(exif_row, textvariable=self.exiftool_path).grid(row=0, column=0, sticky="ew")
        ttk.Button(exif_row, text="...", width=3, command=self.choose_exiftool).grid(row=0, column=1, padx=(4, 0))

        ttk.Label(frame, text="DOCX style map").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(frame, textvariable=self.style_map).grid(row=1, column=1, sticky="ew", pady=3)
        return frame

    def _update_dependency_status(self):
        if MarkItDown is None:
            self.status_text.set("MarkItDown is not installed. Run: pip install -r requirements.txt")
            self.append_log("Missing dependency: markitdown")
        else:
            self.append_log("MarkItDown loaded")
            self.refresh_plugins()

    def choose_input(self):
        path = filedialog.askopenfilename(
            title="Choose input file",
            filetypes=[
                ("Common documents", "*.pdf *.docx *.pptx *.xlsx *.xls *.html *.htm *.csv *.json *.xml *.zip *.epub"),
                ("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff *.webp"),
                ("Audio", "*.wav *.mp3 *.m4a"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        self.input_path.set(path)
        current_output = self.output_path.get().strip()
        if not current_output:
            self.output_path.set(str(Path(path).with_suffix(".md")))

    def choose_output(self):
        path = filedialog.asksaveasfilename(
            title="Choose output Markdown file",
            defaultextension=".md",
            filetypes=[("Markdown", "*.md"), ("Text", "*.txt"), ("All files", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def choose_exiftool(self):
        path = filedialog.askopenfilename(
            title="Choose exiftool executable",
            filetypes=[("Executables", "*.exe *.cmd *.bat"), ("All files", "*.*")],
        )
        if path:
            self.exiftool_path.set(path)

    def start_convert(self):
        if MarkItDown is None:
            messagebox.showerror("Missing dependency", "Install dependencies first: pip install -r requirements.txt")
            return

        input_path = self.input_path.get().strip()
        output_path = self.output_path.get().strip()

        if not input_path:
            messagebox.showwarning("Input required", "Choose an input file or enter a URL first.")
            return
        if not self._is_uri(input_path) and not os.path.exists(input_path):
            messagebox.showerror("Input not found", input_path)
            return
        if not output_path:
            messagebox.showwarning("Output required", "Choose an output Markdown path.")
            return

        try:
            options = self._collect_options()
        except ValueError as exc:
            messagebox.showerror("Invalid option", str(exc))
            return

        self.status_text.set("Converting...")
        self.append_log(f"Converting: {input_path}")
        self.append_log("Options: " + self._describe_options(options))
        threading.Thread(
            target=self.convert_file,
            args=(input_path, output_path, options),
            daemon=True,
        ).start()

    def _collect_options(self):
        extension = self.extension_hint.get().strip()
        if extension:
            if not extension.startswith("."):
                extension = "." + extension
            extension = extension.lower()
        else:
            extension = None

        mime_type = self.mime_type_hint.get().strip()
        if mime_type:
            if mime_type.count("/") != 1:
                raise ValueError(f"Invalid MIME type: {mime_type}")
        else:
            mime_type = None

        charset = self.charset_hint.get().strip()
        if charset:
            try:
                charset = codecs.lookup(charset).name
            except LookupError as exc:
                raise ValueError(f"Invalid charset: {charset}") from exc
        else:
            charset = None

        stream_info = None
        if extension or mime_type or charset:
            stream_info = StreamInfo(extension=extension, mimetype=mime_type, charset=charset)

        llm_client = None
        llm_model = None
        llm_prompt = None
        if self.use_llm.get():
            try:
                from openai import OpenAI
            except ImportError as exc:
                raise ValueError("Install openai first to use LLM image descriptions: pip install openai") from exc

            llm_client = OpenAI()
            llm_model = self.llm_model.get().strip() or "gpt-4o"
            llm_prompt = self.llm_prompt.get().strip() or None

        return {
            "stream_info": stream_info,
            "enable_plugins": self.use_plugins.get(),
            "selected_plugins": self._selected_plugin_names(),
            "keep_data_uris": self.keep_data_uris.get(),
            "llm_client": llm_client,
            "llm_model": llm_model,
            "llm_prompt": llm_prompt,
            "exiftool_path": self.exiftool_path.get().strip() or None,
            "style_map": self.style_map.get().strip() or None,
        }

    def convert_file(self, input_path, output_path, options):
        try:
            init_kwargs = {
                "enable_plugins": False,
            }
            for key in ["llm_client", "llm_model", "llm_prompt", "exiftool_path", "style_map"]:
                if options[key] is not None:
                    init_kwargs[key] = options[key]

            md = MarkItDown(**init_kwargs)
            if options["enable_plugins"]:
                self._enable_selected_plugins(md, options["selected_plugins"], init_kwargs)
            result = md.convert(
                input_path,
                stream_info=options["stream_info"],
                keep_data_uris=options["keep_data_uris"],
            )
            text = result.text_content or ""

            output = Path(output_path)
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(text, encoding="utf-8")

            self.after(0, self.show_result, text, str(output))
        except Exception as exc:
            self.after(0, self.show_error, exc)

    def show_result(self, text, output_path):
        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, text)
        self.status_text.set(f"Saved: {output_path}")
        self.append_log(f"Saved: {output_path}")

    def show_error(self, exc):
        self.status_text.set("Conversion failed")
        self.append_log(f"ERROR: {exc}")
        messagebox.showerror("Conversion failed", str(exc))

    def refresh_plugins(self):
        if not hasattr(self, "plugin_frame"):
            return

        for child in self.plugin_frame.winfo_children():
            child.destroy()
        self.plugin_vars.clear()

        plugins = self._installed_plugins()
        if not plugins:
            ttk.Label(self.plugin_frame, text="No plugins installed").grid(row=0, column=0, sticky="w", pady=(6, 0))
            self.append_log("No MarkItDown plugins installed")
            return

        for row, plugin in enumerate(plugins):
            var = tk.BooleanVar(value=True)
            self.plugin_vars[plugin["name"]] = var
            ttk.Checkbutton(
                self.plugin_frame,
                text=f"{plugin['name']} ({plugin['value']})",
                variable=var,
            ).grid(row=row, column=0, sticky="w", pady=(4 if row == 0 else 2, 0))

        self.append_log("Installed plugins: " + ", ".join(plugin["name"] for plugin in plugins))

    def list_plugins(self):
        try:
            command = [shutil.which("markitdown") or "markitdown", "--list-plugins"]
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
            )
            output = completed.stdout.strip() or completed.stderr.strip() or "No plugin output."
            self.append_log(output)
            messagebox.showinfo("MarkItDown plugins", output)
        except Exception as exc:
            self.append_log(f"ERROR: {exc}")
            messagebox.showerror("Plugin list failed", str(exc))

    def open_output(self):
        output_path = self.output_path.get().strip()
        if not output_path or not os.path.exists(output_path):
            messagebox.showwarning("Output not found", "No output file exists yet.")
            return

        os.startfile(output_path)

    def clear_preview(self):
        self.preview.delete("1.0", tk.END)

    def append_log(self, message):
        self.log.configure(state="normal")
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def _describe_options(self, options):
        labels = []
        if options["enable_plugins"]:
            selected = options["selected_plugins"]
            labels.append("plugins=" + (",".join(selected) if selected else "none-selected"))
        if options["keep_data_uris"]:
            labels.append("keep_data_uris")
        if options["llm_client"] is not None:
            labels.append(f"llm={options['llm_model']}")
        if options["exiftool_path"]:
            labels.append("exiftool_path")
        if options["style_map"]:
            labels.append("style_map")
        if options["stream_info"] is not None:
            labels.append("stream_info")
        return ", ".join(labels) if labels else "default"

    def _installed_plugins(self):
        plugins = []
        for entry_point in entry_points(group="markitdown.plugin"):
            plugins.append({"name": entry_point.name, "value": entry_point.value, "entry_point": entry_point})
        plugins.sort(key=lambda p: p["name"])
        return plugins

    def _selected_plugin_names(self):
        return [name for name, var in self.plugin_vars.items() if var.get()]

    def _enable_selected_plugins(self, markitdown, selected_plugin_names, init_kwargs):
        selected = set(selected_plugin_names)
        installed = self._installed_plugins()

        for plugin_info in installed:
            if plugin_info["name"] not in selected:
                continue
            plugin = plugin_info["entry_point"].load()
            plugin.register_converters(markitdown, **init_kwargs)

    def _is_uri(self, value):
        lowered = value.lower()
        if lowered.startswith(("http://", "https://", "file:", "data:")):
            return True
        guessed_type, _ = mimetypes.guess_type(value)
        return guessed_type is not None and "\n" not in value and "\r" not in value and "://" in value


if __name__ == "__main__":
    app = MarkItDownGui()
    app.mainloop()
