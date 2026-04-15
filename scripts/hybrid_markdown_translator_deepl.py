import os
import re
import sys
import json
import urllib.request
import urllib.error
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import (
        QApplication, QWidget, QMainWindow, QFileDialog, QListWidget, QListWidgetItem,
        QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTextEdit, QSplitter,
        QTableWidget, QTableWidgetItem, QMessageBox, QLineEdit, QComboBox,
        QHeaderView, QInputDialog
    )
except Exception:
    print("PySide6 is required. Install with: python3 -m pip install PySide6")
    raise


@dataclass
class Segment:
    seg_id: str
    kind: str
    source: str
    target: str = ""
    status: str = "new"
    warnings: str = ""


class DeepLClient:
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.getenv("DEEPL_API_KEY", "")
        self.base_url = self._detect_base_url()

    def _detect_base_url(self) -> str:
        if self.api_key.endswith(":fx"):
            return "https://api-free.deepl.com/v2/translate"
        return "https://api.deepl.com/v2/translate"

    def set_api_key(self, api_key: str):
        self.api_key = api_key.strip()
        self.base_url = self._detect_base_url()

    def is_configured(self) -> bool:
        return bool(self.api_key.strip())

    def translate_text(
        self,
        text: str,
        source_lang: str = "IT",
        target_lang: str = "EN-GB",
        context: str = ""
    ) -> str:
        if not self.is_configured():
            raise RuntimeError("DeepL API key is not configured.")

        payload = {
            "text": [text],
            "source_lang": source_lang,
            "target_lang": target_lang,
            "preserve_formatting": True,
            "split_sentences": "1",
            "formality": "prefer_more"
        }

        if context.strip():
            payload["context"] = context[:4000]

        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(self.base_url, data=data, method="POST")
        req.add_header("Authorization", f"DeepL-Auth-Key {self.api_key}")
        req.add_header("Content-Type", "application/json")
        req.add_header("User-Agent", "HybridMarkdownTranslator/2.0")

        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                result = json.loads(resp.read().decode("utf-8"))
            translations = result.get("translations", [])
            if not translations:
                raise RuntimeError("DeepL returned no translation.")
            return translations[0].get("text", "")
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"DeepL HTTP error {e.code}: {body}")
        except urllib.error.URLError as e:
            raise RuntimeError(f"DeepL connection error: {e}")


class MarkdownHybridEngine:
    BRITISH_MAP = {
        "color": "colour",
        "colors": "colours",
        "center": "centre",
        "centers": "centres",
        "analyze": "analyse",
        "analyzed": "analysed",
        "catalog": "catalogue",
        "organization": "organisation",
        "organize": "organise",
        "favor": "favour",
        "honor": "honour",
    }

    TERM_MAP = {
        "codice miniato": "illuminated manuscript",
        "miniatura": "illumination",
        "ciclo pittorico": "pictorial cycle",
        "ciclo di affreschi": "fresco cycle",
        "committente": "patron",
        "tavola": "panel painting",
        "attribuito a": "attributed to",
        "ambito di": "in the circle of",
        "bottega di": "workshop of",
        "iniziale istoriata": "historiated initial",
    }

    def __init__(self, deepl_client: Optional[DeepLClient] = None):
        self.deepl = deepl_client or DeepLClient()

    def parse_markdown(self, text: str) -> List[Segment]:
        segments = []
        lines = text.splitlines(keepends=True)
        in_fence = False
        buffer = []
        idx = 1

        def push(kind: str, content: str):
            nonlocal idx
            segments.append(Segment(seg_id=f"S{idx:05d}", kind=kind, source=content))
            idx += 1

        for line in lines:
            if line.strip().startswith("```"):
                buffer.append(line)
                in_fence = not in_fence
                if not in_fence:
                    push("fence", "".join(buffer))
                    buffer = []
                continue

            if in_fence:
                buffer.append(line)
                continue

            if line.startswith("#"):
                push("heading", line)
            elif re.match(r"^\s*[-*+]\s+", line) or re.match(r"^\s*\d+\.\s+", line):
                push("list", line)
            elif line.strip().startswith(">"):
                push("blockquote", line)
            elif line.strip() == "":
                push("blank", line)
            else:
                push("paragraph", line)

        return segments

    def protect(self, text: str) -> Tuple[str, Dict[str, str]]:
        patterns = [
            re.compile(r"`[^`]+`"),
            re.compile(r"https?://\S+"),
            re.compile(r"\[[^\]]+\]\([^\)]+\)"),
        ]
        mapping = {}
        counter = 1
        out = text

        for pattern in patterns:
            matches = list(pattern.finditer(out))
            for m in matches:
                val = m.group(0)
                if val in mapping.values():
                    continue
                token = f"{{{{PH_{counter:04d}}}}}"
                mapping[token] = val
                out = out.replace(val, token, 1)
                counter += 1

        return out, mapping

    def restore(self, text: str, mapping: Dict[str, str]) -> str:
        out = text
        for token, value in mapping.items():
            out = out.replace(token, value)
        return out

    def apply_glossary(self, text: str) -> str:
        out = text
        for it, en in sorted(self.TERM_MAP.items(), key=lambda x: len(x[0]), reverse=True):
            out = re.sub(rf"\b{re.escape(it)}\b", en, out, flags=re.IGNORECASE)
        return out

    def normalise_british(self, text: str) -> str:
        out = text
        for us, uk in self.BRITISH_MAP.items():
            out = re.sub(rf"\b{re.escape(us)}\b", uk, out, flags=re.IGNORECASE)
        return out

    def build_context(self, segment: Segment) -> str:
        return (
            "Translate Italian Markdown into formal British English for art-historical and scholarly prose. "
            "Preserve placeholders exactly. Preserve proper names, URLs, identifiers, shelfmarks and code. "
            "Use British spelling and a refined academic register."
        )

    def translate_segment(self, segment: Segment) -> Segment:
        if segment.kind in {"fence", "blank"}:
            segment.target = segment.source
            segment.status = "protected"
            return segment

        protected, mapping = self.protect(segment.source)
        prepared = self.apply_glossary(protected)
        translated = self.deepl.translate_text(
            prepared,
            source_lang="IT",
            target_lang="EN-GB",
            context=self.build_context(segment)
        )
        translated = self.normalise_british(translated)
        translated = self.restore(translated, mapping)

        segment.target = translated
        segment.status = "translated"

        if re.search(r"\b(il|lo|la|gli|della|delle|degli|nel|nella|con|per)\b", translated, re.IGNORECASE):
            segment.warnings = "Possible Italian residue"

        return segment

    def rebuild(self, segments: List[Segment]) -> str:
        return "".join(seg.target if seg.target else seg.source for seg in segments)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Hybrid Markdown Translator — DeepL en-GB")
        self.resize(1500, 920)

        self.deepl = DeepLClient()
        self.engine = MarkdownHybridEngine(self.deepl)

        self.project_dir = ""
        self.output_dir = ""
        self.files: List[str] = []
        self.current_segments: List[Segment] = []
        self.current_file = None

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        topbar = QHBoxLayout()
        self.btn_open = QPushButton("Open Folder")
        self.btn_set_key = QPushButton("Set DeepL Key")
        self.btn_translate_file = QPushButton("Translate File")
        self.btn_translate_all = QPushButton("Batch Translate")
        self.btn_export = QPushButton("Export Current")
        self.btn_export_all = QPushButton("Export All")
        self.filter_status = QComboBox()
        self.filter_status.addItems(["all", "new", "translated", "edited", "approved", "flagged", "protected"])
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search files...")

        for w in [
            self.btn_open, self.btn_set_key, self.btn_translate_file,
            self.btn_translate_all, self.btn_export, self.btn_export_all,
            QLabel("Filter:"), self.filter_status, self.search_box
        ]:
            topbar.addWidget(w)

        root.addLayout(topbar)

        splitter = QSplitter(Qt.Horizontal)
        root.addWidget(splitter)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.addWidget(QLabel("Files"))
        self.file_list = QListWidget()
        left_layout.addWidget(self.file_list)
        splitter.addWidget(left)

        center = QWidget()
        center_layout = QVBoxLayout(center)
        center_layout.addWidget(QLabel("Source / Target Segment"))
        self.segment_table = QTableWidget(0, 4)
        self.segment_table.setHorizontalHeaderLabels(["ID", "Type", "Status", "Warnings"])
        self.segment_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        center_layout.addWidget(self.segment_table)

        editors = QSplitter(Qt.Horizontal)
        self.source_editor = QTextEdit()
        self.target_editor = QTextEdit()
        self.source_editor.setPlaceholderText("Italian source segment")
        self.target_editor.setPlaceholderText("British English target segment")
        editors.addWidget(self.source_editor)
        editors.addWidget(self.target_editor)
        center_layout.addWidget(editors)
        splitter.addWidget(center)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.addWidget(QLabel("Glossary / Review"))

        self.glossary_view = QTextEdit()
        self.glossary_view.setReadOnly(True)
        self.glossary_view.setPlainText(self.format_glossary())

        self.status_label = QLabel("No file loaded")
        self.api_label = QLabel("DeepL key: not configured")

        self.btn_save_segment = QPushButton("Save Segment Edit")
        self.btn_approve_segment = QPushButton("Approve Segment")
        self.btn_flag_segment = QPushButton("Flag Segment")

        right_layout.addWidget(self.glossary_view)
        right_layout.addWidget(self.api_label)
        right_layout.addWidget(self.status_label)
        right_layout.addWidget(self.btn_save_segment)
        right_layout.addWidget(self.btn_approve_segment)
        right_layout.addWidget(self.btn_flag_segment)
        right_layout.addStretch(1)

        splitter.addWidget(right)
        splitter.setSizes([260, 930, 310])

        self.btn_open.clicked.connect(self.open_folder)
        self.btn_set_key.clicked.connect(self.set_api_key)
        self.file_list.currentItemChanged.connect(self.load_selected_file)
        self.segment_table.itemSelectionChanged.connect(self.load_selected_segment)
        self.btn_translate_file.clicked.connect(self.translate_current_file)
        self.btn_translate_all.clicked.connect(self.translate_all_files)
        self.btn_save_segment.clicked.connect(self.save_segment_edit)
        self.btn_approve_segment.clicked.connect(self.approve_segment)
        self.btn_flag_segment.clicked.connect(self.flag_segment)
        self.btn_export.clicked.connect(self.export_current)
        self.btn_export_all.clicked.connect(self.export_all)
        self.search_box.textChanged.connect(self.refresh_file_list)

        self.update_api_label()

    def format_glossary(self) -> str:
        return "\n".join(f"{k} -> {v}" for k, v in self.engine.TERM_MAP.items())

    def update_api_label(self):
        if self.deepl.is_configured():
            key = self.deepl.api_key
            masked = (key[:4] + "***" + key[-4:]) if len(key) >= 8 else "configured"
            self.api_label.setText(f"DeepL key: {masked}")
        else:
            self.api_label.setText("DeepL key: not configured")

    def set_api_key(self):
        value, ok = QInputDialog.getText(self, "DeepL API Key", "Paste your DeepL API key:")
        if ok and value.strip():
            self.deepl.set_api_key(value.strip())
            self.update_api_label()
            QMessageBox.information(self, "DeepL", "API key stored for this session.")

    def open_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Markdown Folder")
        if not folder:
            return
        self.project_dir = folder
        self.output_dir = os.path.join(folder, "translated_en_GB")
        os.makedirs(self.output_dir, exist_ok=True)
        self.files = [f for f in os.listdir(folder) if f.lower().endswith(".md")]
        self.files.sort()
        self.refresh_file_list()
        self.status_label.setText(f"Loaded folder: {folder} | {len(self.files)} markdown files")

    def refresh_file_list(self):
        query = self.search_box.text().strip().lower()
        self.file_list.clear()
        for f in self.files:
            if query and query not in f.lower():
                continue
            self.file_list.addItem(QListWidgetItem(f))

    def load_selected_file(self):
        item = self.file_list.currentItem()
        if not item or not self.project_dir:
            return

        path = os.path.join(self.project_dir, item.text())
        self.current_file = path

        with open(path, "r", encoding="utf-8") as fh:
            text = fh.read()

        self.current_segments = self.engine.parse_markdown(text)
        self.populate_segments()
        self.status_label.setText(f"Loaded file: {item.text()} | {len(self.current_segments)} segments")

    def populate_segments(self):
        self.segment_table.setRowCount(len(self.current_segments))
        for row, seg in enumerate(self.current_segments):
            self.segment_table.setItem(row, 0, QTableWidgetItem(seg.seg_id))
            self.segment_table.setItem(row, 1, QTableWidgetItem(seg.kind))
            self.segment_table.setItem(row, 2, QTableWidgetItem(seg.status))
            self.segment_table.setItem(row, 3, QTableWidgetItem(seg.warnings))
        self.source_editor.clear()
        self.target_editor.clear()

    def load_selected_segment(self):
        row = self.segment_table.currentRow()
        if row < 0 or row >= len(self.current_segments):
            return

        seg = self.current_segments[row]
        self.source_editor.setPlainText(seg.source)
        self.target_editor.setPlainText(seg.target)
        self.status_label.setText(f"Segment {seg.seg_id} | {seg.kind} | {seg.status}")

    def translate_current_file(self):
        if not self.current_segments:
            return
        if not self.deepl.is_configured():
            QMessageBox.warning(self, "DeepL", "Please set your DeepL API key first.")
            return

        try:
            self.current_segments = [self.engine.translate_segment(seg) for seg in self.current_segments]
            self.populate_segments()
            self.status_label.setText("Current file translated with DeepL")
        except Exception as e:
            QMessageBox.critical(self, "Translation error", str(e))

    def translate_all_files(self):
        if not self.project_dir:
            return
        if not self.deepl.is_configured():
            QMessageBox.warning(self, "DeepL", "Please set your DeepL API key first.")
            return

        count = 0
        try:
            for fname in self.files:
                path = os.path.join(self.project_dir, fname)
                with open(path, "r", encoding="utf-8") as fh:
                    text = fh.read()

                segments = self.engine.parse_markdown(text)
                segments = [self.engine.translate_segment(seg) for seg in segments]
                output = self.engine.rebuild(segments)

                outpath = os.path.join(self.output_dir, fname.replace(".md", "_en-GB.md"))
                with open(outpath, "w", encoding="utf-8") as fh:
                    fh.write(output)

                count += 1

            QMessageBox.information(self, "Done", f"Batch translated {count} files into {self.output_dir}")
        except Exception as e:
            QMessageBox.critical(self, "Batch translation error", str(e))

    def save_segment_edit(self):
        row = self.segment_table.currentRow()
        if row < 0:
            return

        self.current_segments[row].target = self.target_editor.toPlainText()
        self.current_segments[row].status = "edited"
        self.populate_segments()
        self.segment_table.selectRow(row)

    def approve_segment(self):
        row = self.segment_table.currentRow()
        if row < 0:
            return

        self.current_segments[row].target = self.target_editor.toPlainText()
        self.current_segments[row].status = "approved"
        self.populate_segments()
        self.segment_table.selectRow(row)

    def flag_segment(self):
        row = self.segment_table.currentRow()
        if row < 0:
            return

        self.current_segments[row].target = self.target_editor.toPlainText()
        self.current_segments[row].status = "flagged"
        self.current_segments[row].warnings = "Marked for review"
        self.populate_segments()
        self.segment_table.selectRow(row)

    def export_current(self):
        if not self.current_file or not self.current_segments:
            return

        os.makedirs(self.output_dir, exist_ok=True)
        output = self.engine.rebuild(self.current_segments)
        fname = os.path.basename(self.current_file).replace(".md", "_en-GB.md")
        outpath = os.path.join(self.output_dir, fname)

        with open(outpath, "w", encoding="utf-8") as fh:
            fh.write(output)

        QMessageBox.information(self, "Exported", f"Saved: {outpath}")

    def export_all(self):
        self.translate_all_files()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())