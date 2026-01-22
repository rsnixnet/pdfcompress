import io
import os
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path

import fitz
from PIL import Image
from PySide6 import QtCore, QtGui, QtWidgets


APP_TITLE = "PDF Scan Compressor"


@dataclass
class CompressionSettings:
    dpi: int
    color_mode: str
    jpeg_quality: int
    skip_pages_without_images: bool
    skip_small_images: bool
    output_dir: str | None
    keep_name_in_output_dir: bool


class FileTable(QtWidgets.QTableWidget):
    files_dropped = QtCore.Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.setColumnCount(5)
        self.setHorizontalHeaderLabels([
            "Имя",
            "Размер (до)",
            "Статус",
            "Размер (после)",
            "Экономия %",
        ])
        header = self.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        for column in range(1, 5):
            header.setSectionResizeMode(column, QtWidgets.QHeaderView.ResizeToContents)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if not event.mimeData().hasUrls():
            super().dropEvent(event)
            return
        paths = []
        for url in event.mimeData().urls():
            local = url.toLocalFile()
            if local.lower().endswith(".pdf"):
                paths.append(local)
        if paths:
            self.files_dropped.emit(paths)
        event.acceptProposedAction()


class CompressionWorker(QtCore.QObject):
    progress_overall = QtCore.Signal(int)
    progress_file = QtCore.Signal(int)
    log = QtCore.Signal(str)
    file_started = QtCore.Signal(int)
    file_finished = QtCore.Signal(int, str, int, int, str)
    finished = QtCore.Signal()

    def __init__(self, files, settings: CompressionSettings):
        super().__init__()
        self.files = files
        self.settings = settings
        self._stop_event = threading.Event()

    def stop(self):
        self._stop_event.set()

    def _log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log.emit(f"[{timestamp}] {message}")

    def _determine_output_path(self, input_path: str) -> Path:
        source = Path(input_path)
        if self.settings.output_dir:
            output_dir = Path(self.settings.output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
            if self.settings.keep_name_in_output_dir:
                return output_dir / source.name
            return output_dir / f"{source.stem}_compressed.pdf"
        return source.with_name(f"{source.stem}_compressed.pdf")

    def _colorize(self, image: Image.Image) -> Image.Image:
        if self.settings.color_mode == "Color":
            return image.convert("RGB")
        if self.settings.color_mode == "Grayscale":
            return image.convert("L")
        return image.convert("L").convert("1").convert("L")

    def _downscale(self, image: Image.Image, target_width: int, target_height: int) -> Image.Image:
        width, height = image.size
        if width <= target_width and height <= target_height:
            return image
        return image.resize((target_width, target_height), Image.LANCZOS)

    def _process_page_images(self, doc: fitz.Document, page: fitz.Page):
        images = page.get_images(full=True)
        if not images and self.settings.skip_pages_without_images:
            return
        for image_info in images:
            xref = image_info[0]
            try:
                base = doc.extract_image(xref)
                image_bytes = base.get("image")
                if not image_bytes:
                    continue
                with Image.open(io.BytesIO(image_bytes)) as pil_image:
                    pil_image.load()
                    if self.settings.skip_small_images and pil_image.width < 1000:
                        continue
                    pil_image = self._colorize(pil_image)
                    target_width = max(1, int(page.rect.width / 72 * self.settings.dpi))
                    target_height = max(1, int(page.rect.height / 72 * self.settings.dpi))
                    pil_image = self._downscale(pil_image, target_width, target_height)
                    buffer = io.BytesIO()
                    pil_image.save(
                        buffer,
                        format="JPEG",
                        quality=self.settings.jpeg_quality,
                        optimize=True,
                    )
                    jpeg_bytes = buffer.getvalue()
                try:
                    page.replace_image(xref, stream=jpeg_bytes)
                except Exception:
                    doc.update_stream(xref, jpeg_bytes)
            except Exception as exc:
                self._log(f"Изображение пропущено (xref {xref}): {exc}")
                continue

    @QtCore.Slot()
    def run(self):
        total = len(self.files)
        for index, file_path in enumerate(self.files):
            if self._stop_event.is_set():
                break
            self.file_started.emit(index)
            self._log(f"Открытие: {file_path}")
            input_path = Path(file_path)
            output_path = self._determine_output_path(file_path)
            try:
                before_size = input_path.stat().st_size
                with fitz.open(file_path) as doc:
                    page_count = doc.page_count
                    for page_index in range(page_count):
                        if self._stop_event.is_set():
                            break
                        page = doc.load_page(page_index)
                        self._process_page_images(doc, page)
                        progress = int((page_index + 1) / page_count * 100) if page_count else 0
                        self.progress_file.emit(progress)
                    if self._stop_event.is_set():
                        break
                    doc.save(output_path)
                after_size = output_path.stat().st_size
                savings = 0
                if before_size > 0:
                    savings = round((1 - after_size / before_size) * 100, 1)
                self.file_finished.emit(
                    index,
                    "Готово",
                    before_size,
                    after_size,
                    f"{savings}%",
                )
                self._log(f"Готово: {output_path}")
            except Exception as exc:
                self.file_finished.emit(index, "Error", 0, 0, "-")
                self._log(f"Ошибка: {file_path} — {exc}")
            self.progress_overall.emit(int((index + 1) / total * 100))
        self.finished.emit()


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1100, 700)
        self.worker_thread = None
        self.worker = None
        self.file_paths = []

        self.table = FileTable()
        self.table.files_dropped.connect(self.add_files)

        self.log_output = QtWidgets.QTextEdit()
        self.log_output.setReadOnly(True)

        self.progress_overall = QtWidgets.QProgressBar()
        self.progress_file = QtWidgets.QProgressBar()

        self.add_files_button = QtWidgets.QPushButton("Добавить файлы…")
        self.add_folder_button = QtWidgets.QPushButton("Добавить папку…")
        self.remove_button = QtWidgets.QPushButton("Удалить")
        self.clear_button = QtWidgets.QPushButton("Очистить")
        self.output_folder_button = QtWidgets.QPushButton("Выходная папка…")
        self.start_button = QtWidgets.QPushButton("Старт")
        self.stop_button = QtWidgets.QPushButton("Стоп")
        self.copy_log_button = QtWidgets.QPushButton("Скопировать лог")
        self.exit_button = QtWidgets.QPushButton("Выход")

        self.stop_button.setEnabled(False)

        self.output_folder_edit = QtWidgets.QLineEdit()
        self.output_folder_edit.setPlaceholderText("Папка вывода (опционально)")

        self.keep_name_checkbox = QtWidgets.QCheckBox("Сохранять имя исходного файла")
        self.keep_name_checkbox.setChecked(True)

        self.presets_combo = QtWidgets.QComboBox()
        self.presets_combo.addItems([
            "Max Compression",
            "Balanced",
            "High Quality",
            "Advanced",
        ])

        self.dpi_spin = QtWidgets.QSpinBox()
        self.dpi_spin.setRange(100, 300)
        self.dpi_spin.setValue(200)
        self.color_mode_combo = QtWidgets.QComboBox()
        self.color_mode_combo.addItems(["Color", "Grayscale", "BW"])
        self.jpeg_spin = QtWidgets.QSpinBox()
        self.jpeg_spin.setRange(40, 95)
        self.jpeg_spin.setValue(75)
        self.skip_pages_checkbox = QtWidgets.QCheckBox("Пропускать страницы без картинок")
        self.skip_small_checkbox = QtWidgets.QCheckBox("Не трогать мелкие изображения (<1000px)")

        self.advanced_group = QtWidgets.QGroupBox("Advanced")
        advanced_layout = QtWidgets.QFormLayout(self.advanced_group)
        advanced_layout.addRow("DPI", self.dpi_spin)
        advanced_layout.addRow("Color mode", self.color_mode_combo)
        advanced_layout.addRow("JPEG quality", self.jpeg_spin)
        advanced_layout.addRow("", self.skip_pages_checkbox)
        advanced_layout.addRow("", self.skip_small_checkbox)

        self._build_layout()
        self._connect_signals()
        self._apply_preset()

    def _build_layout(self):
        button_row = QtWidgets.QHBoxLayout()
        for btn in (
            self.add_files_button,
            self.add_folder_button,
            self.remove_button,
            self.clear_button,
            self.output_folder_button,
            self.start_button,
            self.stop_button,
            self.exit_button,
        ):
            button_row.addWidget(btn)

        output_row = QtWidgets.QHBoxLayout()
        output_row.addWidget(self.output_folder_edit)
        output_row.addWidget(self.keep_name_checkbox)

        preset_row = QtWidgets.QHBoxLayout()
        preset_row.addWidget(QtWidgets.QLabel("Preset:"))
        preset_row.addWidget(self.presets_combo)
        preset_row.addStretch()

        progress_layout = QtWidgets.QVBoxLayout()
        progress_layout.addWidget(QtWidgets.QLabel("Общий прогресс"))
        progress_layout.addWidget(self.progress_overall)
        progress_layout.addWidget(QtWidgets.QLabel("Текущий файл"))
        progress_layout.addWidget(self.progress_file)

        log_layout = QtWidgets.QVBoxLayout()
        log_layout.addWidget(QtWidgets.QLabel("Лог"))
        log_layout.addWidget(self.log_output)
        log_layout.addWidget(self.copy_log_button)

        main_layout = QtWidgets.QVBoxLayout()
        main_layout.addLayout(button_row)
        main_layout.addWidget(self.table)
        main_layout.addLayout(output_row)
        main_layout.addLayout(preset_row)
        main_layout.addWidget(self.advanced_group)
        main_layout.addLayout(progress_layout)
        main_layout.addLayout(log_layout)

        container = QtWidgets.QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def _connect_signals(self):
        self.add_files_button.clicked.connect(self.select_files)
        self.add_folder_button.clicked.connect(self.select_folder)
        self.remove_button.clicked.connect(self.remove_selected)
        self.clear_button.clicked.connect(self.clear_files)
        self.output_folder_button.clicked.connect(self.select_output_folder)
        self.start_button.clicked.connect(self.start_processing)
        self.stop_button.clicked.connect(self.stop_processing)
        self.copy_log_button.clicked.connect(self.copy_log)
        self.exit_button.clicked.connect(self.close)
        self.presets_combo.currentIndexChanged.connect(self._apply_preset)

    def _apply_preset(self):
        preset = self.presets_combo.currentText()
        if preset == "Max Compression":
            self.dpi_spin.setValue(150)
            self.color_mode_combo.setCurrentText("Grayscale")
            self.jpeg_spin.setValue(65)
            self.advanced_group.setEnabled(False)
        elif preset == "Balanced":
            self.dpi_spin.setValue(200)
            self.color_mode_combo.setCurrentText("Grayscale")
            self.jpeg_spin.setValue(75)
            self.advanced_group.setEnabled(False)
        elif preset == "High Quality":
            self.dpi_spin.setValue(300)
            self.color_mode_combo.setCurrentText("Color")
            self.jpeg_spin.setValue(85)
            self.advanced_group.setEnabled(False)
        else:
            self.advanced_group.setEnabled(True)

    def log(self, message):
        self.log_output.append(message)

    def copy_log(self):
        QtGui.QGuiApplication.clipboard().setText(self.log_output.toPlainText())

    def add_files(self, paths):
        for path in paths:
            if path in self.file_paths:
                continue
            if not path.lower().endswith(".pdf"):
                continue
            self.file_paths.append(path)
            row = self.table.rowCount()
            self.table.insertRow(row)
            name_item = QtWidgets.QTableWidgetItem(Path(path).name)
            before_size = self._format_size(Path(path).stat().st_size)
            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(before_size))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem("Ожидание"))
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem("-"))
            self.table.setItem(row, 4, QtWidgets.QTableWidgetItem("-"))

    def select_files(self):
        paths, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,
            "Выбрать PDF",
            "",
            "PDF Files (*.pdf)",
        )
        if paths:
            self.add_files(paths)

    def select_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку")
        if not folder:
            return
        pdfs = [str(path) for path in Path(folder).glob("*.pdf")]
        if pdfs:
            self.add_files(pdfs)

    def select_output_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Выходная папка")
        if folder:
            self.output_folder_edit.setText(folder)

    def remove_selected(self):
        rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.table.removeRow(row)
            self.file_paths.pop(row)

    def clear_files(self):
        self.table.setRowCount(0)
        self.file_paths = []

    def _format_size(self, size):
        for unit in ["B", "KB", "MB", "GB"]:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"

    def _collect_settings(self) -> CompressionSettings:
        output_dir = self.output_folder_edit.text().strip() or None
        return CompressionSettings(
            dpi=self.dpi_spin.value(),
            color_mode=self.color_mode_combo.currentText(),
            jpeg_quality=self.jpeg_spin.value(),
            skip_pages_without_images=self.skip_pages_checkbox.isChecked(),
            skip_small_images=self.skip_small_checkbox.isChecked(),
            output_dir=output_dir,
            keep_name_in_output_dir=self.keep_name_checkbox.isChecked(),
        )

    def start_processing(self):
        if not self.file_paths:
            self.log("Нет файлов для обработки.")
            return
        if self.worker_thread and self.worker_thread.isRunning():
            return
        self.progress_overall.setValue(0)
        self.progress_file.setValue(0)
        settings = self._collect_settings()
        self.worker = CompressionWorker(self.file_paths, settings)
        self.worker_thread = QtCore.QThread()
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress_overall.connect(self.progress_overall.setValue)
        self.worker.progress_file.connect(self.progress_file.setValue)
        self.worker.log.connect(self.log)
        self.worker.file_started.connect(self._mark_file_started)
        self.worker.file_finished.connect(self._mark_file_finished)
        self.worker.finished.connect(self._worker_finished)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.worker_thread.start()

    def stop_processing(self):
        if self.worker:
            self.worker.stop()
            self.log("Остановка по запросу...")

    def _mark_file_started(self, row):
        if row < self.table.rowCount():
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem("В работе"))

    def _mark_file_finished(self, row, status, before_size, after_size, savings):
        if row >= self.table.rowCount():
            return
        if before_size:
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(self._format_size(before_size)))
        if after_size:
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(self._format_size(after_size)))
        self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(status))
        self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(savings))

    def _worker_finished(self):
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.progress_file.setValue(0)
        self.log("Готово.")


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
