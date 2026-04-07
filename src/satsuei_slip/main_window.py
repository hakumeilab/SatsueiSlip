from __future__ import annotations

import os
import re
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, as_completed, wait
from datetime import date
from pathlib import Path

from PySide6.QtCore import QDate, QThread, Qt, QUrl, Signal
from PySide6.QtGui import QAction, QDesktopServices, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QProgressDialog,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from satsuei_slip.image_exporter import DeliverySlipImageExporter
from satsuei_slip.models import DeliveryInfo, VideoItem
from satsuei_slip.pdf_exporter import DeliverySlipPdfExporter
from satsuei_slip.settings_store import SettingsStore
from satsuei_slip.updater import UpdateCheckError, check_github_update
from satsuei_slip.video_probe import (
    FFprobeVideoAnalyzer,
    VideoProbeError,
    find_ffprobe_executable,
    iter_video_files,
)


class DropArea(QFrame):
    pathsDropped = Signal(list)

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setObjectName("dropArea")
        self.setFixedHeight(0)
        self.setFrameShape(QFrame.Shape.StyledPanel)

        label = QLabel("")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout = QVBoxLayout(self)
        layout.addWidget(label)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.toLocalFile()]
        if paths:
            self.pathsDropped.emit(paths)
        event.acceptProposedAction()


class DropOverlay(QFrame):
    def __init__(self, parent: QWidget) -> None:
        super().__init__(parent)
        self.setObjectName("dropOverlay")
        self.hide()

        label = QLabel("ここにドロップ")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("color: #4f7dbd; font-size: 22px; font-weight: bold; background: transparent;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 32, 32, 32)
        layout.addWidget(label)


class PresetListDialog(QDialog):
    def __init__(
        self,
        company_names: list[str],
        project_names: list[str],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("候補リスト設定")
        self.resize(520, 420)

        self.company_names_edit = QPlainTextEdit()
        self.company_names_edit.setPlaceholderText("1行につき1件ずつ会社名を入力")
        self.company_names_edit.setPlainText("\n".join(company_names))

        self.project_names_edit = QPlainTextEdit()
        self.project_names_edit.setPlaceholderText("1行につき1件ずつ作品名を入力")
        self.project_names_edit.setPlainText("\n".join(project_names))

        form_layout = QFormLayout()
        form_layout.addRow("会社名候補", self.company_names_edit)
        form_layout.addRow("作品名候補", self.project_names_edit)

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form_layout)
        layout.addWidget(button_box)

    def company_names(self) -> list[str]:
        return self._parse_lines(self.company_names_edit.toPlainText())

    def project_names(self) -> list[str]:
        return self._parse_lines(self.project_names_edit.toPlainText())

    def _parse_lines(self, text: str) -> list[str]:
        values: list[str] = []
        for line in text.splitlines():
            value = line.strip()
            if value and value not in values:
                values.append(value)
        return values


class VideoLoadThread(QThread):
    progressChanged = Signal(int, int, str)
    loadFinished = Signal(list, list, str, str)

    def __init__(
        self,
        raw_paths: list[Path],
        existing_paths: set[Path],
        analyzer: FFprobeVideoAnalyzer,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.raw_paths = raw_paths
        self.existing_paths = existing_paths
        self.analyzer = analyzer

    def run(self) -> None:
        try:
            folder_name = ""
            loaded_items: list[VideoItem] = []
            errors: list[str] = []
            seen_paths: set[Path] = set()
            found_any_video = False
            found_count = 0
            done_count = 0
            max_workers = max(2, min(8, (os.cpu_count() or 4)))
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_map = {}
                for path in iter_video_files(self.raw_paths):
                    if self.isInterruptionRequested():
                        break
                    found_any_video = True
                    if path in self.existing_paths or path in seen_paths:
                        continue
                    seen_paths.add(path)
                    if not folder_name:
                        folder_name = path.parent.name

                    future = executor.submit(self.analyzer.analyze_fast, path)
                    future_map[future] = path
                    found_count += 1

                    if len(future_map) >= max_workers * 2:
                        done_count = self._drain_completed_futures(
                            future_map,
                            loaded_items,
                            errors,
                            done_count,
                            found_count,
                        )

                while future_map and not self.isInterruptionRequested():
                    done_count = self._drain_completed_futures(
                        future_map,
                        loaded_items,
                        errors,
                        done_count,
                        found_count,
                    )

            if not found_any_video:
                self.loadFinished.emit([], [], "対応動画ファイルが見つかりませんでした。", "")
                return
            if found_count == 0:
                self.loadFinished.emit([], [], "新しく追加できる動画ファイルはありませんでした。", "")
                return

            self.loadFinished.emit(loaded_items, errors, "", folder_name)
        except Exception as exc:
            self.loadFinished.emit([], [str(exc)], "動画読み込み中にエラーが発生しました。", "")

    def _drain_completed_futures(
        self,
        future_map: dict,
        loaded_items: list[VideoItem],
        errors: list[str],
        done_count: int,
        found_count: int,
    ) -> int:
        done_futures, _ = wait(future_map.keys(), return_when=FIRST_COMPLETED)
        for future in done_futures:
            path = future_map.pop(future)
            try:
                loaded_items.append(future.result())
            except VideoProbeError as exc:
                errors.append(str(exc))
            except Exception as exc:
                errors.append(f"{path.name}: {exc}")
            done_count += 1
            self.progressChanged.emit(done_count, found_count, path.name)
        return done_count


class FrameCountRefineThread(QThread):
    itemRefined = Signal(str, object)

    def __init__(
        self,
        file_paths: list[Path],
        analyzer: FFprobeVideoAnalyzer,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.file_paths = file_paths
        self.analyzer = analyzer

    def run(self) -> None:
        max_workers = min(len(self.file_paths), max(2, min(8, (os.cpu_count() or 4))))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_map = {
                executor.submit(self.analyzer.analyze, file_path): file_path
                for file_path in self.file_paths
            }
            for future in as_completed(future_map):
                if self.isInterruptionRequested():
                    break
                file_path = future_map[future]
                try:
                    self.itemRefined.emit(str(file_path.resolve()), future.result())
                except Exception:
                    continue


class MainWindow(QMainWindow):
    COLUMNS = ["1-25", "ファイル名", "秒+コマ", "26-50", "ファイル名", "秒+コマ"]
    ROWS_PER_COLUMN = 25
    ITEMS_PER_PAGE = 50

    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle("SatsueiSlip - 納品伝票作成")
        self.setAcceptDrops(True)
        self.settings_store = SettingsStore()
        self.app_settings = self.settings_store.load()
        self.video_items: list[VideoItem] = []
        self.pdf_exporter = DeliverySlipPdfExporter()
        self.image_exporter = DeliverySlipImageExporter()
        self.analyzer: FFprobeVideoAnalyzer | None = None
        self.load_thread: VideoLoadThread | None = None
        self.load_progress: QProgressDialog | None = None
        self.refine_thread: FrameCountRefineThread | None = None

        self._setup_ui()
        self._setup_menu()
        self._restore_settings()
        self._init_analyzer()

    def _setup_ui(self) -> None:
        central = QWidget()
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(10)

        form_widget = QWidget()
        form_layout = QFormLayout(form_widget)
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form_layout.setHorizontalSpacing(12)
        form_layout.setVerticalSpacing(8)

        self.company_edit = QComboBox()
        self.company_edit.setEditable(True)
        self.project_edit = QComboBox()
        self.project_edit.setEditable(True)
        self.delivery_date_edit = QDateEdit()
        self.delivery_date_edit.setCalendarPopup(True)
        self.delivery_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.delivery_date_edit.setDate(QDate.currentDate())
        self.episode_edit = QLineEdit()
        self.folder_name_edit = QLineEdit()
        self.head_trim_spin = QSpinBox()
        self.head_trim_spin.setRange(0, 999)
        self.head_trim_spin.setSuffix(" フレーム")
        self.head_trim_spin.setValue(8)
        self.head_trim_spin.valueChanged.connect(self._refresh_table)
        self.note_edit = QPlainTextEdit()
        self.note_edit.setFixedHeight(70)
        self.sender_footer_edit = QPlainTextEdit()
        self.sender_footer_edit.setFixedHeight(70)

        form_layout.addRow("会社名 *", self.company_edit)
        form_layout.addRow("作品名 *", self.project_edit)
        form_layout.addRow("話数", self.episode_edit)
        form_layout.addRow("フォルダー名", self.folder_name_edit)
        form_layout.addRow("納品日", self.delivery_date_edit)
        form_layout.addRow("頭引き", self.head_trim_spin)
        form_layout.addRow("備考", self.note_edit)
        form_layout.addRow("自社フッター", self.sender_footer_edit)

        self.drop_area = DropArea()
        self.drop_area.pathsDropped.connect(self.handle_dropped_paths)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setHorizontalHeaderLabels(self.COLUMNS)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(False)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 48)
        self.table.setColumnWidth(2, 90)
        self.table.setColumnWidth(3, 48)
        self.table.setColumnWidth(5, 90)
        self.table.setRowCount(self.ROWS_PER_COLUMN)
        self.table.setMinimumHeight(620)

        delete_shortcut = QShortcut(QKeySequence.StandardKey.Delete, self.table)
        delete_shortcut.activated.connect(self.delete_selected_rows)

        self.summary_label = QLabel("0/50")
        self.total_time_label = QLabel("0 + 00")
        footer_summary_layout = QHBoxLayout()
        footer_summary_layout.addWidget(self.summary_label)
        footer_summary_layout.addStretch(1)
        footer_summary_layout.addWidget(self.total_time_label)

        button_layout = QHBoxLayout()
        self.export_button = QPushButton("PDF書き出し")
        self.export_image_button = QPushButton("画像書き出し")
        self.delete_button = QPushButton("選択行削除")
        self.clear_button = QPushButton("一覧クリア")
        self.reload_button = QPushButton("再読み込み")
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.export_image_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addStretch(1)
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.reload_button)

        self.export_button.clicked.connect(self.export_pdf)
        self.export_image_button.clicked.connect(self.export_image)
        self.delete_button.clicked.connect(self.delete_selected_rows)
        self.clear_button.clicked.connect(self.clear_items)
        self.reload_button.clicked.connect(self.reload_items)

        root_layout.addWidget(form_widget)
        root_layout.addWidget(self.drop_area)
        root_layout.addWidget(self.table, 1)
        root_layout.addLayout(footer_summary_layout)
        root_layout.addLayout(button_layout)
        self.setCentralWidget(central)
        self._setup_status_bar()

        self.drop_overlay = DropOverlay(central)

        self.setStyleSheet(
            """
            QMainWindow { background: #ffffff; }
            QLineEdit, QPlainTextEdit, QDateEdit, QTableWidget {
                border: 1px solid #c8c8c8;
                border-radius: 4px;
                padding: 4px;
                background: #ffffff;
            }
            #dropArea {
                border: none;
                border-radius: 8px;
                background: transparent;
            }
            #dropOverlay {
                border: 3px dashed #7ba6dc;
                border-radius: 14px;
                background: rgba(170, 210, 255, 80);
            }
            QPushButton {
                padding: 8px 16px;
                border-radius: 4px;
                border: 1px solid #8d8d8d;
                background: #f5f5f5;
            }
            QPushButton:hover { background: #e9eef5; }
            QStatusBar {
                background: #fafafa;
                border-top: 1px solid #dcdcdc;
            }
            QStatusBar::item { border: none; }
            """
        )

    def _setup_status_bar(self) -> None:
        status_bar = self.statusBar()
        status_bar.setSizeGripEnabled(False)

        self.ffprobe_status_label = QLabel("ffprobe: 未確認")
        self.ffprobe_status_label.setStyleSheet("color: #555;")
        status_bar.addPermanentWidget(self.ffprobe_status_label)

        self._set_status_message("準備完了")

    def _set_status_message(self, message: str, timeout_ms: int = 0) -> None:
        self.statusBar().showMessage(message, timeout_ms)

    def _set_ffprobe_status(self, text: str, tooltip: str = "") -> None:
        self.ffprobe_status_label.setText(text)
        self.ffprobe_status_label.setToolTip(tooltip)

    def _restore_settings(self) -> None:
        self._set_combo_values(self.company_edit, self.app_settings.company_names or [], self.app_settings.company_name)
        self._set_combo_values(self.project_edit, self.app_settings.project_names or [], self.app_settings.project_name)
        self.sender_footer_edit.setPlainText(self.app_settings.sender_footer)
        self.head_trim_spin.setValue(self.app_settings.head_trim_frames)
        if self.app_settings.window_size is not None:
            self.resize(self.app_settings.window_size)
        else:
            self.resize(1180, 780)

    def _setup_menu(self) -> None:
        settings_menu = self.menuBar().addMenu("設定")
        preset_action = QAction("会社名/作品名 候補編集", self)
        preset_action.triggered.connect(self.edit_presets)
        settings_menu.addAction(preset_action)

        reset_sensitive_action = QAction("会社/作品情報をリセット", self)
        reset_sensitive_action.triggered.connect(self.reset_sensitive_settings)
        settings_menu.addAction(reset_sensitive_action)

        help_menu = self.menuBar().addMenu("ヘルプ")
        update_action = QAction("GitHubの更新を確認", self)
        update_action.triggered.connect(self.check_for_updates)
        help_menu.addAction(update_action)

    def _init_analyzer(self) -> None:
        ffprobe_path = find_ffprobe_executable()
        if ffprobe_path:
            self.analyzer = FFprobeVideoAnalyzer(ffprobe_path)
            self._set_ffprobe_status("ffprobe: 利用可能", ffprobe_path)
            self._set_status_message("準備完了")
            return

        self.analyzer = None
        self._set_ffprobe_status("ffprobe: 未検出")
        self._set_status_message("ffprobe が見つかりません。PATH か tools\\ffprobe に配置してください。")

    def closeEvent(self, event) -> None:
        if self.load_thread and self.load_thread.isRunning():
            self.load_thread.requestInterruption()
            self.load_thread.wait(1000)
        if self.refine_thread and self.refine_thread.isRunning():
            self.refine_thread.requestInterruption()
            self.refine_thread.wait(1000)
        self._save_settings()
        super().closeEvent(event)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            self._show_drop_overlay()
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            self._show_drop_overlay()
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self._hide_drop_overlay()
        super().dragLeaveEvent(event)

    def dropEvent(self, event) -> None:
        self._hide_drop_overlay()
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.toLocalFile()]
        if paths:
            self._set_status_message("読み込み準備中...")
            QApplication.processEvents()
            self.handle_dropped_paths(paths)
        event.acceptProposedAction()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        if hasattr(self, "drop_overlay"):
            self.drop_overlay.setGeometry(self.centralWidget().rect().adjusted(12, 12, -12, -12))

    def _show_drop_overlay(self) -> None:
        self.drop_overlay.setGeometry(self.centralWidget().rect().adjusted(12, 12, -12, -12))
        self.drop_overlay.raise_()
        self.drop_overlay.show()

    def _hide_drop_overlay(self) -> None:
        self.drop_overlay.hide()

    def handle_dropped_paths(self, raw_paths: list[Path]) -> None:
        if self.load_thread and self.load_thread.isRunning():
            QMessageBox.information(self, "読み込み中", "現在の動画読み込みが終わるまでお待ちください。")
            return

        if not self.analyzer:
            self._init_analyzer()
        if not self.analyzer:
            QMessageBox.critical(
                self,
                "エラー",
                "ffprobe が見つかりません。PATH に追加するか tools\\ffprobe 配下へ配置してください。",
            )
            return

        self._set_status_message("動画を検索中...")
        self.load_progress = QProgressDialog("動画を検索中...", "キャンセル", 0, 0, self)
        self.load_progress.setWindowModality(Qt.WindowModality.WindowModal)
        self.load_progress.setMinimumDuration(0)
        self.load_progress.canceled.connect(self._cancel_video_loading)
        QApplication.processEvents()

        self.load_thread = VideoLoadThread(
            raw_paths=raw_paths,
            existing_paths={item.file_path for item in self.video_items},
            analyzer=self.analyzer,
            parent=self,
        )
        self.load_thread.progressChanged.connect(self._on_video_load_progress)
        self.load_thread.loadFinished.connect(self._on_video_load_finished)
        self.load_thread.finished.connect(self.load_thread.deleteLater)
        self.load_thread.start()

    def _cancel_video_loading(self) -> None:
        if self.load_thread and self.load_thread.isRunning():
            self.load_thread.requestInterruption()
            self._set_status_message("読み込みをキャンセル中...")

    def _on_video_load_progress(self, done_count: int, total_count: int, file_name: str) -> None:
        if not self.load_progress:
            return
        if self.load_progress.maximum() != total_count:
            self.load_progress.setRange(0, total_count)
        self.load_progress.setValue(done_count)
        self.load_progress.setLabelText(f"動画を解析中... ({done_count}/{total_count})\n{file_name}")
        self._set_status_message(f"動画を解析中... ({done_count}/{total_count}) {file_name}")

    def _on_video_load_finished(
        self,
        loaded_items: list[VideoItem],
        errors: list[str],
        info_message: str,
        folder_name: str,
    ) -> None:
        if self.load_progress:
            self.load_progress.close()
            self.load_progress = None

        self.load_thread = None

        if loaded_items:
            self.video_items.extend(loaded_items)
            self.video_items.sort(key=lambda item: item.file_name.lower())
            if folder_name and not self.folder_name_edit.text().strip():
                self.folder_name_edit.setText(folder_name)
            if not self.episode_edit.text().strip():
                episode_name = self._guess_episode_name(folder_name, loaded_items)
                if episode_name:
                    self.episode_edit.setText(episode_name)
            self._refresh_table()
            self._start_refine_frame_counts([item.file_path for item in loaded_items])
        elif info_message:
            self._set_status_message(info_message)
        elif errors:
            self._set_status_message("動画の読み込みに失敗しました。")
        else:
            self._set_status_message("読み込みを終了しました。")

        if info_message:
            QMessageBox.information(self, "読み込み", info_message)

        if errors:
            if loaded_items:
                message = f"{len(loaded_items)} 件を追加しました。\n\n" + "\n".join(errors[:10])
            else:
                message = "\n".join(errors[:10])
            QMessageBox.warning(self, "一部の動画を読み込めませんでした", message)

    def _start_refine_frame_counts(self, file_paths: list[Path]) -> None:
        if not self.analyzer or not file_paths:
            return
        if self.refine_thread and self.refine_thread.isRunning():
            self.refine_thread.requestInterruption()
            self.refine_thread.wait(1000)

        self._set_status_message("フレーム数をバックグラウンドで精査中...")
        self.refine_thread = FrameCountRefineThread(file_paths, self.analyzer, self)
        self.refine_thread.itemRefined.connect(self._on_frame_count_refined)
        self.refine_thread.finished.connect(self._on_refine_finished)
        self.refine_thread.finished.connect(self.refine_thread.deleteLater)
        self.refine_thread.start()

    def _on_frame_count_refined(self, file_path_text: str, refined_item: object) -> None:
        if not isinstance(refined_item, VideoItem):
            return
        for index, current_item in enumerate(self.video_items):
            if str(current_item.file_path) == file_path_text:
                refined_item.note = current_item.note
                self.video_items[index] = refined_item
                break
        self._refresh_table()

    def _on_refine_finished(self) -> None:
        self.refine_thread = None
        self._init_analyzer()

    def _refresh_table(self) -> None:
        self.table.blockSignals(True)
        self.table.clearContents()
        page_items = self.video_items[: self.ITEMS_PER_PAGE]
        head_trim_frames = self.head_trim_spin.value()
        for row in range(self.ROWS_PER_COLUMN):
            left_item = page_items[row] if row < len(page_items) else None
            right_index = row + self.ROWS_PER_COLUMN
            right_item = page_items[right_index] if right_index < len(page_items) else None
            self._set_sheet_row(row, 0, self._left_display_number(row), left_item, head_trim_frames)
            self._set_sheet_row(row, 3, self._right_display_number(row), right_item, head_trim_frames)
        self.table.blockSignals(False)
        total_frames = sum(item.trimmed_frame_count(head_trim_frames) for item in page_items)
        total_seconds, total_remain = divmod(total_frames, 24)
        self.summary_label.setText(f"{len(page_items)}/50")
        self.total_time_label.setText(f"{total_seconds} + {total_remain:02d}")

    def _set_sheet_row(
        self,
        row: int,
        start_col: int,
        display_index: int,
        item: VideoItem | None,
        head_trim_frames: int,
    ) -> None:
        values = ["", "", ""]
        if item is not None:
            values = [
                str(display_index),
                item.file_name,
                item.cut_duration_text(head_trim_frames, 24),
            ]
        for offset, value in enumerate(values):
            cell = QTableWidgetItem(value)
            cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if offset == 1 and item is not None:
                cell.setToolTip(str(item.file_path))
            if offset in {0, 2}:
                cell.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, start_col + offset, cell)

    def _left_display_number(self, row: int) -> int:
        return row + 1

    def _right_display_number(self, row: int) -> int:
        return row + self.ROWS_PER_COLUMN + 1

    def delete_selected_rows(self) -> None:
        target_indexes: set[int] = set()
        for index in self.table.selectedIndexes():
            if index.column() <= 2:
                target_indexes.add(index.row())
            else:
                target_indexes.add(index.row() + self.ROWS_PER_COLUMN)
        selected_indexes = sorted((i for i in target_indexes if i < len(self.video_items)), reverse=True)
        if not selected_indexes:
            QMessageBox.information(self, "行削除", "削除する行を選択してください。")
            return

        for item_index in selected_indexes:
            del self.video_items[item_index]
        self._refresh_table()

    def clear_items(self) -> None:
        if not self.video_items:
            return
        result = QMessageBox.question(self, "一覧クリア", "一覧をすべてクリアしますか？")
        if result != QMessageBox.StandardButton.Yes:
            return
        self.video_items.clear()
        self._refresh_table()

    def reload_items(self) -> None:
        if not self.video_items:
            QMessageBox.information(self, "再読み込み", "再読み込みする動画がありません。")
            return
        if not self.analyzer:
            self._init_analyzer()
        if not self.analyzer:
            QMessageBox.critical(self, "エラー", "ffprobe が見つかりません。")
            return

        progress = QProgressDialog("動画を再読み込み中...", "キャンセル", 0, len(self.video_items), self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)

        reloaded: list[VideoItem] = []
        errors: list[str] = []
        indexed_items = list(enumerate(self.video_items))
        reloaded_map: dict[int, VideoItem] = {}
        max_workers = min(len(indexed_items), max(2, min(8, (os.cpu_count() or 4))))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_map = {
                executor.submit(self.analyzer.analyze, old_item.file_path): (index, old_item)
                for index, old_item in indexed_items
            }
            for done_count, future in enumerate(as_completed(future_map), start=1):
                index, old_item = future_map[future]
                progress.setValue(done_count - 1)
                progress.setLabelText(f"動画を再読み込み中...\n{old_item.file_name}")
                QApplication.processEvents()
                if progress.wasCanceled():
                    break
                try:
                    new_item = future.result()
                    new_item.note = old_item.note
                    reloaded_map[index] = new_item
                except VideoProbeError as exc:
                    reloaded_map[index] = old_item
                    errors.append(str(exc))
                except Exception as exc:
                    reloaded_map[index] = old_item
                    errors.append(f"{old_item.file_name}: {exc}")

        reloaded = [reloaded_map.get(index, old_item) for index, old_item in indexed_items]

        progress.setValue(len(self.video_items))
        self.video_items = sorted(reloaded, key=lambda item: item.file_name.lower())
        self._refresh_table()

        if errors:
            QMessageBox.warning(
                self,
                "一部の動画を再読み込みできませんでした",
                "\n".join(errors[:10]),
            )

    def export_pdf(self) -> None:
        delivery_info = self._collect_delivery_info()
        validation_error = self._validate_before_export(delivery_info)
        if validation_error:
            QMessageBox.warning(self, "入力エラー", validation_error)
            return

        default_dir = self.app_settings.last_pdf_dir or str(Path.home() / "Desktop")
        default_name = f"{self._default_export_stem(delivery_info)}.pdf"
        default_path = str(Path(default_dir) / default_name)
        output_path_text, _ = QFileDialog.getSaveFileName(
            self,
            "納品伝票PDFを書き出し",
            default_path,
            "PDFファイル (*.pdf)",
        )
        if not output_path_text:
            return

        output_path = Path(output_path_text)
        if output_path.suffix.lower() != ".pdf":
            output_path = output_path.with_suffix(".pdf")

        try:
            self.pdf_exporter.export(output_path, delivery_info, self.video_items)
        except Exception as exc:
            QMessageBox.critical(self, "PDF出力エラー", f"PDFの書き出しに失敗しました。\n{exc}")
            return

        self.app_settings.last_pdf_dir = str(output_path.parent)
        self._save_settings()
        QMessageBox.information(self, "PDF書き出し完了", f"PDFを書き出しました。\n{output_path}")

    def export_image(self) -> None:
        delivery_info = self._collect_delivery_info()
        validation_error = self._validate_before_export(delivery_info)
        if validation_error:
            QMessageBox.warning(self, "入力エラー", validation_error)
            return

        default_dir = self.app_settings.last_pdf_dir or str(Path.home() / "Desktop")
        default_name = f"{self._default_export_stem(delivery_info)}.png"
        default_path = str(Path(default_dir) / default_name)
        output_path_text, _ = QFileDialog.getSaveFileName(
            self,
            "納品伝票画像を書き出し",
            default_path,
            "PNG画像 (*.png)",
        )
        if not output_path_text:
            return

        output_path = Path(output_path_text)
        if output_path.suffix.lower() != ".png":
            output_path = output_path.with_suffix(".png")

        try:
            output_paths = self.image_exporter.export(output_path, delivery_info, self.video_items)
        except Exception as exc:
            QMessageBox.critical(self, "画像出力エラー", f"画像の書き出しに失敗しました。\n{exc}")
            return

        self.app_settings.last_pdf_dir = str(output_path.parent)
        self._save_settings()
        if len(output_paths) == 1:
            message = f"画像を書き出しました。\n{output_paths[0]}"
        else:
            message = (
                f"{len(output_paths)}枚の画像を書き出しました。\n"
                f"{output_paths[0].parent}\\{output_path.stem}_01.png など"
            )
        QMessageBox.information(self, "画像書き出し完了", message)

    def _collect_delivery_info(self) -> DeliveryInfo:
        qdate = self.delivery_date_edit.date()
        return DeliveryInfo(
            company_name=self.company_edit.currentText().strip(),
            project_name=self.project_edit.currentText().strip(),
            delivery_date=date(qdate.year(), qdate.month(), qdate.day()),
            recipient="",
            staff_name="",
            note=self.note_edit.toPlainText().strip(),
            episode_name=self.episode_edit.text().strip(),
            folder_name=self.folder_name_edit.text().strip(),
            head_trim_frames=self.head_trim_spin.value(),
            sender_footer=self.sender_footer_edit.toPlainText().strip(),
        )

    def _validate_before_export(self, delivery_info: DeliveryInfo) -> str | None:
        if not delivery_info.company_name:
            return "会社名を入力してください。"
        if not delivery_info.project_name:
            return "作品名を入力してください。"
        if not self.video_items:
            return "動画を1件以上読み込んでください。"
        return None

    def _default_export_stem(self, delivery_info: DeliveryInfo) -> str:
        stem = delivery_info.folder_name.strip() or delivery_info.project_name.strip() or "納品伝票"
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            stem = stem.replace(char, "_")
        return stem.strip().rstrip(".") or "納品伝票"

    def _guess_episode_name(self, folder_name: str, items: list[VideoItem]) -> str:
        candidates = [folder_name]
        candidates.extend(item.file_path.stem for item in items[:20])
        patterns = [
            r"(?:ep|episode|話数)[_\- ]*(\d{1,3})",
            r"(?:^|[_\- ])(\d{1,3})(?=[_\- ]|$)",
        ]
        for text in candidates:
            if not text:
                continue
            for pattern in patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    return match.group(1)
        return ""

    def _save_settings(self) -> None:
        self.app_settings.company_name = self.company_edit.currentText().strip()
        self.app_settings.project_name = self.project_edit.currentText().strip()
        self.app_settings.company_names = self._merged_combo_values(
            self.company_edit,
            self.app_settings.company_names or [],
        )
        self.app_settings.project_names = self._merged_combo_values(
            self.project_edit,
            self.app_settings.project_names or [],
        )
        self.app_settings.sender_footer = self.sender_footer_edit.toPlainText().strip()
        self.app_settings.head_trim_frames = self.head_trim_spin.value()
        self.app_settings.window_size = self.size()
        self.settings_store.save(self.app_settings)

    def edit_presets(self) -> None:
        dialog = PresetListDialog(
            self._merged_combo_values(self.company_edit, self.app_settings.company_names or []),
            self._merged_combo_values(self.project_edit, self.app_settings.project_names or []),
            self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        company_names = dialog.company_names()
        project_names = dialog.project_names()
        self.app_settings.company_names = company_names
        self.app_settings.project_names = project_names
        self._set_combo_values(self.company_edit, company_names, self.company_edit.currentText().strip())
        self._set_combo_values(self.project_edit, project_names, self.project_edit.currentText().strip())
        self._save_settings()

    def reset_sensitive_settings(self) -> None:
        result = QMessageBox.question(
            self,
            "会社/作品情報をリセット",
            (
                "保存済みの会社名・作品名、候補一覧、前回の書き出し先、"
                "差出人フッターをローカル設定から削除します。\n"
                "動画一覧は削除しません。続行しますか？"
            ),
        )
        if result != QMessageBox.StandardButton.Yes:
            return

        self.app_settings.company_name = ""
        self.app_settings.project_name = ""
        self.app_settings.company_names = []
        self.app_settings.project_names = []
        self.app_settings.last_pdf_dir = ""
        self.app_settings.sender_footer = ""

        self._set_combo_values(self.company_edit, [], "")
        self._set_combo_values(self.project_edit, [], "")
        self.sender_footer_edit.clear()

        self._save_settings()
        self._set_status_message("会社/作品情報をリセットしました。", 5000)
        QMessageBox.information(self, "リセット完了", "ローカル保存していた会社/作品情報を削除しました。")

    def check_for_updates(self) -> None:
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            update_info = check_github_update()
        except UpdateCheckError as exc:
            QMessageBox.warning(self, "更新確認エラー", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        if not update_info.has_update:
            QMessageBox.information(
                self,
                "更新確認",
                f"現在のバージョン {update_info.current_version} は最新です。",
            )
            return

        result = QMessageBox.question(
            self,
            "更新があります",
            (
                f"新しいバージョン {update_info.latest_version} が見つかりました。\n"
                f"現在のバージョン: {update_info.current_version}\n\n"
                "GitHubのリリースページを開きますか？"
            ),
        )
        if result == QMessageBox.StandardButton.Yes:
            QDesktopServices.openUrl(QUrl(update_info.release_url))

    def _set_combo_values(self, combo_box: QComboBox, values: list[str], current_text: str) -> None:
        combo_box.clear()
        unique_values = [value for value in values if value]
        combo_box.addItems(unique_values)
        if current_text:
            if current_text not in unique_values:
                combo_box.insertItem(0, current_text)
            combo_box.setCurrentText(current_text)
        elif unique_values:
            combo_box.setCurrentIndex(0)

    def _merged_combo_values(self, combo_box: QComboBox, base_values: list[str]) -> list[str]:
        values: list[str] = []
        current_text = combo_box.currentText().strip()
        if current_text:
            values.append(current_text)
        for value in base_values:
            clean_value = value.strip()
            if clean_value and clean_value not in values:
                values.append(clean_value)
        for index in range(combo_box.count()):
            clean_value = combo_box.itemText(index).strip()
            if clean_value and clean_value not in values:
                values.append(clean_value)
        return values
