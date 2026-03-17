# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QFileDialog,
    QInputDialog,
    QMainWindow,
    QMessageBox,
    QTableWidgetItem,
)

from data.store import WordStore
from services.app_config import get_ui_language
from services.tts import cancel_all as tts_cancel_all
from services.tts import get_runtime_label, speak_async

from .ui_main_window import Ui_MainWindow


APP_STYLE = """
QWidget { background: #f4f1ea; color: #1f2933; font-family: 'Segoe UI', 'Microsoft YaHei UI'; font-size: 13px; }
QMainWindow { background: #f4f1ea; }
QFrame#headerCard, QFrame#leftCard, QFrame#detailCard, QTabWidget::pane { background: #fffaf3; border: 1px solid #e2d8ca; border-radius: 18px; }
QLabel#titleLabel { font-family: 'Georgia'; font-size: 24px; font-weight: 700; }
QLabel#wordListTitleLabel, QLabel#currentWordTitleLabel, QLabel#studyFocusTitleLabel, QLabel#historyTitleLabel, QLabel#learningToolsTitleLabel { font-family: 'Georgia'; font-size: 18px; font-weight: 700; }
QLabel#subtitleLabel, QLabel#wordListDescLabel, QLabel#toolsHintLabel, QLabel#historyEmptyLabel, QLabel#emptyStateLabel, QLabel#statusLabel, QLabel#reviewStatsValueLabel, QLabel#reviewNoteValueLabel { color: #667085; }
QLabel#currentFileStaticLabel, QLabel#audioBackendStaticLabel, QLabel#wrongCountStaticLabel, QLabel#correctCountStaticLabel, QLabel#lastResultStaticLabel, QLabel#lastSeenStaticLabel { color: #6b7280; font-weight: 600; }
QPushButton { background: #f8f4ee; border: 1px solid #d8cbb8; border-radius: 10px; padding: 8px 12px; }
QPushButton:hover { background: #fff8ee; border-color: #b9835a; }
QPushButton:pressed { background: #efe0c9; }
QLineEdit, QTextEdit, QListWidget, QTableWidget { background: #fffdf9; border: 1px solid #ddd2c2; border-radius: 12px; selection-background-color: #d9b58e; selection-color: #1f2933; }
QHeaderView::section { background: #efe6da; border: none; border-bottom: 1px solid #ddd2c2; padding: 8px; color: #6b7280; font-weight: 600; }
QTabBar::tab { background: #ede4d8; border: 1px solid #ddd2c2; padding: 8px 14px; border-top-left-radius: 10px; border-top-right-radius: 10px; margin-right: 4px; }
QTabBar::tab:selected { background: #fffaf3; border-bottom-color: #fffaf3; }
QSplitter::handle { background: #e2d8ca; width: 2px; }
"""


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.language = get_ui_language()
        self.store = WordStore()
        self.is_dirty = False
        self._table_syncing = False
        self._detail_syncing = False

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self._configure_ui()
        self._connect_signals()
        self._install_shortcuts()
        self._apply_texts()
        self.refresh_after_store_change()
        self.set_status("Qt 主界面已就绪。")

    def _configure_ui(self):
        self.resize(1480, 860)
        self.setMinimumSize(1180, 720)
        self.ui.wordTable.verticalHeader().setVisible(False)
        self.ui.wordTable.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.SelectedClicked
        )
        self.ui.wordTable.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.ui.wordTable.setSelectionMode(QAbstractItemView.SingleSelection)
        self.ui.wordTable.horizontalHeader().setStretchLastSection(True)
        self.ui.wordTable.setColumnWidth(0, 68)
        self.ui.wordTable.setColumnWidth(1, 320)
        self.ui.wordTable.setColumnWidth(2, 280)
        self.ui.selectedNoteEdit.setFixedHeight(110)
        self.ui.mainSplitter.setSizes([720, 700])
        self.ui.statusbar.hide()

    def _apply_texts(self):
        self.setWindowTitle("Word Speaker")
        self.ui.titleLabel.setText("Word Speaker")
        self.ui.subtitleLabel.setText("PySide6 迁移版：主布局尽量贴近旧版 Tk 界面。")
        self.ui.currentFileStaticLabel.setText("当前文件")
        self.ui.audioBackendStaticLabel.setText("当前音源")
        self.ui.wordListTitleLabel.setText("单词表")
        self.ui.wordListDescLabel.setText("导入词表、直接编辑，然后从当前选中单词开始学习。")
        self.ui.emptyStateLabel.setText("还没有单词，先点导入开始。")
        self.ui.currentWordTitleLabel.setText("当前单词")
        self.ui.selectedWordEdit.setPlaceholderText("这里编辑单词内容")
        self.ui.selectedNoteEdit.setPlaceholderText("这里编辑备注、中文或学习提示")
        self.ui.wrongCountStaticLabel.setText("错过次数")
        self.ui.correctCountStaticLabel.setText("正确次数")
        self.ui.lastResultStaticLabel.setText("上次结果")
        self.ui.lastSeenStaticLabel.setText("最近时间")
        self.ui.studyFocusTitleLabel.setText("学习焦点")
        self.ui.historyTitleLabel.setText("历史")
        self.ui.historyEmptyLabel.setText("还没有历史记录。")
        self.ui.learningToolsTitleLabel.setText("学习工具")
        self.ui.toolsHintLabel.setText("提示：先选中单词，再生成例句或做定向语料检索。")
        self.ui.rightTabWidget.setTabText(self.ui.rightTabWidget.indexOf(self.ui.reviewTab), "复习")
        self.ui.rightTabWidget.setTabText(self.ui.rightTabWidget.indexOf(self.ui.historyTab), "历史")
        self.ui.rightTabWidget.setTabText(self.ui.rightTabWidget.indexOf(self.ui.toolsTab), "工具")
        self.ui.wordTable.setHorizontalHeaderLabels(["#", "单词", "备注"])

        title_font = QFont("Georgia", 20)
        title_font.setBold(True)
        self.ui.titleLabel.setFont(title_font)

    def _connect_signals(self):
        self.ui.importButton.clicked.connect(self.import_words)
        self.ui.manualInputButton.clicked.connect(self.manual_input_words)
        self.ui.saveAsButton.clicked.connect(self.save_words_as)
        self.ui.newListButton.clicked.connect(self.new_list)
        self.ui.playButton.clicked.connect(self.speak_selected_word)
        self.ui.speakWordButton.clicked.connect(self.speak_selected_word)
        self.ui.deleteWordButton.clicked.connect(self.delete_selected_row)
        self.ui.markWrongButton.clicked.connect(self.mark_selected_wrong)
        self.ui.openHistoryButton.clicked.connect(self.open_selected_history)
        self.ui.openSourceButton.clicked.connect(self.open_selected_history)

        for button in [
            self.ui.settingsButton,
            self.ui.dictationButton,
            self.ui.generateSentenceButton,
            self.ui.findInCorpusButton,
            self.ui.editPosTranslationButton,
            self.ui.synonymsButton,
            self.ui.inspectAudioCacheButton,
            self.ui.toolSentenceButton,
            self.ui.toolFindButton,
            self.ui.toolPassageButton,
            self.ui.toolSettingsButton,
            self.ui.toolUpdateButton,
            self.ui.exportCacheButton,
            self.ui.importCacheButton,
            self.ui.exportPackButton,
            self.ui.importPackButton,
        ]:
            button.clicked.connect(self.show_classic_only)

        self.ui.wordTable.itemSelectionChanged.connect(self.on_selection_changed)
        self.ui.wordTable.itemChanged.connect(self.on_table_item_changed)
        self.ui.wordTable.itemDoubleClicked.connect(lambda _item: self.speak_selected_word())
        self.ui.historyListWidget.itemSelectionChanged.connect(self.on_history_selection_changed)
        self.ui.historyListWidget.itemDoubleClicked.connect(self.open_selected_history)
        self.ui.selectedWordEdit.textEdited.connect(self.on_detail_word_changed)
        self.ui.selectedNoteEdit.textChanged.connect(self.on_detail_note_changed)

    def _install_shortcuts(self):
        QShortcut(QKeySequence.Open, self, self.import_words)
        QShortcut(QKeySequence.SaveAs, self, self.save_words_as)
        QShortcut(QKeySequence.Save, self, self.save_words)
        QShortcut(QKeySequence.New, self, self.new_list)
        QShortcut(QKeySequence.Delete, self, self.delete_selected_row)
        QShortcut(QKeySequence(Qt.CTRL | Qt.Key_P), self, self.speak_selected_word)

    def set_status(self, text: str):
        self.ui.statusLabel.setText(text)

    def _update_dirty(self, dirty: bool):
        self.is_dirty = bool(dirty)
        self.setWindowTitle("Word Speaker" + (" *" if self.is_dirty else ""))

    def _ensure_note_length(self):
        if len(self.store.notes) < len(self.store.words):
            self.store.notes.extend([""] * (len(self.store.words) - len(self.store.notes)))

    def maybe_save_changes(self, text: str):
        if not self.is_dirty:
            return True
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Question)
        box.setWindowTitle("Word Speaker")
        box.setText(text)
        save_button = box.addButton("保存", QMessageBox.AcceptRole)
        discard_button = box.addButton("不保存", QMessageBox.DestructiveRole)
        cancel_button = box.addButton("取消", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() == cancel_button:
            return False
        if box.clickedButton() == save_button:
            return self.save_words()
        return box.clickedButton() == discard_button

    def refresh_runtime_info(self):
        self.ui.currentFileValueLabel.setText(self.store.get_current_source_path() or "无")
        self.ui.audioBackendValueLabel.setText(get_runtime_label())

    def refresh_history(self):
        history = self.store.load_history()
        current_path = self.store.get_current_source_path()
        self.ui.historyListWidget.clear()
        self.ui.historyEmptyLabel.setVisible(not history)
        for entry in history:
            name = str(entry.get("name") or "")
            path = str(entry.get("path") or "")
            when = str(entry.get("time") or "")
            text = f"{name}  ·  {when}\n{path}" if when else f"{name}\n{path}"
            self.ui.historyListWidget.addItem(text)
        for row, entry in enumerate(history):
            item = self.ui.historyListWidget.item(row)
            item.setData(Qt.UserRole, str(entry.get("path") or ""))
            if current_path and str(entry.get("path") or "") == current_path:
                item.setBackground(QColor("#efe0c9"))
        self.on_history_selection_changed()

    def refresh_table(self, selected_row=None):
        if selected_row is None:
            selected_row = self.current_row()
        self._table_syncing = True
        try:
            self.ui.wordTable.setRowCount(len(self.store.words))
            for row, word in enumerate(self.store.words):
                note = self.store.notes[row] if row < len(self.store.notes) else ""
                idx_item = QTableWidgetItem(f"{row + 1}.")
                idx_item.setFlags((idx_item.flags() | Qt.ItemIsSelectable | Qt.ItemIsEnabled) & ~Qt.ItemIsEditable)
                self.ui.wordTable.setItem(row, 0, idx_item)
                self.ui.wordTable.setItem(row, 1, QTableWidgetItem(str(word or "")))
                self.ui.wordTable.setItem(row, 2, QTableWidgetItem(str(note or "")))
                self.ui.wordTable.setRowHeight(row, 48)
        finally:
            self._table_syncing = False
        self.ui.emptyStateLabel.setVisible(not self.store.words)
        if self.store.words:
            row = 0 if selected_row is None else max(0, min(selected_row, len(self.store.words) - 1))
            self.ui.wordTable.selectRow(row)
        else:
            self.ui.wordTable.clearSelection()
        self.on_selection_changed()

    def refresh_selection_details(self):
        row = self.current_row()
        word = self.current_word()
        note = self.store.notes[row] if row is not None and row < len(self.store.notes) else ""
        stats = self.store.get_dictation_word_stats(word) if word else {}
        last_result = str(stats.get("last_result") or "")
        last_result_text = "正确" if last_result == "correct" else "错误" if last_result == "wrong" else "无"
        self.ui.wrongCountValueLabel.setText(str(stats.get("wrong_count", 0) or 0) if word else "无")
        self.ui.correctCountValueLabel.setText(str(stats.get("correct_count", 0) or 0) if word else "无")
        self.ui.lastResultValueLabel.setText(last_result_text)
        self.ui.lastSeenValueLabel.setText(str(stats.get("last_seen") or "无") if word else "无")
        self.ui.reviewWordValueLabel.setText(word or "无")
        self.ui.reviewNoteValueLabel.setText(note or "无")
        self.ui.reviewStatsValueLabel.setText(
            f"错过次数：{stats.get('wrong_count', 0) or 0}    正确次数：{stats.get('correct_count', 0) or 0}" if word else "无"
        )

    def refresh_after_store_change(self):
        self.refresh_runtime_info()
        self.refresh_history()
        self.refresh_table()

    def current_row(self):
        indexes = self.ui.wordTable.selectionModel().selectedRows() if self.ui.wordTable.selectionModel() else []
        return indexes[0].row() if indexes else None

    def current_word(self):
        row = self.current_row()
        if row is None or row >= len(self.store.words):
            return ""
        return str(self.store.words[row] or "").strip()

    def on_selection_changed(self):
        row = self.current_row()
        self._detail_syncing = True
        try:
            if row is None or row >= len(self.store.words):
                self.ui.selectedWordEdit.setText("")
                self.ui.selectedNoteEdit.setPlainText("")
            else:
                self.ui.selectedWordEdit.setText(str(self.store.words[row] or ""))
                note = self.store.notes[row] if row < len(self.store.notes) else ""
                self.ui.selectedNoteEdit.setPlainText(str(note or ""))
        finally:
            self._detail_syncing = False
        self.refresh_selection_details()

    def on_table_item_changed(self, item):
        if self._table_syncing or item is None:
            return
        row = item.row()
        if row >= len(self.store.words):
            return
        if item.column() == 1:
            self.store.words[row] = str(item.text() or "")
        elif item.column() == 2:
            self._ensure_note_length()
            self.store.notes[row] = str(item.text() or "")
        self._update_dirty(True)
        self.refresh_selection_details()

    def on_detail_word_changed(self, text):
        if self._detail_syncing:
            return
        row = self.current_row()
        if row is None or row >= len(self.store.words):
            return
        self.store.words[row] = str(text or "")
        self._table_syncing = True
        try:
            item = self.ui.wordTable.item(row, 1)
            if item:
                item.setText(self.store.words[row])
        finally:
            self._table_syncing = False
        self._update_dirty(True)
        self.refresh_selection_details()

    def on_detail_note_changed(self):
        if self._detail_syncing:
            return
        row = self.current_row()
        if row is None or row >= len(self.store.words):
            return
        self._ensure_note_length()
        text = self.ui.selectedNoteEdit.toPlainText()
        self.store.notes[row] = text
        self._table_syncing = True
        try:
            item = self.ui.wordTable.item(row, 2)
            if item:
                item.setText(text)
        finally:
            self._table_syncing = False
        self._update_dirty(True)

    def on_history_selection_changed(self):
        item = self.ui.historyListWidget.currentItem()
        path = str(item.data(Qt.UserRole) or "").strip() if item else ""
        self.ui.historyPathValueLabel.setText(path or "无")

    def import_words(self):
        if not self.maybe_save_changes("当前有未保存改动。是否先保存？"):
            return
        path, _ = QFileDialog.getOpenFileName(self, "选择词表文件", str(Path.cwd()), "Word List (*.txt *.csv)")
        if not path:
            return
        try:
            self.store.load_from_file(path)
        except Exception as exc:
            QMessageBox.critical(self, "Word Speaker", f"导入失败：\n{exc}")
            return
        self.refresh_after_store_change()
        self._update_dirty(False)
        self.set_status(f"已载入 {len(self.store.words)} 个词条。")

    def manual_input_words(self):
        text, ok = QInputDialog.getMultiLineText(self, "粘贴 / 输入单词", "每行一个单词；也可以用 Tab 分隔备注。")
        if not ok:
            return
        words, notes = [], []
        for raw_line in str(text or "").replace("\r", "\n").split("\n"):
            line = str(raw_line or "").strip()
            if not line:
                continue
            word, note = (line.split("\t", 1) + [""])[:2] if "\t" in line else (line, "")
            word = word.strip()
            if word:
                words.append(word)
                notes.append(str(note or "").strip())
        if not words:
            QMessageBox.information(self, "Word Speaker", "没有读到有效单词。")
            return
        choices = ["替换当前词表"] + (["追加到当前词表"] if self.store.words else [])
        mode, ok = QInputDialog.getItem(self, "粘贴 / 输入单词", "导入方式", choices, 0, False)
        if not ok:
            return
        if mode == "追加到当前词表" and self.store.words:
            self.store.set_words(list(self.store.words) + words, list(self.store.notes) + notes, preserve_source=True)
        else:
            self.store.set_words(words, notes, preserve_source=False)
        self.refresh_after_store_change()
        self._update_dirty(True)
        self.set_status(f"已添加 {len(words)} 个词条。")

    def _sanitize_words_before_save(self):
        cleaned_words, cleaned_notes = [], []
        for i, raw_word in enumerate(self.store.words):
            word = str(raw_word or "").strip()
            if not word:
                continue
            note = self.store.notes[i] if i < len(self.store.notes) else ""
            cleaned_words.append(word)
            cleaned_notes.append(str(note or "").strip())
        self.store.set_words(cleaned_words, cleaned_notes, preserve_source=True)
        self.refresh_table()

    def save_words(self):
        self._sanitize_words_before_save()
        path = self.store.get_current_source_path()
        if not path:
            return self.save_words_as()
        try:
            self.store.save_to_file(path)
            self.store.add_history(path)
        except Exception as exc:
            QMessageBox.critical(self, "Word Speaker", f"保存失败：\n{exc}")
            return False
        self.refresh_after_store_change()
        self._update_dirty(False)
        self.set_status(f"已保存到 {path}")
        return True

    def save_words_as(self):
        self._sanitize_words_before_save()
        target, _ = QFileDialog.getSaveFileName(
            self,
            "保存词表",
            self.store.get_current_source_path() or str(Path.cwd() / "words.txt"),
            "Word List (*.txt *.csv)",
        )
        if not target:
            return False
        try:
            self.store.save_to_file(target)
            self.store.add_history(target)
        except Exception as exc:
            QMessageBox.critical(self, "Word Speaker", f"保存失败：\n{exc}")
            return False
        self.refresh_after_store_change()
        self._update_dirty(False)
        self.set_status(f"已保存到 {target}")
        return True

    def new_list(self):
        if not self.maybe_save_changes("当前有未保存改动。是否先保存？"):
            return
        self.store.clear()
        self.refresh_after_store_change()
        self._update_dirty(False)
        self.set_status("已创建空白词表。")

    def delete_selected_row(self):
        row = self.current_row()
        if row is None or row >= len(self.store.words):
            QMessageBox.information(self, "Word Speaker", "请先选中一个单词。")
            return
        if QMessageBox.question(self, "Word Speaker", "要删除当前选中的词条吗？") != QMessageBox.Yes:
            return
        word = self.store.words.pop(row)
        if row < len(self.store.notes):
            self.store.notes.pop(row)
        self.refresh_table(min(row, len(self.store.words) - 1))
        self._update_dirty(True)
        self.set_status(f"已删除词条：{word or '空白词条'}")

    def speak_selected_word(self):
        word = self.current_word()
        if not word:
            QMessageBox.information(self, "Word Speaker", "请先选中一个单词。")
            return
        speak_async(word, volume=1.0, rate_ratio=1.0, cancel_before=True, source_path=self.store.get_current_source_path())
        self.set_status(f"正在朗读：{word}")
        self.refresh_runtime_info()

    def mark_selected_wrong(self):
        word = self.current_word()
        if not word:
            QMessageBox.information(self, "Word Speaker", "请先选中一个单词。")
            return
        self.store.add_wrong_word(word)
        self.refresh_table(self.current_row())
        self.refresh_selection_details()
        self.set_status(f"已加入错词：{word}")

    def open_selected_history(self, _item=None):
        item = self.ui.historyListWidget.currentItem()
        if item is None:
            return
        path = str(item.data(Qt.UserRole) or "").strip()
        if not path:
            return
        if not os.path.exists(path):
            QMessageBox.warning(self, "Word Speaker", f"文件不存在或已被移动：\n{path}")
            return
        if not self.maybe_save_changes("当前有未保存改动。是否先保存？"):
            return
        try:
            self.store.load_from_file(path)
        except Exception as exc:
            QMessageBox.critical(self, "Word Speaker", f"导入失败：\n{exc}")
            return
        self.ui.rightTabWidget.setCurrentWidget(self.ui.historyTab)
        self.refresh_after_store_change()
        self._update_dirty(False)
        self.set_status(f"已载入 {len(self.store.words)} 个词条。")

    def show_classic_only(self):
        QMessageBox.information(self, "Word Speaker", "这个功能还没迁到 Qt。\n\n先用 `python app.py --classic` 打开旧界面更稳。")

    def closeEvent(self, event):
        if not self.maybe_save_changes("当前有未保存改动。关闭前是否先保存？"):
            event.ignore()
            return
        tts_cancel_all()
        super().closeEvent(event)


def run():
    app = QApplication.instance()
    owns_app = app is None
    if app is None:
        app = QApplication([])
    app.setApplicationName("Word Speaker")
    app.setOrganizationName("Word Speaker")
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLE)
    window = MainWindow()
    window.show()
    return app.exec() if owns_app else 0

