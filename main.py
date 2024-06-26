from models import Template
from PyQt6.QtWidgets import (
    QTableWidgetItem,
    QApplication,
    QMainWindow,
    QFileDialog,
    QMessageBox,
)
from PyQt6 import QtWidgets, uic
import pandas as pd
import sys
from pathlib import Path
from word_processor import *
import tempfile
from utils import (
    save_templates,
    load_templates,
    save_excel,
    load_excels,
    create_ini_file,
)
from utils import file_watcher


file = Path(__file__).resolve()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(file.parent / "ui.ui", self)
        self.show()
        if not Path("config.ini").exists():
            create_ini_file()
        self.full_data.itemClicked.connect(self.select_row_full_data)
        self.selected_data.itemClicked.connect(self.select_row_selected_data)
        self.excel_btn.clicked.connect(self.get_excel)
        self.remove_btn.clicked.connect(self.remove_from_selected)
        self.add_btn.clicked.connect(self.add_to_selected)
        self.generate_btn.clicked.connect(self.open)
        self.add_template_btn.clicked.connect(self.add_template)
        self.remove_template_btn.clicked.connect(self.remove_template)

        self.template_cb.currentIndexChanged.connect(self.get_template)

        self.search_txt.textChanged.connect(self.search)
        self.original_selected_data = None
        self.excel_file = None
        self.selected_data.setAcceptDrops(True)
        self.selected_data.setDragEnabled(True)
        self.selected_data.viewport().setAcceptDrops(True)
        self.selected_data.setDropIndicatorShown(True)
        self.selected_data.setDragDropMode(
            QtWidgets.QAbstractItemView.DragDropMode.DropOnly
        )
        self.full_data.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.selected_data.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.watcher_started = False
        self.selected_data.dropEvent = self.drop_event

        app.aboutToQuit.connect(self.on_close)
        self.load_templates()
        self.load_excel_file()

    def on_close(self):
        if self.excel_file:
            save_excel(self.excel_file)
        if self.templates:
            save_templates(self.templates)

    def load_excel_file(self):
        file = load_excels()
        if not file:
            return
        self.load_excel(file)

    def load_templates(self):
        self.templates = load_templates()
        for template in self.templates:
            self.template_cb.addItem(template.name)

        self.template_file = self.templates[0].path if self.templates else None

    def add_template(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open Template File", "", "Word Files (*.docx)"
        )
        if not file_name:
            return
        if any(template.path == file_name for template in self.templates):
            QMessageBox.warning(
                self,
                "Error",
                "Template already exists",
                QMessageBox.StandardButton.Ok,
            )
            return
        self.templates.append(Template(Path(file_name).stem, file_name))
        self.template_cb.addItem(Path(file_name).stem)

    def remove_template(self):
        index = self.template_cb.currentIndex()
        if index == -1:
            return
        confirm = QMessageBox.question(
            self,
            "Confirmation",
            "Are you sure you want to remove the template?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if confirm == QMessageBox.StandardButton.No:
            return
        template_name = self.template_cb.itemText(index)
        if not template_name:
            return
        self.template_cb.removeItem(index)
        self.templates = [
            template for template in self.templates if template.name != template_name
        ]

    def get_template(self):
        seleted_template = self.template_cb.currentText()
        self.template_file = next(
            (
                template.path
                for template in self.templates
                if template.name == seleted_template
            ),
            None,
        )

    def drop_event(self, event):
        event.accept()
        source = event.source()

        name_column_index = -1
        for column in range(source.columnCount()):
            if source.horizontalHeaderItem(column).text():
                name_column_index = column
                break
        if name_column_index == -1:
            return

        existing_names = {
            self.selected_data.item(row, name_column_index).text().strip()
            for row in range(self.selected_data.rowCount())
        }

        for index in source.selectedIndexes():
            row_index = index.row()
            name_item = source.item(row_index, name_column_index)
            if name_item:
                name = name_item.text().strip()
                if name not in existing_names:
                    existing_names.add(name)
                    row_position = self.selected_data.rowCount()
                    self.selected_data.insertRow(row_position)
                    for column in range(source.columnCount()):
                        item = source.item(row_index, column)
                        self.selected_data.setItem(
                            row_position,
                            column,
                            QTableWidgetItem(item.text() if item else ""),
                        )

        self.update_original_selected_data()

    def get_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx)"
        )
        if not file_name:
            return
        self.excel_file = file_name
        self.load_excel(file_name)

    def select_row_full_data(self, item):
        row_index = item.row()
        self.full_data.selectRow(row_index)

    def select_row_selected_data(self, item):
        row_index = item.row()
        self.selected_data.selectRow(row_index)

    def add_to_selected(self):
        selected_rows = set(index.row() for index in self.full_data.selectedIndexes())
        if self.selected_data.rowCount() == 0:
            first_row_added = False
        else:
            first_row_added = True

        name_column_index_full = -1
        for column in range(self.full_data.columnCount()):
            if self.full_data.horizontalHeaderItem(column).text():
                name_column_index_full = column
                break
        if name_column_index_full == -1:
            return

        name_column_index_selected = -1
        for column in range(self.selected_data.columnCount()):
            if self.selected_data.horizontalHeaderItem(column).text():
                name_column_index_selected = column
                break
        if name_column_index_selected == -1:
            return

        for row_index in selected_rows:
            name_item = self.full_data.item(row_index, name_column_index_full)
            if name_item:
                name = name_item.text().strip()
                if all(
                    name
                    != self.selected_data.item(row, name_column_index_selected)
                    .text()
                    .strip()
                    for row in range(self.selected_data.rowCount())
                ):
                    selected_row = [
                        (
                            self.full_data.item(row_index, column).text()
                            if self.full_data.item(row_index, column)
                            else ""
                        )
                        for column in range(self.full_data.columnCount())
                    ]
                    if not first_row_added:
                        self.selected_data.setColumnCount(len(selected_row))
                        self.selected_data.setHorizontalHeaderLabels(
                            self.full_data.horizontalHeaderItem(column).text()
                            for column in range(self.full_data.columnCount())
                        )
                        first_row_added = True
                    row_position = self.selected_data.rowCount()
                    self.selected_data.insertRow(row_position)
                    for column, value in enumerate(selected_row):
                        self.selected_data.setItem(
                            row_position, column, QTableWidgetItem(value)
                        )
        self.update_original_selected_data()

    def remove_from_selected(self):
        indexes = self.selected_data.selectionModel().selectedIndexes()
        rows_to_remove = set()
        for index in indexes:
            rows_to_remove.add(index.row())
            self.selected_data.setItem(index.row(), index.column(), None)
        rows_to_remove = sorted(rows_to_remove, reverse=True)
        for row in rows_to_remove:
            self.selected_data.removeRow(row)
        self.update_original_selected_data()

    def update_original_selected_data(self):
        self.original_selected_data = [
            [
                self.selected_data.item(i, j).text()
                for j in range(self.selected_data.columnCount())
            ]
            for i in range(self.selected_data.rowCount())
        ]

    def search(self):
        search_text = self.search_txt.text()
        if search_text:
            search_df = self.df[
                self.df.apply(
                    lambda row: any(
                        search_text.lower() in str(cell).lower() for cell in row
                    ),
                    axis=1,
                )
            ]
            self.full_data.setRowCount(search_df.shape[0])
            self.full_data.setColumnCount(search_df.shape[1])
            self.full_data.setHorizontalHeaderLabels(search_df.columns)
            for i in range(search_df.shape[0]):
                for j in range(search_df.shape[1]):
                    self.full_data.setItem(
                        i, j, QTableWidgetItem(str(search_df.iat[i, j]))
                    )
            self.full_data.setEditTriggers(
                QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers
            )
            self.full_data.setDragDropMode(
                QtWidgets.QAbstractItemView.DragDropMode.DragDrop
            )
            self.full_data.setAcceptDrops(False)
            self.selected_data.setAcceptDrops(True)
        else:
            self.full_data.setRowCount(self.df.shape[0])
            self.full_data.setColumnCount(self.df.shape[1])
            self.full_data.setHorizontalHeaderLabels(self.df.columns)
            for i in range(self.df.shape[0]):
                for j in range(self.df.shape[1]):
                    self.full_data.setItem(
                        i, j, QTableWidgetItem(str(self.df.iat[i, j]))
                    )
            self.full_data.setEditTriggers(
                QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers
            )
            self.full_data.setDragDropMode(
                QtWidgets.QAbstractItemView.DragDropMode.DragDrop
            )
            self.full_data.setAcceptDrops(False)
            self.selected_data.setAcceptDrops(True)

    def load_excel(self, file_path):
        if not file_path:
            QMessageBox.warning(
                self,
                "Error",
                "No file selected",
                QMessageBox.StandardButton.Ok,
            )
            return

        self.df = pd.read_excel(file_path, sheet_name=0)
        for column in self.df.columns:
            if self.df[column].dtype == "datetime64[ns]":
                self.df[column] = self.df[column].apply(
                    lambda x: x.strftime("%Y-%m-%d")
                )
        if self.watcher_started == False:
            file_watcher.start_watching(file_path, lambda: self.load_excel(file_path))
            self.watcher_started = True
        self.full_data.setRowCount(self.df.shape[0])
        self.full_data.setColumnCount(self.df.shape[1])
        self.full_data.setHorizontalHeaderLabels(self.df.columns)
        self.selected_data.setColumnCount(self.df.shape[1])
        self.selected_data.setHorizontalHeaderLabels(self.df.columns)

        for i in range(self.df.shape[0]):
            for j in range(self.df.shape[1]):
                self.full_data.setItem(i, j, QTableWidgetItem(str(self.df.iat[i, j])))
        self.full_data.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.full_data.setDragDropMode(
            QtWidgets.QAbstractItemView.DragDropMode.DragDrop
        )
        self.full_data.setAcceptDrops(False)
        self.selected_data.setAcceptDrops(True)
        self.selected_data.setColumnCount(self.df.shape[1])
        self.selected_data.setRowCount(0)
        self.original_selected_data = None

    def open(self):
        if not self.template_file:
            QMessageBox.warning(
                self,
                "Error",
                "No template file selected",
                QMessageBox.StandardButton.Ok,
            )
            return
        if self.selected_data.rowCount() == 0:
            QMessageBox.warning(
                self,
                "Error",
                "No data selected",
                QMessageBox.StandardButton.Ok,
            )
            return
        docs = []
        for row_index in range(self.selected_data.rowCount()):
            template = read_docx(self.template_file)

            data = {
                f"{{{{{self.df.columns[column]}}}}}": self.selected_data.item(
                    row_index, column
                ).text()
                for column in range(self.selected_data.columnCount())
            }
            replace_placeholder(template, data)
            docs.append(template)
        output = merge_documents(docs)
        with tempfile.TemporaryDirectory(delete=False) as tmp:
            path = Path(tmp) / f"{next(tempfile._get_candidate_names())}.docx"
            save_word(output, path)
            launch_word(path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
