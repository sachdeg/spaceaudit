import os
import sys
import datetime
import warnings

# Suppress qdarkstyle binding warning
warnings.filterwarnings("ignore", category=UserWarning, module="qdarkstyle")
os.environ["QT_API"] = "pyqt6"

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTreeView, QFileDialog, QVBoxLayout, QWidget,
    QToolBar, QSplitter, QLabel, QMessageBox, QSizePolicy
)
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QAction
from PyQt6.QtCore import Qt
import qdarkstyle
import xlsxwriter

# ------------------- Recursive Scan -------------------
def scan_directory(path):
    total_size = 0
    subfolders = {}
    files = []

    try:
        entries = os.listdir(path)
    except PermissionError:
        return {"size": 0, "subfolders": {}, "files": []}

    for entry in entries:
        full_path = os.path.join(path, entry)
        try:
            if os.path.isdir(full_path):
                data = scan_directory(full_path)
                subfolders[entry] = data
                total_size += data["size"]
            else:
                size = os.path.getsize(full_path)
                last_modified = datetime.datetime.fromtimestamp(
                    os.path.getmtime(full_path)
                ).strftime("%Y-%m-%d")
                files.append((entry, size, last_modified))
                total_size += size
        except OSError:
            continue

    return {"size": total_size, "subfolders": subfolders, "files": files}

class SpaceAudit(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Space Audit")
        self.resize(1200, 700)

        # Toolbar
        toolbar = QToolBar("Main Toolbar")
        self.addToolBar(toolbar)

        select_action = QAction("Select Folder", self)
        select_action.triggered.connect(self.select_folder)
        toolbar.addAction(select_action)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(spacer)

        export_action = QAction("Export Excel", self)
        export_action.triggered.connect(self.export_excel)
        toolbar.addAction(export_action)

        # Split layout
        splitter = QSplitter(Qt.Orientation.Vertical)

        # TreeView
        self.tree_view = QTreeView()
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(["Name", "Size", "Percent", "Last Modified"])
        self.tree_view.setModel(self.model)
        self.tree_view.setColumnWidth(0, 400)
        self.tree_view.setColumnWidth(1, 120)
        self.tree_view.setColumnWidth(2, 100)
        self.tree_view.setColumnWidth(3, 150)

        # Status label
        self.status_label = QLabel("Ready")

        splitter.addWidget(self.tree_view)
        splitter.setStretchFactor(0, 1)

        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(toolbar)
        layout.addWidget(splitter)
        layout.addWidget(self.status_label)
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.results = {}
        self.root_size = 0
        self.root_folder = ""

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.status_label.setText("Scan in Progress...")
            QApplication.processEvents()

            self.model.removeRows(0, self.model.rowCount())
            data = scan_directory(folder)
            self.root_size = data["size"]
            self.root_folder = folder
            self.populate_tree(folder, data, None)

            self.tree_view.expand(self.model.index(0, 0))
            self.status_label.setText("Scan Complete")

    def populate_tree(self, name, data, parent_item):
        size_gb = f"{data['size'] / 1_073_741_824:.2f} GB"
        percent = f"{(data['size'] / self.root_size) * 100:.2f}%" if self.root_size else "0%"
        last_modified = datetime.datetime.fromtimestamp(
            os.path.getmtime(name)
        ).strftime("%Y-%m-%d") if os.path.exists(name) else ""

        item_name = QStandardItem(name if parent_item is None else os.path.basename(name))
        item_size = QStandardItem(size_gb)
        item_percent = QStandardItem(percent)
        item_date = QStandardItem(last_modified)

        if parent_item:
            parent_item.appendRow([item_name, item_size, item_percent, item_date])
        else:
            self.model.appendRow([item_name, item_size, item_percent, item_date])

        current_item = item_name

        if data["files"]:
            loose_item = QStandardItem("Loose Files")
            loose_size = sum(f[1] for f in data["files"])
            loose_size_item = QStandardItem(f"{loose_size / 1_073_741_824:.2f} GB")
            current_item.appendRow([loose_item, loose_size_item, QStandardItem(""), QStandardItem("")])
            for f_name, f_size, f_date in data["files"]:
                file_item = QStandardItem(f_name)
                file_size_item = QStandardItem(f"{f_size / 1_073_741_824:.2f} GB")
                loose_item.appendRow([file_item, file_size_item, QStandardItem(""), QStandardItem(f_date)])

        for subfolder, subdata in data["subfolders"].items():
            self.populate_tree(os.path.join(name, subfolder), subdata, current_item)

    def export_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        try:
            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet("Space Audit")

            header_format = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9'})
            worksheet.write_row(0, 0, ["Name", "Size", "Percent", "Last Modified"], header_format)

            row = 1
            row = self.write_tree_to_excel(self.model.invisibleRootItem(), worksheet, row, level=0)

            worksheet.set_column(0, 0, 50)
            worksheet.set_column(1, 3, 15)

            top_data = self.get_top_subfolders()
            if top_data:
                chart = workbook.add_chart({'type': 'pie'})
                labels = [label for label, _ in top_data]
                sizes = [size for _, size in top_data]

                chart_sheet = workbook.add_worksheet("ChartData")
                chart_sheet.write_column(0, 0, labels)
                chart_sheet.write_column(0, 1, sizes)

                chart.add_series({
                    'categories': f'ChartData!A1:A{len(labels)}',
                    'values': f'ChartData!B1:B{len(sizes)}',
                    'data_labels': {'percentage': True}
                })
                chart.set_title({'name': f"Top Folders in {os.path.basename(self.root_folder)}"})
                worksheet.insert_chart('F2', chart)

            workbook.close()
            QMessageBox.information(self, "Export Complete", f"Excel exported to {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Failed", f"Error: {str(e)}")

    def write_tree_to_excel(self, parent_item, worksheet, row, level):
        indent = "    " * level
        for r in range(parent_item.rowCount()):
            name = indent + parent_item.child(r, 0).text()
            size = parent_item.child(r, 1).text()
            percent = parent_item.child(r, 2).text()
            last_modified = parent_item.child(r, 3).text()
            worksheet.write_row(row, 0, [name, size, percent, last_modified])
            row += 1
            row = self.write_tree_to_excel(parent_item.child(r, 0), worksheet, row, level + 1)
        return row

    def get_top_subfolders(self):
        root_item = self.model.item(0, 0)
        if not root_item:
            return []

        subfolders = []
        for r in range(root_item.rowCount()):
            name = root_item.child(r, 0).text()
            size_text = root_item.child(r, 1).text().replace(" GB", "")
            try:
                size_value = float(size_text)
                subfolders.append((name, size_value))
            except ValueError:
                continue

        subfolders.sort(key=lambda x: x[1], reverse=True)
        top_data = subfolders[:9]
        other_total = sum(s for _, s in subfolders[9:])
        if other_total > 0:
            top_data.append(("Others", other_total))
        return top_data

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet())
    window = SpaceAudit()
    window.show()
    sys.exit(app.exec())