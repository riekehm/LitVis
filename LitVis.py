import sys, csv, json, io
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QTabWidget, QPushButton, QHBoxLayout, QInputDialog, QMessageBox,
    QFileDialog, QDialog, QLabel, QTextEdit, QComboBox, QColorDialog,
    QListWidget, QStyledItemDelegate, QLineEdit, QToolBar, QStatusBar,
    QUndoStack, QUndoCommand, QAction, QMenu, QToolButton
)
from PyQt5.QtCore import Qt, QTimer, QSignalBlocker, QSize
from PyQt5.QtGui import QFont, QTextDocument, QAbstractTextDocumentLayout
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog

# ---------------------
# RichTextDelegate:
# Renders HTML in cells and applies conditional formatting rules.
# ---------------------
class RichTextDelegate(QStyledItemDelegate):
    def __init__(self, parent=None, rules=None):
        super().__init__(parent)
        self.rules = rules if rules is not None else {}
        self.defaultFont = QFont()  # Will be updated via updateDefaultFont

    def updateDefaultFont(self, newFont):
        self.defaultFont = newFont

    def paint(self, painter, option, index):
        text = index.data() or ""
        # Apply conditional formatting rules:
        for word, color in self.rules.items():
            text = text.replace(word, f'<span style="color: {color};">{word}</span>')
        doc = QTextDocument()
        doc.setHtml(text)
        doc.setDefaultFont(self.defaultFont)
        doc.setTextWidth(option.rect.width())
        painter.save()
        painter.translate(option.rect.topLeft())
        context = QAbstractTextDocumentLayout.PaintContext()
        doc.documentLayout().draw(painter, context)
        painter.restore()

    def sizeHint(self, option, index):
        text = index.data() or ""
        doc = QTextDocument()
        doc.setHtml(text)
        doc.setDefaultFont(self.defaultFont)
        doc.setTextWidth(option.rect.width())
        return doc.size().toSize()



class BulletTextEdit(QTextEdit):
    def keyPressEvent(self, event):
        # Wenn Enter/Return gedrückt wird
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            cursor = self.textCursor()
            currentBlock = cursor.block().text()
            # Falls die aktuelle Zeile mit einem Bullet (•) beginnt
            if currentBlock.strip().startswith("\u2022"):
                super().keyPressEvent(event)
                # Füge nach dem Zeilenumbruch automatisch ein Bullet und ein Leerzeichen ein.
                cursor.insertText("\u2022 ")
                return
        # Standardverhalten ansonsten
        super().keyPressEvent(event)
# ---------------------
# RichEditDialog:
# Dialog for editing and formatting cell content as rich text.
# ---------------------
class RichEditDialog(QDialog):
    def __init__(self, text, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Cell (Rich Text)")
        self.resize(600, 400)
        mainLayout = QVBoxLayout(self)
        self.zoomFactor = 1.0
        self.defaultFont = self.font()

        # Toolbar für Formatierung
        toolbar = QToolBar("Formatting", self)
        boldBtn = QPushButton("Bold", self)
        italicBtn = QPushButton("Italic", self)
        bulletBtn = QPushButton("Bullet", self)
        colorBtn = QPushButton("Color", self)
        sizeCombo = QComboBox(self)
        for size in [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]:
            sizeCombo.addItem(str(size))
        sizeCombo.setCurrentText("12")
        boldBtn.clicked.connect(self.toggleBold)
        italicBtn.clicked.connect(self.toggleItalic)
        bulletBtn.clicked.connect(self.insertBullet)
        colorBtn.clicked.connect(self.changeColor)
        sizeCombo.currentTextChanged.connect(self.changeFontSize)
        toolbar.addWidget(boldBtn)
        toolbar.addWidget(italicBtn)
        toolbar.addWidget(bulletBtn)
        toolbar.addWidget(colorBtn)
        toolbar.addWidget(QLabel("Size:", self))
        toolbar.addWidget(sizeCombo)
        mainLayout.addWidget(toolbar)

        # Verwende jetzt das BulletTextEdit statt dem Standard QTextEdit
        self.textEdit = BulletTextEdit(self)
        self.textEdit.setAcceptRichText(True)
        self.textEdit.setHtml(text)
        mainLayout.addWidget(self.textEdit)

        # OK/Cancel Buttons
        btnLayout = QHBoxLayout()
        okBtn = QPushButton("OK", self)
        cancelBtn = QPushButton("Cancel", self)
        btnLayout.addWidget(okBtn)
        btnLayout.addWidget(cancelBtn)
        mainLayout.addLayout(btnLayout)
        okBtn.clicked.connect(self.accept)
        cancelBtn.clicked.connect(self.reject)

    def toggleBold(self):
        cursor = self.textEdit.textCursor()
        fmt = cursor.charFormat()
        fmt.setFontWeight(QFont.Bold if fmt.fontWeight() != QFont.Bold else QFont.Normal)
        cursor.mergeCharFormat(fmt)

    def toggleItalic(self):
        cursor = self.textEdit.textCursor()
        fmt = cursor.charFormat()
        fmt.setFontItalic(not fmt.fontItalic())
        cursor.mergeCharFormat(fmt)

    def insertBullet(self):
        cursor = self.textEdit.textCursor()
        cursor.insertText("\u2022 ")

    def changeColor(self):
        color = QColorDialog.getColor(parent=self)
        if color.isValid():
            cursor = self.textEdit.textCursor()
            fmt = cursor.charFormat()
            fmt.setForeground(color)
            cursor.mergeCharFormat(fmt)

    def changeFontSize(self, text):
        try:
            size = int(text)
        except ValueError:
            return
        cursor = self.textEdit.textCursor()
        fmt = cursor.charFormat()
        fmt.setFontPointSize(size)
        cursor.mergeCharFormat(fmt)

    def getText(self):
        return self.textEdit.toHtml()


# ---------------------
# ConditionalFormattingDialog:
# Dialog to define conditional formatting rules.
# ---------------------
class ConditionalFormattingDialog(QDialog):
    def __init__(self, rules=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Conditional Formatting")
        self.resize(400, 300)
        self.rules = rules if rules is not None else {}
        mainLayout = QVBoxLayout(self)

        self.listWidget = QListWidget(self)
        self.refreshList()
        mainLayout.addWidget(self.listWidget)

        btnLayout = QHBoxLayout()
        addBtn = QPushButton("Add", self)
        removeBtn = QPushButton("Remove", self)
        changeBtn = QPushButton("Change Color", self)
        btnLayout.addWidget(addBtn)
        btnLayout.addWidget(removeBtn)
        btnLayout.addWidget(changeBtn)
        mainLayout.addLayout(btnLayout)

        addBtn.clicked.connect(self.addRule)
        removeBtn.clicked.connect(self.removeRule)
        changeBtn.clicked.connect(self.changeRuleColor)

        btnLayout2 = QHBoxLayout()
        okBtn = QPushButton("OK", self)
        cancelBtn = QPushButton("Cancel", self)
        btnLayout2.addWidget(okBtn)
        btnLayout2.addWidget(cancelBtn)
        mainLayout.addLayout(btnLayout2)
        okBtn.clicked.connect(self.accept)
        cancelBtn.clicked.connect(self.reject)

    def refreshList(self):
        self.listWidget.clear()
        for word, color in self.rules.items():
            self.listWidget.addItem(f"{word} = {color}")

    def addRule(self):
        word, ok = QInputDialog.getText(self, "New Rule", "Word to highlight:")
        if ok and word:
            self.rules[word] = "blue"  # Default color
            self.refreshList()

    def removeRule(self):
        idx = self.listWidget.currentRow()
        if idx >= 0:
            key = list(self.rules.keys())[idx]
            del self.rules[key]
            self.refreshList()

    def changeRuleColor(self):
        idx = self.listWidget.currentRow()
        if idx >= 0:
            key = list(self.rules.keys())[idx]
            color = QColorDialog.getColor(parent=self)
            if color.isValid():
                self.rules[key] = color.name()
                self.refreshList()

    def getRules(self):
        return self.rules


# ---------------------
# AdvancedFilterDialog:
# Dialog to define advanced filter conditions.
# ---------------------
class AdvancedFilterDialog(QDialog):
    def __init__(self, headers, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Advanced Filter")
        self.resize(400, 300)
        self.headers = headers  # List of column names
        self.conditions = []  # List: (Column, Operator, Value)
        mainLayout = QVBoxLayout(self)

        formLayout = QHBoxLayout()
        self.columnCombo = QComboBox(self)
        self.columnCombo.addItems(self.headers)
        self.operatorCombo = QComboBox(self)
        self.operatorCombo.addItems(["contains", "starts with", "ends with", "equals"])
        self.valueField = QLineEdit(self)
        addCondBtn = QPushButton("Add Condition", self)
        addCondBtn.clicked.connect(self.addCondition)
        formLayout.addWidget(QLabel("Column:", self))
        formLayout.addWidget(self.columnCombo)
        formLayout.addWidget(QLabel("Operator:", self))
        formLayout.addWidget(self.operatorCombo)
        formLayout.addWidget(QLabel("Value:", self))
        formLayout.addWidget(self.valueField)
        formLayout.addWidget(addCondBtn)
        mainLayout.addLayout(formLayout)

        self.conditionsList = QListWidget(self)
        mainLayout.addWidget(self.conditionsList)
        removeCondBtn = QPushButton("Remove Selected Condition", self)
        removeCondBtn.clicked.connect(self.removeCondition)
        mainLayout.addWidget(removeCondBtn)

        combineLayout = QHBoxLayout()
        combineLayout.addWidget(QLabel("Combine conditions with:", self))
        self.combineCombo = QComboBox(self)
        self.combineCombo.addItems(["AND", "OR"])
        combineLayout.addWidget(self.combineCombo)
        mainLayout.addLayout(combineLayout)

        btnLayout = QHBoxLayout()
        okBtn = QPushButton("OK", self)
        cancelBtn = QPushButton("Cancel", self)
        btnLayout.addWidget(okBtn)
        btnLayout.addWidget(cancelBtn)
        mainLayout.addLayout(btnLayout)
        okBtn.clicked.connect(self.accept)
        cancelBtn.clicked.connect(self.reject)

    def addCondition(self):
        column = self.columnCombo.currentText()
        operator = self.operatorCombo.currentText()
        value = self.valueField.text().strip()
        if value:
            condition = (column, operator, value)
            self.conditions.append(condition)
            self.conditionsList.addItem(f"{column} {operator} '{value}'")
            self.valueField.clear()

    def removeCondition(self):
        row = self.conditionsList.currentRow()
        if row >= 0:
            self.conditions.pop(row)
            self.conditionsList.takeItem(row)

    def getFilterConditions(self):
        return {"combine": self.combineCombo.currentText(), "conditions": self.conditions}


# ---------------------
# Helper: Convert HTML to plain text (for CSV export)
# ---------------------
def htmlToPlainText(html):
    doc = QTextDocument()
    doc.setHtml(html)
    return doc.toPlainText()


# ---------------------
# Helper: Check if a cell's text satisfies a filter condition
# ---------------------
def checkCondition(cellText, operator, value):
    cellText = cellText.lower()
    value = value.lower()
    if operator == "contains":
        return value in cellText
    elif operator == "starts with":
        return cellText.startswith(value)
    elif operator == "ends with":
        return cellText.endswith(value)
    elif operator == "equals":
        return cellText == value
    return False


# ---------------------
# MainWindow:
# Main window in an Excel-like layout.
# then a tabbed function area (with horizontal buttons) at the top,
# and the table is always visible in the lower area.
# ---------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LitVis")
        self.resize(950, 750)
        self.zoomFactor = 1.0
        self.defaultFont = self.font()

        # Auto-Save Timer (every 2 minutes)
        self.autoSaveTimer = QTimer(self)
        self.autoSaveTimer.timeout.connect(self.autoSave)
        self.autoSaveTimer.start(120000)

        # Initialize table (example: 3 rows, 2 columns)
        self.table = QTableWidget(2, 2, self)  # Beispiel: 10 Zeilen, 5 Spalten
        self.table.horizontalHeader().setSectionsMovable(True)
        self.table.verticalScrollBar().setSingleStep(10)
        self.table.horizontalScrollBar().setSingleStep(10)
        self.setCentralWidget(self.table)
        self.headers = ["Title", "Author"]
        self.table.setItem(0, 0, QTableWidgetItem("Sample <b>Title</b>"))
        self.table.setItem(0, 1, QTableWidgetItem("Sample <i>Author</i>"))
        self.table.setItem(1, 0, QTableWidgetItem("Another <font color='red'>Title</font>"))
        self.table.setItem(1, 1, QTableWidgetItem("Another <u>Author</u>"))


        # Enable word wrap and sorting for the table
        self.table.setWordWrap(True)
        self.table.setSortingEnabled(True)

        # Delegate for rich text rendering and conditional formatting
        self.delegate = RichTextDelegate(rules={})
        self.table.setItemDelegate(self.delegate)
        self.table.myDelegate = self.delegate  # for zoom updates

        # Enable cell editing on double-click
        self.table.cellDoubleClicked.connect(self.edit_cell)
        self.table.cellChanged.connect(lambda row, col: self.table.resizeRowsToContents())
        self.table.horizontalHeader().sectionResized.connect(lambda idx, old, new: self.table.resizeRowsToContents())
        self.table.resizeRowsToContents()

        # Dictionary for storing collapsed rows
        self.collapsedRows = {}

        # Create a QTabWidget for function buttons (tabs at the top)
        self.tabWidget = QTabWidget(self)
        self.tabWidget.setFixedHeight(80)

        # Tab 1: Basic Functions – horizontal buttons
        self.tabBasis = QWidget()
        layoutBasis = QHBoxLayout()
        btnAddRow = QPushButton("+ Row", self)
        btnDeleteRow = QPushButton("- Row", self)
        btnAddCol = QPushButton("+ Column", self)
        btnDeleteCol = QPushButton("- Column", self)
        btnRenameCol = QPushButton("Rename Column", self)
        for btn in [btnAddRow, btnDeleteRow, btnAddCol, btnDeleteCol, btnRenameCol]:
            layoutBasis.addWidget(btn)
        self.tabBasis.setLayout(layoutBasis)
        self.tabWidget.addTab(self.tabBasis, "Basic Functions")

        # Tab 2: Project/CSV – horizontal buttons
        self.tabProj = QWidget()
        layoutProj = QHBoxLayout()
        btnExportCSV = QPushButton("Export CSV", self)
        btnImportCSV = QPushButton("Import CSV", self)
        btnSaveProj = QPushButton("Save Project", self)
        btnLoadProj = QPushButton("Load Project", self)
        for btn in [btnExportCSV, btnImportCSV, btnSaveProj, btnLoadProj]:
            layoutProj.addWidget(btn)
        self.tabProj.setLayout(layoutProj)
        self.tabWidget.addTab(self.tabProj, "Project/CSV")

        # Tab 3: Other Functions – horizontal buttons
        self.tabMore = QWidget()
        layoutMore = QHBoxLayout()
        btnCF = QPushButton("Conditional Formatting", self)
        btnAdvFilter = QPushButton("Advanced Filter", self)
        btnCollapse = QPushButton("Collapse Row", self)

        # Erstelle einen QToolButton für die Spalten-Sichtbarkeit
        self.columnsButton = QToolButton(self)
        self.columnsButton.setText("Columns")
        self.columnsButton.setPopupMode(QToolButton.InstantPopup)
        self.columnsButton.setMenu(self.createColumnsMenu())

        btnPrint = QPushButton("Print", self)

        for btn in [btnCF, btnAdvFilter,  btnPrint,btnCollapse, self.columnsButton]:
            layoutMore.addWidget(btn)
            btn.setFixedSize(200, 30)  # oder setMinimumSize(120, 40)
        self.tabMore.setLayout(layoutMore)
        self.tabWidget.addTab(self.tabMore, "Other Functions")

        # Connect tab buttons to functions
        btnAddRow.clicked.connect(self.addRow)
        btnDeleteRow.clicked.connect(self.deleteRow)
        btnAddCol.clicked.connect(self.addColumn)
        btnDeleteCol.clicked.connect(self.deleteColumn)
        btnRenameCol.clicked.connect(self.renameColumn)
        btnExportCSV.clicked.connect(self.exportCSV)
        btnImportCSV.clicked.connect(self.importCSV)
        btnSaveProj.clicked.connect(self.saveProject)
        btnLoadProj.clicked.connect(self.loadProject)
        btnCF.clicked.connect(self.openCFDialog)
        btnAdvFilter.clicked.connect(self.advancedFilter)
        btnCollapse.clicked.connect(self.toggleCollapseRow)
        btnPrint.clicked.connect(self.printTable)

        # Status Bar
        self.statusBar = QStatusBar(self)
        self.setStatusBar(self.statusBar)

        # Main layout: Tabs (function area) at top, table below
        mainLayout = QVBoxLayout()
        mainLayout.addWidget(self.tabWidget)
        mainLayout.addWidget(self.table)
        container = QWidget(self)
        container.setLayout(mainLayout)
        self.setCentralWidget(container)

    def edit_cell(self, row, col):
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem("")
            self.table.setItem(row, col, item)
        oldText = item.text()
        dlg = RichEditDialog(oldText, self)
        if dlg.exec_():
            newText = dlg.getText()
            if newText != oldText:
                item.setText(newText)
                self.table.resizeRowToContents(row)

    def addRow(self):
        rowCount = self.table.rowCount()
        self.table.insertRow(rowCount)
        for col in range(self.table.columnCount()):
            self.table.setItem(rowCount, col, QTableWidgetItem(""))
        self.table.resizeRowsToContents()

    def deleteRow(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)
            self.table.resizeRowsToContents()
        else:
            QMessageBox.warning(self, "Delete Row", "No row selected!")

    def addColumn(self):
        colCount = self.table.columnCount()
        newHeader, ok = QInputDialog.getText(self, "Add Column", "Column Header:")
        if not ok or not newHeader:
            newHeader = f"Column {colCount + 1}"
        self.table.insertColumn(colCount)
        self.headers.append(newHeader)
        self.table.setHorizontalHeaderLabels(self.headers)
        for row in range(self.table.rowCount()):
            self.table.setItem(row, colCount, QTableWidgetItem(""))
        self.table.resizeRowsToContents()

    def deleteColumn(self):
        col, ok = QInputDialog.getInt(self, "Delete Column", "Column index (0-based):", 0, 0,
                                      self.table.columnCount() - 1)
        if ok:
            self.table.removeColumn(col)
            self.headers.pop(col)
            self.table.setHorizontalHeaderLabels(self.headers)
            self.table.resizeRowsToContents()

    def renameColumn(self):
        col = self.table.currentColumn()
        # Überprüfe, ob eine gültige Spalte ausgewählt wurde.
        if col < 0 or col >= self.table.columnCount():
            QMessageBox.warning(self, "Rename Column", "No valid column selected!")
            return

        # Header-Item direkt aus der Tabelle abrufen
        headerItem = self.table.horizontalHeaderItem(col)
        if headerItem is None:
            # Falls kein Header-Item vorhanden ist, erstelle ein leeres
            headerItem = QTableWidgetItem("")
            self.table.setHorizontalHeaderItem(col, headerItem)
        currentName = headerItem.text()

        # Frage den neuen Spaltennamen ab
        newName, ok = QInputDialog.getText(self, "Rename Column", "New column name:", text=currentName)
        if ok and newName:
            headerItem.setText(newName)
            # Aktualisiere auch self.headers, falls vorhanden und gültig
            if hasattr(self, 'headers') and col < len(self.headers):
                self.headers[col] = newName

    def exportCSV(self):
        filePath, _ = QFileDialog.getSaveFileName(self, "Export CSV", "", "CSV Files (*.csv)")
        if not filePath:
            return

        try:
            # Header direkt aus der Tabelle auslesen
            headers = []
            for i in range(self.table.columnCount()):
                header_item = self.table.horizontalHeaderItem(i)
                if header_item is not None:
                    headers.append(header_item.text())
                else:
                    headers.append("")

            # Datei öffnen und CSV schreiben
            with open(filePath, "w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file, delimiter=";", quotechar='"', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(headers)

                # Zeilenweise die Daten schreiben
                for row in range(self.table.rowCount()):
                    rowData = []
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        if item:
                            rowData.append(htmlToPlainText(item.text()))
                        else:
                            rowData.append("")
                    writer.writerow(rowData)
            self.statusBar.showMessage("Export successful!", 3000)
        except Exception as e:
            QMessageBox.warning(self, "Export CSV", f"Error exporting CSV:\n{e}")

    def importCSV(self):
        filePath, _ = QFileDialog.getOpenFileName(self, "Import CSV", "", "CSV Files (*.csv)")
        if not filePath:
            return  # No file selected.

        try:
            with open(filePath, "r", newline="", encoding="utf-8-sig") as f:
                content = f.read()

            import io
            file_io = io.StringIO(content)
            # CSV-Reader immer mit ";" als Delimiter und '"' als Quotechar.
            reader = csv.reader(file_io, delimiter=";", quotechar='"')
            rows = list(reader)
            print("DEBUG: CSV rows =", rows)  # Debug-Ausgabe

            if not rows:
                QMessageBox.warning(self, "Import CSV", "CSV file is empty!")
                return

            headers = rows[0]
            data = rows[1:]

            # Deaktiviere das automatische Sortieren während des Imports.
            self.table.setSortingEnabled(False)

            # Tabelle komplett neu aufbauen:
            self.table.clear()
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)
            self.table.setRowCount(len(data))

            for i, row in enumerate(data):
                for j, cell in enumerate(row):
                    self.table.setItem(i, j, QTableWidgetItem(cell))

            self.table.resizeRowsToContents()
            self.table.resizeColumnsToContents()
            self.table.update()
            self.statusBar.showMessage("CSV import successful!", 3000)

            # Sortierung wieder aktivieren (optional)
            self.table.setSortingEnabled(True)

        except Exception as e:
            QMessageBox.warning(self, "Import CSV", f"Error importing CSV:\n{e}")

    def openCFDialog(self):
        dlg = ConditionalFormattingDialog(self.delegate.rules, self)
        if dlg.exec_():
            new_rules = dlg.getRules()
            self.delegate.rules = new_rules
            self.table.viewport().update()

    def advancedFilter(self):
        # Ermittle die aktuellen Header aus der Tabelle:
        currentHeaders = []
        for i in range(self.table.columnCount()):
            headerItem = self.table.horizontalHeaderItem(i)
            if headerItem is not None:
                currentHeaders.append(headerItem.text())
            else:
                currentHeaders.append("")
        dlg = AdvancedFilterDialog(currentHeaders, self)
        if dlg.exec_():
            filterDict = dlg.getFilterConditions()
            combine = filterDict["combine"]
            conditions = filterDict["conditions"]
            for row in range(self.table.rowCount()):
                rowVisible = True
                for condition in conditions:
                    column, operator, value = condition
                    if column in currentHeaders:
                        col = currentHeaders.index(column)
                        item = self.table.item(row, col)
                        cellText = item.text() if item else ""
                        plain = htmlToPlainText(cellText)
                        condResult = checkCondition(plain, operator, value)
                        if combine == "AND":
                            if not condResult:
                                rowVisible = False
                                break
                        else:  # OR
                            if condResult:
                                rowVisible = True
                                break
                            else:
                                rowVisible = False
                self.table.setRowHidden(row, not rowVisible)

    def toggleCollapseRow(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Toggle Collapse Row", "Please select a row!", parent=self)
            return
        if row in self.collapsedRows:
            originalHeight, originalContent = self.collapsedRows.pop(row)
            self.table.setSpan(row, 0, 1, 1)
            for col, text in enumerate(originalContent):
                self.table.setItem(row, col, QTableWidgetItem(text))
            self.table.setRowHeight(row, originalHeight)
        else:
            originalHeight = self.table.rowHeight(row)
            originalContent = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                originalContent.append(item.text() if item else "")
            self.collapsedRows[row] = (originalHeight, originalContent)
            mergedText = originalContent[0]
            self.table.setSpan(row, 0, 1, self.table.columnCount())
            self.table.setItem(row, 0, QTableWidgetItem(mergedText))
            self.table.setRowHeight(row, 20)

    def createColumnsMenu(self):
        """Erstellt ein QMenu mit checkbaren Aktionen für jede Spalte."""
        menu = QMenu("Columns", self)
        for col in range(self.table.columnCount()):
            header_item = self.table.horizontalHeaderItem(col)
            text = header_item.text() if header_item else f"Column {col}"
            action = menu.addAction(text)
            action.setCheckable(True)
            # Wenn die Spalte sichtbar ist, soll das Häkchen gesetzt sein
            action.setChecked(not self.table.isColumnHidden(col))
            # Speichere den Spaltenindex in den Aktionsdaten
            action.setData(col)
            action.toggled.connect(self.toggleColumnVisibility)
        return menu

    def toggleColumnVisibility(self, checked):
        """Schaltet die Sichtbarkeit der Spalte um, basierend auf dem Toggle des Menüs."""
        action = self.sender()
        if action:
            col = action.data()
            # Wenn 'checked' True ist, soll die Spalte sichtbar sein (also nicht versteckt)
            self.table.setColumnHidden(col, not checked)

    def printTable(self):
        # Erzeuge ein QPrinter-Objekt mit hoher Auflösung
        printer = QPrinter(QPrinter.HighResolution)

        # Öffne einen einfachen QPrintDialog
        printDialog = QPrintDialog(printer, self)
        printDialog.setWindowTitle("Print Table")
        if printDialog.exec_() != QPrintDialog.Accepted:
            return  # Abbruch, wenn der Benutzer nicht druckt

        # Erstelle ein QTextDocument und fülle es mit HTML, das die Tabelle repräsentiert
        doc = QTextDocument()
        html = "<html><head><meta charset='utf-8'></head><body>"
        html += "<table border='1' cellspacing='0' cellpadding='2'>"

        # Headerzeile
        headers = []
        for col in range(self.table.columnCount()):
            header_item = self.table.horizontalHeaderItem(col)
            headers.append(header_item.text() if header_item else "")
        html += "<tr>" + "".join(f"<th>{header}</th>" for header in headers) + "</tr>"

        # Datenzeilen
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue  # Überspringe ausgeblendete Zeilen
            html += "<tr>"
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                cellText = item.text() if item else ""
                html += f"<td>{cellText}</td>"
            html += "</tr>"

        html += "</table></body></html>"
        doc.setHtml(html)

        # Drucke das Dokument
        doc.print_(printer)

    def saveProject(self):
        filePath, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "Project Files (*.json)")
        if not filePath:
            return
        try:
            # Ermittele die aktuellen Header direkt aus der Tabelle.
            headers = []
            for i in range(self.table.columnCount()):
                header_item = self.table.horizontalHeaderItem(i)
                headers.append(header_item.text() if header_item is not None else "")

            # Erstelle eine Liste aller Zeilen, wobei jede Zeile eine Liste der Zelltexte ist.
            rows_data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item is not None else "")
                rows_data.append(row_data)

            # Spaltenbreiten erfassen:
            colWidths = []
            for col in range(self.table.columnCount()):
                colWidths.append(self.table.columnWidth(col))

            # Spaltenreihenfolge erfassen:
            columnOrder = []
            headerView = self.table.horizontalHeader()
            for vis in range(self.table.columnCount()):
                logical = headerView.logicalIndex(vis)
                columnOrder.append(logical)

            projectData = {
                "headers": headers,
                "rows": rows_data,
                "columnWidths": colWidths,
                "columnOrder": columnOrder,
                "rules": self.delegate.rules
            }

            with open(filePath, "w", encoding="utf-8") as f:
                json.dump(projectData, f, indent=2)

            self.statusBar.showMessage("Project saved successfully!", 3000)
        except Exception as e:
            QMessageBox.warning(self, "Save Project", f"Error saving project:\n{e}")

    def loadProject(self):
        filePath, _ = QFileDialog.getOpenFileName(self, "Load Project", "", "Project Files (*.json)")
        if not filePath:
            return
        try:
            with open(filePath, "r", encoding="utf-8") as f:
                projectData = json.load(f)

            # Get headers and row data from the JSON file.
            headers = projectData.get("headers", [])
            rows = projectData.get("rows", [])
            self.delegate.rules = projectData.get("rules", {})

            # Completely clear the table: set both row and column count to zero.
            self.table.setRowCount(0)
            self.table.setColumnCount(0)

            # Set the column count and header labels.
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)

            # Spaltenreihenfolge wiederherstellen:
            columnOrder = projectData.get("columnOrder", [])
            if columnOrder:
                headerView = self.table.horizontalHeader()
                # Für jeden Eintrag in der gespeicherten Reihenfolge:
                for newVisualIndex, logicalIndex in enumerate(columnOrder):
                    currentVisualIndex = headerView.visualIndex(logicalIndex)
                    headerView.moveSection(currentVisualIndex, newVisualIndex)

            # Insert rows one by one.
            for rowData in rows:
                currentRow = self.table.rowCount()  # Get the next row index.
                self.table.insertRow(currentRow)
                # For each header (column), check if rowData has a cell.
                for colIndex in range(len(headers)):
                    cellData = rowData[colIndex] if colIndex < len(rowData) else ""
                    self.table.setItem(currentRow, colIndex, QTableWidgetItem(cellData))


            # Spaltenbreiten wiederherstellen:
            colWidths = projectData.get("columnWidths", [])
            for col in range(len(headers)):
                if col < len(colWidths):
                    self.table.setColumnWidth(col, colWidths[col])

            self.columnsButton.setMenu(self.createColumnsMenu())

            self.statusBar.showMessage("Project loaded successfully!", 3000)
        except Exception as e:
            QMessageBox.warning(self, "Load Project", f"Error loading project:\n{e}")

    def autoSave(self):
        try:
            # Header direkt aus der Tabelle ermitteln
            headers = []
            for i in range(self.table.columnCount()):
                headerItem = self.table.horizontalHeaderItem(i)
                headers.append(headerItem.text() if headerItem is not None else "")

            # Alle Zeilen aus der Tabelle auslesen
            rows_data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item is not None else "")
                rows_data.append(row_data)

            projectData = {
                "headers": headers,
                "rows": rows_data,
                "rules": self.delegate.rules
            }

            with open("autosave_project.json", "w", encoding="utf-8") as f:
                json.dump(projectData, f, indent=2)

            self.statusBar.showMessage("Auto-saved project!", 2000)
        except Exception as e:
            self.statusBar.showMessage(f"Auto-save failed: {e}", 2000)


# ---------------------
# Helper: Convert HTML to plain text
# ---------------------
def htmlToPlainText(html):
    doc = QTextDocument()
    doc.setHtml(html)
    return doc.toPlainText()


# ---------------------
# Helper: Check if a cell's text satisfies a filter condition
# ---------------------
def checkCondition(cellText, operator, value):
    cellText = cellText.lower()
    value = value.lower()
    if operator == "contains":
        return value in cellText
    elif operator == "starts with":
        return cellText.startswith(value)
    elif operator == "ends with":
        return cellText.endswith(value)
    elif operator == "equals":
        return cellText == value
    return False


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    autoSaveTimer = QTimer(window)
    autoSaveTimer.timeout.connect(window.autoSave)
    autoSaveTimer.start(120000)  # Every 2 minutes
    sys.exit(app.exec_())
