import sys, textwrap, os, json
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import loader

APPDATA = QStandardPaths.standardLocations(QStandardPaths.AppDataLocation)[0]
APPDATA = os.path.join(APPDATA, "QSpellbook")

WB_DEFAULT_FILENAME = "Spells.xlsx"
CACHE_FILENAME = os.path.join(APPDATA, "spells.json")
TAGS_FILENAME = os.path.join(APPDATA, "tags.json")

PROGRAM_NAME = "QSpellbook"
PROGRAM_AUTHOR = "Ethan Crooks"

app = QApplication(sys.argv)
screen = app.primaryScreen()
SCREEN_H, SCREEN_W = screen.availableGeometry().height(), screen.availableGeometry().width()

DEFAULT_HEIGHT_RATIO = 0.7
MAIN_HEIGHT_RATIO = 0.8
MAIN_WIDTH_RATIO = 0.9

TABLE_MAX_RESIZE_WIDTH = 1200

COLUMN_LONG = 300
COLUMN_MED = 200
COLUMN_SHORT = 100

TABLE_MAX_ROW_HEIGHT = 50
TOOLTIP_WIDTH = 150

TABLEITEM_FLAGS_NOEDIT = Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable
TABLEITEM_FLAGS_EDIT = Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable

def except_hook(cls, exception, traceback):
    print(cls, exception, traceback)
    sys.__excepthook__(cls, exception, traceback)
    sys.exit(1)

def generateClassStr(spell):
    classes = [x for x in spell.classes if spell.classes[x]]
    class_str = ""
    for cls in classes:
        class_str += cls[:3].upper() + " "
    return class_str[:-1]

def pprintClasses(spell):
    classes = [x for x in spell.classes if spell.classes[x]]
    class_str = ""
    for cls in classes:
        class_str += cls + "\n"
    return class_str[:-1]

def generateTagStr(spell, tags):
    if not hash(spell) in tags: return None
    tag_str = ""
    for tag in tags[hash(spell)]:
        tag_stripped = ""
        for char in tag: # Remove punctuation
            if char.isalnum():
                tag_stripped += char.upper()
            if len(tag_stripped) == 3:
                break
        if len(tag_stripped) < 3:
            tag_stripped = tag[:3].upper() # If punction has been removed and the tag is less than 3 chars, keep the punctuation
        tag_str += tag_stripped + " "
    return tag_str[:-1]

def pprintTags(spell, tags):
    if not hash(spell) in tags: return None
    tag_str = ""
    for tag in tags[hash(spell)]:
        tag_str += tag + "\n"
    return tag_str[:-1]

def addLineBreaks(s):
    newStr = textwrap.fill(s, width=TOOLTIP_WIDTH, break_long_words=False,replace_whitespace=False)
    return newStr

def borderLine():
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setFrameShadow(QFrame.Sunken)
    return line

def pprintComp(spell):
    comp_str = ""
    if spell.components['verbal']:
        comp_str += "Verbal\n"
    if spell.components['somantic']:
        comp_str += "Somantic\n"
    if spell.components['material']:
        comp_str += spell.components['material'] + "\n"
    if comp_str[-1:] == "\n": comp_str = comp_str[:-1]
    if comp_str == "": comp_str = None
    return comp_str

class VisibilityBar(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.initUI()
        self.show()

    def initUI(self):
        titleLabel = QLabel("<h1>Visibility</h1>")

        colLeftVBox = QVBoxLayout()
        colRightVBox = QVBoxLayout()

        columns = [x for x in self.parent.spellheaders]
        midpoint = len(columns) // 2 - 1
        for i, col in enumerate(columns):
            columnCheckBox = QCheckBox(col)
            columnCheckBox.col = col
            columnCheckBox.setChecked(self.parent.spellheaders[col]['enabled'])
            if i <= midpoint:
                colLeftVBox.addWidget(columnCheckBox)
            else:
                colRightVBox.addWidget(columnCheckBox)
            columnCheckBox.stateChanged.connect(self.applyFilters)
        colMainHBox = QHBoxLayout()
        colMainHBox.addLayout(colLeftVBox)
        colMainHBox.addLayout(colRightVBox)

        mainVBox = QVBoxLayout()
        mainVBox.addWidget(titleLabel)
        mainVBox.addWidget(borderLine())
        mainVBox.addLayout(colMainHBox)
        mainVBox.addStretch(0)

        self.colLeftVBox = colLeftVBox
        self.colRightVBox = colRightVBox

        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        self.setLayout(mainVBox)

    def applyFilters(self):
        columns = {}
        for widget in range(self.colLeftVBox.count()):
            colCheckBox = self.colLeftVBox.itemAt(widget).widget()
            columns[colCheckBox.col] = colCheckBox.isChecked()
        for widget in range(self.colRightVBox.count()):
            colCheckBox = self.colRightVBox.itemAt(widget).widget()
            columns[colCheckBox.col] = colCheckBox.isChecked()
        for col in columns:
            self.parent.spellheaders[col]['enabled'] = columns[col]
        self.parent.updateTable(self.parent.spells)
        self.parent.resizeTableCols()

class FilterBar(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.initUI()
        self.show()

    def initUI(self):
        titleLabel = QLabel("<h1>Filters</h1>")
        titleLabel.setMinimumWidth(300)

        nameLabel = QLabel("NAME")
        nameLabel.setAlignment(Qt.AlignHCenter)

        nameEdit = QLineEdit()

        classLabel = QLabel("CLASS")
        classLabel.setAlignment(Qt.AlignHCenter)

        classLeftVBox = QVBoxLayout()
        classRightVBox = QVBoxLayout()

        classes = [x for x in self.parent.spellbook.spells[0].classes]
        midpoint = len(classes)//2 - 1
        for i, cls in enumerate(classes):
            classCheckBox = QCheckBox(cls)
            classCheckBox.cls = cls
            if i <= midpoint:
                classLeftVBox.addWidget(classCheckBox)
            else:
                classRightVBox.addWidget(classCheckBox)
            classCheckBox.stateChanged.connect(self.applyFiltersAutoWrapper)
        classMainHBox = QHBoxLayout()
        classMainHBox.addLayout(classLeftVBox)
        classMainHBox.addLayout(classRightVBox)

        classSelectAllButton = QPushButton("Select All")
        classSelectNoneButton = QPushButton("Select None")
        classSelectHBox = QHBoxLayout()
        classSelectHBox.addWidget(classSelectAllButton)
        classSelectHBox.addWidget(classSelectNoneButton)

        levelLabel = QLabel("LEVEL 0")
        levelLabel.setAlignment(Qt.AlignHCenter)
        levelCheckBox = QCheckBox()
        maxLevel = max((x.level for x in self.parent.spellbook.spells))
        levelSlider = QSlider(Qt.Horizontal)
        levelSlider.setMinimum(0)
        levelSlider.setMaximum(maxLevel)
        levelSlider.setEnabled(False)

        minLevelLabel = QLabel("0")
        maxLevelLabel = QLabel(str(maxLevel))

        levelHBox = QHBoxLayout()
        levelHBox.addWidget(levelCheckBox)
        levelHBox.addWidget(minLevelLabel)
        levelHBox.addWidget(levelSlider)
        levelHBox.addWidget(maxLevelLabel)

        applyButton = QPushButton("Update")
        applyButton.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Fixed)
        autoCheckBox = QCheckBox("Auto Update")
        applyHBox = QHBoxLayout()
        applyHBox.addWidget(applyButton)
        applyHBox.addWidget(autoCheckBox)

        mainVBox = QVBoxLayout()
        mainVBox.addWidget(titleLabel)
        mainVBox.addWidget(borderLine())
        mainVBox.addWidget(nameLabel)
        mainVBox.addWidget(nameEdit)
        mainVBox.addWidget(borderLine())
        mainVBox.addWidget(classLabel)
        mainVBox.addLayout(classMainHBox)
        mainVBox.addLayout(classSelectHBox)
        mainVBox.addWidget(borderLine())
        mainVBox.addWidget(levelLabel)
        mainVBox.addLayout(levelHBox)
        mainVBox.addWidget(borderLine())
        mainVBox.addLayout(applyHBox)
        mainVBox.addStretch(0)

        levelCheckBox.stateChanged.connect(lambda state: levelSlider.setEnabled(state))
        levelSlider.valueChanged.connect(lambda value: levelLabel.setText("LEVEL " + str(value)))

        #nameEdit.editingFinished.connect(self.applyFilters)
        nameEdit.textChanged.connect(self.applyFiltersAutoWrapper)
        levelCheckBox.stateChanged.connect(self.applyFiltersAutoWrapper)
        levelSlider.valueChanged.connect(self.applyFiltersAutoWrapper)
        autoCheckBox.stateChanged.connect(self.applyFiltersAutoWrapper)
        autoCheckBox.stateChanged.connect(lambda state: applyButton.setEnabled(not state))

        classSelectAllButton.clicked.connect(lambda: self.classesSetEnabled(True))
        classSelectNoneButton.clicked.connect(lambda: self.classesSetEnabled(False))
        applyButton.clicked.connect(self.applyFilters)

        self.nameEdit = nameEdit
        self.classLeftVBox = classLeftVBox
        self.classRightVBox = classRightVBox
        self.levelCheckBox = levelCheckBox
        self.levelSlider = levelSlider
        self.autoCheckBox = autoCheckBox

        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.MinimumExpanding)
        self.setLayout(mainVBox)

    def applyFiltersAutoWrapper(self):
        if self.autoCheckBox.isChecked():
            self.applyFilters()

    def collectClasses(self):
        classes = []
        for widget in range(self.classLeftVBox.count()):
            classCheckBox = self.classLeftVBox.itemAt(widget).widget()
            if classCheckBox.isChecked():
                classes.append(classCheckBox.cls)
        for widget in range(self.classRightVBox.count()):
            classCheckBox = self.classRightVBox.itemAt(widget).widget()
            if classCheckBox.isChecked():
                classes.append(classCheckBox.cls)
        return classes

    def classesSetEnabled(self, enabled):
        enabled = 2 if enabled else 0
        for widget in range(self.classLeftVBox.count()):
            classCheckBox = self.classLeftVBox.itemAt(widget).widget()
            classCheckBox.setCheckState(enabled)
        for widget in range(self.classRightVBox.count()):
            classCheckBox = self.classRightVBox.itemAt(widget).widget()
            classCheckBox.setCheckState(enabled)

    def applyFilters(self):
        conditions = []
        if self.nameEdit.text().strip() != "":
            conditions.append(lambda x: self.nameEdit.text().strip().lower() in x.name.lower())
        if self.levelCheckBox.isChecked():
            conditions.append(lambda x: x.level == self.levelSlider.value())
        classes = self.collectClasses()
        for cls in classes:
            conditions.append(lambda x, cls=cls: x.classes[cls])
        self.parent.filterCondition = lambda spell: all([condition(spell) for condition in conditions])
        self.parent.applyFilters()

class TagBar(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.allTags = self.parent.tags
        self.initUI()

    def initUI(self):
        titleLabel = QLabel("<h1>Tags</h1>")

        tagMainHBox = self.generateTagBox()

        mainVBox = QVBoxLayout()
        mainVBox.addWidget(titleLabel)
        mainVBox.addWidget(borderLine())
        mainVBox.addWidget(tagMainHBox)
        mainVBox.addStretch(0)

        self.tagMainHBox = tagMainHBox
        self.mainVBox = mainVBox

        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        self.setLayout(mainVBox)

    def reupTagBox(self):
        self.allTags = self.parent.tags
        self.mainVBox.itemAt(2).widget().hide()
        self.mainVBox.itemAt(2).widget().setParent(None)
        self.tagMainHBox = self.generateTagBox()
        self.mainVBox.insertWidget(2, self.tagMainHBox)

    def generateTagBox(self):
        tagLeftVBox = QVBoxLayout()
        tagRightVBox = QVBoxLayout()

        tags = {self.allTags[x][y] for x in self.allTags for y in range(len(self.allTags[x]))}
        if tags:
            midpoint = len(tags) // 2 - 1
            for i, tag in enumerate(tags):
                tagCheckBox = QCheckBox(tag)
                tagCheckBox.tag = tag
                if i <= midpoint:
                    tagLeftVBox.addWidget(tagCheckBox)
                else:
                    tagRightVBox.addWidget(tagCheckBox)
                tagCheckBox.stateChanged.connect(self.applyFilters)

            tagMainHBox = QHBoxLayout()
            tagMainHBox.addLayout(tagLeftVBox)
            tagMainHBox.addLayout(tagRightVBox)

            self.tagLeftVBox = tagLeftVBox
            self.tagRightVBox = tagRightVBox

            widgetWrapper = QWidget()
            widgetWrapper.setLayout(tagMainHBox)
            widgetWrapper.setContentsMargins(QMargins(0, 0, 0, 0))
            return widgetWrapper
        else:
            return QLabel("There are no tags")

    def applyFilters(self):
        tags = []
        for widget in range(self.tagLeftVBox.count()):
            tagCheckBox = self.tagLeftVBox.itemAt(widget).widget()
            if tagCheckBox.isChecked(): tags.append(tagCheckBox.tag)
        for widget in range(self.tagRightVBox.count()):
            tagCheckBox = self.tagRightVBox.itemAt(widget).widget()
            if tagCheckBox.isChecked(): tags.append(tagCheckBox.tag)

        #  - There are no tags selected, or
        #   - The spell is tagged
        #   - Every tag that is in tags is also in self.allTag[hash(spell)]
        # There is probably a better way to represent the last one in a lambda function
        # Lol look at this fucking abomination
        self.parent.tagCondition = lambda spell, allTags=self.allTags, tags=tags: \
            tags == [] or (
            hash(spell) in allTags and
            [tag for tag in tags if tag in self.allTags[hash(spell)]] == tags)
        self.parent.applyFilters()

class TagDialog(QDialog): # If remove=False, adding a tag. If remove=True, removing a tag
    def __init__(self, tags, selectedSpell=None, remove=False, bulk=False):
        super().__init__()
        self.selectedSpell = selectedSpell
        self.tags = tags
        self.tag = None
        self.remove = remove
        self.bulk = bulk
        self.initUI()

    def getAllTags(self):
        return {self.tags[x][y] for x in self.tags for y in range(len(self.tags[x]))}

    def initUI(self):
        tagsList = QListWidget()
        if self.remove:
            titleLabel = QLabel("Remove a tag")
            if self.bulk:
                tagsList.addItems(self.getAllTags())
            else:
                tagsList.addItems(self.tags[hash(self.selectedSpell)])
            self.setWindowTitle("Remove a Tag")
        else:
            titleLabel = QLabel("Add a tag")
            tagsList.addItems(self.getAllTags())
            tagsList.addItem("Add New...")
            self.setWindowTitle("Add a Tag")

        doneButton = QPushButton("Done")
        cancelButton = QPushButton("Cancel")
        buttonHBox = QHBoxLayout()
        buttonHBox.addWidget(doneButton)
        buttonHBox.addWidget(cancelButton)

        mainVBox = QVBoxLayout()
        mainVBox.addWidget(titleLabel)
        mainVBox.addWidget(tagsList)
        mainVBox.addLayout(buttonHBox)

        self.tagsList = tagsList

        doneButton.clicked.connect(self.closeDialog)
        cancelButton.clicked.connect(self.reject)
        tagsList.setCurrentRow(tagsList.count()-1)

        self.setLayout(mainVBox)
        self.setWindowModality(Qt.ApplicationModal)
        self.show()

    def closeDialog(self):
        self.tag = self.tagsList.currentItem().text()
        self.accept()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initDataFiles()
        self.settings = QSettings(PROGRAM_AUTHOR, PROGRAM_NAME)
        self.spellspreadsheet = self.settings.value("spreadsheet", WB_DEFAULT_FILENAME)
        if not os.path.isfile(self.spellspreadsheet):
            #QMessageBox.information(self, "Select Spellbook","Please select your Excel spreadsheet spellbook.")
            result = self.setSpellbook()
            if not result: sys.exit(1)
        if os.path.isfile(CACHE_FILENAME):
            self.spellbook = loader.Spellbook.from_cache(CACHE_FILENAME)
        else:
            self.spellbook = loader.Spellbook.from_workbook(self.spellspreadsheet)
            self.spellbook.to_cache(CACHE_FILENAME)
        self.filterCondition = lambda spell: True
        self.tagCondition = lambda spell: True
        self.spells = []
        self.tags = {}
        self.initUI()
        self.initDockWidgets()
        self.initMenu()
        self.initStatusBar()
        self.restoreTags()
        self.show()

    def initUI(self):
        self.spellheaders = {
            "Name": {
                "value": lambda spell: spell.name,
                "tooltip": None,
                "size": COLUMN_MED,
                "enabled": True
            },
            "Level": {
                "value": lambda spell: spell.level,
                "tooltip": None,
                "size": COLUMN_SHORT,
                "enabled": True
            },
            "Classes": {
                "value": lambda spell: generateClassStr(spell),
                "tooltip": lambda spell: pprintClasses(spell),
                "size": COLUMN_SHORT,
                "enabled": True
            },
            "Origin": {
                "value": lambda spell: spell.origin,
                "tooltip": None,
                "size": COLUMN_SHORT,
                "enabled": True
            },
            "School": {
                "value": lambda spell: spell.school,
                "tooltip": None,
                "size": COLUMN_SHORT,
                "enabled": True
            },
            "Ritual": {
                "value": lambda spell: "Yes" if spell.ritual else "No",
                "tooltip": None,
                "size": COLUMN_SHORT,
                "enabled": False
            },
            "Time": {
                "value": lambda spell: spell.time,
                "tooltip": None,
                "size": COLUMN_SHORT,
                "enabled": True
            },
            "Range": {
                "value": lambda spell: spell.range,
                "tooltip": None,
                "size": COLUMN_SHORT,
                "enabled": True
            },
            "Comp": {
                "value": lambda spell: spell.compstr,
                "tooltip": lambda spell: pprintComp(spell),
                "size": COLUMN_SHORT,
                "enabled": False
            },
            "Duration": {
                "value": lambda spell: spell.duration,
                "tooltip": None,
                "size": COLUMN_SHORT,
                "enabled": True
            },
            "Tag": {
                "value": lambda spell: generateTagStr(spell, self.tags),
                "tooltip": lambda spell: pprintTags(spell, self.tags),
                "size": COLUMN_SHORT,
                "enabled": True
            },
            "Description": {
                "value": lambda spell: spell.description.replace("\n", " ") if spell.description else None,
                "tooltip": lambda spell: addLineBreaks(str(spell.description)),
                "size": COLUMN_LONG,
                "enabled": True
            }
        }

        table = QTableWidget(1, len(self.spellheaders))
        table.setHorizontalHeaderLabels(list(self.spellheaders.keys()))
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.NoSelection)
        table.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        table.setWordWrap(True)
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(self.showTableContextMenu)

        self.table = table
        self.table.sortByColumn(0, Qt.AscendingOrder)

        self.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.MinimumExpanding)
        self.setWindowTitle("QSpellbook")
        self.setCentralWidget(table)

    def setSpellbook(self):
        dialog = QFileDialog()
        dialog.setWindowTitle("Select Excel Spellbook")
        dialog.setAcceptMode(QFileDialog.AcceptOpen)
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setNameFilter("*.xlsx")
        dialog.setDirectory(os.getcwd())
        if dialog.exec() and len(dialog.selectedFiles()) > 0:
            filepath = dialog.selectedFiles()[0]
            filepath = os.path.relpath(filepath)
            if os.path.exists(filepath):
                self.spellspreadsheet = filepath
                self.settings.setValue("spreadsheet", filepath)
                return True
        return False

    def reloadSpellbook(self):
        if not os.path.exists(self.spellspreadsheet):
            QMessageBox.critical(self, "Reload Error", "The currently loaded spreadsheet no longer exists.\nPlease select a new spreadsheet.")
            self.setSpellbook()
        self.filterCondition = lambda spell: True
        self.tagCondition = lambda spell: True
        os.remove(CACHE_FILENAME)
        self.spellbook = loader.Spellbook.from_workbook(self.spellspreadsheet)
        self.spellbook.to_cache(CACHE_FILENAME)
        self.updateTable(self.spellbook.spells)
        self.dirLabel.setText(self.spellspreadsheet + " ") # Space for padding
        self.resizeTableRows()
        self.resizeTableCols()

    def reloadFromFileWrapper(self):
        result = self.setSpellbook()
        if result: self.reloadSpellbook()

    def initDataFiles(self):
        head, tail = os.path.split(CACHE_FILENAME)
        if head and not os.path.isdir(head): os.makedirs(head)
        head, tail = os.path.split(TAGS_FILENAME)
        if head and not os.path.isdir(head): os.makedirs(head)
        if not os.path.isfile(TAGS_FILENAME):
            with open(TAGS_FILENAME, "w") as f:
                f.write("{}")

    def saveTags(self):
        with open(TAGS_FILENAME, "w") as f:
            f.write(json.dumps(self.tags))

    def restoreTags(self):
        with open(TAGS_FILENAME) as f:
            data = json.loads(f.read())
            self.tags = {int(key):data[key] for key in data}
        self.tagBar.widget().reupTagBox()
        self.updateTable(self.spells)
        self.resizeTableCols()

    def applyFilters(self):
        mainCondition = lambda spell: self.filterCondition(spell) and self.tagCondition(spell)
        spells = self.spellbook.search(mainCondition)
        self.updateTable(spells)
        self.resizeTableCols()
        self.resizeTableRows()

    def addTag(self, row=None, bulk=False):
        spell = self.spells[row] if not bulk else None
        dialog = TagDialog(self.tags, spell, remove=False, bulk=bulk)
        if dialog.exec():
            tag = dialog.tag
            if tag == "Add New...":
                tag, state = QInputDialog.getText(self, "New Tag", "Enter a new tag:", QLineEdit.Normal, "")
                if not state: return
            if bulk:
                for spell in self.spells: # Loop through and tag all spells rather than selected spell
                    if hash(spell) in self.tags:
                        self.tags[hash(spell)].append(tag)
                    else:
                        self.tags[hash(spell)] = [tag]
            else:
                if hash(spell) in self.tags:
                    self.tags[hash(spell)].append(tag)
                else:
                    self.tags[hash(spell)] = [tag]
            self.updateTable(self.spells)
            self.resizeTableCols()
            self.tagBar.widget().reupTagBox()
            self.saveTags()

    def removeTag(self, row=None, bulk=False):
        spell = self.spells[row] if not bulk else None
        dialog = TagDialog(self.tags, spell, remove=True, bulk=bulk)
        if dialog.exec():
            tag = dialog.tag
            if bulk:
                for spell in self.spells: # Loop through and untag all spells rather than selected spell
                    if hash(spell) in self.tags and tag in self.tags[hash(spell)]:
                        self.tags[hash(spell)].remove(tag)
                        if self.tags[hash(spell)] == []:
                            self.tags.pop(hash(spell))
            else:
                self.tags[hash(spell)].remove(tag)
                if self.tags[hash(spell)] == []:
                    self.tags.pop(hash(spell))
            self.updateTable(self.spells)
            self.resizeTableCols()
            self.tagBar.widget().reupTagBox()
            self.saveTags()

    def wipeTags(self, visible=False):
        if visible:
            for spell in self.spells:
                if hash(spell) in self.tags:
                    self.tags.pop(hash(spell))
        else:
            self.tags = {}
        self.updateTable(self.spells)
        self.resizeTableCols()
        self.tagBar.widget().reupTagBox()
        self.saveTags()

    def initMenu(self):
        menuBar = self.menuBar()

        fileMenu = menuBar.addMenu("&File")
        openNewAction = fileMenu.addAction("&Open New File")
        reloadAction = fileMenu.addAction("&Reload Current File")
        quitAction = fileMenu.addAction("&Quit")

        tagMenu = menuBar.addMenu("&Tags")
        saveTagsAction = tagMenu.addAction("&Save Tags")
        restoreTagsAction = tagMenu.addAction("&Restore Tags")
        tagMenu.addSeparator()
        tagAllVisibleAction = tagMenu.addAction("&Tag All Visible")
        untagAllVisibleAction = tagMenu.addAction("&Untag All Visible")
        tagMenu.addSeparator()
        wipeTagsAction = tagMenu.addAction("&Wipe All Tags")

        windowMenu = menuBar.addMenu("&Window")
        stateSaveAction = windowMenu.addAction("&Save Position")
        stateRestoreAction = windowMenu.addAction("&Restore Position")
        windowMenu.addSeparator()
        filterBarAction = windowMenu.addAction("&Filters")
        filterBarAction.setCheckable(True)
        visBarAction = windowMenu.addAction("&Visibility")
        visBarAction.setCheckable(True)
        tagBarAction = windowMenu.addAction("&Tags")
        tagBarAction.setCheckable(True)

        openNewAction.triggered.connect(self.reloadFromFileWrapper)
        reloadAction.triggered.connect(self.reloadSpellbook)
        quitAction.triggered.connect(lambda: app.exit(0))

        saveTagsAction.triggered.connect(self.saveTags)
        restoreTagsAction.triggered.connect(self.restoreTags)
        tagAllVisibleAction.triggered.connect(lambda: self.addTag(bulk=True))
        untagAllVisibleAction.triggered.connect(lambda: self.removeTag(bulk=True))
        wipeTagsAction.triggered.connect(lambda: self.wipeTags(visible=False))

        stateSaveAction.triggered.connect(self.save)
        stateRestoreAction.triggered.connect(self.restore)

        filterBarAction.changed.connect(lambda: self.filterBar.setHidden(not filterBarAction.isChecked()))
        self.filterBar.visibilityChanged.connect(lambda: filterBarAction.setChecked(not self.filterBar.isHidden()))
        filterBarAction.setChecked(not self.filterBar.isHidden())

        visBarAction.changed.connect(lambda: self.visBar.setHidden(not visBarAction.isChecked()))
        self.visBar.visibilityChanged.connect(lambda: visBarAction.setChecked(not self.visBar.isHidden()))
        visBarAction.setChecked(not self.visBar.isHidden())

        tagBarAction.changed.connect(lambda: self.tagBar.setHidden(not tagBarAction.isChecked()))
        self.tagBar.visibilityChanged.connect(lambda: tagBarAction.setChecked(not self.tagBar.isHidden()))
        tagBarAction.setChecked(not self.tagBar.isHidden())

    def showTableContextMenu(self, pos):
        item = self.table.itemAt(pos)
        if not item: return
        row = item.unsortedRow
        print(row)
        contextMenu = QMenu()

        addTagAction = contextMenu.addAction("&Add Tag")
        addTagAction.triggered.connect(lambda: self.addTag(row))
        if hash(self.spells[row]) in self.tags:
            removeTagAction = contextMenu.addAction("&Remove Tag")
            removeTagAction.triggered.connect(lambda: self.removeTag(row))

        contextMenu.exec(QCursor.pos())

    def closeEvent(self, *args, **kwargs):
        self.save()
        return super().closeEvent(*args, **kwargs)

    def save(self):
        #self.settings.setValue("dockState", self.saveState())
        self.settings.setValue("geometryState", self.saveGeometry())

    def restore(self):
        #stateData = self.settings.value("dockState", None)
        geometryData = self.settings.value("geometryState", None)
        #if stateData:
        #    self.restoreState(stateData)
        if geometryData:
            self.restoreGeometry(geometryData)

    def initStatusBar(self):
        statusBar = self.statusBar()
        dirLabel = QLabel(self.spellspreadsheet + " ") # Space for padding
        countLabel = QLabel("Count: 0")
        statusBar.addPermanentWidget(dirLabel)
        statusBar.addPermanentWidget(countLabel)
        #statusBar.setSizeGripEnabled(False)
        self.dirLabel = dirLabel
        self.countLabel = countLabel

    def initDockWidgets(self):
        self.setDockOptions(
            QMainWindow.AnimatedDocks | QMainWindow.AllowNestedDocks
        )

        scrollArea = QScrollArea()
        scrollArea.setWidget(FilterBar(self))
        scrollArea.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scrollArea.setWidgetResizable(True)

        filterBar = QDockWidget("Filters")
        filterBar.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        filterBar.setWidget(scrollArea)

        visBar = QDockWidget("Visibility")
        visBar.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        visBar.setWidget(VisibilityBar(self))

        tagBar = QDockWidget("Tags")
        tagBar.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        tagBar.setWidget(TagBar(self))

        self.filterBar = filterBar
        self.visBar = visBar
        self.tagBar = tagBar

        self.addDockWidget(Qt.LeftDockWidgetArea, filterBar)
        self.addDockWidget(Qt.LeftDockWidgetArea, visBar)
        self.addDockWidget(Qt.RightDockWidgetArea, tagBar)

    def updateTable(self, spells):
        self.spells = spells
        self.table.setSortingEnabled(False)
        self.table.clearContents()
        spellheaders = {x:self.spellheaders[x] for x in self.spellheaders if self.spellheaders[x]['enabled']}
        self.table.setColumnCount(len(spellheaders))
        self.table.setHorizontalHeaderLabels(spellheaders.keys())
        self.table.setRowCount(len(spells))
        for x, spell in enumerate(spells):
            row = [x for x in spellheaders.values()]
            for y, cell in enumerate(row):
                item = QTableWidgetItem(str(row[y]['value'](spell)))
                item.setFlags(TABLEITEM_FLAGS_NOEDIT)
                item.unsortedRow = x
                if row[y]["tooltip"] != None:
                    item.setToolTip(row[y]['tooltip'](spell))
                self.table.setItem(x, y, item)
        self.countLabel.setText("Count: "+str(len(spells)))
        self.table.setSortingEnabled(True)
        #self.table.sortByColumn(0, Qt.AscendingOrder)

    def resizeTableCols(self, resizeTable=False):
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(self.table.columnCount() - 1, QHeaderView.Stretch)
        totalSize = 0
        self.table.resizeColumnsToContents()
        #for col in range(self.table.columnCount()):
        #    self.table.setColumnWidth(col, self.table.sizeHintForColumn(col))
        spellheaders = {x: self.spellheaders[x] for x in self.spellheaders if self.spellheaders[x]['enabled']}
        for col in range(self.table.columnCount()):
            key = list(spellheaders.keys())[col]
            if self.table.columnWidth(col) > spellheaders[key]['size']:
                self.table.setColumnWidth(col, spellheaders[key]['size'])
            totalSize += self.table.columnWidth(col)
        totalSize += self.table.verticalScrollBar().width()
        totalSize += self.table.verticalHeader().width()
        if resizeTable:
            self.table.resize(max(totalSize, self.table.width()), self.table.height())

    def resizeTableRows(self):
        self.table.resizeRowsToContents()
        for row in range(self.table.rowCount()):
            if self.table.rowHeight(row) >= TABLE_MAX_ROW_HEIGHT:
                self.table.setRowHeight(row, TABLE_MAX_ROW_HEIGHT)

    def layoutCleanup(self):
        self.updateTable(self.spellbook.spells)
        self.resizeTableCols(True)
        self.resizeTableRows()
        self.resize(min(self.table.width() + COLUMN_LONG, SCREEN_W*MAIN_HEIGHT_RATIO), min(SCREEN_H*DEFAULT_HEIGHT_RATIO, SCREEN_H*MAIN_WIDTH_RATIO))
        self.restore()

def main():
    sys.excepthook = except_hook
    #sys.setprofile(tracefunc)
    win = MainWindow()
    win.layoutCleanup()
    sys.exit(app.exec_())

if __name__ == "__main__": main()