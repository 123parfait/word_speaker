# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'main_window.ui'
##
## Created by: Qt User Interface Compiler version 6.10.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QFrame, QGridLayout,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMenuBar,
    QPushButton, QSizePolicy, QSpacerItem, QSplitter,
    QStatusBar, QTabWidget, QTableWidget, QTableWidgetItem,
    QTextEdit, QVBoxLayout, QWidget)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(1480, 860)
        MainWindow.setMinimumSize(QSize(1180, 720))
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.rootLayout = QVBoxLayout(self.centralwidget)
        self.rootLayout.setSpacing(12)
        self.rootLayout.setObjectName(u"rootLayout")
        self.rootLayout.setContentsMargins(18, 18, 18, 18)
        self.headerCard = QFrame(self.centralwidget)
        self.headerCard.setObjectName(u"headerCard")
        self.headerGrid = QGridLayout(self.headerCard)
        self.headerGrid.setObjectName(u"headerGrid")
        self.headerGrid.setHorizontalSpacing(16)
        self.headerGrid.setContentsMargins(20, 18, 20, 18)
        self.titleLabel = QLabel(self.headerCard)
        self.titleLabel.setObjectName(u"titleLabel")

        self.headerGrid.addWidget(self.titleLabel, 0, 0, 1, 2)

        self.subtitleLabel = QLabel(self.headerCard)
        self.subtitleLabel.setObjectName(u"subtitleLabel")
        self.subtitleLabel.setWordWrap(True)

        self.headerGrid.addWidget(self.subtitleLabel, 1, 0, 1, 2)

        self.currentFileStaticLabel = QLabel(self.headerCard)
        self.currentFileStaticLabel.setObjectName(u"currentFileStaticLabel")

        self.headerGrid.addWidget(self.currentFileStaticLabel, 0, 2, 1, 1)

        self.currentFileValueLabel = QLabel(self.headerCard)
        self.currentFileValueLabel.setObjectName(u"currentFileValueLabel")
        self.currentFileValueLabel.setWordWrap(True)

        self.headerGrid.addWidget(self.currentFileValueLabel, 1, 2, 1, 1)

        self.audioBackendStaticLabel = QLabel(self.headerCard)
        self.audioBackendStaticLabel.setObjectName(u"audioBackendStaticLabel")

        self.headerGrid.addWidget(self.audioBackendStaticLabel, 0, 3, 1, 1)

        self.audioBackendValueLabel = QLabel(self.headerCard)
        self.audioBackendValueLabel.setObjectName(u"audioBackendValueLabel")
        self.audioBackendValueLabel.setWordWrap(True)

        self.headerGrid.addWidget(self.audioBackendValueLabel, 1, 3, 1, 1)


        self.rootLayout.addWidget(self.headerCard)

        self.headerSeparator = QFrame(self.centralwidget)
        self.headerSeparator.setObjectName(u"headerSeparator")
        self.headerSeparator.setFrameShape(QFrame.HLine)
        self.headerSeparator.setFrameShadow(QFrame.Sunken)

        self.rootLayout.addWidget(self.headerSeparator)

        self.mainSplitter = QSplitter(self.centralwidget)
        self.mainSplitter.setObjectName(u"mainSplitter")
        self.mainSplitter.setOrientation(Qt.Horizontal)
        self.leftCard = QFrame(self.mainSplitter)
        self.leftCard.setObjectName(u"leftCard")
        self.leftLayout = QVBoxLayout(self.leftCard)
        self.leftLayout.setSpacing(10)
        self.leftLayout.setObjectName(u"leftLayout")
        self.leftLayout.setContentsMargins(16, 16, 16, 16)
        self.wordListTitleLabel = QLabel(self.leftCard)
        self.wordListTitleLabel.setObjectName(u"wordListTitleLabel")

        self.leftLayout.addWidget(self.wordListTitleLabel)

        self.wordListDescLabel = QLabel(self.leftCard)
        self.wordListDescLabel.setObjectName(u"wordListDescLabel")
        self.wordListDescLabel.setWordWrap(True)

        self.leftLayout.addWidget(self.wordListDescLabel)

        self.topButtonRow = QHBoxLayout()
        self.topButtonRow.setSpacing(8)
        self.topButtonRow.setObjectName(u"topButtonRow")
        self.importButton = QPushButton(self.leftCard)
        self.importButton.setObjectName(u"importButton")

        self.topButtonRow.addWidget(self.importButton)

        self.manualInputButton = QPushButton(self.leftCard)
        self.manualInputButton.setObjectName(u"manualInputButton")

        self.topButtonRow.addWidget(self.manualInputButton)

        self.saveAsButton = QPushButton(self.leftCard)
        self.saveAsButton.setObjectName(u"saveAsButton")

        self.topButtonRow.addWidget(self.saveAsButton)

        self.newListButton = QPushButton(self.leftCard)
        self.newListButton.setObjectName(u"newListButton")

        self.topButtonRow.addWidget(self.newListButton)


        self.leftLayout.addLayout(self.topButtonRow)

        self.wordTable = QTableWidget(self.leftCard)
        if (self.wordTable.columnCount() < 3):
            self.wordTable.setColumnCount(3)
        __qtablewidgetitem = QTableWidgetItem()
        self.wordTable.setHorizontalHeaderItem(0, __qtablewidgetitem)
        __qtablewidgetitem1 = QTableWidgetItem()
        self.wordTable.setHorizontalHeaderItem(1, __qtablewidgetitem1)
        __qtablewidgetitem2 = QTableWidgetItem()
        self.wordTable.setHorizontalHeaderItem(2, __qtablewidgetitem2)
        self.wordTable.setObjectName(u"wordTable")
        self.wordTable.setAlternatingRowColors(True)
        self.wordTable.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.wordTable.setSelectionMode(QAbstractItemView.SingleSelection)

        self.leftLayout.addWidget(self.wordTable)

        self.emptyStateLabel = QLabel(self.leftCard)
        self.emptyStateLabel.setObjectName(u"emptyStateLabel")
        self.emptyStateLabel.setWordWrap(True)

        self.leftLayout.addWidget(self.emptyStateLabel)

        self.playerButtonRow = QHBoxLayout()
        self.playerButtonRow.setSpacing(8)
        self.playerButtonRow.setObjectName(u"playerButtonRow")
        self.playButton = QPushButton(self.leftCard)
        self.playButton.setObjectName(u"playButton")

        self.playerButtonRow.addWidget(self.playButton)

        self.settingsButton = QPushButton(self.leftCard)
        self.settingsButton.setObjectName(u"settingsButton")

        self.playerButtonRow.addWidget(self.settingsButton)

        self.dictationButton = QPushButton(self.leftCard)
        self.dictationButton.setObjectName(u"dictationButton")

        self.playerButtonRow.addWidget(self.dictationButton)


        self.leftLayout.addLayout(self.playerButtonRow)

        self.statusLabel = QLabel(self.leftCard)
        self.statusLabel.setObjectName(u"statusLabel")
        self.statusLabel.setWordWrap(True)

        self.leftLayout.addWidget(self.statusLabel)

        self.mainSplitter.addWidget(self.leftCard)
        self.rightPanel = QWidget(self.mainSplitter)
        self.rightPanel.setObjectName(u"rightPanel")
        self.rightLayout = QVBoxLayout(self.rightPanel)
        self.rightLayout.setSpacing(12)
        self.rightLayout.setObjectName(u"rightLayout")
        self.rightLayout.setContentsMargins(0, 0, 0, 0)
        self.detailCard = QFrame(self.rightPanel)
        self.detailCard.setObjectName(u"detailCard")
        self.detailLayout = QVBoxLayout(self.detailCard)
        self.detailLayout.setSpacing(10)
        self.detailLayout.setObjectName(u"detailLayout")
        self.detailLayout.setContentsMargins(16, 16, 16, 16)
        self.currentWordTitleLabel = QLabel(self.detailCard)
        self.currentWordTitleLabel.setObjectName(u"currentWordTitleLabel")

        self.detailLayout.addWidget(self.currentWordTitleLabel)

        self.selectedWordEdit = QLineEdit(self.detailCard)
        self.selectedWordEdit.setObjectName(u"selectedWordEdit")

        self.detailLayout.addWidget(self.selectedWordEdit)

        self.selectedNoteEdit = QTextEdit(self.detailCard)
        self.selectedNoteEdit.setObjectName(u"selectedNoteEdit")

        self.detailLayout.addWidget(self.selectedNoteEdit)

        self.statsGrid = QGridLayout()
        self.statsGrid.setObjectName(u"statsGrid")
        self.wrongCountStaticLabel = QLabel(self.detailCard)
        self.wrongCountStaticLabel.setObjectName(u"wrongCountStaticLabel")

        self.statsGrid.addWidget(self.wrongCountStaticLabel, 0, 0, 1, 1)

        self.wrongCountValueLabel = QLabel(self.detailCard)
        self.wrongCountValueLabel.setObjectName(u"wrongCountValueLabel")

        self.statsGrid.addWidget(self.wrongCountValueLabel, 0, 1, 1, 1)

        self.correctCountStaticLabel = QLabel(self.detailCard)
        self.correctCountStaticLabel.setObjectName(u"correctCountStaticLabel")

        self.statsGrid.addWidget(self.correctCountStaticLabel, 1, 0, 1, 1)

        self.correctCountValueLabel = QLabel(self.detailCard)
        self.correctCountValueLabel.setObjectName(u"correctCountValueLabel")

        self.statsGrid.addWidget(self.correctCountValueLabel, 1, 1, 1, 1)

        self.lastResultStaticLabel = QLabel(self.detailCard)
        self.lastResultStaticLabel.setObjectName(u"lastResultStaticLabel")

        self.statsGrid.addWidget(self.lastResultStaticLabel, 2, 0, 1, 1)

        self.lastResultValueLabel = QLabel(self.detailCard)
        self.lastResultValueLabel.setObjectName(u"lastResultValueLabel")

        self.statsGrid.addWidget(self.lastResultValueLabel, 2, 1, 1, 1)

        self.lastSeenStaticLabel = QLabel(self.detailCard)
        self.lastSeenStaticLabel.setObjectName(u"lastSeenStaticLabel")

        self.statsGrid.addWidget(self.lastSeenStaticLabel, 3, 0, 1, 1)

        self.lastSeenValueLabel = QLabel(self.detailCard)
        self.lastSeenValueLabel.setObjectName(u"lastSeenValueLabel")

        self.statsGrid.addWidget(self.lastSeenValueLabel, 3, 1, 1, 1)


        self.detailLayout.addLayout(self.statsGrid)

        self.detailActionGrid = QGridLayout()
        self.detailActionGrid.setObjectName(u"detailActionGrid")
        self.speakWordButton = QPushButton(self.detailCard)
        self.speakWordButton.setObjectName(u"speakWordButton")

        self.detailActionGrid.addWidget(self.speakWordButton, 0, 0, 1, 1)

        self.generateSentenceButton = QPushButton(self.detailCard)
        self.generateSentenceButton.setObjectName(u"generateSentenceButton")

        self.detailActionGrid.addWidget(self.generateSentenceButton, 0, 1, 1, 1)

        self.findInCorpusButton = QPushButton(self.detailCard)
        self.findInCorpusButton.setObjectName(u"findInCorpusButton")

        self.detailActionGrid.addWidget(self.findInCorpusButton, 1, 0, 1, 1)

        self.editPosTranslationButton = QPushButton(self.detailCard)
        self.editPosTranslationButton.setObjectName(u"editPosTranslationButton")

        self.detailActionGrid.addWidget(self.editPosTranslationButton, 1, 1, 1, 1)

        self.synonymsButton = QPushButton(self.detailCard)
        self.synonymsButton.setObjectName(u"synonymsButton")

        self.detailActionGrid.addWidget(self.synonymsButton, 2, 0, 1, 1)

        self.inspectAudioCacheButton = QPushButton(self.detailCard)
        self.inspectAudioCacheButton.setObjectName(u"inspectAudioCacheButton")

        self.detailActionGrid.addWidget(self.inspectAudioCacheButton, 2, 1, 1, 1)

        self.deleteWordButton = QPushButton(self.detailCard)
        self.deleteWordButton.setObjectName(u"deleteWordButton")

        self.detailActionGrid.addWidget(self.deleteWordButton, 3, 0, 1, 2)


        self.detailLayout.addLayout(self.detailActionGrid)


        self.rightLayout.addWidget(self.detailCard)

        self.rightTabWidget = QTabWidget(self.rightPanel)
        self.rightTabWidget.setObjectName(u"rightTabWidget")
        self.reviewTab = QWidget()
        self.reviewTab.setObjectName(u"reviewTab")
        self.reviewLayout = QVBoxLayout(self.reviewTab)
        self.reviewLayout.setSpacing(10)
        self.reviewLayout.setObjectName(u"reviewLayout")
        self.reviewLayout.setContentsMargins(12, 12, 12, 12)
        self.studyFocusTitleLabel = QLabel(self.reviewTab)
        self.studyFocusTitleLabel.setObjectName(u"studyFocusTitleLabel")

        self.reviewLayout.addWidget(self.studyFocusTitleLabel)

        self.reviewWordValueLabel = QLabel(self.reviewTab)
        self.reviewWordValueLabel.setObjectName(u"reviewWordValueLabel")

        self.reviewLayout.addWidget(self.reviewWordValueLabel)

        self.reviewNoteValueLabel = QLabel(self.reviewTab)
        self.reviewNoteValueLabel.setObjectName(u"reviewNoteValueLabel")
        self.reviewNoteValueLabel.setWordWrap(True)

        self.reviewLayout.addWidget(self.reviewNoteValueLabel)

        self.reviewStatsValueLabel = QLabel(self.reviewTab)
        self.reviewStatsValueLabel.setObjectName(u"reviewStatsValueLabel")
        self.reviewStatsValueLabel.setWordWrap(True)

        self.reviewLayout.addWidget(self.reviewStatsValueLabel)

        self.reviewActionRow = QHBoxLayout()
        self.reviewActionRow.setObjectName(u"reviewActionRow")
        self.openSourceButton = QPushButton(self.reviewTab)
        self.openSourceButton.setObjectName(u"openSourceButton")

        self.reviewActionRow.addWidget(self.openSourceButton)

        self.markWrongButton = QPushButton(self.reviewTab)
        self.markWrongButton.setObjectName(u"markWrongButton")

        self.reviewActionRow.addWidget(self.markWrongButton)


        self.reviewLayout.addLayout(self.reviewActionRow)

        self.reviewSpacer = QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)

        self.reviewLayout.addItem(self.reviewSpacer)

        self.rightTabWidget.addTab(self.reviewTab, "")
        self.historyTab = QWidget()
        self.historyTab.setObjectName(u"historyTab")
        self.historyLayout = QVBoxLayout(self.historyTab)
        self.historyLayout.setObjectName(u"historyLayout")
        self.historyLayout.setContentsMargins(12, 12, 12, 12)
        self.historyTitleLabel = QLabel(self.historyTab)
        self.historyTitleLabel.setObjectName(u"historyTitleLabel")

        self.historyLayout.addWidget(self.historyTitleLabel)

        self.historyListWidget = QListWidget(self.historyTab)
        self.historyListWidget.setObjectName(u"historyListWidget")

        self.historyLayout.addWidget(self.historyListWidget)

        self.historyEmptyLabel = QLabel(self.historyTab)
        self.historyEmptyLabel.setObjectName(u"historyEmptyLabel")
        self.historyEmptyLabel.setWordWrap(True)

        self.historyLayout.addWidget(self.historyEmptyLabel)

        self.historyButtonRow = QHBoxLayout()
        self.historyButtonRow.setObjectName(u"historyButtonRow")
        self.openHistoryButton = QPushButton(self.historyTab)
        self.openHistoryButton.setObjectName(u"openHistoryButton")

        self.historyButtonRow.addWidget(self.openHistoryButton)

        self.historyPathValueLabel = QLabel(self.historyTab)
        self.historyPathValueLabel.setObjectName(u"historyPathValueLabel")
        self.historyPathValueLabel.setWordWrap(True)

        self.historyButtonRow.addWidget(self.historyPathValueLabel)


        self.historyLayout.addLayout(self.historyButtonRow)

        self.rightTabWidget.addTab(self.historyTab, "")
        self.toolsTab = QWidget()
        self.toolsTab.setObjectName(u"toolsTab")
        self.toolsLayout = QVBoxLayout(self.toolsTab)
        self.toolsLayout.setSpacing(10)
        self.toolsLayout.setObjectName(u"toolsLayout")
        self.toolsLayout.setContentsMargins(12, 12, 12, 12)
        self.learningToolsTitleLabel = QLabel(self.toolsTab)
        self.learningToolsTitleLabel.setObjectName(u"learningToolsTitleLabel")

        self.toolsLayout.addWidget(self.learningToolsTitleLabel)

        self.toolsGrid = QGridLayout()
        self.toolsGrid.setObjectName(u"toolsGrid")
        self.toolSentenceButton = QPushButton(self.toolsTab)
        self.toolSentenceButton.setObjectName(u"toolSentenceButton")

        self.toolsGrid.addWidget(self.toolSentenceButton, 0, 0, 1, 1)

        self.toolFindButton = QPushButton(self.toolsTab)
        self.toolFindButton.setObjectName(u"toolFindButton")

        self.toolsGrid.addWidget(self.toolFindButton, 0, 1, 1, 1)

        self.toolPassageButton = QPushButton(self.toolsTab)
        self.toolPassageButton.setObjectName(u"toolPassageButton")

        self.toolsGrid.addWidget(self.toolPassageButton, 1, 0, 1, 1)

        self.toolSettingsButton = QPushButton(self.toolsTab)
        self.toolSettingsButton.setObjectName(u"toolSettingsButton")

        self.toolsGrid.addWidget(self.toolSettingsButton, 1, 1, 1, 1)

        self.toolUpdateButton = QPushButton(self.toolsTab)
        self.toolUpdateButton.setObjectName(u"toolUpdateButton")

        self.toolsGrid.addWidget(self.toolUpdateButton, 2, 0, 1, 2)

        self.exportCacheButton = QPushButton(self.toolsTab)
        self.exportCacheButton.setObjectName(u"exportCacheButton")

        self.toolsGrid.addWidget(self.exportCacheButton, 3, 0, 1, 1)

        self.importCacheButton = QPushButton(self.toolsTab)
        self.importCacheButton.setObjectName(u"importCacheButton")

        self.toolsGrid.addWidget(self.importCacheButton, 3, 1, 1, 1)

        self.exportPackButton = QPushButton(self.toolsTab)
        self.exportPackButton.setObjectName(u"exportPackButton")

        self.toolsGrid.addWidget(self.exportPackButton, 4, 0, 1, 1)

        self.importPackButton = QPushButton(self.toolsTab)
        self.importPackButton.setObjectName(u"importPackButton")

        self.toolsGrid.addWidget(self.importPackButton, 4, 1, 1, 1)


        self.toolsLayout.addLayout(self.toolsGrid)

        self.toolsHintLabel = QLabel(self.toolsTab)
        self.toolsHintLabel.setObjectName(u"toolsHintLabel")
        self.toolsHintLabel.setWordWrap(True)

        self.toolsLayout.addWidget(self.toolsHintLabel)

        self.toolsSpacer = QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)

        self.toolsLayout.addItem(self.toolsSpacer)

        self.rightTabWidget.addTab(self.toolsTab, "")

        self.rightLayout.addWidget(self.rightTabWidget)

        self.mainSplitter.addWidget(self.rightPanel)

        self.rootLayout.addWidget(self.mainSplitter)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)

        self.rightTabWidget.setCurrentIndex(0)


        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"Word Speaker", None))
        self.titleLabel.setText(QCoreApplication.translate("MainWindow", u"Word Speaker", None))
        self.subtitleLabel.setText(QCoreApplication.translate("MainWindow", u"PySide6 migration shell", None))
        self.currentFileStaticLabel.setText(QCoreApplication.translate("MainWindow", u"Current File", None))
        self.currentFileValueLabel.setText(QCoreApplication.translate("MainWindow", u"No file", None))
        self.audioBackendStaticLabel.setText(QCoreApplication.translate("MainWindow", u"Audio Backend", None))
        self.audioBackendValueLabel.setText(QCoreApplication.translate("MainWindow", u"None", None))
        self.wordListTitleLabel.setText(QCoreApplication.translate("MainWindow", u"\u5355\u8bcd\u8868", None))
        self.wordListDescLabel.setText(QCoreApplication.translate("MainWindow", u"\u5bfc\u5165\u8bcd\u8868\u3001\u76f4\u63a5\u7f16\u8f91\uff0c\u7136\u540e\u4ece\u5f53\u524d\u9009\u4e2d\u5355\u8bcd\u5f00\u59cb\u5b66\u4e60\u3002", None))
        self.importButton.setText(QCoreApplication.translate("MainWindow", u"\u5bfc\u5165", None))
        self.manualInputButton.setText(QCoreApplication.translate("MainWindow", u"\u7c98\u8d34 / \u8f93\u5165", None))
        self.saveAsButton.setText(QCoreApplication.translate("MainWindow", u"\u53e6\u5b58\u4e3a", None))
        self.newListButton.setText(QCoreApplication.translate("MainWindow", u"\u65b0\u5efa\u8bcd\u8868", None))
        ___qtablewidgetitem = self.wordTable.horizontalHeaderItem(0)
        ___qtablewidgetitem.setText(QCoreApplication.translate("MainWindow", u"#", None));
        ___qtablewidgetitem1 = self.wordTable.horizontalHeaderItem(1)
        ___qtablewidgetitem1.setText(QCoreApplication.translate("MainWindow", u"\u5355\u8bcd", None));
        ___qtablewidgetitem2 = self.wordTable.horizontalHeaderItem(2)
        ___qtablewidgetitem2.setText(QCoreApplication.translate("MainWindow", u"\u5907\u6ce8", None));
        self.emptyStateLabel.setText(QCoreApplication.translate("MainWindow", u"\u8fd8\u6ca1\u6709\u5355\u8bcd\uff0c\u5148\u70b9\u5bfc\u5165\u5f00\u59cb\u3002", None))
        self.playButton.setText(QCoreApplication.translate("MainWindow", u"\u25b6 \u64ad\u653e", None))
        self.settingsButton.setText(QCoreApplication.translate("MainWindow", u"\u8bbe\u7f6e", None))
        self.dictationButton.setText(QCoreApplication.translate("MainWindow", u"\u542c\u5199", None))
        self.statusLabel.setText(QCoreApplication.translate("MainWindow", u"Not started", None))
        self.currentWordTitleLabel.setText(QCoreApplication.translate("MainWindow", u"\u5f53\u524d\u5355\u8bcd", None))
        self.selectedWordEdit.setPlaceholderText(QCoreApplication.translate("MainWindow", u"\u8fd9\u91cc\u7f16\u8f91\u5355\u8bcd\u5185\u5bb9", None))
        self.selectedNoteEdit.setPlaceholderText(QCoreApplication.translate("MainWindow", u"\u8fd9\u91cc\u7f16\u8f91\u5907\u6ce8\u3001\u4e2d\u6587\u6216\u5b66\u4e60\u63d0\u793a", None))
        self.wrongCountStaticLabel.setText(QCoreApplication.translate("MainWindow", u"\u9519\u8fc7\u6b21\u6570", None))
        self.wrongCountValueLabel.setText(QCoreApplication.translate("MainWindow", u"0", None))
        self.correctCountStaticLabel.setText(QCoreApplication.translate("MainWindow", u"\u6b63\u786e\u6b21\u6570", None))
        self.correctCountValueLabel.setText(QCoreApplication.translate("MainWindow", u"0", None))
        self.lastResultStaticLabel.setText(QCoreApplication.translate("MainWindow", u"\u4e0a\u6b21\u7ed3\u679c", None))
        self.lastResultValueLabel.setText(QCoreApplication.translate("MainWindow", u"\u65e0", None))
        self.lastSeenStaticLabel.setText(QCoreApplication.translate("MainWindow", u"\u6700\u8fd1\u65f6\u95f4", None))
        self.lastSeenValueLabel.setText(QCoreApplication.translate("MainWindow", u"\u65e0", None))
        self.speakWordButton.setText(QCoreApplication.translate("MainWindow", u"\u6717\u8bfb\u5355\u8bcd", None))
        self.generateSentenceButton.setText(QCoreApplication.translate("MainWindow", u"\u751f\u6210\u4f8b\u53e5", None))
        self.findInCorpusButton.setText(QCoreApplication.translate("MainWindow", u"\u8bed\u6599\u68c0\u7d22", None))
        self.editPosTranslationButton.setText(QCoreApplication.translate("MainWindow", u"\u7f16\u8f91\u8bcd\u6027 / \u4e2d\u6587", None))
        self.synonymsButton.setText(QCoreApplication.translate("MainWindow", u"\u8fd1\u4e49\u8bcd", None))
        self.inspectAudioCacheButton.setText(QCoreApplication.translate("MainWindow", u"\u67e5\u8be2\u97f3\u9891\u7f13\u5b58", None))
        self.deleteWordButton.setText(QCoreApplication.translate("MainWindow", u"\u5220\u9664\u5355\u8bcd", None))
        self.studyFocusTitleLabel.setText(QCoreApplication.translate("MainWindow", u"\u5b66\u4e60\u7126\u70b9", None))
        self.reviewWordValueLabel.setText(QCoreApplication.translate("MainWindow", u"-", None))
        self.reviewNoteValueLabel.setText(QCoreApplication.translate("MainWindow", u"-", None))
        self.reviewStatsValueLabel.setText(QCoreApplication.translate("MainWindow", u"-", None))
        self.openSourceButton.setText(QCoreApplication.translate("MainWindow", u"\u6253\u5f00\u5386\u53f2", None))
        self.markWrongButton.setText(QCoreApplication.translate("MainWindow", u"\u624b\u52a8\u52a0\u5165\u9519\u8bcd", None))
        self.rightTabWidget.setTabText(self.rightTabWidget.indexOf(self.reviewTab), QCoreApplication.translate("MainWindow", u"\u590d\u4e60", None))
        self.historyTitleLabel.setText(QCoreApplication.translate("MainWindow", u"\u6700\u8fd1\u6253\u5f00", None))
        self.historyEmptyLabel.setText(QCoreApplication.translate("MainWindow", u"\u8fd8\u6ca1\u6709\u5386\u53f2\u8bb0\u5f55\u3002", None))
        self.openHistoryButton.setText(QCoreApplication.translate("MainWindow", u"\u6253\u5f00\u5386\u53f2", None))
        self.historyPathValueLabel.setText(QCoreApplication.translate("MainWindow", u"-", None))
        self.rightTabWidget.setTabText(self.rightTabWidget.indexOf(self.historyTab), QCoreApplication.translate("MainWindow", u"\u5386\u53f2", None))
        self.learningToolsTitleLabel.setText(QCoreApplication.translate("MainWindow", u"\u5b66\u4e60\u5de5\u5177", None))
        self.toolSentenceButton.setText(QCoreApplication.translate("MainWindow", u"\u751f\u6210\u4f8b\u53e5", None))
        self.toolFindButton.setText(QCoreApplication.translate("MainWindow", u"\u8bed\u6599\u53e5\u5b50\u68c0\u7d22", None))
        self.toolPassageButton.setText(QCoreApplication.translate("MainWindow", u"\u751f\u6210 IELTS \u7bc7\u7ae0", None))
        self.toolSettingsButton.setText(QCoreApplication.translate("MainWindow", u"\u97f3\u6e90 / \u6a21\u578b\u8bbe\u7f6e", None))
        self.toolUpdateButton.setText(QCoreApplication.translate("MainWindow", u"\u66f4\u65b0\u7a0b\u5e8f", None))
        self.exportCacheButton.setText(QCoreApplication.translate("MainWindow", u"\u5bfc\u51fa\u5171\u4eab\u7f13\u5b58", None))
        self.importCacheButton.setText(QCoreApplication.translate("MainWindow", u"\u5bfc\u5165\u5171\u4eab\u7f13\u5b58", None))
        self.exportPackButton.setText(QCoreApplication.translate("MainWindow", u"\u5bfc\u51fa\u8bcd\u8868\u8d44\u6e90\u5305", None))
        self.importPackButton.setText(QCoreApplication.translate("MainWindow", u"\u5bfc\u5165\u8bcd\u8868\u8d44\u6e90\u5305", None))
        self.toolsHintLabel.setText(QCoreApplication.translate("MainWindow", u"\u63d0\u793a\uff1a\u5148\u9009\u4e2d\u5355\u8bcd\uff0c\u518d\u751f\u6210\u4f8b\u53e5\u6216\u505a\u5b9a\u5411\u8bed\u6599\u68c0\u7d22\u3002", None))
        self.rightTabWidget.setTabText(self.rightTabWidget.indexOf(self.toolsTab), QCoreApplication.translate("MainWindow", u"\u5de5\u5177", None))
    # retranslateUi

