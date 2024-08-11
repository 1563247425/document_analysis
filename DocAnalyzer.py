from ui.DocAnalyze_ui import Ui_MainWindow
from PyQt6.QtWidgets import QApplication, QMessageBox, QAbstractItemView, QMdiSubWindow, QPlainTextEdit, QLabel
from PyQt6.QtGui import QIcon, QPixmap, QFileSystemModel,QStandardItemModel, QStandardItem, QTextDocumentWriter
from PyQt6.QtCore import Qt, QDir
import sys
import os
import re
# Word文档操作库(pip install python-docx)
from docx import Document
#朗读库(pip install pyttsx3)
import pyttsx3
#分词库(pip install zhon)
from zhon.hanzi import punctuation
import jieba
#生成词云库
from wordcloud import WordCloud
#爬取信息库(pip install beautifulsoup4)
from bs4 import BeautifulSoup
import json
#识别文字库(pip install pytesseract)
import pytesseract
from PIL import Image
class MyDocAnalyzer(Ui_MainWindow):
    def __init__(self):
        super(MyDocAnalyzer, self).__init__()
        self.setupUi(self)
        self.initUi()
    def initUi(self):
        self.setWindowIcon(QIcon('icon.ico'))
        self.setWindowTitle('我的文档')
        self.setWindowFlags(Qt.WindowType.MSWindowsFixedSizeDialogHint)
        self.mdiArea.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.mdiArea.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.mdiArea.subWindowActivated.connect(self.updateMenuBar)
        self.actSave.setIcon(QIcon('./image/save.png'))
        self.actSave.setShortcut('Ctrl+S')
        self.actSave.triggered.connect(self.saveDoc)
        self.actSpeak.setIcon(QIcon('./image/speak.png'))
        self.actSpeak.setShortcut('Ctrl+R')
        self.actSpeak.setEnabled(False)
        self.actWord.setShortcut('Ctrl+W')
        self.actWord.setEnabled(False)
        self.actWord.triggered.connect(self.cutWord)
        self.actCloud.setIcon(QIcon('./image/cloud.png'))
        self.actCloud.setEnabled(False)
        self.actCloud.triggered.connect(self.generCloud)
        self.actCrawl.setIcon(QIcon('./image/crawl.png'))
        self.actCrawl.setEnabled(False)
        self.actCrawl.triggered.connect(self.titleCrawl)
        self.actRecog.setEnabled(False)
        self.actRecog.triggered.connect(self.textRecog)
        self.actClose.triggered.connect(self.closeDoc)
        self.actCloseAll.triggered.connect(self.closeAllDocs)
        self.actTile.triggered.connect(self.tileDocs)
        self.actCasCade.triggered.connect(self.cascadeDocs)
        self.actNext.triggered.connect(self.nextDoc)
        self.actPrev.triggered.connect(self.prevDoc)
        self.actAbout.triggered.connect(self.aboutApp)
        # 文档管理(树状视图)
        self.dirModel = QFileSystemModel()
        self.dirModel.setRootPath('')
        #从操作系统根目录开始显示,所有驱动器都能看到
        self.dirModel.setFilter(QDir.Filter.AllDirs | QDir.Filter.NoDotAndDotDot)
        self.trvOSDirs.setModel(self.dirModel)
        self.trvOSDirs.setHeaderHidden(True)
        for col in range(1, 4):
            self.trvOSDirs.setColumnHidden(col, True)
        self.trvOSDirs.doubleClicked.connect(self.showFiles)
        self.curPath = 'D:/PyQt6'
        self.curFile = ''
        dirList = self.curPath.split('/')
        defPath = ''
        for dir in dirList:
            if len(defPath) > 0:
                dir = '/' + dir
                defPath += dir
                self.trvOSDirs.setExpanded(self.dirModel.index(defPath),1)  # 逐层展开至默认目录
        self.fileModel = QStandardItemModel()
        self.trvDocFiles.setModel(self.fileModel)
        self.trvDocFiles.setHeaderHidden(True)
        self.trvDocFiles.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.trvDocFiles.doubleClicked.connect(self.showContent)
        self.initFileModel()
        self.resText = '' #分析出的结果文本

    def initFileModel(self):
        self.fileModel.clear()
        self.textType = QStandardItem('文本')
        self.fileModel.appendRow(self.textType)
        self.wordType = QStandardItem('Word文档')
        self.fileModel.appendRow(self.wordType)
        self.htmlType = QStandardItem('网页')
        self.fileModel.appendRow(self.htmlType)
        self.picType = QStandardItem('图片')
        self.fileModel.appendRow(self.picType)
        self.curFile = ''
        self.updateStatus()

    def showFiles(self, index):
        self.initFileModel()
        self.curPath = self.dirModel.filePath(index)
        fileSet = os.listdir(self.curPath)
        for i in range(len(fileSet)):
            if os.path.isdir(self.curPath + '\\' + fileSet[i]) == False:
                fileItem = QStandardItem(fileSet[i])
                type = fileSet[i].split('.')[1]
                if type == 'txt':
                    fileItem.setIcon(QIcon('image/text.jpg'))
                    self.textType.appendRow(fileItem)
                elif type == 'docx':
                    fileItem.setIcon(QIcon('image/word.jpg'))
                    self.wordType.appendRow(fileItem)
                elif type == 'htm' or type == 'html':
                    fileItem.setIcon(QIcon('image/html.jpg'))
                    self.htmlType.appendRow( fileItem)
                elif type == 'jpg' or type == 'jpeg' or type == 'png' or type == 'gif' or type == 'ico' or type == 'bmp':
                    fileItem.setIcon(QIcon('image/pic.jpg'))
                    self.picType.appendRow(fileItem)
        self.trvDocFiles.expandAll()
        self.updateStatus()

    def showContent(self, index):
        self.curFile = self.fileModel.itemData(index)[0]
        self.updateStatus()
        path = self.curPath + '/' + self.curFile
        type = self.curFile.split('.')[1]
        if type == 'txt' or type == 'docx' or type == 'htm' or type == 'html':
            content = ''
            if type == 'txt' or type == 'htm' or type == 'html':
                with open(path, 'r', encoding='utf-8') as f:
                    content = f.read()
            elif type == 'docx':
                doc = Document(path)
                for p in doc.paragraphs:
                    content += p.text
                    content += '\r\n'
            textDoc = QMdiSubWindow(self)
            textDoc.setWindowTitle(path)
            teContent = QPlainTextEdit(textDoc)
            teContent.setPlainText(content)
            textDoc.setWidget(teContent)
            self.mdiArea.addSubWindow(textDoc)
            textDoc.show()
        elif type == 'jpg' or type == 'jpeg' or type == 'png' or type == 'gif'or type == 'ico' or type == 'bmp':
            picDoc = QMdiSubWindow(self)
            picDoc.setWindowTitle(path)
            lbContent = QLabel(picDoc)
            lbContent.setPixmap(QPixmap(path))
            picDoc.setWidget(lbContent)
            self.mdiArea.addSubWindow(picDoc)
            picDoc.show()

    def updateStatus(self):
        self.statusbar.showMessage(self.curPath + '/' + self.curFile)

    def saveDoc(self):
        type = self.curFile.split('.')[1]
        if type == 'txt' or type == 'docx' or type == 'htm' or type == 'html':
            docName = self.mdiArea.currentSubWindow().windowTitle() + '.txt'
            writer = QTextDocumentWriter(docName)
            if writer.write(self.mdiArea.activeSubWindow().widget().document()):
                QMessageBox.information(self, '提示','已保存。')

        def quitApp(self):
            app = QApplication.instance()
            app.quit()

        def readSpeak(self):
            content = self.mdiArea.currentSubWindow().widget().toPlainText()
            engine = pyttsx3.init()
            engine.say(content)
            engine.runAndWait()

    def cutWord(self):
        content = self.mdiArea.currentSubWindow().widget().toPlainText()
        content = re.sub('[%s]+' % punctuation, '', content)
        content = re.sub('[%s]+' % '\r\n', '', content)
        jieba.load_userdict('dict.txt')
        self.resText = str(jieba.lcut(content))
        # result = jieba.cut(content)
        # for word in result:
        # self.resText += word
        # self.resText += ', '
        # self.resText += '\n'
        self.showResult('分词')

    def generCloud(self):
        content = self.mdiArea.currentSubWindow().widget().toPlainText()
        content = re.sub('[%s]+' % punctuation, '', content)
        content = re.sub('[%s]+' % '\r\n', '', content)
        jieba.load_userdict('dict.txt')
        content = ' '.join(jieba.lcut(content))
        # #将文本中",.?':!"符号替换成空格
        # for ch in ",.?':!":
        # content = content.replace(ch, ' ')
        cloud = WordCloud(font_path='simsun.ttc').generate(content)
        path = self.curPath + '/' + '词频云图.png'
        cloud.to_file(path)
        picDoc = QMdiSubWindow(self)
        picDoc.setWindowTitle(self.mdiArea.currentSubWindow().windowTitle()+ '-词云')
        lbResult = QLabel(picDoc)
        lbResult.setPixmap(QPixmap(path))
        picDoc.setWidget(lbResult)
        self.mdiArea.addSubWindow(picDoc)
        picDoc.show()

    def titleCrawl(self):
        content = self.mdiArea.currentSubWindow().widget().toPlainText()
        soup = BeautifulSoup(content, 'html.parser')
        links = []
        for div in soup.find_all('div', {'data-tools': re.compile('title')},{'data-tools': re.compile('url')}):
            data = div.attrs['data-tools'] #获取data-tools属性值
            data = str(data).replace("'", '"')
            d = json.loads(data)
            links.append(d['title'] + ': ' + d['url'])
            count = 1
            self.resText = ''
            for i in links:
                self.resText += '[{:^3}]{}'.format(count, i) + '\r\n'
                count += 1
            self.showResult('主题链接')

    def textRecog(self):
        path = self.curPath + '/' + self.curFile
        image = Image.open(path)
        self.resText = pytesseract.image_to_string(image, lang='chi_sim')
        self.showResult('识别文字')

    def showResult(self, mode):
        textDoc = QMdiSubWindow(self)
        textDoc.setWindowTitle(self.mdiArea.currentSubWindow().windowTitle() + '-' + mode)
        teResult = QPlainTextEdit(textDoc)
        teResult.setPlainText(self.resText)
        textDoc.setWidget(teResult)
        self.mdiArea.addSubWindow(textDoc)
        textDoc.show()

    def updateMenuBar(self):
        self.actSpeak.setEnabled(False)
        self.actWord.setEnabled(False)
        self.actCloud.setEnabled(False)
        self.actCrawl.setEnabled(False)
        self.actRecog.setEnabled(False)
        type = self.curFile.split('.')[1]
        if type == 'txt' or type == 'docx':
            self.actSpeak.setEnabled(True)
            self.actWord.setEnabled(True)
            self.actCloud.setEnabled(True)
        elif type == 'htm' or type == 'html':
            self.actCrawl.setEnabled(True)
        elif type == 'jpg' or type == 'jpeg' or type == 'png' or type == 'gif' or type == 'ico' or type == 'bmp':
            self.actRecog.setEnabled(True)

    def closeDoc(self):
        self.mdiArea.closeActiveSubWindow()

    def closeAllDocs(self):
        self.mdiArea.closeAllSubWindows()

    def tileDocs(self):
        self.mdiArea.tileSubWindows()

    def cascadeDocs(self):
        self.mdiArea.cascadeSubWindows()

    def nextDoc(self):
        self.mdiArea.activateNextSubWindow()

    def prevDoc(self):
        self.mdiArea.activatePreviousSubWindow()

    def aboutApp(self):
        QMessageBox.about(self, '关于',
                          '这是一个基于PyQt6实现的文档可视化分析软件\r\n可对文档进行朗读、分词、生成词云，另外还能\r\n爬取网页中的主题链接、识别图片中的文字。')
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyDocAnalyzer()
    window.show()
    sys.exit(app.exec())
