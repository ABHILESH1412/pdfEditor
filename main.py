import sys
import os
from PyQt5.QtWidgets import (
  QApplication,
  QWidget,
  QPushButton,
  QLabel, QFileDialog,
  QMainWindow,
  QVBoxLayout,
  QHBoxLayout,
  QStackedLayout,
  QLineEdit,
  QListWidget,
)

from pdf2docx import parse
from docx2pdf import convert
from PyPDF2 import PdfMerger
from PIL import Image

class Win(QMainWindow):
  def __init__(self):
    super().__init__()

    self.setGeometry(750, 250, 400, 400)
    self.setWindowTitle("PDF Editor")
    self.setFixedSize(400, 400)

    self.count = 0
    self.inputFilePath = ""
    self.inputFileName = ""
    self.outputPath = ""
    self.outputFileName = ""
    self.outputDirectoryLineEdit = QLineEdit()
    self.outputDirectoryLineEdit1 = QLineEdit()
    self.outputDirectoryLineEdit.textEdited.connect(self.textEdit)
    self.outputDirectoryLineEdit1.textEdited.connect(self.textEdit)
    self.inputFiles = []

    self.pageLayout = QStackedLayout()
    self.mainPageLayout = QVBoxLayout()
    self.outputLayout1 = QVBoxLayout()
    self.outputLayout2 = QVBoxLayout()
    self.buttonLayout1 = QHBoxLayout()
    self.buttonLayout2 = QHBoxLayout()
    self.buttonLayout3 = QHBoxLayout()
    self.buttonLayout4 = QHBoxLayout()
    self.dirSelLayout = QHBoxLayout()
    self.dirSelLayout1 = QHBoxLayout()
    self.fileInputLayout = QStackedLayout()

    self.editPdfButton = QPushButton("Edit PDF")
    self.editPdfButton.setObjectName("editPdfButton")
    self.pdfToWordButton = QPushButton("PDF to WORD")
    self.wordToPdfButton = QPushButton("WORD to PDF")
    self.mergePdfsButton = QPushButton("Merge PDFs")
    self.imageToPdfButton = QPushButton("Image to PDF")
    self.inputBrowseButton = QPushButton("Browse")
    self.inputBrowseButton.setObjectName("inputBrowse")
    self.inputMultipleFileBrowseButton = QPushButton("Browse")
    self.inputMultipleFileBrowseButton.setObjectName("inputBrowse")
    self.inputMultipleFileBrowseButton1 = QPushButton("Add More")
    self.outputDirectoryBrowseButton = QPushButton("Browse")
    self.outputDirectoryBrowseButton1 = QPushButton("Browse")
    self.saveAsOriginalButton = QPushButton("Save as Original")
    self.saveAsCopyButton = QPushButton("Save as Copy")
    self.startButton = QPushButton("Start")
    self.deleteButton = QPushButton("Delete")

    self.listOfInputs = QListWidget()
    self.label1 = QLabel("Output file directory: ")
    self.label1.setObjectName("label1")
    self.mainPageLayout.addLayout(self.buttonLayout1)
    self.mainPageLayout.addLayout(self.buttonLayout2)
    self.mainPageLayout.addLayout(self.fileInputLayout)
    self.dirSelLayout.addWidget(self.outputDirectoryLineEdit)
    self.dirSelLayout.addWidget(self.outputDirectoryBrowseButton)
    self.dirSelLayout1.addWidget(self.outputDirectoryLineEdit1)
    self.dirSelLayout1.addWidget(self.outputDirectoryBrowseButton1)
    self.outputLayout1.addWidget(self.label1)
    self.outputLayout1.addLayout(self.dirSelLayout)
    self.outputLayout1.addLayout(self.buttonLayout3)
    self.outputLayout2.addLayout(self.buttonLayout4)
    self.outputLayout2.addWidget(QLabel("Output file directory:", self))
    self.outputLayout2.addLayout(self.dirSelLayout1)
    self.outputLayout2.addWidget(self.listOfInputs)
    self.outputLayout2.addWidget(self.startButton)

    self.buttonLayout1.addWidget(self.editPdfButton)
    self.buttonLayout1.addWidget(self.pdfToWordButton)
    self.buttonLayout1.addWidget(self.wordToPdfButton)
    self.buttonLayout2.addWidget(self.mergePdfsButton)
    self.buttonLayout2.addWidget(self.imageToPdfButton)
    self.buttonLayout3.addWidget(self.saveAsOriginalButton)
    self.buttonLayout3.addWidget(self.saveAsCopyButton)
    self.buttonLayout4.addWidget(self.inputMultipleFileBrowseButton1)
    self.buttonLayout4.addWidget(self.deleteButton)
    
    self.fileInputLayout.addWidget(self.inputBrowseButton)
    self.fileInputLayout.addWidget(self.inputMultipleFileBrowseButton)

    self.editPdfButton.clicked.connect(self.editPdfButtonClicked)
    self.pdfToWordButton.clicked.connect(self.pdfToWordButtonClicked)
    self.wordToPdfButton.clicked.connect(self.wordToPdfButtonClicked)
    self.mergePdfsButton.clicked.connect(self.mergePdfsButtonClicked)
    self.imageToPdfButton.clicked.connect(self.imageToPdfButtonClicked)
    self.inputBrowseButton.clicked.connect(self.inputBrowseButtonClicked)
    self.inputMultipleFileBrowseButton.clicked.connect(self.inputMultipleFileBrowseButtonClicked)
    self.inputMultipleFileBrowseButton1.clicked.connect(self.inputMultipleFileBrowseButtonClicked)
    self.outputDirectoryBrowseButton.clicked.connect(self.outputDirectoryBrowseButtonClicked)
    self.outputDirectoryBrowseButton1.clicked.connect(self.outputDirectoryBrowseButtonClicked)
    self.saveAsOriginalButton.clicked.connect(self.saveAsOriginalButtonClicked)
    self.saveAsCopyButton.clicked.connect(self.saveAsCopyButtonClicked)
    self.deleteButton.clicked.connect(self.deleteButtonClicked)
    self.startButton.clicked.connect(self.startButtonClicked)

    pages = QWidget()
    page0 = QWidget()
    page1 = QWidget()
    page2 = QWidget()
    page0.setLayout(self.mainPageLayout)
    page1.setLayout(self.outputLayout1)
    page2.setLayout(self.outputLayout2)
    self.pageLayout.addWidget(page0)
    self.pageLayout.addWidget(page1)
    self.pageLayout.addWidget(page2)
    pages.setLayout(self.pageLayout)
    self.setCentralWidget(pages)

  def textEdit(self, s):
    self.outputPath = s

  def button(self, buttonText):
    tempButton = QPushButton(self)
    tempButton.setText(f"{buttonText}")

    return tempButton

  ############################## Button Click Events #############################
  def inputBrowseButtonClicked(self):
    filePath = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf)")
    if(filePath[0]):
      count = 0

      for ch in reversed(filePath[0]):
        if ch == '/':
          break

        self.inputFileName = ch + self.inputFileName
        count += 1

      self.inputFilePath = filePath[0][0 : len(filePath[0]) - count]
      self.outputPath = self.inputFilePath
      self.pdfToDocxConverter(filePath[0])
      os.system(f'start {self.outputPath + self.outputFileName}')

      self.outputDirectoryLineEdit.setText(self.outputPath)
      self.pageLayout.setCurrentIndex(1)
    else:
      pass

  def outputDirectoryBrowseButtonClicked(self):
    path = QFileDialog.getExistingDirectory(self, "Choose Directory", "")
    if(len(path) != 3):
      path += "/"
    
    self.outputPath = path
    if self.count == 0:
      self.outputDirectoryLineEdit.setText(self.outputPath)
    elif self.count == range(1,5):
      self.outputDirectoryLineEdit1.setText(self.outputPath)

  def saveAsCopyButtonClicked(self):
    self.docxToPdfConverter("copy")
    os.system(f"rm {self.inputFilePath + self.outputFileName}")

    self.inputFilePath = ""
    self.inputFileName = ""
    self.outputPath = ""
    self.outputFileName = ""
    self.pageLayout.setCurrentIndex(0)

  def saveAsOriginalButtonClicked(self):
    self.docxToPdfConverter("original")
    os.system(f"rm {self.inputFilePath + self.outputFileName}")

    self.inputFilePath = ""
    self.inputFileName = ""
    self.outputPath = ""
    self.outputFileName = ""
    self.pageLayout.setCurrentIndex(0)

  def editPdfButtonClicked(self):
    self.fileInputLayout.setCurrentIndex(0)
    self.count = 0
    self.colorButtons()

  def pdfToWordButtonClicked(self):
    self.fileInputLayout.setCurrentIndex(1)
    self.count = 1
    self.colorButtons()

  def wordToPdfButtonClicked(self):
    self.fileInputLayout.setCurrentIndex(1)
    self.count = 2
    self.colorButtons()

  def mergePdfsButtonClicked(self):
    self.fileInputLayout.setCurrentIndex(1)
    self.count = 3
    self.colorButtons()

  def imageToPdfButtonClicked(self):
    self.fileInputLayout.setCurrentIndex(1)
    self.count = 4
    self.colorButtons()

  def colorButtons(self):
    self.editPdfButton.setStyleSheet("background: rgb(61, 153, 245); color: white")
    self.pdfToWordButton.setStyleSheet("background: rgb(61, 153, 245); color: white")
    self.wordToPdfButton.setStyleSheet("background: rgb(61, 153, 245); color: white")
    self.mergePdfsButton.setStyleSheet("background: rgb(61, 153, 245); color: white")
    self.imageToPdfButton.setStyleSheet("background: rgb(61, 153, 245); color: white")

    self.editPdfButton.setStyleSheet("""
        QPushButton:hover {
          background: rgb(97, 174, 250);         
        }
    """)
    self.pdfToWordButton.setStyleSheet("""
        QPushButton:hover {
          background: rgb(97, 174, 250);         
        }
    """)
    self.wordToPdfButton.setStyleSheet("""
        QPushButton:hover {
          background: rgb(97, 174, 250);         
        }
    """)
    self.mergePdfsButton.setStyleSheet("""
        QPushButton:hover {
          background: rgb(97, 174, 250);         
        }
    """)
    self.imageToPdfButton.setStyleSheet("""
        QPushButton:hover {
          background: rgb(97, 174, 250);         
        }
    """)
    

    if self.count == 0:
      self.editPdfButton.setStyleSheet("background: white; color: black; border: 2px solid black")
    elif self.count == 1:
      self.pdfToWordButton.setStyleSheet("background: white; color: black; border: 2px solid black")
    elif self.count == 2:
      self.wordToPdfButton.setStyleSheet("background: white; color: black; border: 2px solid black")
    elif self.count == 3:
      self.mergePdfsButton.setStyleSheet("background: white; color: black; border: 2px solid black")
    elif self.count == 4:
      self.imageToPdfButton.setStyleSheet("background: white; color: black; border: 2px solid black")




  def inputMultipleFileBrowseButtonClicked(self):
    path = ""
    if self.count == 1 or self.count == 3:
      path = QFileDialog.getOpenFileNames(self, "Select PDF Files", "", "PDF Files (*.pdf)")

    elif self.count == 2:
      path = QFileDialog.getOpenFileNames(self, "Select WORD Files", "", "WORD Files (*.docx)")

    elif self.count == 4:
      path = QFileDialog.getOpenFileNames(self, "Select Image Files", "", "Images (*.jpg, *.jpeg, *.png)")

    if(len(path[0]) == 0):
      return
    lenOfList = self.listOfInputs.count()
    for i in range(len(path[0])):
      self.listOfInputs.insertItem(lenOfList, path[0][i])
      lenOfList += 1

    if len(self.outputPath) == 0:
      count = 0
      path = self.listOfInputs.item(0).text()

      for ch in reversed(path):
        if(ch == "/"):
          break
        else:
          count += 1

      for i in range(len(path)-count):
        self.outputPath += path[i]

      self.outputDirectoryLineEdit1.setText(self.outputPath)


    self.pageLayout.setCurrentIndex(2)

  def deleteButtonClicked(self):
    self.listOfInputs.takeItem(self.listOfInputs.currentRow())

  def startButtonClicked(self):
    if self.count == 1:
      for i in range(self.listOfInputs.count()):
        self.getName(self.listOfInputs.item(i).text())
        self.pdfToDocxConverter(self.listOfInputs.item(i).text())

    elif self.count == 2:
      for i in range(self.listOfInputs.count()):
        self.getName(self.listOfInputs.item(i).text())
        self.outputFileName = self.inputFileName.replace(".docx", ".pdf")
        self.outputFileName = self.outputFileName.replace(" ", "_")
        convert(self.listOfInputs.item(i).text(), self.outputPath + self.outputFileName)
      
    elif self.count == 3:
      self.outputFileName = "Merged_Pdfs.pdf"
      merger = PdfMerger()

      for i in range(self.listOfInputs.count()):
        merger.append(self.listOfInputs.item(i).text())

      merger.write(self.outputPath + self.outputFileName)
      merger.close()

    elif self.count == 4:
      imgs = []

      for i in range(1, self.listOfInputs.count()):
        img = Image.open(self.listOfInputs.item(i).text())
        imgs.append(img.convert("RGB"))

      img = Image.open(self.listOfInputs.item(0).text())
      img = img.convert("RGB")

      img.save(self.outputPath + "Result.pdf", save_all = True, append_images = imgs)

    self.inputFilePath = ""
    self.inputFileName = ""
    self.outputPath = ""
    self.outputFileName = ""
    self.listOfInputs.clear()
    self.pageLayout.setCurrentIndex(0)

  def getName(self, path):
    fileName = ""
    for ch in reversed(path):
      if(ch == '/'):
        break
      
      fileName = ch + fileName
    
    self.inputFileName = fileName
    
  #################################################################################

  ################################### Converters ##################################
  def pdfToDocxConverter(self, filePath):
    self.outputFileName = self.inputFileName.replace(".pdf", "_converted.docx")
    self.outputFileName = self.outputFileName.replace(" ", "_")

    parse(filePath, self.outputPath + self.outputFileName)

  def docxToPdfConverter(self, format):
    path = ""
    if format == "copy":
      path = self.outputPath + self.outputFileName.replace(".docx", ".pdf")
    else:
      path = self.inputFilePath + f"{self.inputFileName}"
      
    convert(self.inputFilePath + self.outputFileName, path)
  #################################################################################


def main():
  app = QApplication(sys.argv)

  style = """
    QPushButton{
      font-size: 12px;
      font-weight: bold;
      color: white;
      background: rgb(61, 153, 245);
      padding: 10;
      border-radius: 10;
    }
    QPushButton#inputBrowse{
      border: 2px dashed black;
      font-size: 16px;
    }
    QPushButton:hover{
      background: rgb(97, 174, 250);
    }
    QLineEdit{
      font-size: 12px;
      padding: 8;
      border-radius: 10;
      border: 1px solid black;
    }
    QListWidget{
      font-size: 12px;
    }
    QLabel{
      font-size: 12px;
      font-weight: bold;
    }
    QLabel#label1{
      max-height: 20px;
    }
  """
  app.setStyleSheet(style)

  Window = Win()
  Window.show()

  sys.exit(app.exec_())

main()