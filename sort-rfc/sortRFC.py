import sys
import os
import re
import PyPDF2
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
import xml.etree.cElementTree as ET
from PyQt4 import QtGui, QtCore


class Example(QtGui.QWidget):
    def __init__(self):
        super(Example, self).__init__()
        self.initUI()

    def initUI(self):
        lineOfButtonsX = 200
        lineOfButtonsY = 300
        self.directorio = os.getcwd()
        self.readScrTxt()  # initialize self.src and self.dist

        QtGui.QToolTip.setFont(QtGui.QFont('SansSerif', 10))
        self.setToolTip('Aplicación ordenar')
        self.srcDir = QtGui.QLabel('Dirección de los Archivos:')
        self.distDir = QtGui.QLabel('Dirección de los Clientes:')
        self.srcDirEdit = QtGui.QLabel(self.src)
        self.distDirEdit = QtGui.QLabel(self.dist)
        self.buttonChangeScr = QtGui.QPushButton('Cambiar', self)
        self.buttonChangeScr.clicked.connect(self.cambiarDirScr)
        self.buttonChangeDist = QtGui.QPushButton('Cambiar', self)
        self.buttonChangeDist.clicked.connect(self.cambiarDirDist)

        self.grid = QtGui.QGridLayout()
        self.grid.setSpacing(2)
        self.grid.addWidget(self.srcDir, 0, 0)
        self.grid.addWidget(self.srcDirEdit, 0, 1)
        self.grid.addWidget(self.buttonChangeScr, 0, 3)
        self.grid.addWidget(self.distDir, 1, 0)
        self.grid.addWidget(self.distDirEdit, 1, 1)
        self.grid.addWidget(self.buttonChangeDist, 1, 3)

        self.buttonSort = QtGui.QPushButton('Ordenar', self)
        self.buttonCancelar = QtGui.QPushButton('Cancelar', self)
        self.grid.addWidget(self.buttonSort, 2, 1)
        self.grid.addWidget(self.buttonCancelar, 2, 2)
        self.setLayout(self.grid)
        self.buttonSort.setToolTip(
            'Al presionar movera cada archivo<b> PDF y XML </b>que se encuentra en esta carpeta a las carpeta del cliente siempre y cuando el <b>RFC</b> se encuentre escrito dentro del archivo')
        self.buttonSort.resize(self.buttonSort.sizeHint())
        self.buttonCancelar.move(lineOfButtonsX + 100, lineOfButtonsY)
        self.buttonCancelar.clicked.connect(QtCore.QCoreApplication.instance().quit)
        self.buttonSort.clicked.connect(
            self.ordenador)  # if you want to pass a param you need to put lambda:self.prueba('pdf')
        self.setGeometry(200, 200, 400, 200)
        self.setWindowTitle('Ordenador de Archivos')
        self.show()

    def cambiarDirDist(self):
        text, ok = QtGui.QInputDialog.getText(self, 'Carpeta destino',
                                              'Inserte directorio de carpeta destino:')
        if ok:
            if str(text) != '':
                self.modifyDir(self.dist, str(text))
                self.readScrTxt()
                self.distDirEdit.setText(str(text))

    def cambiarDirScr(self):
        text, ok = QtGui.QInputDialog.getText(self, 'Carpeta origen',
                                              'Inserte directorio de carpeta origen:')
        if ok:
            if str(text) != '':
                self.modifyDir(self.src, str(text))
                self.readScrTxt()
                self.srcDirEdit.setText(str(text))

    # get all the Filenames of the extension of the CWD
    def getNames(self, extension):
        # extension = 'pdf'
        Names = []
        for filename in os.listdir(self.src):
            if filename.endswith('.' + extension):
                Names.append(filename)
        return Names

    def getTextPdf(self, filename):
        try:
            pdfFileObj = open(filename, 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            numeroDePaginas = pdfReader.numPages
            pdfFileObj.close()
            print(numeroDePaginas)
            if numeroDePaginas > 4:
                return ''
            else:
                file = open(filename, 'rb')
                parser = PDFParser(file)
                doc = PDFDocument()
                parser.set_document(doc)
                doc.set_parser(parser)
                doc.initialize('')
                rsrcmgr = PDFResourceManager()
                laparams = LAParams()
                laparams.char_margin = 1.0
                laparams.word_margin = 1.0
                device = PDFPageAggregator(rsrcmgr, laparams=laparams)
                interpreter = PDFPageInterpreter(rsrcmgr, device)
                extracted_text = ''
                for page in doc.get_pages():
                    interpreter.process_page(page)
                    layout = device.get_result()
                    for lt_obj in layout:
                        if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                            extracted_text += lt_obj.get_text()
                return extracted_text
        except:
            return ''

    def getTextXml(self, filename):
        # forma correcta de obtener una instancia del xml
        tree = ET.ElementTree(file=filename)
        root = tree.getroot()
        # RFC Emisor
        RFCEmisor = root[0].attrib['rfc']
        # RFC Receptor
        RFCReceptor = root[1].attrib['rfc']
        # getRFC ya compara los RFC por Regex, solo es necesario pasarle un string con los rfc
        return RFCEmisor + ' ' + RFCReceptor

    def rFCClientes(self):
        rFCClientes = []
        os.chdir(self.dist)
        for archivo in os.listdir('.'):
            if os.path.isdir(archivo):
                rFCClientes.append(archivo)
        os.chdir(self.src)
        return rFCClientes

    # encuentra que RFC(cliente) tiene el archivo
    # (text, filename):
    def getRFC(self, text):
        # extrae todos los RFC
        clientes = self.rFCClientes()
        for rFC in clientes:
            # (pattern,string,flag)
            cliente = re.search(rFC, text)
            if cliente:
                # regresa el RFC del Cliente en el archivo filename
                return cliente.group(0)
            # en caso de no encontrarlo
        return ''

    def ordenador(self):
        self.fileMover('pdf')
        self.fileMover('xml')
        self.showFinish()

    def fileMover(self, extension):
        # diccionario que guardara el nombre del archivo y a que cliente le pertence
        registro = {}
        clientes = self.rFCClientes()
        # crea un diccionario de dos niveles con todos los RFC de los directorios
        for cliente in clientes:
            registro[cliente] = []
        for filename in self.getNames(extension):
            if extension == 'pdf':
                try:
                    text = self.getTextPdf(filename)
                except:
                    pass
            else:
                try:
                    text = self.getTextXml(filename)
                except:
                    pass
            cliente = self.getRFC(text)
            if cliente != '':
                self.fileManage(cliente, filename)
                # es seguro que el cliente realmente existe en registro (se creo a partir de alli) y guarda el nombre de archivo
                registro[cliente].append(filename)
        return registro

    def fileManage(self, RFCName, filename):
        src = self.src + '\\' + filename
        dist = self.dist + '\\' + RFCName + '\\' + filename
        try:
            os.rename(src, dist)
        except:
            self.showError(filename, RFCName)

    def showError(self, filename, RFCName):
        problema = 'ya existe: ' + filename + ' en ' + self.dist + '\\' + RFCName
        problemaMsg = QtGui.QLabel(problema)
        ventanaError = QtGui.QDialog()
        buttonOk = QtGui.QPushButton("Ok", ventanaError)
        buttonOk.clicked.connect(ventanaError.accept)
        buttonOk.setDefault(True)
        ventanaError.layout = QtGui.QGridLayout(ventanaError)
        ventanaError.layout.addWidget(problemaMsg, 0, 0, 1, 3)
        ventanaError.layout.addWidget(buttonOk, 1, 1)
        ventanaError.setWindowTitle('Error, archivo ya existente')
        ventanaError.exec_()

    def showFinish(self):
        finalizacion = 'Se han transladado todos los archivos posibles'
        problemaMsg = QtGui.QLabel(finalizacion)
        ventanaFinal = QtGui.QDialog()
        buttonOk = QtGui.QPushButton("Ok", ventanaFinal)
        buttonOk.clicked.connect(ventanaFinal.accept)
        buttonOk.setDefault(True)
        ventanaFinal.layout = QtGui.QGridLayout(ventanaFinal)
        ventanaFinal.layout.addWidget(problemaMsg, 0, 0, 1, 3)
        ventanaFinal.layout.addWidget(buttonOk, 1, 1)
        ventanaFinal.setWindowTitle('Ordenación Terminada')
        ventanaFinal.exec_()

    def readScrTxt(self):
        if self.directorio == os.getcwd():
            patron = '(?<=\")(.*?)(?=\")'
            filename = 'directories.txt'
            document = open(filename)
            content = document.readlines()
            src = re.search(patron, content[0])
            dist = re.search(patron, content[1])
            document.close()
            self.src = src.group()
            self.dist = dist.group()
        else:
            os.chdir(self.directorio)
            patron = '(?<=\")(.*?)(?=\")'
            filename = 'directories.txt'
            document = open(filename)
            content = document.readlines()
            src = re.search(patron, content[0])
            dist = re.search(patron, content[1])
            document.close()
            self.src = src.group()
            self.dist = dist.group()
            os.chdir(self.src)

    def modifyDir(self, textTarget, textReplace):
        if self.directorio == os.getcwd():
            # Read in the file
            with open('directories.txt', 'r') as file:
                filedata = file.read()
            # Replace the target string
            filedata = filedata.replace(textTarget, textReplace, 1)
            file.close()
            # Write the file out again
            with open('directories.txt', 'w') as file:
                file.write(filedata)
                file.close()
        else:
            os.chdir(self.directorio)
            # Read in the file
            with open('directories.txt', 'r') as file:
                filedata = file.read()
            # Replace the target string
            filedata = filedata.replace(textTarget, textReplace, 1)
            file.close()
            # Write the file out again
            with open('directories.txt', 'w') as file:
                file.write(filedata)
                file.close()
            os.chdir(self.src)


def main():
    app = QtGui.QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
