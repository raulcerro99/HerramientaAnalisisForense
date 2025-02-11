import PySide6
import spacy
import sys
import os
#import fitz
import matplotlib.pyplot as plt
import io
import numpy as np
import json
import html

from docx import Document  # python-docx para archivos Word
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt
from PySide6.QtGui import *
from PySide6.QtPrintSupport import QPrinter
from collections import Counter
from scipy.stats import mannwhitneyu

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle

nlp = spacy.load("es_core_news_sm")


class Principal(QMainWindow):

    def __init__(self):
        super().__init__()
        
        self.initUI()
        
    def initUI(self):
        
        self.html = ""
        self.resultados_por_texto = []
        self.resultados_agregados = []
        self.setWindowTitle("Herramienta para el análisis forense de textos")
        self.setGeometry(192, 108, 1920, 1080)

        self.mygrid = QGridLayout(self)
        self.mygrid.setSpacing(0)

        self.stacked_widget = QTabWidget()
        self.stacked_widget.addTab(QWidget(), "Análisis")
        self.stacked_widget.addTab(QWidget(), "Contraste")

        self.text_html_analisis = QTextEdit(self)
        self.contraste = QTextEdit(self)
        self.text_html_comparacion = QTextEdit(self)
        self.text_html_analaisis_corpus_2 = QTextEdit(self)
        self.text_archivo = QTextEdit(self)
        self.tabla_seleccion = QTableWidget(self)

        self.text_html_analisis.setReadOnly(True)
        self.contraste.setReadOnly(True)
        self.text_html_comparacion.setReadOnly(True)
        self.text_html_analaisis_corpus_2.setReadOnly(True)

        self.setupToolbar()
        self.setupPages()
        self.setupButtons()
        self.setupTable()
        self.setupLayout()

        central_widget = QWidget()
        central_widget.setLayout(self.mygrid)
        self.setCentralWidget(central_widget)

        self.analasis_informe.clicked.connect(self.ejecutar_algoritmo)
        self.analasis_comparacion.clicked.connect(self.cargar_archivo_único)

        self.text_file_list = [[]]
        self.index = 0
        self.actual_index = 0

    def setupToolbar(self):
        menu = self.menuBar()
        File = menu.addMenu("&Subir Archivos")
        Export = menu.addMenu("&Exportar Datos")
        Options = menu.addMenu("&Options")
        Tools = menu.addMenu("&Tools")
        Window = menu.addMenu("&Window")
        Help = menu.addMenu("&Help")

        button_subir_archivo_analisis = QAction("&Subir archivos", self)
        button_subir_archivo_analisis.triggered.connect(self.cargar_archivo)
        File.addAction(button_subir_archivo_analisis)

        button_exportar_informe = QAction("&Exportar informe", self)
        button_exportar_informe.triggered.connect(self.extraer_html_informe)
        Export.addAction(button_exportar_informe)

        button_exportar_comparacion = QAction("&Exportar comparación", self)
        button_exportar_comparacion.triggered.connect(self.extraer_html_comparacion)
        Export.addAction(button_exportar_comparacion)

    def setupPages(self):
        self.page1_frame = QFrame()
        self.page1_frame.setFrameShape(QFrame.Box)
        self.page1_frame.setLineWidth(1)
        self.page1_frame.setStyleSheet("border: 1px solid black;")
        self.page1_layout = QHBoxLayout(self.page1_frame)
        self.page1_layout.addWidget(self.text_archivo)
        self.page1_layout.addWidget(self.text_html_analisis)
        self.page1_layout.setStretch(0, 1)
        self.page1_layout.setStretch(1, 1)
        self.stacked_widget.widget(0).setLayout(self.page1_layout)

        self.page2_frame = QFrame()
        self.page2_frame.setFrameShape(QFrame.Box)
        self.page2_frame.setLineWidth(1)
        self.page2_frame.setStyleSheet("border: 1px solid black;")
        self.page2_layout = QHBoxLayout(self.page2_frame)
        self.page2_layout.addWidget(self.text_html_comparacion)
        self.page2_layout.addWidget(self.text_html_analaisis_corpus_2)
        self.page2_layout.addWidget(self.contraste)
        self.page2_layout.setStretch(0, 1)
        self.page2_layout.setStretch(1, 1)
        self.page2_layout.setStretch(2, 1)
        self.stacked_widget.widget(1).setLayout(self.page2_layout)

    def setupButtons(self):
        self.analasis_informe = QPushButton("Analizar textos", self)
        self.analasis_comparacion = QPushButton("Comparar análisis", self)
        self.analasis_comparacion.setEnabled(False)
        self.analasis_informe.setStyleSheet("border: 1px solid black;")
        self.analasis_comparacion.setStyleSheet("border: 1px solid black;")

    def setupTable(self):
        self.tabla_seleccion.setColumnCount(3)
        self.tabla_seleccion.setHorizontalHeaderLabels(['Seleccion', 'Archivo'])
        self.tabla_seleccion.setShowGrid(True)
        self.tabla_seleccion.setStyleSheet("QTableWidget::item { border: 1px solid #A9A9A9; }")
        self.tabla_seleccion.setColumnHidden(0, True)
        self.tabla_seleccion.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabla_seleccion.setColumnHidden(2, True)
        self.tabla_seleccion.horizontalHeader().setDefaultAlignment(Qt.AlignLeft)
        self.tabla_seleccion.itemSelectionChanged.connect(self.cargar_archivo_seleccionado)

    def setupLayout(self):
        spacer_label = QLabel("")
        self.mygrid.addWidget(spacer_label, 1, 0, 1, 2)
        self.mygrid.addWidget(self.stacked_widget, 5, 12, 43, 66)
        self.mygrid.addWidget(self.analasis_comparacion, 51, 70, 2, 8)
        self.mygrid.addWidget(self.analasis_informe, 51, 2, 2, 8)
        self.mygrid.addWidget(self.tabla_seleccion, 6, 2, 42, 10)
        self.mygrid.setContentsMargins(0, 0, 35, 0)


    def onMyToolBarButtonClick(self, s):
            print("click", s)

            
    def cargar_archivo_seleccionado(self):
        inormacion_carga = ""
        lista = {}
        filas_seleccionadas = self.tabla_seleccion.selectedItems()
        filas_unicas = set()

        for fila_total in filas_seleccionadas:
            filas_unicas.add(fila_total.row())
        if len(filas_seleccionadas) >= 0:
            for fila in filas_unicas:
                ruta_completa = self.tabla_seleccion.item(fila, 2).text()  
                contenido = self.obtener_contenido(ruta_completa)
                titulo = f"<h2>{os.path.basename(ruta_completa)}</h2></p>"
                inormacion_carga += f"{titulo}\n\n{contenido}"

                with open(ruta_completa, 'r', encoding='utf-8') as archivo:
                    contenido = archivo.read()
                
                nombre_archivo = os.path.basename(ruta_completa)
                lista[nombre_archivo] = contenido
            
            self.mostrar_documentos(inormacion_carga, self.text_archivo)

    def obtener_contenido(self,ruta_completa):
        _, extension = os.path.splitext(ruta_completa.lower()) 
        contenido = ""

        if extension == ".txt":
            with open(ruta_completa, 'r', encoding='utf-8') as archivo:
                contenido = archivo.read()
                contenido = html.escape(contenido)

        elif extension == ".pdf":
            with fitz.open(ruta_completa) as pdf_document:
                for pagina in pdf_document.pages():
                    contenido += pagina.get_text()
                contenido = html.escape(contenido)

        elif extension == ".docx":
            doc = Document(ruta_completa)
            for paragraph in doc.paragraphs:
                    contenido += paragraph.text + "\n"

        return contenido
        # Mostrar el contenido en el QTextEdit


    
    def cargar_archivo(self):
        informacion = ""
        cuadro_dialogo = QFileDialog()
        cuadro_dialogo.setWindowTitle("Seleccionar Archivos")
        cuadro_dialogo.setFileMode(QFileDialog.ExistingFiles)
        cuadro_dialogo.setNameFilter("Archivos de Texto (*.txt);;Archivos PDF (*.pdf);;Archivos Word (*.docx);;Todos los archivos (*.*)")
        archivos, _ = cuadro_dialogo.getOpenFileNames(self)
        archivos = [archivo for archivo in archivos if archivo.lower().endswith(('.txt', '.pdf', '.docx'))]

        
        if archivos:
            for archivo in archivos:
                nombre_archivo = os.path.basename(archivo)
                nombre_carpeta = os.path.basename(os.path.dirname(archivo))
                ruta_completa = os.path.abspath(archivo)
                fila_seleccion = self.tabla_seleccion.rowCount()
                self.tabla_seleccion.insertRow(fila_seleccion)
                item_seleccion = QTableWidgetItem()
                item_seleccion.setFlags(item_seleccion.flags() | Qt.ItemIsUserCheckable)
                item_seleccion.setCheckState(Qt.Checked)
                self.tabla_seleccion.setItem(fila_seleccion, 0, item_seleccion)
                item_documento = QTableWidgetItem(nombre_archivo)
                self.tabla_seleccion.setItem(fila_seleccion, 1, item_documento)
                self.tabla_seleccion.verticalHeader().setVisible(False) 
                self.tabla_seleccion.setShowGrid(False) 
                item_ruta = QTableWidgetItem(ruta_completa)
                self.tabla_seleccion.setItem(fila_seleccion, 2, item_ruta)

                contenido = self.obtener_contenido(ruta_completa)
                titulo = f"<h2>{os.path.basename(ruta_completa)}</h2></p>"
                informacion += f"{titulo}\n\n{contenido}"
            
    
    def cargar_archivo_único(self):
        lista = {}
        filas_seleccionadas = self.tabla_seleccion.selectedItems()
        filas_unicas = set()

        for fila_total in filas_seleccionadas:
            filas_unicas.add(fila_total.row())

        if len(filas_seleccionadas) >= 0:
            for fila in filas_unicas:
                ruta_completa = self.tabla_seleccion.item(fila, 2).text()  
                contenido = self.obtener_contenido(ruta_completa)

                with open(ruta_completa, 'r', encoding='utf-8') as archivo:
                    contenido = archivo.read()
                
                nombre_archivo = os.path.basename(ruta_completa)
                lista[nombre_archivo] = contenido

            self.algoritmo(lista, "comparacion")
            self.analasis_comparacion.setEnabled(True)
            self.test_de_Wilcoxon()


    def mostrar_documentos(self, informacion, boton):
        boton.setHtml(informacion)
        #boton.setHtml(informacion)
            
    def cargar_proyecto(self):
        carpeta = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta")

        if carpeta:
            extensiones_permitidas = ['.txt', '.pdf', '.docx']
            archivos = [os.path.join(carpeta, archivo) for archivo in os.listdir(carpeta) if os.path.isfile(os.path.join(carpeta, archivo))and any(archivo.endswith(ext) for ext in extensiones_permitidas)]

            for archivo in archivos:
                nombre_archivo = os.path.basename(archivo)
                nombre_carpeta = os.path.basename(os.path.dirname(archivo))
                ruta_completa = os.path.abspath(archivo)
                fila_seleccion = self.tabla_seleccion.rowCount()
                self.tabla_seleccion.insertRow(fila_seleccion)
                item_seleccion = QTableWidgetItem()
                item_seleccion.setFlags(item_seleccion.flags() | Qt.ItemIsUserCheckable)
                item_seleccion.setCheckState(Qt.Checked)
                self.tabla_seleccion.setItem(fila_seleccion, 0, item_seleccion)
                item_documento = QTableWidgetItem(nombre_archivo)
                self.tabla_seleccion.setItem(fila_seleccion, 1, item_documento)
                item_ruta = QTableWidgetItem(ruta_completa)
                self.tabla_seleccion.setItem(fila_seleccion, 2, item_ruta)
                self.tabla_seleccion.verticalHeader().setVisible(False)
                self.tabla_seleccion.setShowGrid(False)      
            ruta_completa = os.path.abspath(nombre_carpeta)
            self.text_direccion.setPlainText(ruta_completa)

            
    def ejecutar_algoritmo(self):
        lista = {}
        filas_seleccionadas = self.tabla_seleccion.selectedItems()
        filas_unicas = set()

        for fila_total in filas_seleccionadas:
            filas_unicas.add(fila_total.row())
        if len(filas_seleccionadas) >= 0:
            for fila in filas_unicas:
                ruta_completa = self.tabla_seleccion.item(fila, 2).text()  
                contenido = self.obtener_contenido(ruta_completa)

                with open(ruta_completa, 'r', encoding='utf-8') as archivo:
                    contenido = archivo.read()
                
                nombre_archivo = os.path.basename(ruta_completa)
                lista[nombre_archivo] = contenido

            self.algoritmo(lista, "informe")
            self.analasis_comparacion.setEnabled(True)
            
    def algoritmo_individual(self, texto):
        palabras_funcionales = [
            "el", "la", "los", "las", "un", "una", "unos", "unas", "de", "del", "al",
            "en", "y", "o", "a", "que", "es", "con", "por", "para", "como", "se", 
            "su", "lo", "no", "sí", "si", "ya", "le", "les", "me", "mi", "mis", 
            "te", "tu", "tus", "nos", "nuestro", "nuestra", "nuestros", "nuestras", 
            "vosotros", "vosotras", "vuestro", "vuestra", "vuestros", "vuestras", 
            "ellos", "ellas", "su", "sus", "esta", "estas", "este", "estos", 
            "ese", "esos", "esa", "esas", "aquel", "aquellos", "aquella", 
            "aquellas", "esto", "eso", "aquello", "mi", "mis", "tu", "tus", 
            "su", "sus", "nuestro", "nuestra", "nuestros", "nuestras", 
            "vuestro", "vuestra", "vuestros", "vuestras"
        ]
        emoticons = [
            ":)", ":-)", ":))", ":-))", ":D", ":-D", ":))", ":)))", "XD", ":')", ":'D", "^_^", "^__^", "^___^", "=)", "=]", ":]",
            ":(", ":-(", ":'(", ":'-(", ":(", ":-(((", "D:", "D-:", "T_T", "T.T", ";_;", "=_=",
            ":o", ":-o", ":O", ":-O", ":0", "8O", "8-O", "O_O", "o_o", "0_0", "0.o", "O.o",
            ":P", ":-P", ":p", ":-p", ":Þ", ":-Þ", ":þ", ":-þ", ":P", ":-P", "xP", "x-p",
            ";)", ";-)", ";]", ";D", ";^)",
            "<3", "<33", "<333", ":*", ":-*", ":-x", ":X", "xoxo",
            ":/", ":-/", ":\\", ":-\\", ":S", ":-S", "o_O", "o.O", "O.o",
            ">:( ", ">:(", "D:<", "D:<", "D-:<", ">:["
        ]


        doc = nlp(texto)
        palabras_por_frase = np.mean([len(sent) for sent in doc.sents])
        parrafos = texto.split('\n')
        palabras_por_parrafo = np.mean([len(nlp(parrafo)) for parrafo in parrafos])
        palabras_por_texto = len(doc)
        frases_por_parrafo = np.mean(([len(list(nlp(parrafo).sents)) for parrafo in parrafos]))
        frases_por_texto = len(list(doc.sents))
        numero_de_parrafos = len(parrafos)
        riqueza_lexica = len(set([token.lemma_ for token in doc])) / palabras_por_texto if palabras_por_texto > 0 else 0 
        palabras_unicas = len(set([token.text.lower() for token in doc if token.is_alpha]))
        palabras_compartidas_una_vez = self.calculo_palabras_compartidas_una_vez(doc)
        palabras_compartidas_mas_de_una_vez = self.calculo_palabras_compartidas_mas_de_una_vez(doc)
        lemas_unicos = self.calculo_lemas_unicos(doc)
        lemas_compartidos_una_vez = self.calculo_lemas_compartidos_una_vez(doc)
        lemas_compartidos_mas_de_una_vez = self.calculo_lemas_compartidos_mas_de_una_vez(doc)
        signos_de_puntuacion = len([token.text for token in doc if token.is_punct])
        ngramas_2 = self.calculo_ngramas(doc, 2)
        ngramas_3 = self.calculo_ngramas(doc, 3)
        ngramas_4 = self.calculo_ngramas(doc, 4)
        function_words = len([token.text.lower() for token in doc if token.text.lower() in palabras_funcionales])
        emoticonos = sum(doc._.emoji, len([token.text for token in doc if token.text in emoticons]))
        palabras_mayusculas = len([token.text for token in doc if token.text.isupper()])
        palabras_con_mayusculas = len([token.text for token in doc if any(c.isupper() for c in token.text)])
        palabras_con_estilos_de_fuente =len([token.text for token in doc if token.is_title])

        resultados = {
            "Palabras por frase": palabras_por_frase,
            "Palabras por párrafo": palabras_por_parrafo,
            "Palabras por texto": palabras_por_texto,
            "Frases por párrafo": frases_por_parrafo,
            "Frases por texto": frases_por_texto,
            "Número de párrafos": numero_de_parrafos,
            "Riqueza léxica": np.around(riqueza_lexica,2),
            "Palabras únicas": palabras_unicas,
            "Palabras compartidas una vez": palabras_compartidas_una_vez,
            "Palabras compartidas más de una vez": palabras_compartidas_mas_de_una_vez,
            "Lemas únicos": lemas_unicos,
            "Lemas compartidos una vez": lemas_compartidos_una_vez,
            "Lemas compartidos más de una vez": lemas_compartidos_mas_de_una_vez,
            "Signos de puntuación": signos_de_puntuacion,
            "N-gramas de tamaño 2": ngramas_2,
            "N-gramas de tamaño 3": ngramas_3,
            "N-gramas de tamaño 4": ngramas_4,
            "Function words": function_words,
            "Emoticonos": emoticonos,
            "Palabras en mayúsculas": palabras_mayusculas,
            "Palabras con mayúsculas": palabras_con_mayusculas,
            "Palabras con estilos de fuente": palabras_con_estilos_de_fuente
        }
        return resultados
    
    def algoritmo(self, diccionario_textos, nombre_JSON):
        resultados = {}
        resultados_agregados = {
            "Media de Palabras por frase": [],
            "Media de Palabras por párrafo": [],
            "Media de Palabras por texto": [],
            "Media de Frases por párrafo": [],
            "Media de Frases por texto": [],
            "Media de Número de párrafos": [],
            "Media de Riqueza léxica": [],
            "Media de Palabras únicas": [],
            "Media de Palabras compartidas una vez": [],
            "Media de Palabras compartidas más de una vez": [],
            "Media de Lemas únicos": [],
            "Media de Lemas compartidos una vez": [],
            "Media de Lemas compartidos más de una vez": [],
            "Media de Signos de puntuación": [],
            "Media de N-gramas de tamaño 2": [],
            "Media de N-gramas de tamaño 3": [],
            "Media de N-gramas de tamaño 4": [],
            "Media de Function words": [],
            "Media de Emoticonos": [],
            "Media de Palabras en mayúsculas": [],
            "Media de Palabras con mayúsculas": [],
            "Media de Palabras con estilos de fuente": [],
            "Desviación de Palabras por párrafo": [],
            "Desviación de Palabras por frase": [],
            "Desviación de Palabras por texto": [],
            "Desviación de Frases por párrafo": [],
            "Desviación de Frases por texto": [],
            "Desviación de Número de párrafos": [],
            "Desviación de Riqueza léxica": [],
            "Desviación de Palabras únicas": [],
            "Desviación de Palabras compartidas una vez": [],
            "Desviación de Palabras compartidas más de una vez": [],
            "Desviación de Lemas únicos": [],
            "Desviación de Lemas compartidos una vez": [],
            "Desviación de Lemas compartidos más de una vez": [],
            "Desviación de Signos de puntuación": [],
            "Desviación de N-gramas de tamaño 2": [],
            "Desviación de N-gramas de tamaño 3": [],
            "Desviación de N-gramas de tamaño 4": [],
            "Desviación de Function words": [],
            "Desviación de Emoticonos": [],
            "Desviación de Palabras en mayúsculas": [],
            "Desviación de Palabras con mayúsculas": [],
            "Desviación de Palabras con estilos de fuente": [],
        }
        resultados_por_texto = {
            "Palabras por frase": {},
            "Palabras por párrafo": {},
            "Palabras por texto": {},
            "Frases por párrafo": {},
            "Frases por texto": {},
            "Número de párrafos": {},
            "Riqueza léxica": {},
            "Palabras únicas": {},
            "Palabras compartidas una vez": {},
            "Palabras compartidas más de una vez": {},
            "Lemas únicos": {},
            "Lemas compartidos una vez": {},
            "Lemas compartidos más de una vez": {},
            "Signos de puntuación": {},
            "N-gramas de tamaño 2": {},
            "N-gramas de tamaño 3": {},
            "N-gramas de tamaño 4": {},
            "Function words": {},
            "Emoticonos": {},
            "Palabras en mayúsculas": {},
            "Palabras con mayúsculas": {},
            "Palabras con estilos de fuente": {}
        }


        for nombre, texto in diccionario_textos.items():
            resultado_individual = self.algoritmo_individual(texto)
            resultados[nombre] = resultado_individual

            for tipo, valor in resultado_individual.items():
                if tipo in resultados_por_texto:
                    resultados_por_texto[tipo][nombre] = valor
            self._agregar_resultados_agregados(resultado_individual, resultados_agregados)

        datos = {"resultados_agregados": resultados_agregados,"resultados_por_texto": resultados_por_texto}

        self._calcular_media_desviacion(resultados_agregados)
        self.JSON(datos, nombre_JSON)
        self.tabla_analisis(nombre_JSON)

        self.resultados_por_texto = resultados_por_texto
        self.resultados_agregados = resultados_agregados

    def _agregar_resultados_agregados(self, resultado_individual, resultados_agregados):
        for key in resultado_individual.keys():
            if isinstance(resultado_individual[key], (int, float)):
                resultados_agregados[f"Media de {key}"].append(resultado_individual[key])
            else:
                resultados_agregados[f"Media de {key}"].extend(resultado_individual[key])

    def _calcular_media_desviacion(self, resultados_agregados):
        for key in list(resultados_agregados.keys()):
            valores = resultados_agregados[key]
            if 'Media' in key:
                media = np.around(np.mean(valores),2)
                resultados_agregados[key] = media
                desviacion_key = key.replace('Media', 'Desviación')
                resultados_agregados[desviacion_key] = np.around(np.std(valores),2)

    def calculo_palabras_compartidas_una_vez(self, doc):
        todas_las_palabras = [token.text.lower() for token in doc if token.is_alpha]

        frecuencia_palabras = Counter(todas_las_palabras)

        palabras_compartidas_una_vez = [palabra for palabra, frecuencia in frecuencia_palabras.items() if frecuencia == 1]

        return len(palabras_compartidas_una_vez)
    
    def calculo_palabras_compartidas_mas_de_una_vez(self, doc):
        todas_las_palabras = [token.text.lower() for token in doc if token.is_alpha]

        frecuencia_palabras = Counter(todas_las_palabras)

        palabras_compartidas_mas_de_una_vez = [palabra for palabra, frecuencia in frecuencia_palabras.items() if frecuencia > 1]

        return len(palabras_compartidas_mas_de_una_vez)
        
    def calculo_lemas_unicos(self, doc):
        todos_los_lemas = [token.lemma_.lower() for token in doc if token.is_alpha]

        frecuencia_lemas = Counter(todos_los_lemas)

        lemas_unicos = len(frecuencia_lemas)

        return lemas_unicos


    def calculo_lemas_compartidos_una_vez(self, doc):
        todos_los_lemas = [token.lemma_.lower() for token in doc if token.is_alpha]

        frecuencia_lemas = Counter(todos_los_lemas)

        lemas_compartidos_una_vez = [lema for lema, frecuencia in frecuencia_lemas.items() if frecuencia == 1]

        return len(lemas_compartidos_una_vez)

    def calculo_lemas_compartidos_mas_de_una_vez(self, doc):
        todos_los_lemas = [token.lemma_.lower() for token in doc if token.is_alpha]

        frecuencia_lemas = Counter(todos_los_lemas)

        lemas_compartidos_mas_de_una_vez = [lema for lema, frecuencia in frecuencia_lemas.items() if frecuencia > 1]

        return len(lemas_compartidos_mas_de_una_vez)
        
    def calculo_ngramas(self, doc, n):
        
        def get_ngrams(tokens, n):
            ngrams = zip(*[tokens[i:] for i in range(n)])
            return [' '.join(ngram) for ngram in ngrams]

        ngramas = list(get_ngrams([token.text.lower() for token in doc if token.is_alpha], n))
        return len(ngramas)
    
    def JSON(self, datos, nombre):
        
        with open(f'{nombre}.json', 'w', encoding='utf-8') as f:
            json.dump(datos, f, indent=4)

    def abrir_datos_desde_json(self, nombre):
        archivo_json = f'{nombre}.json'
        
        try:
            with open(archivo_json, 'r', encoding='utf-8') as f:
                datos = json.load(f)
            if nombre == "chatGPT":
                return datos
            else:
                resultados_agregados = datos.get('resultados_agregados', {})
                resultados_por_texto = datos.get('resultados_por_texto', {})
            return resultados_agregados, resultados_por_texto
        
        except FileNotFoundError:
            self.text_html_analisis.setText(f"El archivo '{archivo_json}' no se encontró.")
        except json.JSONDecodeError as e:
            self.text_html_analisis.setText(f"Error al decodificar JSON: {str(e)}")


    def tabla_analisis(self, tipo):
        resultados_agregados,resultados_individuales = self.abrir_datos_desde_json(tipo)
        valores = list(resultados_individuales.items())
        midpoint = len(valores) // 2
        diccionario1 = dict(valores[:midpoint])
        diccionario2 = dict(valores[midpoint:])
        # Crear el HTML para la tabla
        valores1 = list(diccionario1.items())
        midpoint1 = len(valores1) // 2
        subdic1_1 = dict(valores1[:midpoint1])
        subdic1_2 = dict(valores1[midpoint1:])

        valores2 = list(diccionario2.items())
        midpoint2 = len(valores2) // 2
        subdic2_1 = dict(valores2[:midpoint2])
        subdic2_2 = dict(valores2[midpoint2:])

        def generar_tabla(diccionario, titulo):
            html = f"<h1>{titulo}</h1>"
            html += """
            <style>
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 20px;
                }
                th, td {
                    border: 1px solid black;
                    padding: 8px;
                    text-align: center;
                }
                th {
                    background-color: #f2f2f2;
                }
            </style>
            <table>
                <thead>
                    <tr>
                        <th>Nombre del Texto</th>
            """
            
            for variable in diccionario.keys():
                html += f"<th>{variable}</th>"
            
            html += "</tr></thead><tbody>"
            
            primer_variable = next(iter(diccionario.keys()))
            for texto in diccionario[primer_variable].keys():
                html += f"<tr><td>{texto}</td>"
                for variable in diccionario.keys():
                    resultado = str(round(diccionario[variable].get(texto, "N/A"), 2))
                    html += f"<td>{resultado}</td>"
                html += "</tr>"
            
            html += "</tbody></table>"
            return html

        html1_1 = generar_tabla(subdic1_1,"Tabla Datos")
        html1_2 = generar_tabla(subdic1_2,"Tabla Datos")
        html2_1 = generar_tabla(subdic2_1,"Tabla Datos")
        html2_2 = generar_tabla(subdic2_2,"Tabla Datos")

        html3 = """
            <h1>Tabla Comparativa</h1>
            <style>
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 20px;
                }
                th, td {
                    border: 1px solid black;
                    padding: 8px;
                    text-align: center;
                }
                th {
                    background-color: #f2f2f2;
                }
                .p_valor {
                height: {len(resultados_agregados_informe) * 50}px;
                vertical-align: middle;
            }
            </style>
            <table>
                <thead>
                    <tr>
                        <th>Tipo de dato</th>
                        <th>Media</th>
                        <th>Desviacion</th>
                    </tr>
                </thead>
                <tbody>
            """

        # Agregar filas de datos
        diccionario_media = []
        diccionario_desviacion = []
        keys = list(resultados_agregados.keys())
        for i, key in enumerate(keys):
            if "Media" in key:       
                diccionario_media.append(key)
            elif "Desviación" in key:
                diccionario_desviacion.append(key)
        keys = list(resultados_agregados.keys())

        for media, desviacion in zip(diccionario_media, diccionario_desviacion):
            key = media.replace("Media de ","")
            html3 += f"<tr><td>{key}</td><td>{resultados_agregados[media]}</td><td>{resultados_agregados[desviacion]}</td></tr>"
        html3 += "</tbody></table>"

        self.html = html1_1 +  html1_2 +  html2_1 +  html2_2 +  html3
        # Insertar el HTML en el QTextEdit
        imagenes = self.analisisGraficas(resultados_individuales)
        if tipo == "informe":
            self.text_html_analisis.setHtml("")
            self.text_html_comparacion.setHtml("")
            cursor1 = self.text_html_analisis.textCursor()
            cursor2 = self.text_html_comparacion.textCursor()
            for imagen in imagenes:
                image_format = QTextImageFormat()
                image_format.setName(imagen)  
                cursor1.insertImage(image_format)
                cursor1.insertText("\n\n")  
                cursor2.insertImage(image_format)
                cursor2.insertText("\n\n")  
            cursor1.insertText("\n\n\n\n\n\n\n\n")  
            cursor1.insertHtml(self.html)
            cursor2.insertHtml(self.html)
        else:
            self.text_html_analaisis_corpus_2.setHtml("")
            cursor = self.text_html_analaisis_corpus_2.textCursor()
            for imagen in imagenes:
                image_format = QTextImageFormat()
                image_format.setName(imagen)  
                cursor.insertImage(image_format)
                cursor.insertText("\n\n")  
            cursor.insertText("\n\n")  
            cursor.insertHtml(self.html)
        


    def tabla_comparacion(self, resultados_agregados_informe, resultados_agregados_comparacion, p_valor_final):

            html = """
            <h1>Tabla Comparativa</h1>
            <style>
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 20px;
                }
                th, td {
                    border: 1px solid black;
                    padding: 8px;
                    text-align: center;
                }
                th {
                    background-color: #f2f2f2;
                }
                .p_valor {
                height: {len(resultados_agregados_informe) * 50}px;
                vertical-align: middle;
            }
            </style>
            <table>
                <thead>
                    <tr>
                        <th>Tipo de dato</th>
                        <th>Corpus 1</th>
                        <th>Corpus 2</th>
                        <th>Diferencia</th>
                        <th>p_valor Mann-Whitney U</th>
                    </tr>
                </thead>
                <tbody>
            """
            keys = list(resultados_agregados_comparacion.keys())
            for i, key in enumerate(keys):
                if "Media" in key:
                    if resultados_agregados_comparacion[key] < resultados_agregados_informe[key]:
                        if resultados_agregados_comparacion[key] != 0:
                            diferencia = ((resultados_agregados_informe[key]-resultados_agregados_comparacion[key])/(resultados_agregados_comparacion[key]))*100
                        else:
                            diferencia = 0
                    else:
                        if resultados_agregados_informe[key] != 0:
                            diferencia = ((resultados_agregados_comparacion[key]-resultados_agregados_informe[key])/(resultados_agregados_informe[key]))*100
                        else:
                            diferencia = 0
                    html += f"<tr><td>{key}</td><td>{resultados_agregados_informe[key]}</td><td>{resultados_agregados_comparacion[key]}</td> <td>{diferencia:.2f} %</td> <td>{p_valor_final[key].pvalue:.2f}</td></tr>"

            html += "</tbody></table>"


            self.contraste.setHtml(html)
            self.crear_graficos_comparativos(resultados_agregados_informe, resultados_agregados_comparacion)



    def test_de_Wilcoxon(self):
        p_valor_final = {}
        resultados_agregados_informe,resultados_por_texto_informe = self.abrir_datos_desde_json("informe")
        resultados_agregados_comparacion,resultados_por_texto_comparacion = self.abrir_datos_desde_json("comparacion")

        for key in list(resultados_por_texto_informe.keys()):
            p_valor_final[f"Media de {key}"] = mannwhitneyu( 
                list(resultados_por_texto_informe[key].values()), 
                list(resultados_por_texto_comparacion[key].values()), 
                alternative='two-sided')
        self.tabla_comparacion(resultados_agregados_informe, resultados_agregados_comparacion, p_valor_final)

    def analisisGraficas(self, resultados_por_texto):
        colores = plt.cm.get_cmap('tab20', len(resultados_por_texto))
        imagenes = [] 
        
        for idx, (key, valores) in enumerate(resultados_por_texto.items()):
            fig, ax = plt.subplots(figsize=(6, 4)) 
            archivos = list(valores.keys())
            puntajes = list(valores.values())
            color = colores(idx)
            ax.bar(archivos, puntajes, color=color)
            ax.set_title(key)
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()  
            
            
            temp_file = f"temp_{key}.png"
            plt.savefig(temp_file, format='png', bbox_inches='tight')
            plt.close(fig)
            imagenes.append(temp_file)
        
        return imagenes
        

    def crear_graficos_comparativos(self, resultados_agregados1, resultados_agregados2):
        claves = list(resultados_agregados1.keys())

        for clave in claves:
            if "Media" in clave:
                valor1 = resultados_agregados1[clave]
                valor2 = resultados_agregados2[clave]
                fig, ax = plt.subplots(figsize=(6, 4))
                ax.bar(["Corpus 1"], [valor1], label='Corpus 1', color='blue')
                ax.bar(["Corpus 2"], [valor2], label='Corpus 2', color='orange')
                ax.set_title(clave)
                ax.legend()
                plt.tight_layout()
                buf = io.BytesIO()
                plt.savefig(buf, format='png')
                plt.close(fig)
                buf.seek(0)
                qimage = QImage()
                qimage.loadFromData(buf.getvalue(), "PNG")
                cursor = self.contraste.textCursor()
                cursor.insertImage(QPixmap.fromImage(qimage).toImage())
                cursor.insertText('\n\n') 
        cursor.insertText('\n\n') 



    
    def extraer_html_informe(self):
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        archivo, _ = QFileDialog.getSaveFileName(None, "Guardar informe como PDF", "", "PDF Files (*.pdf);;All Files (*)", options=options)
        
        if archivo: 
            if not archivo.lower().endswith('.pdf'):
                archivo += f'.pdf'
            printer.setOutputFileName(archivo)
            self.text_html_analisis.document().print_(printer)
            
            print(f"Contenido del informe guardado en {archivo} con contenido {self.text_html_analisis}")
        else:
            print("No se seleccionó ningún directorio o archivo")

    def extraer_html_comparacion(self):
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        archivo, _ = QFileDialog.getSaveFileName(None, "Guardar informe como PDF", "", "PDF Files (*.pdf);;All Files (*)", options=options)
        
        if archivo: 
            if not archivo.lower().endswith('.pdf'):
                archivo += f'.pdf'
            printer.setOutputFileName(archivo)
            self.contraste.document().print_(printer)
            print(f"Contenido del informe guardado en {archivo} con contenido {self.contraste}")
        else:
            print("No se seleccionó ningún directorio o archivo")

                
if __name__ == "__main__":
    nlp.add_pipe("emoji", first=True)
    if not QApplication.instance():
        app = QApplication(sys.argv)

    else:
        app = QApplication.instance()
        app.setStyle(QStyleFactory.create('Cleanlooks'))

    ventana_principal = Principal()
    ventana_principal.show()
    sys.exit(app.exec())
