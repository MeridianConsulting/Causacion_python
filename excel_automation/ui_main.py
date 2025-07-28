#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interfaz Gráfica Principal - Automatización Excel
"""

import sys
from pathlib import Path
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                              QHBoxLayout, QLabel, QPushButton, QFrame, 
                              QMessageBox, QProgressBar, QTextEdit, QFileDialog)
from PySide6.QtCore import Qt, QThread, Signal, QMimeData, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont, QPalette, QColor, QDragMoveEvent

from .excel_processor import ExcelProcessor

class DropArea(QWidget):
    """Widget interno para el área de drag & drop"""
    
    file_dropped = Signal(str)
    
    def __init__(self, parent_zone):
        super().__init__()
        self.parent_zone = parent_zone
        self.drag_active = False
        self.setup_ui()
        
    def setup_ui(self):
        """Configurar la interfaz del área de drop"""
        self.setMinimumHeight(150)
        self.setAcceptDrops(True)
        
        # Estilo visual para el área de drop
        self.setStyleSheet("""
            QWidget {
                border: 3px dashed #cccccc;
                border-radius: 12px;
                background-color: #f8f9fa;
                color: #666666;
                margin: 0px;
                padding: 15px;
            }
            QWidget:hover {
                border-color: #007acc;
                background-color: #e7f3ff;
            }
        """)
        
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)
        
        # Label principal con instrucciones claras
        self.file_label = QLabel("🖱️ Arrastra un archivo Excel aquí\n📁 o haz clic para seleccionar\n\n💡 Formatos: .xlsx, .xls")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setStyleSheet("color: #999999; font-style: italic; font-size: 13px; line-height: 1.6;")
        
        # Botón para seleccionar archivo (como alternativa)
        self.select_button = QPushButton("📁 Buscar archivo")
        self.select_button.setMinimumHeight(35)
        self.select_button.clicked.connect(self.select_file)
        self.select_button.setStyleSheet("""
            QPushButton {
                background-color: #6c757d;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-size: 12px;
                font-weight: normal;
                margin-top: 5px;
            }
            QPushButton:hover {
                background-color: #5a6268;
            }
        """)
        
        layout.addWidget(self.file_label)
        layout.addWidget(self.select_button)
        
        self.setLayout(layout)
    
    def select_file(self):
        """Abrir diálogo para seleccionar archivo"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            f"Seleccionar archivo {self.parent_zone.title}",
            "",
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*.*)"
        )
        
        if file_path:
            self.handle_file_selection(file_path)
    
    def handle_file_selection(self, file_path):
        """Manejar la selección de archivo (desde botón o drag & drop)"""
        try:
            self.parent_zone.file_path = file_path
            file_name = Path(file_path).name
            
            # Actualizar UI con estilo visual mejorado
            self.file_label.setText(f"📄 {file_name}\n✅ Archivo cargado correctamente")
            self.file_label.setStyleSheet("color: #28a745; font-weight: bold; font-size: 13px;")
            
            self.select_button.setText("🔄 Cambiar archivo")
            self.select_button.setStyleSheet("""
                QPushButton {
                    background-color: #28a745;
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    border-radius: 6px;
                    font-size: 12px;
                    font-weight: normal;
                    margin-top: 5px;
                }
                QPushButton:hover {
                    background-color: #218838;
                }
            """)
            
            # Cambiar estilo del widget cuando hay archivo cargado
            self.setStyleSheet("""
                QWidget {
                    border: 3px solid #28a745;
                    border-radius: 12px;
                    background-color: #d4edda;
                    color: #155724;
                    margin: 0px;
                    padding: 15px;
                }
            """)
            
            # Emitir señal
            self.file_dropped.emit(file_path)
            print(f"✅ Archivo seleccionado: {file_name}")
            
        except Exception as e:
            print(f"❌ Error al seleccionar archivo: {e}")
            QMessageBox.warning(self, "Error", f"Error al procesar el archivo: {e}")

    def dragEnterEvent(self, event: QDragEnterEvent):
        """Evento cuando un archivo entra en la zona de drop"""
        print("🔍 DRAG ENTER - Iniciando detección...")
        
        # Verificar si tiene URLs
        if not event.mimeData().hasUrls():
            print("❌ No hay URLs en el drag")
            event.ignore()
            return
            
        urls = event.mimeData().urls()
        print(f"📂 URLs detectadas: {len(urls)}")
        
        if not urls:
            print("❌ Lista de URLs vacía")
            event.ignore()
            return
            
        # Obtener el primer archivo
        file_url = urls[0]
        file_path = file_url.toLocalFile()
        print(f"📁 Archivo detectado: {file_path}")
        
        # Verificar si es un archivo Excel
        if not file_path or not file_path.lower().endswith(('.xlsx', '.xls')):
            print(f"❌ No es archivo Excel: {file_path}")
            event.ignore()
            return
            
        # Aceptar el drag
        print("✅ ARCHIVO EXCEL VÁLIDO - Aceptando drag")
        self.drag_active = True
        event.setDropAction(Qt.CopyAction)
        event.accept()
        
        # Cambiar estilo visual cuando se detecta archivo Excel
        self.setStyleSheet("""
            QWidget {
                border: 3px dashed #28a745;
                border-radius: 12px;
                background-color: #d4edda;
                color: #155724;
                margin: 0px;
                padding: 15px;
            }
        """)
        
    def dragMoveEvent(self, event: QDragMoveEvent):
        """Evento cuando se mueve el archivo sobre la zona"""
        print("🔄 DRAG MOVE")
        
        if self.drag_active and event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith(('.xlsx', '.xls')):
                event.setDropAction(Qt.CopyAction)
                event.accept()
                return
                
        event.ignore()
                
    def dragLeaveEvent(self, event):
        """Evento cuando un archivo sale de la zona de drop"""
        print("🚪 DRAG LEAVE")
        self.drag_active = False
        
        # Restaurar estilo solo si no hay archivo seleccionado
        if not self.parent_zone.file_path:
            self.restore_normal_style()
        
    def dropEvent(self, event: QDropEvent):
        """Evento cuando se suelta un archivo"""
        print("🎯 DROP EVENT - Procesando archivo...")
        
        self.drag_active = False
        
        if not event.mimeData().hasUrls():
            print("❌ No hay URLs en drop")
            event.ignore()
            return
            
        urls = event.mimeData().urls()
        if not urls:
            print("❌ Lista de URLs vacía en drop")
            event.ignore()
            return
            
        file_path = urls[0].toLocalFile()
        file_path = str(Path(file_path).resolve())
        print(f"📁 Archivo para procesar: {file_path}")
        
        if file_path and file_path.lower().endswith(('.xlsx', '.xls')):
            print("✅ PROCESANDO ARCHIVO EXCEL")
            
            # Procesar archivo
            self.handle_file_selection(file_path)
            event.setDropAction(Qt.CopyAction)
            event.accept()
            
            print("🎉 DROP COMPLETADO EXITOSAMENTE")
        else:
            print("❌ Archivo no válido en drop")
            self.restore_normal_style()
            event.ignore()
            
    def restore_normal_style(self):
        """Restaurar el estilo normal del widget"""
        self.setStyleSheet("""
            QWidget {
                border: 3px dashed #cccccc;
                border-radius: 12px;
                background-color: #f8f9fa;
                color: #666666;
                margin: 0px;
                padding: 15px;
            }
            QWidget:hover {
                border-color: #007acc;
                background-color: #e7f3ff;
            }
        """)

class DropZone(QWidget):
    """Widget para arrastrar y soltar archivos"""
    
    file_dropped = Signal(str)
    
    def __init__(self, title: str, description: str):
        super().__init__()
        self.title = title
        self.description = description
        self.file_path = None
        self.setup_ui()
        
    def setup_ui(self):
        """Configurar la interfaz de la zona de drop"""
        # Configuración básica del widget (SIN drag & drop)
        self.setMinimumHeight(200)
        self.setMinimumWidth(350)
        
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(8, 8, 8, 8)
        
        # Título (NO arrastrable)
        title_label = QLabel(self.title)
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                background-color: #ecf0f1;
                padding: 8px;
                border-radius: 6px;
                margin-bottom: 5px;
            }
        """)
        
        # Descripción (NO arrastrable)
        desc_label = QLabel(self.description)
        desc_label.setAlignment(Qt.AlignCenter)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("""
            QLabel {
                color: #7f8c8d;
                font-size: 12px;
                margin-bottom: 10px;
            }
        """)
        
        # Área de drop (SÍ arrastrable)
        self.drop_area = DropArea(self)
        self.drop_area.file_dropped.connect(self.file_dropped.emit)
        
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        layout.addWidget(self.drop_area)
        
        self.setLayout(layout)

class ProcessingThread(QThread):
    """Hilo para procesar archivos sin bloquear la UI"""
    
    progress = Signal(str)
    finished = Signal(bool, str)
    
    def __init__(self, token_file: str, movement_file: str):
        super().__init__()
        self.token_file = token_file
        self.movement_file = movement_file
        
    def run(self):
        """Ejecutar el procesamiento"""
        try:
            processor = ExcelProcessor()
            
            self.progress.emit("Leyendo archivo Token DIAN...")
            token_df = processor.read_excel(Path(self.token_file))
            
            self.progress.emit("Leyendo archivo Movimiento Contable...")
            movement_df = processor.read_excel(Path(self.movement_file))
            
            self.progress.emit("Procesando datos...")
            # Aquí puedes agregar tu lógica específica de procesamiento
            processed_token = processor.process_data(token_df)
            processed_movement = processor.process_data(movement_df)
            
            self.progress.emit("Guardando resultados...")
            from config import Config
            
            output_token = Config.OUTPUT_PATH / "token_dian_procesado.xlsx"
            output_movement = Config.OUTPUT_PATH / "movimiento_contable_procesado.xlsx"
            
            processor.write_excel(processed_token, output_token, "Token_DIAN")
            processor.write_excel(processed_movement, output_movement, "Movimiento_Contable")
            
            self.finished.emit(True, "Procesamiento completado exitosamente")
            
        except Exception as e:
            self.finished.emit(False, f"Error durante el procesamiento: {str(e)}")

class MainWindow(QMainWindow):
    """Ventana principal de la aplicación"""
    
    def __init__(self):
        super().__init__()
        self.token_file = None
        self.movement_file = None
        self.processing_thread = None
        self.setup_ui()
        
    def setup_ui(self):
        """Configurar la interfaz principal"""
        self.setWindowTitle("Automatización Excel - Token DIAN & Movimiento Contable")
        self.setMinimumSize(800, 600)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QPushButton {
                background-color: #007acc;
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 6px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #005a9f;
            }
            QPushButton:pressed {
                background-color: #004080;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QVBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(30, 30, 30, 30)
        
        # Título
        title_label = QLabel("🔄 Automatización de Procesos Excel")
        title_font = QFont()
        title_font.setPointSize(20)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 20px;")
        
        # Layout horizontal para las zonas de drop
        drop_layout = QHBoxLayout()
        drop_layout.setSpacing(20)
        
        # Zona de drop para Token DIAN
        self.token_drop = DropZone(
            "Token DIAN",
            "Archivo con información de tokens DIAN"
        )
        self.token_drop.file_dropped.connect(self.on_token_file_dropped)
        
        # Zona de drop para Movimiento Contable
        self.movement_drop = DropZone(
            "Movimiento Contable",
            "Archivo con movimientos contables"
        )
        self.movement_drop.file_dropped.connect(self.on_movement_file_dropped)
        
        drop_layout.addWidget(self.token_drop)
        drop_layout.addWidget(self.movement_drop)
        
        # Botón procesar
        self.process_btn = QPushButton("🚀 Procesar Archivos")
        self.process_btn.setMinimumHeight(50)
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #cccccc;
                border-radius: 5px;
                text-align: center;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #007acc;
                border-radius: 3px;
            }
        """)
        
        # Área de log
        self.log_area = QTextEdit()
        self.log_area.setMaximumHeight(150)
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            QTextEdit {
                border: 1px solid #cccccc;
                border-radius: 5px;
                background-color: #f8f9fa;
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 12px;
            }
        """)
        self.log_area.setVisible(False)
        
        # Agregar widgets al layout principal
        main_layout.addWidget(title_label)
        main_layout.addLayout(drop_layout)
        main_layout.addWidget(self.process_btn)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.log_area)
        main_layout.addStretch()
        
        central_widget.setLayout(main_layout)
        
    def on_token_file_dropped(self, file_path: str):
        """Manejar archivo Token DIAN seleccionado"""
        self.token_file = file_path
        self.log_message(f"Token DIAN cargado: {Path(file_path).name}")
        self.check_ready_to_process()
        
    def on_movement_file_dropped(self, file_path: str):
        """Manejar archivo Movimiento Contable seleccionado"""
        self.movement_file = file_path
        self.log_message(f"Movimiento Contable cargado: {Path(file_path).name}")
        self.check_ready_to_process()
        
    def check_ready_to_process(self):
        """Verificar si ambos archivos están cargados"""
        if self.token_file and self.movement_file:
            self.process_btn.setEnabled(True)
            self.process_btn.setText("🚀 Procesar Archivos")
        else:
            self.process_btn.setEnabled(False)
            missing = []
            if not self.token_file:
                missing.append("Token DIAN")
            if not self.movement_file:
                missing.append("Movimiento Contable")
            self.process_btn.setText(f"⏳ Faltan: {', '.join(missing)}")
            
    def log_message(self, message: str):
        """Agregar mensaje al log"""
        self.log_area.setVisible(True)
        self.log_area.append(f"• {message}")
        
    def process_files(self):
        """Iniciar el procesamiento de archivos"""
        if not self.token_file or not self.movement_file:
            QMessageBox.warning(self, "Archivos faltantes", 
                              "Por favor, selecciona ambos archivos antes de procesar.")
            return
            
        # Configurar UI para procesamiento
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Barra de progreso indeterminada
        self.log_area.setVisible(True)
        self.log_area.clear()
        
        # Iniciar procesamiento en hilo separado
        self.processing_thread = ProcessingThread(self.token_file, self.movement_file)
        self.processing_thread.progress.connect(self.log_message)
        self.processing_thread.finished.connect(self.on_processing_finished)
        self.processing_thread.start()
        
    def on_processing_finished(self, success: bool, message: str):
        """Manejar finalización del procesamiento"""
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        self.process_btn.setText("🚀 Procesar Archivos")
        
        if success:
            self.log_message("✅ " + message)
            QMessageBox.information(self, "Éxito", message)
        else:
            self.log_message("❌ " + message)
            QMessageBox.critical(self, "Error", message)

def run_app():
    """Ejecutar la aplicación"""
    app = QApplication(sys.argv)
    app.setApplicationName("Automatización Excel")
    app.setApplicationVersion("1.0.0")
    
    window = MainWindow()
    window.show()
    
    return app.exec() 