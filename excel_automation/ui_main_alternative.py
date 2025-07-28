#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interfaz Gráfica Alternativa - Sin Drag & Drop
Para mejor compatibilidad en todos los sistemas
"""

import sys
from pathlib import Path
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                              QHBoxLayout, QLabel, QPushButton, QFrame, 
                              QMessageBox, QProgressBar, QTextEdit, QFileDialog)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont

from .excel_processor import ExcelProcessor

class FileSelector(QFrame):
    """Widget selector de archivos con botones"""
    
    file_selected = Signal(str)
    
    def __init__(self, title: str, description: str, file_type: str):
        super().__init__()
        self.title = title
        self.description = description
        self.file_type = file_type
        self.file_path = None
        self.setup_ui()
        
    def setup_ui(self):
        """Configurar la interfaz del selector"""
        self.setFrameStyle(QFrame.Box)
        self.setLineWidth(2)
        self.setMinimumHeight(200)
        self.setStyleSheet("""
            QFrame {
                border: 2px solid #e0e0e0;
                border-radius: 10px;
                background-color: #fafafa;
                color: #333333;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 15px 25px;
                border-radius: 8px;
                font-size: 14px;
                font-weight: bold;
                margin: 10px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(15)
        
        # Título
        title_label = QLabel(self.title)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 10px;")
        
        # Descripción
        desc_label = QLabel(self.description)
        desc_label.setAlignment(Qt.AlignCenter)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #7f8c8d; font-size: 12px;")
        
        # Botón para seleccionar archivo
        self.select_btn = QPushButton(f"📁 Seleccionar {self.file_type}")
        self.select_btn.clicked.connect(self.select_file)
        
        # Label de archivo seleccionado
        self.file_label = QLabel("Ningún archivo seleccionado")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("color: #95a5a6; font-style: italic; font-size: 11px;")
        
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        layout.addWidget(self.select_btn)
        layout.addWidget(self.file_label)
        
        self.setLayout(layout)
        
    def select_file(self):
        """Abrir diálogo para seleccionar archivo"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            f"Seleccionar archivo {self.file_type}",
            str(Path.home()),
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*.*)"
        )
        
        if file_path:
            self.file_path = file_path
            file_name = Path(file_path).name
            
            # Actualizar UI
            self.file_label.setText(f"✅ {file_name}")
            self.file_label.setStyleSheet("color: #27ae60; font-weight: bold; font-size: 12px;")
            
            self.select_btn.setText(f"✅ {self.file_type} Seleccionado")
            self.select_btn.setStyleSheet("""
                QPushButton {
                    background-color: #27ae60;
                    color: white;
                    border: none;
                    padding: 15px 25px;
                    border-radius: 8px;
                    font-size: 14px;
                    font-weight: bold;
                    margin: 10px;
                }
                QPushButton:hover {
                    background-color: #229954;
                }
            """)
            
            # Cambiar estilo del frame
            self.setStyleSheet("""
                QFrame {
                    border: 2px solid #27ae60;
                    border-radius: 10px;
                    background-color: #f0f8f0;
                    color: #333333;
                }
            """ + self.select_btn.styleSheet())
            
            self.file_selected.emit(file_path)

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
            
            self.progress.emit("📖 Leyendo archivo Token DIAN...")
            token_df = processor.read_excel(Path(self.token_file))
            
            self.progress.emit("📖 Leyendo archivo Movimiento Contable...")
            movement_df = processor.read_excel(Path(self.movement_file))
            
            self.progress.emit("⚙️ Procesando datos...")
            processed_token = processor.process_data(token_df)
            processed_movement = processor.process_data(movement_df)
            
            self.progress.emit("💾 Guardando resultados...")
            from config import Config
            
            output_token = Config.OUTPUT_PATH / "token_dian_procesado.xlsx"
            output_movement = Config.OUTPUT_PATH / "movimiento_contable_procesado.xlsx"
            
            processor.write_excel(processed_token, output_token, "Token_DIAN")
            processor.write_excel(processed_movement, output_movement, "Movimiento_Contable")
            
            self.finished.emit(True, "✅ Procesamiento completado exitosamente")
            
        except Exception as e:
            self.finished.emit(False, f"❌ Error durante el procesamiento: {str(e)}")

class AlternativeMainWindow(QMainWindow):
    """Ventana principal alternativa sin drag & drop"""
    
    def __init__(self):
        super().__init__()
        self.token_file = None
        self.movement_file = None
        self.processing_thread = None
        self.setup_ui()
        
    def setup_ui(self):
        """Configurar la interfaz principal"""
        self.setWindowTitle("🔄 Automatización Excel - Token DIAN & Movimiento Contable")
        self.setMinimumSize(900, 700)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QPushButton#process_btn {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 8px;
                font-size: 16px;
                font-weight: bold;
                margin: 20px 0px;
            }
            QPushButton#process_btn:hover {
                background-color: #2980b9;
            }
            QPushButton#process_btn:pressed {
                background-color: #21618c;
            }
            QPushButton#process_btn:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QVBoxLayout()
        main_layout.setSpacing(25)
        main_layout.setContentsMargins(40, 40, 40, 40)
        
        # Título principal
        title_label = QLabel("🔄 Automatización de Procesos Excel")
        title_font = QFont()
        title_font.setPointSize(24)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 30px;")
        
        # Subtítulo
        subtitle_label = QLabel("Procesamiento de Token DIAN y Movimiento Contable")
        subtitle_font = QFont()
        subtitle_font.setPointSize(12)
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setStyleSheet("color: #7f8c8d; margin-bottom: 20px;")
        
        # Layout horizontal para los selectores
        selectors_layout = QHBoxLayout()
        selectors_layout.setSpacing(30)
        
        # Selector para Token DIAN
        self.token_selector = FileSelector(
            "🏛️ Token DIAN", 
            "Selecciona el archivo Excel\ncon información de tokens DIAN",
            "Token DIAN"
        )
        self.token_selector.file_selected.connect(self.on_token_file_selected)
        
        # Selector para Movimiento Contable
        self.movement_selector = FileSelector(
            "📊 Movimiento Contable",
            "Selecciona el archivo Excel\ncon movimientos contables",
            "Movimiento Contable"
        )
        self.movement_selector.file_selected.connect(self.on_movement_file_selected)
        
        selectors_layout.addWidget(self.token_selector)
        selectors_layout.addWidget(self.movement_selector)
        
        # Botón procesar
        self.process_btn = QPushButton("🚀 Procesar Archivos")
        self.process_btn.setObjectName("process_btn")
        self.process_btn.setMinimumHeight(60)
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(25)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                text-align: center;
                font-weight: bold;
                font-size: 12px;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 6px;
            }
        """)
        
        # Área de log
        self.log_area = QTextEdit()
        self.log_area.setMaximumHeight(180)
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            QTextEdit {
                border: 2px solid #ecf0f1;
                border-radius: 8px;
                background-color: #f8f9fa;
                font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
                font-size: 12px;
                padding: 10px;
            }
        """)
        self.log_area.setVisible(False)
        
        # Agregar widgets al layout principal
        main_layout.addWidget(title_label)
        main_layout.addWidget(subtitle_label)
        main_layout.addLayout(selectors_layout)
        main_layout.addWidget(self.process_btn)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.log_area)
        main_layout.addStretch()
        
        central_widget.setLayout(main_layout)
        
        # Mensaje de bienvenida
        self.log_message("💡 Selecciona ambos archivos Excel para comenzar el procesamiento")
        
    def on_token_file_selected(self, file_path: str):
        """Manejar archivo Token DIAN seleccionado"""
        self.token_file = file_path
        self.log_message(f"🏛️ Token DIAN cargado: {Path(file_path).name}")
        self.check_ready_to_process()
        
    def on_movement_file_selected(self, file_path: str):
        """Manejar archivo Movimiento Contable seleccionado"""
        self.movement_file = file_path
        self.log_message(f"📊 Movimiento Contable cargado: {Path(file_path).name}")
        self.check_ready_to_process()
        
    def check_ready_to_process(self):
        """Verificar si ambos archivos están cargados"""
        if self.token_file and self.movement_file:
            self.process_btn.setEnabled(True)
            self.process_btn.setText("🚀 Procesar Archivos")
            self.log_message("✅ Ambos archivos cargados. ¡Listo para procesar!")
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
            QMessageBox.warning(self, "⚠️ Archivos faltantes", 
                              "Por favor, selecciona ambos archivos antes de procesar.")
            return
            
        # Configurar UI para procesamiento
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Barra de progreso indeterminada
        self.log_area.clear()
        self.log_message("🚀 Iniciando procesamiento...")
        
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
            self.log_message(message)
            QMessageBox.information(self, "🎉 Éxito", 
                                  message + f"\n\n📂 Los archivos se guardaron en:\n{Path('data/output').resolve()}")
        else:
            self.log_message(message)
            QMessageBox.critical(self, "❌ Error", message)

def run_alternative_app():
    """Ejecutar la aplicación alternativa"""
    app = QApplication(sys.argv)
    app.setApplicationName("Automatización Excel")
    app.setApplicationVersion("1.0.0")
    
    window = AlternativeMainWindow()
    window.show()
    
    return app.exec() 