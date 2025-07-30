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
from .causacion_processor import CausacionProcessor

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
    """Hilo para procesar archivos de causación sin bloquear la UI"""
    
    progress = Signal(str)
    finished = Signal(bool, str, dict)  # Agregar estadísticas al signal
    
    def __init__(self, dian_file: str, contable_file: str):
        super().__init__()
        self.dian_file = dian_file
        self.contable_file = contable_file
        self.stats = {}
        self._is_running = False
        
    def run(self):
        """Ejecutar el procesamiento de causación completo"""
        self._is_running = True
        try:
            # Inicializar procesador de causación
            self.progress.emit("🔧 Inicializando procesador de causación...")
            processor = CausacionProcessor()
            
            # Cargar archivo DIAN
            self.progress.emit("📄 Cargando archivo DIAN...")
            dian_df = processor.load_dian_file(self.dian_file)
            self.progress.emit(f"✅ Archivo DIAN cargado: {len(dian_df)} registros")
            
            # Cargar archivo contable
            self.progress.emit("📄 Cargando archivo contable...")
            contable_df = processor.load_contable_file(self.contable_file)
            self.progress.emit(f"✅ Archivo contable cargado: {len(contable_df)} registros")
            
            # Validar archivos
            self.progress.emit("🔍 Validando archivos...")
            is_valid, errors = processor.validate_files()
            if not is_valid:
                raise Exception(f"Error en validación: {', '.join(errors)}")
            self.progress.emit("✅ Archivos validados correctamente")
            
            # Realizar matching de datos
            self.progress.emit("🔗 Realizando cruce de datos...")
            matching_result = processor.perform_data_matching(dian_df, contable_df)
            matches_df = matching_result['matches']
            non_matches_df = matching_result['non_matches']
            self.progress.emit(f"✅ Cruce completado: {len(matches_df)} coincidencias, {len(non_matches_df)} no coincidencias")
            
            # Generar DataFrames estructurados
            self.progress.emit("📊 Generando DataFrames de resultado...")
            coincidencias_df = processor.create_coincidencias_dataframe(matches_df)
            no_coincidencias_df = processor.create_no_coincidencias_dataframe(non_matches_df)
            self.progress.emit("✅ DataFrames estructurados creados")
            
            # Calcular estadísticas
            self.progress.emit("📈 Calculando estadísticas...")
            stats = processor.calculate_statistics(coincidencias_df, no_coincidencias_df)
            self.stats = stats
            self.progress.emit(f"✅ Estadísticas calculadas - Calidad: {stats['resumen_ejecutivo']['calidad_general']}")
            
            # Crear archivo Excel con formato avanzado
            self.progress.emit("📋 Creando archivo Excel profesional...")
            from config import Config
            output_dir = Config.OUTPUT_PATH
            output_dir.mkdir(exist_ok=True)
            
            excel_path = processor.create_excel_file(
                coincidencias_df=coincidencias_df,
                no_coincidencias_df=no_coincidencias_df,
                output_path=output_dir,
                stats=stats
            )
            
            self.progress.emit(f"✅ Archivo Excel creado: {Path(excel_path).name}")
            
            # Mensaje de éxito con estadísticas
            success_message = (
                f"Procesamiento de causación completado exitosamente\n\n"
                f"📊 Resumen:\n"
                f"• Total procesado: {stats['total_registros']} registros\n"
                f"• Coincidencias: {stats['total_coincidencias']} ({stats['porcentaje_coincidencias']:.1f}%)\n"
                f"• No coincidencias: {stats['total_no_coincidencias']} ({stats['porcentaje_no_coincidencias']:.1f}%)\n"
                f"• Calidad general: {stats['resumen_ejecutivo']['calidad_general']}\n"
                f"• Archivo generado: {Path(excel_path).name}"
            )
            
            self.finished.emit(True, success_message, stats)
            
        except Exception as e:
            error_message = f"Error durante el procesamiento de causación: {str(e)}"
            self.progress.emit(f"❌ {error_message}")
            self.finished.emit(False, error_message, {})
        finally:
            self._is_running = False
    
    def stop(self):
        """Detener el hilo de forma segura"""
        self._is_running = False
        self.wait()  # Esperar a que termine
    
    def is_running(self):
        """Verificar si el hilo está ejecutándose"""
        return self._is_running

class MainWindow(QMainWindow):
    """Ventana principal de la aplicación"""
    
    def __init__(self):
        super().__init__()
        self.dian_file = None
        self.contable_file = None
        self.processing_thread = None
        self.stats = {}
        self.setup_ui()
    
    def closeEvent(self, event):
        """Manejar el cierre de la ventana"""
        if self.processing_thread and self.processing_thread.is_running():
            self.processing_thread.stop()
        event.accept()
        
    def setup_ui(self):
        """Configurar la interfaz principal"""
        self.setWindowTitle("Sistema de Causación - DIAN & Contabilidad")
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
        title_label = QLabel("🔗 Sistema de Causación DIAN-Contable")
        title_font = QFont()
        title_font.setPointSize(20)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 20px;")
        
        # Layout horizontal para las zonas de drop
        drop_layout = QHBoxLayout()
        drop_layout.setSpacing(20)
        
        # Zona de drop para Archivo DIAN
        self.dian_drop = DropZone(
            "Archivo DIAN",
            "Archivo con facturas/registros DIAN"
        )
        self.dian_drop.file_dropped.connect(self.on_dian_file_dropped)
        
        # Zona de drop para Archivo Contable
        self.contable_drop = DropZone(
            "Archivo Contable",
            "Archivo con movimientos contables"
        )
        self.contable_drop.file_dropped.connect(self.on_contable_file_dropped)
        
        drop_layout.addWidget(self.dian_drop)
        drop_layout.addWidget(self.contable_drop)
        
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
        
    def on_dian_file_dropped(self, file_path: str):
        """Manejar archivo DIAN seleccionado"""
        self.dian_file = file_path
        self.log_message(f"📄 Archivo DIAN cargado: {Path(file_path).name}")
        self.check_ready_to_process()
        
    def on_contable_file_dropped(self, file_path: str):
        """Manejar archivo contable seleccionado"""
        self.contable_file = file_path
        self.log_message(f"📄 Archivo contable cargado: {Path(file_path).name}")
        self.check_ready_to_process()
        
    def check_ready_to_process(self):
        """Verificar si ambos archivos están cargados"""
        if self.dian_file and self.contable_file:
            self.process_btn.setEnabled(True)
            self.process_btn.setText("🚀 Iniciar Causación")
        else:
            self.process_btn.setEnabled(False)
            missing = []
            if not self.dian_file:
                missing.append("Archivo DIAN")
            if not self.contable_file:
                missing.append("Archivo Contable")
            self.process_btn.setText(f"⏳ Faltan: {', '.join(missing)}")
            
    def log_message(self, message: str):
        """Agregar mensaje al log"""
        self.log_area.setVisible(True)
        self.log_area.append(f"• {message}")
        
    def process_files(self):
        """Iniciar el procesamiento de causación"""
        if not self.dian_file or not self.contable_file:
            QMessageBox.warning(self, "Archivos faltantes", 
                              "Por favor, selecciona ambos archivos antes de iniciar la causación.")
            return
            
        # Configurar UI para procesamiento
        self.process_btn.setEnabled(False)
        self.process_btn.setText("⏳ Procesando...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Barra de progreso indeterminada
        self.log_area.setVisible(True)
        self.log_area.clear()
        
        # Iniciar procesamiento en hilo separado
        self.processing_thread = ProcessingThread(self.dian_file, self.contable_file)
        self.processing_thread.progress.connect(self.log_message)
        self.processing_thread.finished.connect(self.on_processing_finished)
        self.processing_thread.start()
        
    def on_processing_finished(self, success: bool, message: str, stats: dict = {}):
        """Manejar finalización del procesamiento de causación"""
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        self.process_btn.setText("🚀 Iniciar Causación")
        
        if success:
            self.stats = stats
            self.log_message("✅ Procesamiento de causación completado")
            
            # Mostrar estadísticas detalladas
            if stats:
                self.log_message("📊 Estadísticas del proceso:")
                self.log_message(f"   • Total registros: {stats.get('total_registros', 0)}")
                self.log_message(f"   • Coincidencias: {stats.get('total_coincidencias', 0)} ({stats.get('porcentaje_coincidencias', 0):.1f}%)")
                self.log_message(f"   • No coincidencias: {stats.get('total_no_coincidencias', 0)} ({stats.get('porcentaje_no_coincidencias', 0):.1f}%)")
                self.log_message(f"   • Calidad general: {stats.get('resumen_ejecutivo', {}).get('calidad_general', 'N/A')}")
            
            QMessageBox.information(self, "✅ Causación Completada", message)
        else:
            self.log_message("❌ Error en el procesamiento")
            QMessageBox.critical(self, "❌ Error de Causación", message)

def run_app():
    """Ejecutar la aplicación de causación"""
    app = QApplication(sys.argv)
    app.setApplicationName("Sistema de Causación DIAN-Contable")
    app.setApplicationVersion("2.0.0")
    
    # Configurar información de la aplicación
    app.setOrganizationName("Sistema de Causación")
    app.setOrganizationDomain("causacion.com")
    
    # Validar dependencias
    try:
        from .causacion_processor import CausacionProcessor
        processor = CausacionProcessor()
        print("✅ Procesador de causación inicializado correctamente")
    except Exception as e:
        print(f"❌ Error al inicializar procesador de causación: {e}")
        QMessageBox.critical(None, "Error de Inicialización", 
                           f"No se pudo inicializar el procesador de causación:\n{str(e)}")
        return 1
    
    # Crear y mostrar ventana principal
    window = MainWindow()
    window.show()
    
    return app.exec() 