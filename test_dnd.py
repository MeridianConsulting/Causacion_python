#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test mínimo de Drag & Drop para verificar funcionamiento en el entorno
"""

import sys
from PySide6.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget
from PySide6.QtCore import Qt
from PySide6.QtGui import QDragEnterEvent, QDropEvent

class DropLabel(QLabel):
    def __init__(self):
        super().__init__("Arrastra aquí un archivo desde el Explorador")
        self.setAlignment(Qt.AlignCenter)
        self.setAcceptDrops(True)
        self.setMinimumSize(400, 200)
        self.setStyleSheet("""
            QLabel {
                border: 3px dashed #007acc;
                border-radius: 10px;
                background-color: #f0f0f0;
                padding: 20px;
                font-size: 14px;
            }
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        print("=" * 60)
        print("[TEST] dragEnterEvent DISPARADO")
        print(f"[TEST] mime formats: {event.mimeData().formats()}")
        print(f"[TEST] hasUrls: {event.mimeData().hasUrls()}")
        
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            print(f"[TEST] URLs encontradas: {len(urls)}")
            for i, url in enumerate(urls):
                print(f"[TEST]   URL {i+1}: {url.toLocalFile()}")
            event.acceptProposedAction()
            self.setStyleSheet("""
                QLabel {
                    border: 3px solid #28a745;
                    border-radius: 10px;
                    background-color: #d4edda;
                    padding: 20px;
                    font-size: 14px;
                }
            """)
        else:
            print("[TEST] No hay URLs, ignorando evento")
            event.ignore()

    def dragLeaveEvent(self, event):
        print("[TEST] dragLeaveEvent DISPARADO")
        self.setStyleSheet("""
            QLabel {
                border: 3px dashed #007acc;
                border-radius: 10px;
                background-color: #f0f0f0;
                padding: 20px;
                font-size: 14px;
            }
        """)

    def dropEvent(self, event: QDropEvent):
        print("=" * 60)
        print("[TEST] dropEvent DISPARADO")
        print(f"[TEST] mime formats: {event.mimeData().formats()}")
        print("[TEST] URLs recibidas:")
        
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            print(f"[TEST]   Archivo: {file_path}")
            self.setText(f"Archivo recibido:\n{file_path}")
        
        event.acceptProposedAction()
        print("[TEST] dropEvent completado exitosamente")
        print("=" * 60)

class TestWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test de Drag & Drop")
        self.setMinimumSize(500, 300)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        info_label = QLabel(
            "INSTRUCCIONES:\n\n"
            "1. Abre el Explorador de Windows\n"
            "2. Arrastra un archivo Excel (.xlsx o .xls)\n"
            "3. Suéltalo sobre el recuadro de abajo\n"
            "4. Revisa la consola para ver los mensajes de debug\n\n"
            "Si ves mensajes [TEST] en la consola, el drag & drop funciona.\n"
            "Si NO ves ningún mensaje, hay un problema de permisos/entorno."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("padding: 10px; background-color: #fff3cd; border-radius: 5px;")
        
        self.drop_area = DropLabel()
        
        layout.addWidget(info_label)
        layout.addWidget(self.drop_area)
        
        self.setLayout(layout)

if __name__ == "__main__":
    print("=" * 60)
    print("TEST DE DRAG & DROP - Iniciando aplicación")
    print("=" * 60)
    print("\nIMPORTANTE: Asegúrate de ejecutar esto SIN permisos de administrador")
    print("Si ves mensajes [TEST] al arrastrar, el drag & drop funciona.\n")
    
    app = QApplication(sys.argv)
    app.setApplicationName("Test Drag & Drop")
    
    window = TestWindow()
    window.show()
    
    sys.exit(app.exec())

