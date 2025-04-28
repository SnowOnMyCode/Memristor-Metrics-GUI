# main.py
# Entry point for Memristor Metrics GUI
# Developed by Lek Ibrahimi


import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QSlider, QFileDialog, QComboBox, QLabel, QHBoxLayout, QMainWindow)
from PyQt5.QtCore import Qt
from class_IVSweeps import IVSweepsAnalyser
from class_pulses import PulsesViewer
from class_Retention import Retention_viewer
from class_volatile import VolatileSweepsAnalyser

class MainAppWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Choose measurement type")
        self.setMinimumSize(500, 300)  # Set a larger size for the main window
        
        layout = QVBoxLayout()
        hbox1 = QHBoxLayout()
        hbox2 = QHBoxLayout()

        self.iv_sweeps_button = QPushButton("IV Sweeps")
        self.iv_sweeps_button.clicked.connect(self.open_iv_sweeps)
        self.iv_sweeps_button.setFixedSize(150, 50)

        self.pulses_button = QPushButton("Pulses")
        self.pulses_button.clicked.connect(self.open_pulses)
        self.pulses_button.setFixedSize(150, 50)

        self.retention_button = QPushButton("Retention")
        self.retention_button.clicked.connect(self.open_retention)
        self.retention_button.setFixedSize(150, 50)

        self.volatile_button = QPushButton("Volatile")
        self.volatile_button.clicked.connect(self.open_volatile)
        self.volatile_button.setFixedSize(150, 50)

         # Add buttons to the horizontal box layout
        hbox1.addStretch(1)
        hbox1.addWidget(self.iv_sweeps_button)
        hbox1.addWidget(self.pulses_button)
        hbox1.addStretch(1)
        hbox2.addStretch(1)
        hbox2.addWidget(self.retention_button)
        hbox2.addWidget(self.volatile_button)
        hbox2.addStretch(1)


        # Add the horizontal box layout to the vertical box layout
        layout.addStretch(1)
        layout.addLayout(hbox1)
        layout.addLayout(hbox2)
        layout.addStretch(1)

        self.setLayout(layout)
        
    def open_iv_sweeps(self):
        self.iv_window = IVSweepsAnalyser()
        self.iv_window.show()

    def open_pulses(self):
        self.pulses_window = PulsesViewer()
        self.pulses_window.show()
    
    def open_retention(self):
        self.retention_window = Retention_viewer()
        self.retention_window.show()

    def open_volatile(self):
        self.volatile_window = VolatileSweepsAnalyser()
        self.volatile_window.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet("""
        QWidget {
            background-color: #2E3440;
            color: #D8DEE9;
            font-family: 'Arial';
        }
        QPushButton {
            background-color: #4C566A;
            border-radius: 10px;
            padding: 10px;
            color: #D8DEE9;
            font-size: 16px;
            margin: 5px;
        }
        QPushButton:hover {
            background-color: #5E81AC;
        }
        QSlider::groove:horizontal {
            height: 8px;
            background: #4C566A;
            border: 1px solid #5E81AC;
            border-radius: 4px;
        }
        QSlider::handle:horizontal {
            background: #88C0D0;
            border: 1px solid #81A1C1;
            width: 18px;
            margin: -5px 0;
            border-radius: 9px;
        }
        QLabel {
            font-size: 14px;
            color: #D8DEE9;
        }
        QComboBox {
            background-color: #4C566A;
            color: #D8DEE9;
            border: 1px solid #5E81AC;
            border-radius: 5px;
            padding: 5px;
        }
        QProgressBar {
            text-align: center;
            color: #D8DEE9;
            background-color: #4C566A;
            border-radius: 5px;
            border: 1px solid #5E81AC;
        }
        QProgressBar::chunk {
            background-color: #88C0D0;
            border-radius: 5px;
        }
    """)
    viewer = MainAppWindow()
    viewer.show()
    sys.exit(app.exec_())