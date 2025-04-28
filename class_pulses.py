import sys
from PyQt5.QtWidgets import (QApplication, QLineEdit, QInputDialog, QCheckBox, QWidget, QVBoxLayout, QPushButton, QSlider, QFileDialog, QComboBox, QLabel, QHBoxLayout, QMainWindow, QMessageBox, QProgressBar,QDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.ticker import FuncFormatter, AutoLocator, LogFormatterMathtext, ScalarFormatter
import matplotlib.ticker as ticker
import mplcursors
from scipy.signal import find_peaks
from PyQt5.QtWidgets import QColorDialog

class FileLoader(QThread):
    progress = pyqtSignal(int)

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths
        self.sweeps = []

    def run(self):
        total_sheets = sum(len(pd.ExcelFile(file_path).sheet_names) for file_path in self.file_paths)
        sheets_processed = 0

        for file_path in self.file_paths:
            excel_file = pd.ExcelFile(file_path)
            sheet_name_list = [sheet for sheet in excel_file.sheet_names if sheet not in ["Calc", "Settings"]]
            for sheet in sheet_name_list:
                df = pd.read_excel(file_path, sheet_name=sheet)
                self.sweeps.append(df)
                sheets_processed += 1
                self.progress.emit(int((sheets_processed / (total_sheets-2*len(self.file_paths))) * 100))

class PulsesViewer(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Pulses Viewer")

        # Layout
        main_layout = QVBoxLayout()

        options_layout = QHBoxLayout()
        main_layout.addLayout(options_layout)
        channels_layout = QVBoxLayout()
        options_layout.addLayout(channels_layout)
        colors_layout = QVBoxLayout()
        options_layout.addLayout(colors_layout)
        labels_layout = QVBoxLayout()
        options_layout.addLayout(labels_layout)

        # Label to display Ion
        self.ion_label = QLabel("Ion: N/A")
        self.ion_label.setAlignment(Qt.AlignRight)
        labels_layout.addWidget(self.ion_label)

        # Label to display Ioff
        self.ioff_label = QLabel("Ioff: N/A")
        self.ioff_label.setAlignment(Qt.AlignRight)
        labels_layout.addWidget(self.ioff_label)

        # Label to display Ron
        self.Ron_label = QLabel("Ron: N/A")
        self.Ron_label.setAlignment(Qt.AlignRight)
        labels_layout.addWidget(self.Ron_label)
        
        # Label to display Roff
        self.Roff_label = QLabel("Roff: N/A")
        self.Roff_label.setAlignment(Qt.AlignRight)
        labels_layout.addWidget(self.Roff_label)

        # Label to display ton
        self.ton_label = QLabel("ton: N/A")
        self.ton_label.setAlignment(Qt.AlignRight)
        labels_layout.addWidget(self.ton_label)

        # Label to display toff
        self.toff_label = QLabel("toff: N/A")
        self.toff_label.setAlignment(Qt.AlignRight)
        labels_layout.addWidget(self.toff_label)

        # Color Picker for Voltage Line
        voltage_color_layout = QHBoxLayout()
        voltage_color_label = QLabel("Pick Voltage Line Color")
        #voltage_color_layout.addWidget(voltage_color_label)
        self.voltage_color_button = QPushButton("Choose Voltage Color")
        self.voltage_color_button.setFixedSize(300, 50)
        self.voltage_color_button.clicked.connect(self.pick_voltage_color)
        self.voltage_color_button.adjustSize()
        voltage_color_layout.addWidget(self.voltage_color_button)
        colors_layout.addLayout(voltage_color_layout)

        # Color Picker for Current Line
        current_color_layout = QHBoxLayout()
        current_color_label = QLabel("Pick Current Line Color")
        #current_color_layout.addWidget(current_color_label)
        self.current_color_button = QPushButton("Choose Current Color")
        self.current_color_button.setFixedSize(300, 50)
        self.current_color_button.clicked.connect(self.pick_current_color)
        self.current_color_button.adjustSize()
        current_color_layout.addWidget(self.current_color_button, stretch=2)
        colors_layout.addLayout(current_color_layout)

        # Add a checkbox for grid control
        self.grid_checkbox = QCheckBox("Grid On/Off")
        self.grid_checkbox.setChecked(False)  # Default is grid off
        self.grid_checkbox.stateChanged.connect(self.update_plot)  # Connect to grid toggling method
        main_layout.addWidget(self.grid_checkbox, alignment=Qt.AlignRight)

        # Slider
        self.slider = QSlider(Qt.Horizontal)
        self.slider.setMinimum(0)
        self.slider.valueChanged.connect(self.update_plot)
        main_layout.addWidget(self.slider)

        # Plot area
        self.figure, self.ax1 = plt.subplots(figsize=(14, 6))
        self.canvas = FigureCanvas(self.figure)
        main_layout.addWidget(self.canvas)

        self.setLayout(main_layout)

        self.sweeps = []
        self.current_index = 0
        self.VChannel = "VMeasCh2"  # Default voltage channel
        self.IChannel = "IMeasCh1"  # Default current channel
        self.TimeChannel = "TimeOutput" # Default time channel
        self.current_color = 'blue' # Default current color
        self.voltage_color = 'red'  # Default voltage color

        #Making text bigger
        self.setStyleSheet("""
            QLabel {
                font-family: Arial;
                font-size: 14pt;
                color: White;
            }
            QCheckBox {
                font-family: Arial;
                font-size: 12pt;
                color: White;
            }
            QComboBox {
                font-size: 16px;           
            }
        """)

        
        buttons_layout = QHBoxLayout()
        main_layout.addLayout(buttons_layout)
        buttons_layout.setAlignment(Qt.AlignRight)

        # Save All Plots button
        self.save_button = QPushButton("Save All Plots")
        self.save_button.clicked.connect(self.save_all_plots)
        self.save_button.setFixedSize(150, 50)
        buttons_layout.addWidget(self.save_button)

        # Upload button
        self.upload_button = QPushButton("Upload Files")
        self.upload_button.clicked.connect(self.upload_files)
        self.upload_button.setFixedSize(150, 50)
        buttons_layout.addWidget(self.upload_button)

    def upload_files(self):
        # Ask for voltage channel name
        VChannel, ok1 = QInputDialog.getText(self, "Voltage Channel", "Enter the voltage channel name:", QLineEdit.Normal, "VMeasCh2")
        if not ok1:
            return
        elif not VChannel:
            VChannel = "VMeasCh2"  # Default value if canceled or empty

        # Ask for current channel name
        IChannel, ok2 = QInputDialog.getText(self, "Current Channel", "Enter the current channel name:", QLineEdit.Normal, "IMeasCh1")
        if not ok2:
            return
        elif not IChannel:
            IChannel = "IMeasCh1"  # Default value if canceled or empty
        
        # Assign the user-input column names
        self.VChannel = VChannel
        self.IChannel = IChannel

        options = QFileDialog.Options()
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Upload Excel Files", "", "Excel Files (*.xls *.xlsx)", options=options)
        
        if not file_paths:
            return
        
        self.sweeps.clear()

        progress_layout = QVBoxLayout()

        # Create progress bar dialog
        self.progress_dialog = QDialog()
        self.progress_dialog.setWindowTitle("Loading Files")
        #self.progress_dialog.setText("Please wait while the files are being loaded.")
        self.label = QLabel("Please wait while the files are being loaded.")
        #self.progress_dialog.setStandardButtons(QMessageBox.NoButton)
        
        self.progress_bar = QProgressBar()
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.label)
        #self.progress_dialog.layout().addChildLayout(progress_layout)
        self.progress_dialog.setLayout(progress_layout)
        self.progress_dialog.show()

        # Start file loader thread
        self.file_loader = FileLoader(file_paths)
        self.file_loader.progress.connect(self.progress_bar.setValue)
        self.file_loader.finished.connect(self.on_files_loaded)
        self.file_loader.start()

    def on_files_loaded(self):
        self.sweeps = self.file_loader.sweeps

        if self.sweeps:
            self.slider.setMaximum(len(self.sweeps) - 1)
            self.current_index = 0
            self.update_plot()

        self.progress_dialog.hide()

    def update_plot(self, *args):
        self.current_index = self.slider.value()
        if self.sweeps:
            df = self.sweeps[self.current_index]
            self.plot_data(df)

    def calculate_ion(self, df):
        """
        Calculate Ion as the average current in the second half of the voltage pulse.
        """
        # Interpolate the current and voltage channels
        df[self.IChannel] = df[self.IChannel].interpolate(method='linear')
        df[self.VChannel] = df[self.VChannel].interpolate(method='linear')

        peaks,_ = find_peaks(df[self.VChannel])
        #height_threshold = np.percentile(df[VChannel][peaks], 80)
        height_threshold = 0.9*max(df[self.VChannel])
        high_peaks = peaks[df[self.VChannel][peaks] > height_threshold]

        if len(high_peaks)<2 or (high_peaks[1]-high_peaks[0] > 20):
            return False, False

        #In some measurement profiles the current is inverted and in some not, check for that here
        if self.IChannel == "Imeas":
            pulse_current = df[self.IChannel][high_peaks[0]:high_peaks[-1]]
        else:
            pulse_current = abs(df[self.IChannel])[high_peaks[0]:high_peaks[-1]]

        pulse_voltage = df[self.VChannel][high_peaks[0]:high_peaks[-1]]
        pulse_time = df[self.TimeChannel][high_peaks[0]:high_peaks[-1]]
        Ion_interval = pulse_current[(len(pulse_current)//2):]
        ton_interval = pulse_time[(len(pulse_time)//2):]
        Von_interval = pulse_voltage[(len(pulse_voltage)//2):]
        Ion_avg = np.mean(Ion_interval)
        Von_avg = np.mean(Von_interval)
        Ron = abs(Von_avg/Ion_avg)

        # Find 90% of Ion_avg
        target_current = 0.9 * Ion_avg
        
        # Locate the first time the current exceeds the target
        idx_90 = np.where(pulse_current >= target_current)[0][0]  # Index in the second-half array
        time_90 = pulse_time.iloc[idx_90]  # Corresponding time
        
        # Average current in the second half
        return Ion_avg, Ron, time_90
    
    def calculate_ioff(self, df):
        peaks,_ = find_peaks(df[self.VChannel])
        height_threshold = 0.9*max(df[self.VChannel])
        high_peaks = peaks[df[self.VChannel][peaks] > height_threshold]
        if len(high_peaks)<2 or (high_peaks[1]-high_peaks[0] > 20):
            return False, False
        idx_left = peaks[np.where(peaks==high_peaks[-1])[0]+1][0]
        
        #Again, check if current inverted or not
        if self.IChannel == "Imeas":
            off_current = df[self.IChannel][idx_left:peaks[-1]]
        else:
            df[self.IChannel] = abs(df[self.IChannel])
            off_current = df[self.IChannel][idx_left:peaks[-1]]

        off_voltage = df[self.VChannel][idx_left:peaks[-1]]
        off_time = df[self.TimeChannel][idx_left:peaks[-1]]
        Ioff_interval = off_current[(len(off_current)//2):]
        toff_interval = off_time[(len(off_time)//2):]
        Ioff_avg = np.mean(Ioff_interval)
        Voff_avg = np.mean(off_voltage[(len(off_voltage)//2):])
        Roff = abs(Voff_avg/Ioff_avg)
        Ioff_delta = (0.1 * (self.calculate_ion(df)[0] - Ioff_avg)) + Ioff_avg
        idx_10 = np.where(df[self.IChannel][len(df[self.IChannel])//2:] <= Ioff_delta)[0][0]  # Index in the second-half array
        time_10 = df[self.TimeChannel].iloc[idx_10 + (len(df[self.IChannel])//2)-1]


        return Ioff_avg, Roff, time_10

    def save_all_plots(self):
        # Ask for a directory to save the plots
        save_dir = QFileDialog.getExistingDirectory(self, "Select Directory to Save Plots")
        if not save_dir:
            return  # If no directory selected, do nothing
        first_file_name = self.file_loader.file_paths[0].split('/')[-1].rsplit('.', 1)[0]
        # Iterate over uploaded files and corresponding DataFrames
        for i, df in enumerate(self.sweeps):
            if len(self.file_loader.file_paths) == 1:
                file_name = f"{first_file_name}_Sheet{i+1}"
            else:
                file_name = self.file_loader.file_paths[i].split('/')[-1].rsplit('.', 1)[0]
            save_path = f"{save_dir}/{file_name}.png"

            # Plot data and save to the specified path
            self.plot_data(df, save_path=save_path)
        
        QMessageBox.information(self, "Success", f"Plots saved successfully to {save_dir}")

    def plot_data(self, df, save_path=None):
        try:
            if self.VChannel not in df.columns:
                raise KeyError("Chosen voltage column not found")
            if self.IChannel not in df.columns:
                raise KeyError("Chosen current column not found")
            
            if "Time" in self.sweeps[0].columns:
                self.TimeChannel = "Time"
                df[self.IChannel] = abs(df[self.IChannel])
            else:
                self.TimeChannel = "TimeOutput"

            # Clear the previous plot
            self.figure.clear()

            # Create new axes
            self.ax1 = self.figure.add_subplot(111)

            # Left axis for Voltage Channel
            self.ax1.set_xlabel(r'$\it{Time}\ (s)$')
            self.ax1.set_ylabel(r'$\it{V}\ (V)$', color=self.voltage_color)
            line1, = self.ax1.plot(df[self.TimeChannel], df[self.VChannel], color=self.voltage_color, linewidth=1.5, label=self.VChannel)
            self.ax1.tick_params(axis='y', labelcolor=self.voltage_color)

            if self.grid_checkbox.isChecked():
                self.ax1.grid(True, which='both', linestyle='--', linewidth=0.5)

            self.ax1.xaxis.set_minor_locator(ticker.AutoMinorLocator())
            self.ax1.yaxis.set_minor_locator(ticker.AutoMinorLocator())

            # Format y-axis1 with scientific notation
            self.ax1.yaxis.set_major_formatter(ScalarFormatter())

            # Right axis for Current Channel
            ax2 = self.ax1.twinx()
            ax2.set_ylabel(r'$\it{I}\ (A)$', color=self.current_color)
            line2, = ax2.plot(df[self.TimeChannel], abs(df[self.IChannel]), color=self.current_color, linewidth=1.5, label=self.IChannel)
            ax2.tick_params(axis='y', labelcolor=self.current_color)
            ax2.grid(False)

            # Format y-axis2 with scientific notation
            ax2.yaxis.set_major_formatter(ScalarFormatter())

            # Get handles and labels from both axes
            lines1, labels1 = self.ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()

            # Combine both and create a single legend
            self.ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right')

            # Format right y-axis with scientific notation
            #self.ax1.yaxis.set_major_formatter(LogFormatterMathtext(base=10))

            # Add interactive cursor
            mplcursors.cursor([line1, line2], hover=2)

            self.figure.tight_layout()
            self.canvas.draw()

            if save_path:
                self.figure.savefig(save_path, dpi=300)

            on_results = self.calculate_ion(df)
            
            # Calculate Ion
            ion = on_results[0]
            if ion != False:
                self.ion_label.setText(f"Ion: {ion:.2e} A")
                # Add horizontal dotted line for Ion
                ax2.axhline(ion, color='blue', linestyle='--', linewidth=1.0, label='Ion')
            
            # Calculate Ron
            Ron = on_results[1]
            if Ron != False:
                self.Ron_label.setText(f"Ron: {Ron:.2e} Ohm")

            # Calculate ton
            ton = on_results[2]
            if ton != False:
                self.ton_label.setText(f"ton: {ton:.2e} s")

            off_results = self.calculate_ioff(df)
            # Calculate Ioff
            ioff = off_results[0]
            if ioff != False:
                self.ioff_label.setText(f"Ioff: {ioff:.2e} A")
                # Add horizontal dotted line for Ioff
                ax2.axhline(ioff, color='red', linestyle='--', linewidth=1.0, label='Ioff')

            # Calculate Roff
            Roff = off_results[1]
            if Roff != False:
                self.Roff_label.setText(f"Roff: {Roff:.2e} Ohm")

            # Calculate toff
            toff = off_results[2]
            if toff != False:
                self.toff_label.setText(f"toff: {toff:.2e} s")

            # Update the legend to include Ion and Ioff
            lines1, labels1 = self.ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            self.ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right')

        except KeyError as e:
            QMessageBox.critical(self, "Error", f"Column doesn't exist: {str(e)}. Please try another file.")

    def pick_voltage_color(self):
        # Open color picker dialog and set the selected color
        color = QColorDialog.getColor()
        if color.isValid():
            self.voltage_color = color.name()  # Convert selected color to HEX code (e.g., "#RRGGBB")
            self.voltage_color_button.setStyleSheet(f"background-color: {self.voltage_color};")  # Update button background color
        self.update_plot()

    def pick_current_color(self):
        # Open color picker dialog and set the selected color
        color = QColorDialog.getColor()
        if color.isValid():
            self.current_color = color.name()  # Convert selected color to HEX code (e.g., "#RRGGBB")
            self.current_color_button.setStyleSheet(f"background-color: {self.current_color};")  # Update button background color
        self.update_plot()

    def update_grid(self, grid_enabled):
        """Update the grid visibility."""
        self.grid_enabled = grid_enabled
        self.update_plot()