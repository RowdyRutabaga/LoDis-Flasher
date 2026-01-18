import sys
import os
import time
import serial.tools.list_ports
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QComboBox, QPushButton, QProgressBar, QMessageBox,
    QFileDialog, QTextEdit,QLineEdit
)
from PySide6.QtCore import QThread, Signal, QObject, Slot
from PySide6.QtGui import QFont

# Determine the base path for resources (like the 'bin' directory)
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    # Running as a bundled app (.app)
    # The path needs to go up from .../Flasher.app/Contents/MacOS/Flasher
    base_path = os.path.abspath(os.path.join(os.path.dirname(sys.executable), '..', '..', '..'))
else:
    # Running as a normal python script
    base_path = os.path.dirname(os.path.abspath(__file__))

BIN_DIR = os.path.join(base_path, 'bin')  # Directory where the binary files are located

class StdoutEmitter(QObject):
    textWritten = Signal(str)

    def write(self, text):
        self.textWritten.emit(str(text))

    def flush(self):
        pass

class EsptoolWorker(QObject):
    """
    Worker thread for running esptool as a Python library to avoid freezing the GUI.
    """
    output = Signal(str)
    finished = Signal(int)

    def __init__(self, args):
        super().__init__()
        self.args = args

    def run(self):
        """
        Executes the esptool command by calling its main function and
        redirecting stdout to capture output in real-time.
        """
        original_stdout = sys.stdout
        original_stderr = sys.stderr
        
        emitter = StdoutEmitter()
        # Connect the emitter's signal to the worker's output signal
        emitter.textWritten.connect(lambda text: self.output.emit(text.strip()))

        sys.stdout = emitter
        sys.stderr = emitter
        
        exit_code = 0
        try:
            import esptool
            import serial
            esptool.main(self.args)
        except SystemExit as e:
            # esptool calls sys.exit() on completion. 0 is success.
            exit_code = e.code if e.code is not None else 0
        except Exception as e:
            # Print any other exceptions to our redirected output
            print(f"An error occurred while running esptool:\n{str(e)}")
            exit_code = 1
        finally:
            # Always restore the original stdout/stderr
            sys.stdout = original_stdout
            sys.stderr = original_stderr

        self.finished.emit(exit_code)

    def stop(self):
        pass


class PortMonitor(QObject):
    """Monitors serial port connections in a background thread."""
    ports_changed = Signal()

    def __init__(self):
        super().__init__()
        self._running = True
        self._previous_ports = set()

    def run(self):
        while self._running:
            try:
                ports = set(p.device for p in serial.tools.list_ports.comports())
                if ports != self._previous_ports:
                    self._previous_ports = ports
                    self.ports_changed.emit()
            except Exception:
                # Ignore errors during port scanning
                pass
            time.sleep(1)

    def stop(self):
        self._running = False


class ESPFlasherApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LoDis Signal Processor Flasher")
        self.setGeometry(100, 100, 700, 600)

        self.esptool_thread = None
        self.esptool_worker = None
        self.current_version_files = {}  # Store the files for the selected version

        self.create_widgets()
        self.refresh_versions()
        self.refresh_ports()

        self.start_port_monitor()

    def create_widgets(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        # Port selection
        port_group = QGroupBox("Step 1: Select COM Port")
        port_layout = QHBoxLayout(port_group)
        self.port_combo = QComboBox()
        self.port_combo.setMinimumWidth(300)
        refresh_ports_button = QPushButton("Refresh")
        refresh_ports_button.clicked.connect(self.refresh_ports)
        port_layout.addWidget(self.port_combo)
        port_layout.addWidget(refresh_ports_button)
        main_layout.addWidget(port_group)

        # Firmware version selection
        version_group = QGroupBox("Step 2: Select Firmware Version")
        version_layout = QHBoxLayout(version_group)
        self.version_combo = QComboBox()
        self.version_combo.setMinimumWidth(300)
        self.version_combo.currentTextChanged.connect(self.on_version_changed)
        refresh_versions_button = QPushButton("Refresh")
        refresh_versions_button.clicked.connect(self.refresh_versions)
        version_layout.addWidget(self.version_combo)
        version_layout.addWidget(refresh_versions_button)
        main_layout.addWidget(version_group)

        # Flash button and progress bar
        action_group = QGroupBox("Step 3: Flash Firmware")
        action_layout = QHBoxLayout(action_group)
        self.flash_button = QPushButton("Flash ESP32")
        self.flash_button.clicked.connect(self.flash_esp32)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.progress_bar.hide()
        action_layout.addWidget(self.flash_button)
        action_layout.addWidget(self.progress_bar)
        main_layout.addWidget(action_group)

        # Flash button and progress bar
        configure_group = QGroupBox("Step 4: Set Name and ID")
        configure_layout = QVBoxLayout(configure_group)
        self.signal_name_label = QLabel("Signal Name:")
        self.signal_ID_label = QLabel("Signal ID:")
        self.signal_name = QLineEdit()
        self.signal_ID = QLineEdit()
        self.signal_ID.setMaxLength(3)
        self.configure_button = QPushButton("Set Name and ID")
        self.configure_button.clicked.connect(self.configure)
        configure_layout.addWidget(self.signal_name_label)
        configure_layout.addWidget(self.signal_name)
        configure_layout.addWidget(self.signal_ID_label)
        configure_layout.addWidget(self.signal_ID)
        configure_layout.addWidget(self.configure_button)
        main_layout.addWidget(configure_group)

        # Output console
        output_group = QGroupBox("Output")
        output_layout = QVBoxLayout(output_group)
        self.output_console = QTextEdit()
        self.output_console.setReadOnly(True)
        self.output_console.setFont(QFont("Courier", 10))
        output_layout.addWidget(self.output_console)
        main_layout.addWidget(output_group)

        # Status label
        self.status_label = QLabel("Ready")
        main_layout.addWidget(self.status_label)
        self.on_version_changed("1.0.0")

    def create_file_selection(self, parent_layout, label_text):
        row_layout = QHBoxLayout()
        label = QLabel(label_text)
        label.setMinimumWidth(80)
        combobox = QComboBox()
        browse_button = QPushButton("Browse...")
        
        row_layout.addWidget(label)
        row_layout.addWidget(combobox)
        row_layout.addWidget(browse_button)
        parent_layout.addLayout(row_layout)

        browse_button.clicked.connect(lambda: self.browse_file(combobox))
        return combobox

    def browse_file(self, combobox):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Binary File", BIN_DIR, "Binary Files (*.bin)")
        if file_path:
            if combobox.findText(file_path) == -1:
                 combobox.addItem(file_path)
            combobox.setCurrentText(file_path)

    @Slot()
    def refresh_versions(self):
        """Scan the bin directory for firmware version folders."""
        if not os.path.exists(BIN_DIR):
            os.makedirs(BIN_DIR)
        
        versions = []
        try:
            for item in sorted(os.listdir(BIN_DIR)):
                item_path = os.path.join(BIN_DIR, item)
                if os.path.isdir(item_path):
                    versions.append(item)
        except Exception:
            pass
        
        current_selection = self.version_combo.currentText()
        self.version_combo.clear()
        self.version_combo.addItems(versions)
        
        if current_selection in versions:
            self.version_combo.setCurrentText(current_selection)
        elif versions:
            self.version_combo.setCurrentIndex(0)

    @Slot(str)
    def on_version_changed(self, version):
        """Update binary files when version is changed."""
        if not version:
            self.current_version_files = {}
            return
        
        version_dir = os.path.join(BIN_DIR, str(version))
        if not os.path.exists(version_dir):
            self.current_version_files = {}
            return
        
        # Scan for required binary files in the version folder
        files_found = {}
        try:
            for filename in os.listdir(version_dir):
                filepath = os.path.join(version_dir, filename)
                if os.path.isfile(filepath) and filename.endswith('.bin'):
                    if 'bootloader' in filename.lower():
                        files_found['bootloader'] = filepath
                    elif 'partition' in filename.lower():
                        files_found['partition'] = filepath
                    elif 'boot_app0' in filename.lower():
                        files_found['ota_data'] = filepath
                    else:
                        # Any other .bin file is the application firmware
                        files_found['app'] = filepath
        except Exception:
            pass
        
        self.current_version_files = files_found

    def start_port_monitor(self):
        self.port_monitor_thread = QThread()
        self.port_monitor = PortMonitor()
        self.port_monitor.moveToThread(self.port_monitor_thread)
        self.port_monitor.ports_changed.connect(self.refresh_ports)
        self.port_monitor_thread.started.connect(self.port_monitor.run)
        self.port_monitor_thread.start()

    @Slot()
    def refresh_ports(self):
        ports = serial.tools.list_ports.comports()
        port_list = [f"{port.device} - {port.description}" for port in ports]
        current_selection = self.port_combo.currentText()
        self.port_combo.clear()
        self.port_combo.addItems(port_list)
        if current_selection in port_list:
            self.port_combo.setCurrentText(current_selection)

    def closeEvent(self, event):
        self.port_monitor.stop()
        self.port_monitor_thread.quit()
        self.port_monitor_thread.wait()
        
        # The esptool function call cannot be forcefully stopped.
        # We just wait for the thread to finish its work if it's running.
        if self.esptool_thread and self.esptool_thread.isRunning():
            self.esptool_thread.quit()
            self.esptool_thread.wait()

        super().closeEvent(event)
    
    def configure(self):
        name = self.signal_name.text()
        id = self.signal_ID.text()
        selected_port_desc = self.port_combo.currentText()
        if not selected_port_desc:
            QMessageBox.critical(self, "Error", "Please select a COM port.")
            return
        
        port = selected_port_desc.split(' - ')[0]
        
        try:
            ser = serial.Serial(port, 115200, timeout=2)
            
            # Set Signal ID
            cmd = f"signal_id:{id}\n"
            self.output_console.append(f"Sending: {cmd.strip()}")
            ser.write(cmd.encode())
            response = ser.readline().decode().strip()
            self.output_console.append(f"Response: {response}")
            response = ser.readline().decode().strip()
            self.output_console.append(f"Response: {response}")
            
            # Set Signal Name
            cmd = f"signal_name:{name}\n"
            self.output_console.append(f"Sending: {cmd.strip()}")
            ser.write(cmd.encode())
            response = ser.readline().decode().strip()
            self.output_console.append(f"Response: {response}")
            response = ser.readline().decode().strip()
            self.output_console.append(f"Response: {response}")
            
            ser.close()

            port = selected_port_desc.split(' - ')[0]

            esptool_args = [
                '--chip', 'esp32s3',
                '--port', port,
                '--baud', '115200',
                '--after', 'hard-reset',
                'chip-id']

            try:
                import esptool
                esptool.main(argv=esptool_args)
            except SystemExit as e:
                # esptool calls sys.exit() on completion. 0 is success.
                exit_code = e.code if e.code is not None else 0
                if exit_code ==0:
                    QMessageBox.information(self, "Success", "Name and ID set successfully!")
                else:
                    QMessageBox.critical(self, "Error", "Failed to reset device after configuration.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to configure device: {str(e)}")

    def flash_esp32(self):
        selected_port_desc = self.port_combo.currentText()
        selected_version = self.version_combo.currentText()

        if not all([selected_port_desc, selected_version]) or not self.current_version_files:
            QMessageBox.critical(self, "Error", "A COM port and firmware version must be selected, and all required binary files must be present in the version folder.")
            return

        # Check if all required files are present
        required_files = ['bootloader', 'partition', 'ota_data', 'app']
        missing_files = [f for f in required_files if f not in self.current_version_files]
        if missing_files:
            QMessageBox.critical(self, "Error", f"Missing required files in {selected_version} folder: {', '.join(missing_files)}")
            return

        self.flash_button.setEnabled(False)
        self.progress_bar.show()
        self.status_label.setText("Flashing in progress...")
        self.output_console.clear()

        port = selected_port_desc.split(' - ')[0]

        esptool_args = [
            '--chip', 'esp32s3',
            '--port', port,
            '--baud', '115200',
            '--before', 'default-reset',
            '--after', 'hard-reset',
            'write_flash',
            '--flash-mode', 'keep',
            '--flash-freq', 'keep',
            '--flash-size', 'keep',
            '-z',
            '0x0', self.current_version_files['bootloader'],
            '0x8000', self.current_version_files['partition'],
            '0xe000', self.current_version_files['ota_data'],
            '0x10000', self.current_version_files['app']
        ]
        
        self.esptool_thread = QThread()
        self.esptool_worker = EsptoolWorker(esptool_args)
        self.esptool_worker.moveToThread(self.esptool_thread)

        self.esptool_worker.output.connect(self.append_output)
        self.esptool_worker.finished.connect(self.on_flash_finished)
        self.esptool_thread.started.connect(self.esptool_worker.run)
        
        # Clean up thread and worker
        self.esptool_worker.finished.connect(self.esptool_thread.quit)
        self.esptool_worker.finished.connect(self.esptool_worker.deleteLater)
        self.esptool_thread.finished.connect(self.esptool_thread.deleteLater)

        self.esptool_thread.start()

    @Slot(str)
    def append_output(self, text):
        self.output_console.append(text)

    @Slot(int)
    def on_flash_finished(self, exit_code):
        self.progress_bar.hide()
        self.flash_button.setEnabled(True)
        
        if exit_code == 0:
            self.status_label.setText("Flashing completed successfully!")
            QMessageBox.information(self, "Success", "Flashing completed successfully!")
        else:
            self.status_label.setText("Flashing failed!")
            QMessageBox.critical(self, "Error", "Flashing failed. Check the output console for details.")

        self.refresh_ports()
        self.refresh_versions()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ESPFlasherApp()
    window.show()
    sys.exit(app.exec())
