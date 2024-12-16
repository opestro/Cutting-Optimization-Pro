# Create a new file: gui.py

import sys
import json
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QFileDialog,
                           QVBoxLayout, QHBoxLayout, QWidget, QLabel, QSpinBox,
                           QTableWidget, QTableWidgetItem, QHeaderView, QProgressBar,
                           QComboBox, QAction, QToolButton, QMenu, QGroupBox, QLineEdit,
                           QMessageBox, QTextEdit, QDialog, QSplitter, QPlainTextEdit)
from PyQt5.QtCore import Qt, QTimer, QSize, QThread
from PyQt5.QtGui import QColor, QPalette, QFont, QIcon, QPixmap
import co
import pandas as pd
import webbrowser
import os
import requests
from packaging import version
import logging
import traceback
from datetime import datetime

# Set up logging at the start of the file
def setup_logger():
    """Set up logging configuration"""
    try:
        # Create logs directory if it doesn't exist
        if not os.path.exists('logs'):
            os.makedirs('logs')
        
        # Create log filename with timestamp
        log_filename = os.path.join('logs', f'debug_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
        
        # Configure logging
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                # File handler with UTF-8 encoding
                logging.FileHandler(log_filename, encoding='utf-8'),
                # Console handler
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        logger = logging.getLogger(__name__)
        logger.info("=== Application Starting ===")
        logger.info(f"Python version: {sys.version}")
        logger.info(f"Operating System: {os.name}")
        return logger
        
    except Exception as e:
        # If logging setup fails, write to emergency log file
        with open('emergency_log.txt', 'a') as f:
            f.write(f"{datetime.now()} - Error setting up logger: {str(e)}\n")
        raise

logger = setup_logger()

COLORS = {
    'dark': {
        'primary': '#0078d4',
        'background': '#1e1e1e',
        'surface': '#252526',
        'border': '#3d3d3d',
        'text': '#ffffff',
    },
    'light': {
        'primary': '#0078d4',
        'background': '#f0f0f0',
        'surface': '#ffffff',
        'border': '#d0d0d0',
        'text': '#000000',
    }
}

class SplashScreen(QWidget):
    def __init__(self):
        try:
            super().__init__()
            logger.debug("Initializing SplashScreen")
            
            self.setFixedSize(600, 300)
            self.setWindowFlag(Qt.FramelessWindowHint)
            logger.debug("SplashScreen window properties set")
            
            layout = QVBoxLayout()
            
            # Title
            title = QLabel("Cutting Optimizer Pro")
            title.setAlignment(Qt.AlignCenter)
            title.setStyleSheet("""
                QLabel {
                    color: #2c3e50;
                    font-size: 24px;
                    font-weight: bold;
                }
            """)
            
            # Progress Bar
            self.progress = QProgressBar()
            self.progress.setStyleSheet("""
                QProgressBar {
                    border: 2px solid #2c3e50;
                    border-radius: 5px;
                    text-align: center;
                }
                QProgressBar::chunk {
                    background-color: #3498db;
                }
            """)
            
            # Developer Info
            dev_info = QLabel("Developed with ❤️ by Mehdi Harzallah")
            dev_info.setAlignment(Qt.AlignCenter)
            dev_info.setStyleSheet("color: #7f8c8d;")
            
            contact = QLabel("Contact: +213 778191078")
            contact.setAlignment(Qt.AlignCenter)
            contact.setStyleSheet("color: #7f8c8d;")
            
            layout.addStretch()
            layout.addWidget(title)
            layout.addWidget(self.progress)
            layout.addWidget(dev_info)
            layout.addWidget(contact)
            layout.addStretch()
            
            self.setLayout(layout)
            
            # Start progress with delay
            self.counter = 0
            self.timer = QTimer()
            self.timer.timeout.connect(self.progress_update)
            self.timer.start(30)
            logger.debug("SplashScreen timer started")
            
        except Exception as e:
            logger.error(f"Error in SplashScreen initialization: {str(e)}")
            logger.error(traceback.format_exc())
            raise

    def progress_update(self):
        try:
            self.progress.setValue(self.counter)
            logger.debug(f"Progress update: {self.counter}%")
            
            if self.counter >= 100:
                logger.info("Progress reached 100%, creating main window")
                self.timer.stop()
                
                try:
                    logger.debug("Creating main window")
                    self.main_window = CuttingOptimizerGUI()
                    logger.debug("Main window created successfully")
                    self.main_window.show()
                    logger.debug("Main window shown")
                    self.close()
                    logger.debug("Splash screen closed")
                    
                except Exception as e:
                    logger.error(f"Error creating main window: {str(e)}")
                    logger.error(traceback.format_exc())
                    raise
                    
            self.counter += 1
            
        except Exception as e:
            logger.error(f"Error in progress update: {str(e)}")
            logger.error(traceback.format_exc())
            raise

class DebugWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(parent.tr("optimization_details"))
        self.setGeometry(200, 200, 1000, 700)
        
        # Create layout
        layout = QVBoxLayout()
        
        # Create splitter for results and debug output
        splitter = QSplitter()
        
        # Results text area (left side)
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        self.results_text.setStyleSheet("""
            QTextEdit {
                background-color: #2b2b2b;
                color: #a9b7c6;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
                border: 1px solid #3c3f41;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        
        # Debug output area (right side)
        self.debug_text = QPlainTextEdit()
        self.debug_text.setReadOnly(True)
        self.debug_text.setStyleSheet("""
            QPlainTextEdit {
                background-color: #1e1e1e;
                color: #00ff00;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
                border: 1px solid #3c3f41;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        
        # Add text areas to splitter
        splitter.addWidget(self.results_text)
        splitter.addWidget(self.debug_text)
        
        # Set splitter sizes to be equal
        splitter.setSizes([500, 500])
        
        # Close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #3c3f41;
                color: white;
                border: none;
                padding: 5px 20px;
                border-radius: 4px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #4c5052;
            }
        """)
        
        # Add widgets to layout
        layout.addWidget(splitter)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)
        
        self.setLayout(layout)
    
    def append_results(self, text):
        self.results_text.append(text)
    
    def append_debug(self, text, delay=False):
        cursor = self.debug_text.textCursor()
        cursor.movePosition(cursor.End)
        self.debug_text.setTextCursor(cursor)
        self.debug_text.insertPlainText(text + "\n")
        self.debug_text.verticalScrollBar().setValue(
            self.debug_text.verticalScrollBar().maximum()
        )
        if delay:
            QApplication.processEvents()  # Update UI
            QThread.msleep(100)  # Small delay for visual effect

class UpdateChecker:
    def __init__(self):
        try:
            # Get the directory where the executable/script is located
            if getattr(sys, 'frozen', False):
                # If running as compiled executable
                application_path = os.path.dirname(sys.executable)
            else:
                # If running as script
                application_path = os.path.dirname(os.path.abspath(__file__))
            
            version_file = os.path.join(application_path, 'version.json')
            
            if os.path.exists(version_file):
                with open(version_file, 'r') as f:
                    version_data = json.load(f)
                    self.current_version = version_data['version']
                    self.min_required = version_data['min_required']
                    logger.info(f"Loaded version: {self.current_version} (min required: {self.min_required})")
            else:
                logger.warning("version.json not found, creating default version file")
                self.current_version = "1.1.1"
                self.min_required = "1.1.1"
                default_version_data = {
                    "version": self.current_version,
                    "min_required": self.min_required
                }
                try:
                    with open(version_file, 'w') as f:
                        json.dump(default_version_data, f, indent=4)
                    logger.info(f"Created default version.json with version {self.current_version}")
                except Exception as e:
                    logger.error(f"Failed to create version.json: {str(e)}")
        except Exception as e:
            logger.error(f"Error reading version.json: {str(e)}")
            self.current_version = "1.1.1"  # Fallback version
            self.min_required = "1.1.1"
        
        # Update with your GitHub username and repo
        self.github_api_url = "https://api.github.com/repos/opestro/Cutting-Optimization-Pro/releases/latest"
        self.update_url = "https://github.com/opestro/Cutting-Optimization-Pro/releases/latest"

    def check_for_updates(self):
        try:
            response = requests.get(self.github_api_url)
            if response.status_code == 200:
                latest_version = response.json()["tag_name"].replace("v", "")
                return version.parse(latest_version) > version.parse(self.current_version)
        except Exception as e:
            logger.error(f"Error checking for updates: {e}")
            return False

class CuttingOptimizerGUI(QMainWindow):
    def __init__(self):
        try:
            super().__init__()
            logger.debug("Initializing CuttingOptimizerGUI")
            self.current_language = "en"
            self.dark_mode = True  # Default to dark mode
            self.profiles = {}
            
            try:
                self.update_checker = UpdateChecker()
            except Exception as e:
                logger.error(f"Failed to initialize update checker: {str(e)}")
                # Continue without update checker
            
            try:
                self.load_translations()
            except Exception as e:
                logger.error(f"Failed to load translations: {str(e)}")
                # Set default translations
                self.translations = {"en": {}}
            
            self.initUI()
            self.apply_theme()
            
            # Only check for updates if update_checker was initialized
            if hasattr(self, 'update_checker'):
                self.check_for_updates()
            
            logger.debug("CuttingOptimizerGUI initialized successfully")
        except Exception as e:
            logger.error(f"Error in CuttingOptimizerGUI initialization: {str(e)}")
            logger.error(traceback.format_exc())
            raise

    def get_dark_theme(self):
        return """
            /* Main Window */
            QMainWindow {
                background-color: #1e1e1e;
            }
            
            /* Toolbar */
            QToolBar {
                background-color: #2d2d2d;
                border-bottom: 2px solid #0078d4;
                spacing: 3px;
                padding: 3px;
            }
            
            /* Menu Bar */
            QMenuBar {
                background-color: #2d2d2d;
                color: #ffffff;
            }
            
            QMenuBar::item:selected {
                background-color: #0078d4;
            }
            
            /* Group Boxes */
            QGroupBox {
                background-color: #252526;
                border: 1px solid #3d3d3d;
                border-radius: 6px;
                margin-top: 12px;
                padding: 15px;
                font-weight: bold;
            }
            
            QGroupBox::title {
                color: #0078d4;
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 5px 10px;
                background-color: #252526;
            }
            
            /* Input Fields */
            QLineEdit, QSpinBox {
                background-color: #333333;
                color: #ffffff;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 5px;
                min-height: 16px;
                font-size: 13px;
            }
            
            QLineEdit:focus, QSpinBox:focus {
                border: 1px solid #0078d4;
            }
            
            /* Buttons */
            QPushButton {
                background-color: #2d2d2d;
                color: #ffffff;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 5px 16px;
                min-height: 18px;
                font-size: 13px;
                font-weight: bold;
            }
            
            QPushButton:hover {
                background-color: #3d3d3d;
                border: 1px solid #0078d4;
            }
            
            QPushButton#runButton {
                background-color: #0078d4;
                border: none;
            }
            
            QPushButton#runButton:hover {
                background-color: #1e90ff;
            }
            
            /* Table */
            QTableWidget {
                background-color: #252526;
                alternate-background-color: #2d2d2d;
                border: 1px solid #3d3d3d;
                gridline-color: #3d3d3d;
                color: #ffffff;
                selection-background-color: #0078d4;
            }
            
            QHeaderView::section {
                background-color: #2d2d2d;
                color: #0078d4;
                padding: 5px;
                border: none;
                font-weight: normal;
            }
            
            /* Labels */
            QLabel {
                color: #e1e1e1;
                font-size: 12px;
            }
            
            /* Status Bar */
            QStatusBar {
                background-color: #2d2d2d;
                color: #e1e1e1;
            }
        """

    def get_light_theme(self):
        return """
            /* Main Window */
            QMainWindow {
                background-color: #f0f0f0;
            }
            
            /* Toolbar */
            QToolBar {
                background-color: #ffffff;
                border-bottom: 2px solid #0078d4;
                spacing: 3px;
                padding: 3px;
            }
            
            /* Group Boxes */
            QGroupBox {
                background-color: #ffffff;
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                margin-top: 12px;
                padding: 15px;
                font-weight: bold;
            }
            
            QGroupBox::title {
                color: #0078d4;
                background-color: #ffffff;
            }
            
            /* Input Fields */
            QLineEdit, QSpinBox {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
                min-height: 16px;
                font-size: 13px;
            }
            
            QLineEdit:focus, QSpinBox:focus {
                border: 1px solid #0078d4;
            }
            
            /* Buttons */
            QPushButton {
                background-color: #f5f5f5;
                color: #000000;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px 16px;
                min-height: 18px;
                font-size: 13px;
                font-weight: bold;
            }
            
            QPushButton:hover {
                background-color: #e5e5e5;
                border: 1px solid #0078d4;
            }
            
            QPushButton#runButton {
                background-color: #0078d4;
                color: white;
                border: none;
            }
            
            /* Table */
            QTableWidget {
                background-color: #ffffff;
                alternate-background-color: #f5f5f5;
                border: 1px solid #d0d0d0;
                gridline-color: #d0d0d0;
                color: #000000;
            }
            
            QHeaderView::section {
                background-color: #f5f5f5;
                color: #0078d4;
                padding: 5px;
                border: none;
                font-weight: normal;
            }
            
            /* Labels */
            QLabel {
                color: #000000;
                font-size: 12px;
            }
        """

    def apply_theme(self):
        if self.dark_mode:
            self.setStyleSheet(self.get_dark_theme())
        else:
            self.setStyleSheet(self.get_light_theme())

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.apply_theme()

    def initUI(self):
        # Initialize icons with both dark and light versions
        self.icons = {
            'dark': {
                'app': QIcon('icons/dark/app.png'),
                'theme': QIcon('icons/dark/theme.png'),
                'file': QIcon('icons/dark/file.png'),
                'settings': QIcon('icons/dark/settings.png'),
                'run': QIcon('icons/dark/run.png'),
                'export': QIcon('icons/dark/export.png'),
                'add': QIcon('icons/dark/add.png'),
                'delete': QIcon('icons/dark/delete.png'),
                'help': QIcon('icons/dark/help.png'),
            },
            'light': {
                'app': QIcon('icons/light/app.png'),
                'theme': QIcon('icons/light/theme.png'),
                'file': QIcon('icons/light/file.png'),
                'settings': QIcon('icons/light/settings.png'),
                'run': QIcon('icons/light/run.png'),
                'export': QIcon('icons/light/export.png'),
                'add': QIcon('icons/light/add.png'),
                'delete': QIcon('icons/light/delete.png'),
                'help': QIcon('icons/light/help.png'),
            }
        }

        # Create menu bar
        menubar = self.menuBar()
        view_menu = menubar.addMenu(self.tr('View'))
        
        # Add theme toggle action
        theme_action = QAction(QIcon('icons/theme.png'), self.tr('Toggle Theme'), self)
        theme_action.setShortcut('Ctrl+T')
        theme_action.triggered.connect(self.toggle_theme)
        view_menu.addAction(theme_action)

        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Create toolbar
        toolbar = self.addToolBar('Main Toolbar')
        toolbar.setIconSize(QSize(24, 24))

        # Language selector
        lang_button = QToolButton()
        lang_button.setIcon(QIcon('icons/language.png'))
        lang_button.setPopupMode(QToolButton.InstantPopup)
        lang_menu = QMenu()
        lang_menu.addAction("English", lambda: self.change_language("en"))
        lang_menu.addAction("Français", lambda: self.change_language("fr"))
        lang_menu.addAction("العربية", lambda: self.change_language("ar"))
        lang_button.setMenu(lang_menu)
        toolbar.addWidget(lang_button)

        # File selection section
        file_layout = QHBoxLayout()
        self.work_file_btn = QPushButton(self.tr('select_work_file'))
        self.work_file_btn.clicked.connect(self.select_work_file)
        self.work_file_label = QLabel(self.tr('no_file_selected'))
        file_layout.addWidget(self.work_file_btn)
        file_layout.addWidget(self.work_file_label)
        layout.addLayout(file_layout)

        # Settings section
        settings_layout = QHBoxLayout()
        self.default_length_label = QLabel(self.tr('default_length'))
        self.default_length_spin = QSpinBox()
        self.default_length_spin.setRange(1000, 20000)
        self.default_length_spin.setValue(12000)
        self.default_length_spin.setSuffix(" mm")
        
        self.kerf_width_label = QLabel(self.tr('kerf_width'))
        self.kerf_width_spin = QSpinBox()
        self.kerf_width_spin.setRange(1, 100)
        self.kerf_width_spin.setValue(3)
        self.kerf_width_spin.setSuffix(" mm")
        
        settings_layout.addWidget(self.default_length_label)
        settings_layout.addWidget(self.default_length_spin)
        settings_layout.addWidget(self.kerf_width_label)
        settings_layout.addWidget(self.kerf_width_spin)
        layout.addLayout(settings_layout)

        # Profile group
        self.profile_group = QGroupBox(self.tr('profile_lengths'))
        profile_layout = QVBoxLayout()
        
        # Profile input section
        input_layout = QHBoxLayout()
        self.profile_name_label = QLabel(self.tr('profile_name'))
        self.profile_name = QLineEdit()
        self.profile_length_label = QLabel(self.tr('profile_length'))
        self.profile_length = QSpinBox()
        self.profile_length.setRange(1, 20000)
        self.profile_length.setSuffix(" mm")
        self.add_profile_btn = QPushButton(self.tr('add_profile'))
        self.add_profile_btn.clicked.connect(self.add_profile)
        
        input_layout.addWidget(self.profile_name_label)
        input_layout.addWidget(self.profile_name)
        input_layout.addWidget(self.profile_length_label)
        input_layout.addWidget(self.profile_length)
        input_layout.addWidget(self.add_profile_btn)
        profile_layout.addLayout(input_layout)

        # Profile table
        self.profile_table = QTableWidget()
        self.profile_table.setColumnCount(5)
        headers = [
            self.tr('profile'),
            self.tr('length'),
            self.tr('quantity'),
            self.tr('stock_length'),
            ''  # For delete button
        ]
        self.profile_table.setHorizontalHeaderLabels(headers)
        self.profile_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        profile_layout.addWidget(self.profile_table)
        
        self.profile_group.setLayout(profile_layout)
        layout.addWidget(self.profile_group)

        # Run button and status
        bottom_layout = QHBoxLayout()
        self.run_btn = QPushButton(self.tr('run_optimization'))
        self.run_btn.clicked.connect(self.run_optimization)
        self.status_label = QLabel("")
        bottom_layout.addWidget(self.run_btn)
        bottom_layout.addWidget(self.status_label)
        layout.addLayout(bottom_layout)

        # Set window properties
        self.setWindowTitle(self.tr('app_title'))
        self.setGeometry(100, 100, 800, 600)
        self.updateUI()

    def updateUI(self):
        """Update all UI elements with current language"""
        self.setWindowTitle(self.tr('app_title'))
        self.work_file_btn.setText(self.tr('select_work_file'))
        self.work_file_label.setText(self.tr('no_file_selected'))
        self.profile_group.setTitle(self.tr('profile_lengths'))
        self.default_length_label.setText(self.tr('default_length'))
        self.kerf_width_label.setText(self.tr('kerf_width'))
        self.profile_name_label.setText(self.tr('profile_name'))
        self.profile_length_label.setText(self.tr('profile_length'))
        self.add_profile_btn.setText(self.tr('add_profile'))
        self.run_btn.setText(self.tr('run_optimization'))

        # Update table headers
        headers = [
            self.tr('profile'),
            self.tr('length'),
            self.tr('quantity'),
            self.tr('stock_length'),
            ''  # For delete button
        ]
        self.profile_table.setHorizontalHeaderLabels(headers)

    def select_work_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, 'Select Work File', '',
            'Excel Files (*.xlsx *.xls)'
        )
        if filename:
            self.work_file_label.setText(filename)
            self.detect_profiles(filename)
            
    def detect_profiles(self, filename):
        try:
            data = pd.read_excel(filename, engine='openpyxl')
            data = data.iloc[2:]  # Skip first two rows
            data.columns = data.iloc[0]  # Use the first row as headers
            data = data[1:]  # Skip the header row
            
            # Reset profiles
            self.profiles = {}
            
            # Convert data types and handle NaN values
            data['Qté'] = pd.to_numeric(data['Qté'], errors='coerce').fillna(1).astype(int)
            data['Long.'] = pd.to_numeric(data['Long.'], errors='coerce').fillna(0).astype(int)
            
            # Group by profile and collect all lengths
            for _, row in data.iterrows():
                profile = str(row['Profil'])  # Convert to string
                if profile and 'Total' not in profile:
                    if profile not in self.profiles:
                        self.profiles[profile] = []
                    
                    self.profiles[profile].append({
                        'length': int(row['Long.']),  # Convert to int
                        'qty': int(row['Qté'])        # Convert to int
                    })
            
            self.update_profile_table()
            self.status_label.setText(f'Detected {len(self.profiles)} profiles')
            
        except Exception as e:
            self.status_label.setText(f'Error detecting profiles: {str(e)}')
            print(f"Error details: {str(e)}")  # For debugging

    def update_profile_table(self):
        # Store current values before clearing
        current_values = {}
        for row in range(self.profile_table.rowCount()):
            profile_item = self.profile_table.item(row, 0)
            length_item = self.profile_table.item(row, 1)
            
            if profile_item and length_item:
                profile_name = profile_item.text()
                length = int(length_item.text())
                qty = self.profile_table.cellWidget(row, 2).value()
                stock = self.profile_table.cellWidget(row, 3).value()
                
                # Store the current values
                current_values[(profile_name, length)] = (qty, stock)
                
                # Update the profiles dictionary
                if profile_name in self.profiles:
                    for length_data in self.profiles[profile_name]:
                        if length_data['length'] == length:
                            length_data['qty'] = qty
                            length_data['stock_length'] = stock

        # Clear and setup table
        self.profile_table.setRowCount(0)
        self.profile_table.setColumnCount(5)
        self.profile_table.setHorizontalHeaderLabels([
            self.tr('Profile'),
            self.tr('Length (mm)'),
            self.tr('Quantity'),
            self.tr('Stock Length (mm)'),
            self.tr('Action')
        ])
        
        # Add all profiles
        for profile_name, lengths in self.profiles.items():
            for length_data in lengths:
                current_row = self.profile_table.rowCount()
                self.profile_table.insertRow(current_row)
                
                # Profile name
                self.profile_table.setItem(current_row, 0, QTableWidgetItem(str(profile_name)))
                
                # Length - use the stored length
                self.profile_table.setItem(current_row, 1, QTableWidgetItem(str(length_data['length'])))
                
                # Get stored or previous values
                prev_qty, prev_stock = current_values.get(
                    (profile_name, length_data['length']),
                    (length_data.get('qty', 1), length_data.get('stock_length', self.default_length_spin.value()))
                )
                
                # Quantity
                qty_spin = QSpinBox()
                qty_spin.setRange(1, 1000)
                qty_spin.setValue(prev_qty)
                self.profile_table.setCellWidget(current_row, 2, qty_spin)
                
                # Stock length
                stock_spin = QSpinBox()
                stock_spin.setRange(1000, 20000)
                stock_spin.setValue(prev_stock)
                stock_spin.setSuffix(" mm")
                self.profile_table.setCellWidget(current_row, 3, stock_spin)
                
                # Delete button
                delete_btn = QPushButton("🗑️")
                delete_btn.clicked.connect(lambda checked, r=current_row: self.delete_profile_row(r))
                self.profile_table.setCellWidget(current_row, 4, delete_btn)

    def add_profile(self):
        name = self.profile_name.text()
        length = self.profile_length.value()
        
        if not name:
            QMessageBox.warning(self, "Warning", self.tr("Please enter a profile name"))
            return
        
        # Initialize profile if it doesn't exist
        if name not in self.profiles:
            self.profiles[name] = []
        
        # Check if length already exists for this profile
        length_exists = any(l['length'] == length for l in self.profiles[name])
        
        if not length_exists:
            # Add new length to profile with user-set values
            self.profiles[name].append({
                'length': length,
                'qty': 1,  # Default quantity
                'stock_length': self.default_length_spin.value()  # Store stock length
            })
            
            self.update_profile_table()
            # Only reset length input, keep the profile name
            self.profile_length.setValue(0)

    def delete_profile_row(self, row):
        profile_name = self.profile_table.item(row, 0).text()
        length = int(self.profile_table.item(row, 1).text())  # Convert to int
        
        # Remove specific length entry from profile
        if profile_name in self.profiles:
            self.profiles[profile_name] = [
                l for l in self.profiles[profile_name] 
                if l['length'] != length
            ]
            
            # Remove profile if no lengths remain
            if not self.profiles[profile_name]:
                del self.profiles[profile_name]
        
        self.update_profile_table()

    def get_profile_settings(self):
        settings = {}
        for row in range(self.profile_table.rowCount()):
            profile = self.profile_table.item(row, 0).text()
            length = self.profile_table.cellWidget(row, 3).value()
            settings[profile] = length
        return settings

    def update_all_lengths(self):
        default_length = self.default_length_spin.value()
        for row in range(self.profile_table.rowCount()):
            length_spin = self.profile_table.cellWidget(row, 3)
            if length_spin:
                length_spin.setValue(default_length)
                
    def run_optimization(self):
        try:
            debug_window = DebugWindow(self)
            debug_window.show()
            
            debug_window.append_debug(self.tr("starting_optimization"), delay=True)
            debug_window.append_debug(f"{self.tr('default_stock_length')}: {self.default_length_spin.value()}mm", delay=True)
            debug_window.append_debug(f"{self.tr('kerf_width')}: {self.kerf_width_spin.value()}mm\n", delay=True)
            
            settings_data = {
                'Profile': [],
                'Stock Length': [],
                'Total Needed': []
            }
            
            debug_window.append_debug(self.tr("collecting_profile_data"), delay=True)
            
            # Collect data with progress updates
            for row in range(self.profile_table.rowCount()):
                profile_item = self.profile_table.item(row, 0)
                if profile_item and profile_item.text():
                    profile = profile_item.text()
                    length_item = self.profile_table.item(row, 1)
                    qty_spin = self.profile_table.cellWidget(row, 2)
                    stock_spin = self.profile_table.cellWidget(row, 3)
                    
                    if length_item and qty_spin and stock_spin:
                        piece_length = int(length_item.text())
                        quantity = qty_spin.value()
                        stock_length = stock_spin.value()
                        
                        debug_window.append_debug(
                            f"\n{self.tr('profile')}: {profile}\n"
                            f"  {self.tr('length')}: {piece_length}mm\n"
                            f"  {self.tr('quantity')}: {quantity}\n"
                            f"  {self.tr('stock_length')}: {stock_length}mm",
                            delay=True
                        )
                        
                        settings_data['Profile'].append(profile)
                        settings_data['Stock Length'].append(stock_length)
                        settings_data['Total Needed'].append(piece_length * quantity)
            
            settings_df = pd.DataFrame(settings_data)
            debug_window.append_debug("\nCreating optimization data...")
            
            # Create optimization data
            optimization_data = []
            for row in range(self.profile_table.rowCount()):
                profile_item = self.profile_table.item(row, 0)
                if profile_item and profile_item.text():
                    profile = profile_item.text()
                    length_item = self.profile_table.item(row, 1)
                    qty_spin = self.profile_table.cellWidget(row, 2)
                    
                    if length_item and qty_spin:
                        optimization_data.append({
                            'Profil': profile,
                            'Long.': int(length_item.text()),
                            'Qté': qty_spin.value()
                        })
            
            data_df = pd.DataFrame(optimization_data)
            debug_window.append_debug("\nRunning optimization algorithm...")
            
            # Run optimization
            waste_stats = co.main(
                data_df, 
                settings_df,
                'optimized_cutting_plan.xlsx',
                'cutting_plan.png',
                self.default_length_spin.value()
            )
            
            if waste_stats:
                results_text = self.tr('optimization_results') + "\n\n"
                for profile, stats in waste_stats.items():
                    results_text += self.tr('results_profile').format(profile) + "\n"
                    results_text += self.tr('results_used_length').format(stats['used_length']) + "\n"
                    results_text += self.tr('results_total_stock').format(stats['total_stock']) + "\n"
                    results_text += self.tr('results_waste').format(stats['waste_percentage']) + "\n\n"
                
                debug_window.append_results(results_text)
                debug_window.append_debug("\nOptimization completed successfully!")
                
                self.status_label.setText("Optimization completed!")
            
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            self.status_label.setText(error_msg)
            if 'debug_window' in locals():
                debug_window.append_debug(f"\nERROR: {error_msg}")
            print(f"Error details: {str(e)}")

    def change_language(self, language):
        """Change application language"""
        self.current_language = language
        
        # Handle RTL for Arabic
        if language == "ar":
            self.setLayoutDirection(Qt.RightToLeft)
        else:
            self.setLayoutDirection(Qt.LeftToRight)
        
        self.updateUI()
        
    def load_translations(self):
        with open('translations.json', 'r', encoding='utf-8') as f:
            self.translations = json.load(f)
            
    def tr(self, key):
        return self.translations[self.current_language].get(key, key)

    def get_current_icons(self):
        return self.icons['dark' if self.dark_mode else 'light']

    def check_for_updates(self):
        """Check for software updates"""
        update_available = self.update_checker.check_for_updates()
        
        if update_available:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle(self.tr('Update Available'))
            msg.setText(self.tr('A new version is available!'))
            msg.setInformativeText(self.tr('Would you like to download the update?'))
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            
            if msg.exec_() == QMessageBox.Yes:
                webbrowser.open(self.update_checker.update_url)
        else:
            # Only show if manually checked (not on startup)
            if hasattr(self, 'isVisible') and self.isVisible():
                QMessageBox.information(
                    self,
                    self.tr('No Updates'),
                    self.tr('You are using the latest version.')
                )

    def auto_check_updates(self):
        """Automatically check for updates periodically"""
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.check_for_updates)
        # Check every 24 hours (in milliseconds)
        self.update_timer.start(24 * 60 * 60 * 1000)

def main():
    try:
        logger = setup_logger()
        logger.info("Starting application")
        
        app = QApplication(sys.argv)
        logger.info("QApplication created")
        
        try:
            splash = SplashScreen()
            logger.info("SplashScreen created")
            splash.show()
            logger.info("SplashScreen shown")
            
            # Add a small delay to ensure splash screen is visible
            QTimer.singleShot(100, lambda: logger.info("Splash screen timer triggered"))
            
            result = app.exec_()
            logger.info(f"Application exited with code: {result}")
            sys.exit(result)
            
        except Exception as e:
            logger.error(f"Error in main application loop: {str(e)}")
            logger.error(traceback.format_exc())
            raise
            
    except Exception as e:
        # If logger isn't set up, write to emergency file
        with open('emergency_log.txt', 'a') as f:
            f.write(f"{datetime.now()} - Critical error in main: {str(e)}\n")
            f.write(traceback.format_exc())
        sys.exit(1)

if __name__ == '__main__':
    main()