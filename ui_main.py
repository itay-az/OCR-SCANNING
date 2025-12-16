"""
ממשק משתמש ראשי לאפליקציית עיבוד PDF
"""
import os
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLineEdit, QLabel, QTextEdit,
                             QFileDialog, QMessageBox, QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
import pdf_processor
import scanner_module


class ProcessingThread(QThread):
    """Thread לעיבוד PDFs ברקע כדי לא לחסום את ה-UI"""
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(dict)
    
    def __init__(self, source_folder, destination_folder, regex_pattern):
        super().__init__()
        self.source_folder = source_folder
        self.destination_folder = destination_folder
        self.regex_pattern = regex_pattern
    
    def run(self):
        """מריץ את עיבוד התיקייה"""
        def log_callback(message):
            self.log_signal.emit(message)
        
        stats = pdf_processor.process_folder_with_destination(
            self.source_folder,
            self.destination_folder,
            self.regex_pattern,
            log_callback
        )
        self.finished_signal.emit(stats)


class ScanningThread(QThread):
    """Thread לסריקה ברקע"""
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool)
    
    def __init__(self, output_folder, regex_pattern):
        super().__init__()
        self.output_folder = output_folder
        self.regex_pattern = regex_pattern
    
    def run(self):
        """מריץ את הסריקה"""
        def log_callback(message):
            self.log_signal.emit(message)
        
        try:
            result = scanner_module.scan_and_process(
                self.output_folder,
                self.regex_pattern,
                log_callback
            )
            self.finished_signal.emit(result is not None)
        except Exception as e:
            log_callback(f"שגיאה: {e}")
            self.finished_signal.emit(False)


class MainWindow(QMainWindow):
    """חלון ראשי של האפליקציה"""
    
    def __init__(self):
        super().__init__()
        self.selected_folder = ""  # תיקיית יעד (לסריקה)
        self.source_folder = ""  # תיקיית מקור (לעיבוד)
        self.processing_thread = None
        self.scan_thread = None
        self.init_ui()
        self.check_tesseract_on_startup()
    
    def init_ui(self):
        """אתחול ממשק המשתמש"""
        self.setWindowTitle("PDF Rename Tool - שינוי שמות קבצי PDF")
        self.setGeometry(100, 100, 800, 600)
        
        # Widget מרכזי
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout ראשי
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # כותרת
        title_label = QLabel("כלי שינוי שמות קבצי PDF")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # בחירת תיקיית מקור
        source_folder_layout = QHBoxLayout()
        source_folder_label = QLabel("תיקיית מקור:")
        source_folder_label.setMinimumWidth(80)
        self.source_folder_path_edit = QLineEdit()
        self.source_folder_path_edit.setPlaceholderText("בחר תיקיית מקור...")
        self.source_folder_path_edit.setReadOnly(True)
        browse_source_button = QPushButton("בחר תיקיית מקור")
        browse_source_button.clicked.connect(self.browse_source_folder)
        
        source_folder_layout.addWidget(source_folder_label)
        source_folder_layout.addWidget(self.source_folder_path_edit)
        source_folder_layout.addWidget(browse_source_button)
        main_layout.addLayout(source_folder_layout)
        
        # בחירת תיקיית יעד
        destination_folder_layout = QHBoxLayout()
        destination_folder_label = QLabel("תיקיית יעד:")
        destination_folder_label.setMinimumWidth(80)
        self.folder_path_edit = QLineEdit()
        self.folder_path_edit.setPlaceholderText("בחר תיקיית יעד...")
        self.folder_path_edit.setReadOnly(True)
        browse_button = QPushButton("בחר תיקיית יעד")
        browse_button.clicked.connect(self.browse_folder)
        
        destination_folder_layout.addWidget(destination_folder_label)
        destination_folder_layout.addWidget(self.folder_path_edit)
        destination_folder_layout.addWidget(browse_button)
        main_layout.addLayout(destination_folder_layout)
        
        # שדה REGEX
        regex_layout = QHBoxLayout()
        regex_label = QLabel("REGEX:")
        regex_label.setMinimumWidth(80)
        self.regex_edit = QLineEdit()
        self.regex_edit.setPlaceholderText("לדוגמה: \\d{9} (9 ספרות רצופות)")
        regex_layout.addWidget(regex_label)
        regex_layout.addWidget(self.regex_edit)
        main_layout.addLayout(regex_layout)
        
        # --- הוסף את השורה הזו כאן: ---
        self.regex_edit.setText(r'\b\d{9}\b')  # הגדרת ברירת מחדל לתעודת זהות
        # -----------------------------
        
        # כפתור הרצה
        self.run_button = QPushButton("הרץ עיבוד")
        self.run_button.setMinimumHeight(40)
        self.run_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.run_button.clicked.connect(self.start_processing)
        main_layout.addWidget(self.run_button)
        
        # כפתור סריקה
        scan_layout = QHBoxLayout()
        self.scan_button = QPushButton("סרוק מסמך")
        self.scan_button.setMinimumHeight(40)
        self.scan_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.scan_button.clicked.connect(self.start_scanning)
        scan_layout.addWidget(self.scan_button)
        main_layout.addLayout(scan_layout)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # אזור לוג
        log_label = QLabel("לוג פעילות:")
        main_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Courier", 9))
        self.log_text.setPlaceholderText("הלוג יופיע כאן...")
        main_layout.addWidget(self.log_text)
        
        # הודעת ברוכים הבאים
        self.log_text.append("ברוכים הבאים לכלי שינוי שמות קבצי PDF")
        self.log_text.append("1. בחר תיקיית מקור (מכילה קבצי PDF searchable)")
        self.log_text.append("2. בחר תיקיית יעד (לשמירת הקבצים המסודרים)")
        self.log_text.append("3. הזן תבנית REGEX לחיפוש")
        self.log_text.append("4. לחץ על 'הרץ עיבוד' לעיבוד תיקייה או 'סרוק מסמך' לסריקה ישירה\n")
    
    def check_tesseract_on_startup(self):
        """בודק זמינות Tesseract OCR בהתחלה"""
        try:
            import pytesseract
            pytesseract.get_tesseract_version()
            self.log_text.append("✓ Tesseract OCR זמין - תמיכה ב-OCR מופעלת\n")
        except Exception:
            self.log_text.append("⚠ Tesseract OCR לא נמצא - רק PDFs searchable יעובדו\n")
            self.log_text.append("  (להורדת Tesseract: https://github.com/UB-Mannheim/tesseract/wiki)\n")
    
    def browse_source_folder(self):
        """פתיחת דיאלוג לבחירת תיקיית מקור"""
        folder = QFileDialog.getExistingDirectory(
            self,
            "בחר תיקיית מקור",
            "",
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )
        if folder:
            self.source_folder = folder
            self.source_folder_path_edit.setText(folder)
            self.log_text.append(f"נבחרה תיקיית מקור: {folder}\n")
    
    def browse_folder(self):
        """פתיחת דיאלוג לבחירת תיקיית יעד"""
        folder = QFileDialog.getExistingDirectory(
            self,
            "בחר תיקיית יעד",
            "",
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )
        if folder:
            self.selected_folder = folder
            self.folder_path_edit.setText(folder)
            self.log_text.append(f"נבחרה תיקיית יעד: {folder}\n")
    
    def validate_inputs(self):
        """בדיקת תקינות הקלט"""
        if not self.source_folder:
            QMessageBox.warning(
                self,
                "שגיאה",
                "אנא בחר תיקיית מקור"
            )
            return False
        
        if not self.selected_folder:
            QMessageBox.warning(
                self,
                "שגיאה",
                "אנא בחר תיקיית יעד"
            )
            return False
        
        if not self.regex_edit.text().strip():
            QMessageBox.warning(
                self,
                "שגיאה",
                "אנא הזן תבנית REGEX"
            )
            return False
        
        # בדיקת תקינות REGEX
        import re
        try:
            re.compile(self.regex_edit.text().strip())
        except re.error as e:
            QMessageBox.warning(
                self,
                "שגיאת REGEX",
                f"תבנית REGEX לא תקינה:\n{e}"
            )
            return False
        
        return True
    
    def start_processing(self):
        """התחלת עיבוד הקבצים"""
        if not self.validate_inputs():
            return
        
        # בדיקה אם כבר רץ עיבוד
        if self.processing_thread and self.processing_thread.isRunning():
            QMessageBox.warning(
                self,
                "עיבוד בתהליך",
                "עיבוד כבר רץ. אנא המתן לסיום."
            )
            return
        
        # בדיקת הרשאות כתיבה לתיקיית יעד
        if not os.access(self.selected_folder, os.W_OK):
            QMessageBox.warning(
                self,
                "שגיאת הרשאות",
                "אין הרשאות כתיבה לתיקיית היעד שנבחרה. אנא בחר תיקייה אחרת."
            )
            return
        
        # בדיקת הרשאות קריאה מתיקיית מקור
        if not os.access(self.source_folder, os.R_OK):
            QMessageBox.warning(
                self,
                "שגיאת הרשאות",
                "אין הרשאות קריאה מתיקיית המקור שנבחרה. אנא בחר תיקייה אחרת."
            )
            return
        
        # ניקוי לוג
        self.log_text.clear()
        self.log_text.append("מתחיל עיבוד...\n")
        
        # השבתת כפתור
        self.run_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        
        # יצירת thread לעיבוד
        regex_pattern = self.regex_edit.text().strip()
        self.processing_thread = ProcessingThread(
            self.source_folder,
            self.selected_folder,
            regex_pattern
        )
        self.processing_thread.log_signal.connect(self.append_log)
        self.processing_thread.finished_signal.connect(self.processing_finished)
        self.processing_thread.start()
    
    def start_scanning(self):
        """התחלת סריקה ישירה"""
        if not self.selected_folder:
            QMessageBox.warning(
                self,
                "שגיאה",
                "אנא בחר תיקייה לשמירת הקבצים"
            )
            return
        
        if not self.regex_edit.text().strip():
            QMessageBox.warning(
                self,
                "שגיאה",
                "אנא הזן תבנית REGEX"
            )
            return
        
        # בדיקת תקינות REGEX
        import re
        try:
            re.compile(self.regex_edit.text().strip())
        except re.error as e:
            QMessageBox.warning(
                self,
                "שגיאת REGEX",
                f"תבנית REGEX לא תקינה:\n{e}"
            )
            return
        
        # בדיקה אם יש סורקים זמינים
        try:
            scanners = scanner_module.get_scanners()
            if not scanners:
                QMessageBox.warning(
                    self,
                    "לא נמצא סורק",
                    "לא נמצא סורק זמין. ודא שהסורק מחובר ומופעל."
                )
                return
        except Exception as e:
            QMessageBox.warning(
                self,
                "שגיאה",
                f"שגיאה בחיבור לסורק:\n{e}\n\nודא שהסורק מחובר ומופעל."
            )
            return
        
        # השבתת כפתור
        self.scan_button.setEnabled(False)
        self.log_text.append("\n=== מתחיל סריקה ===\n")
        
        # הרצת סריקה ב-thread נפרד
        regex_pattern = self.regex_edit.text().strip()
        self.scan_thread = ScanningThread(
            self.selected_folder,
            regex_pattern
        )
        self.scan_thread.log_signal.connect(self.append_log)
        self.scan_thread.finished_signal.connect(self.scanning_finished)
        self.scan_thread.start()
    
    def scanning_finished(self, success):
        """טיפול בסיום הסריקה"""
        self.scan_button.setEnabled(True)
        if success:
            QMessageBox.information(
                self,
                "סריקה הושלמה",
                "הסריקה הושלמה בהצלחה!"
            )
        else:
            QMessageBox.warning(
                self,
                "שגיאה",
                "אירעה שגיאה בסריקה. בדוק את הלוג לפרטים."
            )
    
    def append_log(self, message):
        """הוספת הודעה ללוג"""
        self.log_text.append(message)
        # גלילה אוטומטית למטה
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def processing_finished(self, stats):
        """טיפול בסיום העיבוד"""
        self.progress_bar.setVisible(False)
        self.run_button.setEnabled(True)
        
        # הצגת הודעה
        success_msg = f"העיבוד הושלם!\n\n"
        success_msg += f"הושלמו בהצלחה: {stats['success_count']}\n"
        success_msg += f"לא נמצאה התאמה: {stats['failed_count']}"
        
        if stats['errors']:
            success_msg += f"\nשגיאות: {len(stats['errors'])}"
        
        QMessageBox.information(
            self,
            "עיבוד הושלם",
            success_msg
        )

