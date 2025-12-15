"""
נקודת כניסה ראשית לאפליקציית PDF Rename Tool
"""
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt
import ui_main


def main():
    """פונקציה ראשית"""
    # יצירת אפליקציית Qt
    app = QApplication(sys.argv)
    
    # הגדרת שם האפליקציה
    app.setApplicationName("PDF Rename Tool")
    
    # תמיכה ב-RTL (עברית)
    app.setLayoutDirection(Qt.RightToLeft)
    
    # יצירת חלון ראשי
    window = ui_main.MainWindow()
    window.show()
    
    # הרצת לולאת האירועים
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

