import sys
from PySide6.QtWidgets import QApplication
from gui import NameplateGeneratorGUI

def main():
    app = QApplication(sys.argv)
    window = NameplateGeneratorGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()