import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QMessageBox
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import winreg

class CheckJeppViewApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('JeppView 允许其他盘符安装工具')
        self.resize(781, 487)  # 设置窗口大小为 781x487 像素
        self.setWindowIcon(QIcon('C:\\Users\\18310\\Desktop\\新建文件夹 (5)\\蔡元培.jpg'))  # 设置窗口图标
        self.layout = QVBoxLayout()
        
        # 更新信息标签的文本及字体大小
        self.infoLabel = QLabel(
            "Powered By XiaoChennng 仅供学习交流使用\n"
            "此工具用来解决JeppView只能在C盘安装的问题\n"
            "您可以在安装JeppView时选择其他盘符进行安装 安装完成后运行本程序即可解决报错问题\n"
            "该方案同样适用于已经安装完毕的JeppView\n"
            "PS:由于需要修改注册表以实现JeppView的其他盘符安装\n"
            "     若有杀软拦截软件修改行为 请您放行 本软件绝对无任何其他功能", self)
        self.infoLabel.setFont(QFont('Arial', 12))
        self.layout.addWidget(self.infoLabel)
        
        # 设置检查按钮的字体大小
        self.checkButton = QPushButton('开始检查')
        self.checkButton.setFont(QFont('MiSans', 14))
        self.checkButton.clicked.connect(self.confirmation)
        self.layout.addWidget(self.checkButton)
        
        self.setStyleSheet("""
        QPushButton {
            background-color: #5e81ac;
            color: white;
            padding: 10px;
            margin: 10px;
            border-radius: 5px;
        }
        QPushButton:hover {
            background-color: #81a1c1;
        }
        """)
        
        self.setLayout(self.layout)

    def confirmation(self):
        reply = QMessageBox.question(self, '确认', '确定要开始检查吗？',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.checkReg()

    def checkReg(self):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Jeppesen\JeppView for Windows\Paths", 0, winreg.KEY_READ) as key:
                themes_path, _ = winreg.QueryValueEx(key, "Themes")
                
                if themes_path != r"C:\ProgramData\Jeppesen\JeppView for Windows\Themes\\":
                    self.updateReg(themes_path)
                else:
                    QMessageBox.information(self, "结果", "已设置正确。")
                
        except FileNotFoundError:
            QMessageBox.warning(self, "错误", "无法找到对应值，请检查JeppView是否正确安装。")
        except PermissionError:
            QMessageBox.warning(self, "错误", "无权限访问，请以管理员权限运行。")

    def updateReg(self, current_path):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Jeppesen\JeppView for Windows\Paths", 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "Themes", 0, winreg.REG_SZ, r"C:\ProgramData\Jeppesen\JeppView for Windows\Themes" "\\" )
                QMessageBox.information(self, "成功", "注册表更新成功")
        except Exception as e:
            QMessageBox.critical(self, "错误", "更新注册表时遇到问题。\n" + str(e))

def main():
    app = QApplication(sys.argv)
    ex = CheckJeppViewApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()