import os
import openpyxl
import sys
from PyQt5 import QtWidgets, QtGui, QtCore


class Student:
    def __init__(self, name, roll_number, age, std_id):
        self.name = name
        self.roll_number = roll_number
        self.age = age
        self.std_id = std_id

    def __str__(self):
        return f"{self.name} - {self.roll_number} - {self.age} - {self.std_id}"


def load_students():
    students = []
    if os.path.exists("students.txt"):
        with open("students.txt", "r") as file:
            for line in file.readlines():
                name, roll_number, age, std_id = line.strip().split(",")
                students.append(Student(name, roll_number, age, std_id))
    return students


def save_students(students):
    with open("students.txt", "w") as file:
        for student in students:
            file.write(f"{student.name},{student.roll_number},{student.age},{student.std_id}\n")


def export_to_excel(students):
    if not students:
        QtWidgets.QMessageBox.warning(None, "Export Failed", "No students to export.")
        return

    file_name = "students.xlsx"

    if os.path.exists(file_name):
        workbook = openpyxl.load_workbook(file_name)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Students"
        sheet.append(["Name", "Roll Number", "Age", "Student ID"])

    for student in students:
        sheet.append([student.name, student.roll_number, student.age, student.std_id])

    workbook.save(file_name)
    QtWidgets.QMessageBox.information(None, "Success", f"Data successfully exported to '{file_name}'.")
    os.startfile(file_name)


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(480, 620)

        # Background gradient
        self.background = QtWidgets.QLabel(Dialog)
        self.background.setGeometry(0, 0, 480, 620)
        self.background.setStyleSheet("background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, stop:0 rgba(0, 255, 187, 255), stop:1 rgba(0, 123, 255, 255));")

        # Login panel
        self.panel = QtWidgets.QFrame(Dialog)
        self.panel.setGeometry(QtCore.QRect(40, 150, 400, 320))
        self.panel.setStyleSheet("background-color: rgba(255, 255, 255, 0.8); border-radius: 20px;")

        # Login label
        self.label = QtWidgets.QLabel(self.panel)
        self.label.setGeometry(QtCore.QRect(140, 20, 120, 40))
        self.label.setStyleSheet("font: 24pt 'Segoe UI'; color: #2E3440;")
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setText("LOGIN")

        # Username field
        self.username = QtWidgets.QLineEdit(self.panel)
        self.username.setGeometry(QtCore.QRect(50, 90, 300, 50))
        self.username.setPlaceholderText("Username")
        self.username.setStyleSheet(
            "font: 16pt 'Segoe UI'; padding: 10px; background-color: white; border: 2px solid rgba(0, 0, 0, 0.2); border-radius: 10px;"
        )

        # Password field
        self.password = QtWidgets.QLineEdit(self.panel)
        self.password.setGeometry(QtCore.QRect(50, 160, 300, 50))
        self.password.setPlaceholderText("Password")
        self.password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.password.setStyleSheet(
            "font: 16pt 'Segoe UI'; padding: 10px; background-color: white; border: 2px solid rgba(0, 0, 0, 0.2); border-radius: 10px;"
        )

        # Show Password button
        self.show_password_button = QtWidgets.QPushButton(self.panel)
        self.show_password_button.setGeometry(QtCore.QRect(360, 160, 40, 50))
        self.show_password_button.setText("👁")
        self.show_password_button.setStyleSheet("font: bold 14pt 'Segoe UI'; color: black; background-color: white; border-radius: 5px;")
        self.show_password_button.setCheckable(True)
        self.show_password_button.clicked.connect(self.toggle_password_visibility)

        # Login button
        self.login_button = QtWidgets.QPushButton(self.panel)
        self.login_button.setGeometry(QtCore.QRect(50, 240, 300, 50))
        self.login_button.setText("SIGN IN")
        self.login_button.setStyleSheet(
            "font: bold 16pt 'Segoe UI'; color: white; background-color: #4CAF50; border-radius: 10px;"
        )

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def toggle_password_visibility(self):
        if self.show_password_button.isChecked():
            self.password.setEchoMode(QtWidgets.QLineEdit.Normal)
        else:
            self.password.setEchoMode(QtWidgets.QLineEdit.Password)

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle("Login")


class MainWin(QtWidgets.QDialog):
    def __init__(self):
        super(MainWin, self).__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.login_button.clicked.connect(self.check_login)
        self.show()

    def check_login(self):
        username = self.ui.username.text()
        password = self.ui.password.text()

        if username == "nilay123@gmail.com" and password == "nilay@123":
            self.open_student_management()
        else:
            QtWidgets.QMessageBox.warning(self, "Error", "Incorrect username or password.")

    def open_student_management(self):
        self.student_management_window = StudentManagementWindow()
        self.student_management_window.show()
        self.close()


class StudentManagementWindow(QtWidgets.QWidget):
    def __init__(self):
        super(StudentManagementWindow, self).__init__()
        self.setWindowTitle("Student Management System")
        self.setGeometry(100, 100, 600, 500)
        self.setStyleSheet("background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, stop:0 rgba(46, 52, 64, 255), stop:1 rgba(67, 76, 84, 255)); color: #ECEFF4;")
        self.students = load_students()
        self.init_ui()

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout()

        self.add_student_button = QtWidgets.QPushButton("Add Student", self)
        self.view_students_button = QtWidgets.QPushButton("View Students", self)
        self.delete_student_button = QtWidgets.QPushButton("Delete Student", self)
        self.export_button = QtWidgets.QPushButton("Export to Excel", self)

        for button in [self.add_student_button, self.view_students_button, self.delete_student_button, self.export_button]:
            button.setStyleSheet(
                "background-color: #88C0D0; color: #2E3440; font-size: 16px; font-weight: bold; border: none; border-radius: 8px; padding: 10px;"
            )
            button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))

        self.add_student_button.clicked.connect(self.add_student)
        self.view_students_button.clicked.connect(self.view_students)
        self.delete_student_button.clicked.connect(self.delete_student)
        self.export_button.clicked.connect(self.export_to_excel)

        layout.addWidget(self.add_student_button)
        layout.addWidget(self.view_students_button)
        layout.addWidget(self.delete_student_button)
        layout.addWidget(self.export_button)

        self.setLayout(layout)

    def add_student(self):
        name, ok1 = QtWidgets.QInputDialog.getText(self, "Add Student", "Enter name:")
        roll_number, ok2 = QtWidgets.QInputDialog.getText(self, "Add Student", "Enter roll number:")
        age, ok3 = QtWidgets.QInputDialog.getText(self, "Add Student", "Enter age:")
        std_id, ok4 = QtWidgets.QInputDialog.getText(self, "Add Student", "Enter student ID:")

        if ok1 and ok2 and ok3 and ok4:
            student = Student(name, roll_number, age, std_id)
            self.students.append(student)
            save_students(self.students)
            QtWidgets.QMessageBox.information(self, "Success", "Student added successfully.")

    def view_students(self):
        students_str = "\n".join(str(student) for student in self.students)
        QtWidgets.QMessageBox.information(self, "View Students", students_str)

    def delete_student(self):
        roll_number, ok = QtWidgets.QInputDialog.getText(self, "Delete Student", "Enter roll number to delete:")
        if ok:
            for student in self.students:
                if student.roll_number == roll_number:
                    self.students.remove(student)
                    save_students(self.students)
                    QtWidgets.QMessageBox.information(self, "Success", "Student deleted successfully.")
                    return
            QtWidgets.QMessageBox.warning(self, "Error", "Student not found.")

    def export_to_excel(self):
        export_to_excel(self.students)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWin = MainWin()
    sys.exit(app.exec_())