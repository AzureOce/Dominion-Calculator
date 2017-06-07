# -*- coding: utf-8 -*-
import os
import sys
import threading
import time
from datetime import datetime

import xlsxwriter
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox

import calc
import gui


def tests():
    v = threading.Thread(target=progress_start)
    v.start()
    if v.isAlive:
        v.join()


def progress_start():
    ui.progressBar.setMaximum(100)
    ui.progressBar.setValue(0)
    count = 0
    while count < 100:
        count += 1
        time.sleep(1)
        ui.progressBar.setValue(count)
        QApplication.processEvents()


def select_origin_dir():
    path = QFileDialog.getExistingDirectory(None, '选择源文件目录')
    ui.lineEdit.setText(path)
    print(ui.lineEdit.text())


def select_result_dir():
    path = QFileDialog.getExistingDirectory(None, '选择保存目录')
    ui.lineEdit_2.setText(path)
    print(ui.lineEdit_2.text())


def show_error_dialog():
    ui.textBrowser.setText("参数检查不通过，校验失败")
    ui.pushButton.setEnabled(False)
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText("Error")
    msg.setInformativeText("请检查下文件目录是否正确选择，保存文件名是否填写。")
    msg.setWindowTitle("瞬间爆炸")
    msg.exec()
    ui.pushButton.setEnabled(True)


def add_log(log):
    ui.textBrowser.append(log)


def exec_calc():
    ui.pushButton.setEnabled(False)
    add_log("校验通过，开始执行计算")
    list_file = calc.list_files(str(ui.lineEdit.text()))
    add_log("读取目录，文件个数:%s" % len(list_file))
    file_num = len(list_file)
    if file_num <= 0:
        add_log("指定目录下没有找到符合要求需要处理的文件，处理完成。")
        ui.progressBar.setMaximum(100)
        ui.progressBar.setValue(100)
        ui.pushButton.setEnabled(True)
        return
    start_time = datetime.now()
    ui.progressBar.setMaximum(file_num + 1)
    ui.progressBar.setValue(0)
    add_log("初始化Excel")
    work_book = xlsxwriter.Workbook(os.path.join(str(ui.lineEdit_2.text()), "%s.xlsx" % str(ui.lineEdit_3.text())))
    work_sheet = work_book.add_worksheet()
    work_sheet.set_column('A:A', 40)
    work_sheet.set_column('B:A', 30)
    work_sheet.set_column('C:A', 30)

    style_bold = work_book.add_format({'bold': True, 'font_size': 16})
    style_normal = work_book.add_format({'font_size': 14})
    style_normal.set_align('center')
    style_normal.set_align('vcenter')
    style_bold.set_align('center')
    style_bold.set_align('vcenter')
    work_sheet.write(0, 0, "文件名字", style_bold)
    work_sheet.write(0, 1, "层数等级", style_bold)
    work_sheet.write(0, 2, "主权总数", style_bold)
    add_log("初始化Excel完成")
    index = 1
    for txt in list_file:
        add_log("解析文件：%s" % txt)
        dict_step = calc.parser_txt_content(txt)
        dict_index = calc.pattern_dict_privacy(dict_step)
        level = calc.calc_level(dict_index)
        add_log("写入Excel:  %s ==> %s" % (str(os.path.splitext(os.path.basename(txt))[0]), str(level)))
        work_sheet.write(index, 0, os.path.splitext(os.path.basename(txt))[0], style_normal)
        work_sheet.write(index, 1, level[0], style_normal)
        work_sheet.write(index, 2, level[1], style_normal)
        index += 1
        ui.progressBar.setValue(index)
        QApplication.processEvents()

    end_time = datetime.now()
    work_sheet.write(index + 1, 3, "处理耗时：%s" % str(end_time - start_time), style_normal)
    add_log("保存Excel...")
    work_book.close()
    add_log("保存Excel完成")
    ui.progressBar.setValue(file_num + 1)
    add_log("处理完成")
    ui.pushButton.setEnabled(True)


def del_with_file():
    ui.textBrowser.setText("")
    ui.progressBar.setValue(0)
    if str(ui.lineEdit.text()) != '' and str(ui.lineEdit_2.text()) != '' and str(ui.lineEdit_3.text()) != '':
        exec_calc()
    else:
        show_error_dialog()


def new_thread():
    try:
        run = threading.Thread(target=del_with_file())
        run.start()
    except Exception as e:
        ui.pushButton.setEnabled(True)
        add_log("发生严重错误: " + repr(e))
        add_log("异常终止")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = gui.Ui_MainWindow()
    ui.setupUi(MainWindow)

    ui.toolButton.clicked.connect(select_origin_dir)
    ui.toolButton_2.clicked.connect(select_result_dir)
    ui.pushButton.clicked.connect(new_thread)

    MainWindow.show()
    sys.exit(app.exec_())
