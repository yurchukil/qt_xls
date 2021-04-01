#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include<QtWidgets>

class MainWindow : public QMainWindow
{
    Q_OBJECT
protected:

  QLabel *name_lb, *job_lb;
QTextEdit *name_txt, *job_txt;

QPushButton *open_btn;

public slots: void from_xls();

public:
    MainWindow(QWidget *parent = 0);
    ~MainWindow();
};

#endif // MAINWINDOW_H
