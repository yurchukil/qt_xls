#include "mainwindow.h"
#include <QAxObject>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    name_lb = new QLabel (this);
    name_lb->setText("Name and Soname");
    name_lb->move(20,10);
    name_txt = new QTextEdit (this);
    name_txt->setGeometry(20,50,150,25);
    //---
    job_lb=new QLabel ("What did",this);
    job_lb->move(20,90);
    job_txt=new QTextEdit (this);
    job_txt->setGeometry(20,130,300,300);
    //---
    open_btn=new QPushButton ("Open daa",this);
    open_btn->move(350,10);


    //---

connect(open_btn,SIGNAL(clicked()),this,SLOT(from_xls()));

}

MainWindow::~MainWindow()
{

}

void MainWindow::from_xls()
{
    // получаем указатель на Excel
        QAxObject *mExcel = new QAxObject( "Excel.Application",this);
        // на книги
        QAxObject *workbooks = mExcel->querySubObject( "Workbooks" );
        // на директорию, откуда грузить книгу
        QAxObject *workbook = workbooks->querySubObject( "Open(const QString&)", "D:\\qt_xls.xls" );
        // на листы (снизу вкладки)
        QAxObject *mSheets = workbook->querySubObject("Worksheets(const QVariant&)", 1);
        // указываем, какой лист выбрать. У меня он называется topic.
        //QAxObject *StatSheet = mSheets->querySubObject( "Item(const QVariant&)", QVariant("first") );

        // получение указателя на ячейку [row][col] ((!)нумерация с единицы)
                       QAxObject* cell = mSheets->querySubObject("Cells(int,int)", 6, 4);
                       // получение содержимого
                       QVariant result = cell->dynamicCall("Value()");
        name_txt->setText(result.toString());

        // получение указателя на ячейку [row][col] ((!)нумерация с единицы)
                       QAxObject* cell1 = mSheets->querySubObject("Cells(int,int)", 7, 7);
                       // получение содержимого
                       QVariant result1 = cell1->dynamicCall("Value()");
        job_txt->setPlainText(result1.toString());
}
