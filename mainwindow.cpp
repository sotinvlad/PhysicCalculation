#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QMessageBox>
#include <QDebug>
#include <QtGui>
#include <ActiveQt/qaxobject.h>
#include <ActiveQt/qaxbase.h>
#include <algorithm>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}




void MainWindow::on_addResult_clicked()
{
    QAxObject* excel = new QAxObject( "Excel.Application", 0 );
    excel->setProperty("Visible", false);
    QAxObject* workbooks = excel->querySubObject( "Workbooks" );
    QAxObject* book = workbooks->querySubObject( "Open(const QString&)", QDir::currentPath()+"/database.xlsx");
    QAxObject* sheets = book->querySubObject("Sheets");
    QAxObject* sheet = sheets->querySubObject( "Item( int )", 1 );

    QAxObject* rows = sheet ->querySubObject( "Rows" );
    int rowCount = rows ->dynamicCall( "Count()" ).toInt();

    int mainRow = 0;
    for(int row = 1; row < rowCount; row++){
        QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row, 1);
        QVariant content = cell->property("Value").toString();
        if (content == ""){
            mainRow = row;
            QAxObject* cell_fio = sheet->querySubObject( "Cells( int, int )", row, 1);
            cell_fio->dynamicCall( "SetValue(const QVariant&)", ui->fio->text());
            delete cell_fio;

            QAxObject* cell_group = sheet->querySubObject( "Cells( int, int )", row, 2);
            cell_group->dynamicCall( "SetValue(const QVariant&)", ui->group->text());
            delete cell_group;

            QAxObject* cell_podt = sheet->querySubObject( "Cells( int, int )", row, 3);
            cell_podt->dynamicCall( "SetValue(const QVariant&)", ui->podt->text());
            delete cell_podt;

            QAxObject* cell_three = sheet->querySubObject( "Cells( int, int )", row, 4);
            cell_three->dynamicCall( "SetValue(const QVariant&)", ui->three->text());
            delete cell_three;

            QAxObject* cell_oneHundred = sheet->querySubObject( "Cells( int, int )", row, 5);
            cell_oneHundred->dynamicCall( "SetValue(const QVariant&)", ui->oneHundred->text());
            delete cell_oneHundred;

            break;
        }

        delete cell;

    }

    QAxObject* excel_2 = new QAxObject( "Excel.Application", 0 );
    excel_2->setProperty("Visible", false);
    QAxObject* workbooks_2 = excel_2->querySubObject( "Workbooks" );
    QAxObject* book_2 = workbooks_2->querySubObject( "Open(const QString&)", QDir::currentPath()+"/norms.xlsx");
    QAxObject* sheets_2 = book_2->querySubObject("Sheets");
    QAxObject* sheet_2 = sheets_2->querySubObject( "Item( int )", 1 );
    double ans_podt = 0;
    double ans_three = 0;
    double ans_oneHundred = 0;
    for(int row = 2; row <= 101; row++){

        QAxObject* cell_result = sheet_2->querySubObject("Cells(QVariant,QVariant)", row, 1);
        QVariant content_result = cell_result->property("Value").toString();
        double result = content_result.toDouble();

        QAxObject* cell_podt = sheet_2->querySubObject("Cells(QVariant,QVariant)", row, 2);
        if (cell_podt->property("Value").toString() != ""){
            QVariant content_podt = cell_podt->property("Value").toString();
            double podt = content_podt.toDouble();
            if (ui->podt->text().toInt() >= podt){
                if (result > ans_podt){
                    ans_podt = result;
                }

            }
        }

        QAxObject* cell_three = sheet_2->querySubObject("Cells(QVariant,QVariant)", row, 3);
        if (cell_three->property("Value").toString() != ""){
            QVariant content_three = cell_three->property("Value").toString();
            double three = content_three.toDouble();
            if (ui->three->text().toDouble() <= three){
                if (result > ans_three){
                    ans_three = result;
                }

            }
        }

        QAxObject* cell_oneHundred = sheet_2->querySubObject("Cells(QVariant,QVariant)", row, 4);
        if (cell_oneHundred->property("Value").toString() != ""){
            QVariant content_oneHundred = cell_oneHundred->property("Value").toString();
            double oneHundred = content_oneHundred.toDouble();
            if (ui->oneHundred->text().toDouble() <= oneHundred){
                if (result > ans_oneHundred){
                    ans_oneHundred = result;
                }

            }
        }


    }
    QAxObject* cell = sheet->querySubObject( "Cells( int, int )", mainRow, 6);
    cell->dynamicCall( "SetValue(const QVariant&)", "-");
    if (ans_podt + ans_three + ans_oneHundred >= 210){
        QAxObject* cell = sheet->querySubObject( "Cells( int, int )", mainRow, 6);
        cell->dynamicCall( "SetValue(const QVariant&)", "III");
    }
    if (ans_podt + ans_three + ans_oneHundred >= 230){
        QAxObject* cell = sheet->querySubObject( "Cells( int, int )", mainRow, 6);
        cell->dynamicCall( "SetValue(const QVariant&)", "II");
    }
    if (ans_podt + ans_three + ans_oneHundred >= 240){
        QAxObject* cell = sheet->querySubObject( "Cells( int, int )", mainRow, 6);
        cell->dynamicCall( "SetValue(const QVariant&)", "I");
    }
    if (ans_podt + ans_three + ans_oneHundred >= 250){
        QAxObject* cell = sheet->querySubObject( "Cells( int, int )", mainRow, 6);
        cell->dynamicCall( "SetValue(const QVariant&)", "Высший");
    }

//    qDebug() << ans_podt << " " << ans_three << " " << ans_oneHundred;
    book->querySubObject("Save");
    workbooks->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    book_2->querySubObject("Save");
    workbooks_2->dynamicCall("Close()");
    excel_2->dynamicCall("Quit()");
    QMessageBox::about(this, "Добавление результата", "Результат успешно добавлен");
}

void MainWindow::on_otchet_clicked()
{
    QAxObject* excel = new QAxObject( "Excel.Application", 0 );
    excel->setProperty("Visible", false);
    QAxObject* workbooks = excel->querySubObject( "Workbooks" );
    QAxObject* book = workbooks->querySubObject( "Open(const QString&)", QDir::currentPath()+"/database.xlsx");
    QAxObject* sheets = book->querySubObject("Sheets");
    QAxObject* sheet = sheets->querySubObject( "Item( int )", 1 );

    QAxObject* word = new QAxObject("Word.Application");
    QAxObject* doco = word->querySubObject("Documents");
    QAxObject* doc = doco->querySubObject("Add()");
    word->setProperty("Visible", true);

    QAxObject* rows = sheet ->querySubObject( "Rows" );
    int rowCount = rows ->dynamicCall( "Count()" ).toInt();

    int index = 1;
    for(int row = 1; row < rowCount; row++){
        QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row, 6);
        QVariant content = cell->property("Value").toString();
        if (content == "Высший"){
            QAxObject* cell_fio = sheet->querySubObject( "Cells( int, int )", row, 1);
            QString fio = cell_fio->property("Value").toString();

            QAxObject* cell_group = sheet->querySubObject( "Cells( int, int )", row, 2);
            QString group = cell_group->property("Value").toString();

            QAxObject* cell_uroven = sheet->querySubObject( "Cells( int, int )", row, 6);
            QString uroven = cell_uroven->property("Value").toString();
            QString result = QString::number(index) + ". " + fio + ", уч.гр. " + group + " - " + uroven;
            QAxObject* range = doc->querySubObject("Range()");
            range->dynamicCall("SetRange(int,int)", row*100, row*100+100);
            range->setProperty("Text",result);
            range->dynamicCall("InsertParagraphAfter()");
            index += 1;
        }
        if (content == "") break;
    }

}
