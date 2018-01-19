#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include <QAction>
#include <QDialog>
#include <QFileDialog>
#include <QMessageBox>
#include <QWidget>
#include<cstdlib>


MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    connect(ui->actionOpen,SIGNAL(triggered()),this,SLOT(openFile()));
    connect(ui->actionSave_as, SIGNAL(triggered()),this,SLOT(saveAs()));
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::openFile()
{
    fileName=QFileDialog::getOpenFileName(0, "Open", "", "*.xlsx");
    QXlsx::Document doc(fileName);
    QXlsx::CellRange cells;
    cells=doc.dimension();
    ui->tableWidget->setRowCount(cells.rowCount());
    ui->tableWidget->setColumnCount(cells.columnCount()+1);
    for(int i=cells.firstRow();i<cells.rowCount();i++)
        ui->tableWidget->setRowHeight(i,20);
    for(int i=cells.firstColumn();i<cells.columnCount();i++)
        ui->tableWidget->setColumnWidth(i,60);
    for(int i=cells.firstRow();i<cells.lastRow()+1;i++)
        for(int j=cells.firstColumn();j<cells.lastColumn()+1;j++)
        {
            ui->tableWidget->setItem(i-1,j-1,new QTableWidgetItem);
            ui->tableWidget->setHorizontalHeaderItem(j-1, new QTableWidgetItem(doc.read(1,j).toString()));
            ui->tableWidget->item(i-1,j-1)->setText(doc.read(i+1,j).toString());
        }
    ui->tableWidget->removeRow(cells.lastRow()-1);
    ui->tableWidget->removeColumn(cells.lastColumn());
    ui->tableWidget->removeColumn(cells.lastColumn());
    /*for(int i=cells.firstColumn();i<cells.lastColumn()+1;i++)
        variables.append(doc.read(1,i).toString());
    emit sendVariables(variables);*/
    ui->tableWidget->resizeColumnsToContents();
    ui->tableWidget->resizeRowsToContents();
}

/*void MainWindow::save()
{
    //QString savedFileName=QFileDialog::getSaveFileName(0, "Save As", "", "*.xlsx");
    QXlsx::Document doc;
    QXlsx::CellRange cells;
    cells=doc.dimension();
    for(int i=0;i<ui->tableWidget->rowCount();i++)
        for(int j=0;j<ui->tableWidget->columnCount();j++)
        {
            doc.write(i+1,j+1,ui->tableWidget->item(i,j)->text());
        }
    doc.saveAs(fileName);
}*/

void MainWindow::saveAs()
{
    QString savedFileName=QFileDialog::getSaveFileName(0, "Save As", "", "*.xlsx");
    QXlsx::Document doc;
    QXlsx::CellRange cells;
    cells=doc.dimension();
    for(int i=0;i<ui->tableWidget->rowCount();i++)
        for(int j=0;j<ui->tableWidget->columnCount();j++)
        {
            doc.write(i+1,j+1,ui->tableWidget->item(i,j)->text());
        }
    doc.saveAs(savedFileName);
}
