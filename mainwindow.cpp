#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include <QAction>
#include <QDialog>
#include <QFileDialog>
#include <QMessageBox>
#include <QWidget>
#include <QGraphicsScene>
#include <QGraphicsView>
#include <cstdlib>
#include <QPainter>


MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    connect(ui->actionOpen,SIGNAL(triggered()),this,SLOT(openFile()));
    connect(ui->actionSave_as, SIGNAL(triggered()),this,SLOT(saveAs()));
    //test_draw();
    //connect(ui->pushButton,SIGNAL(triggered()),this,SLOT(drawPieChart()));

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

//void MainWindow::test_draw()
//{
//    double a = -1; //Начало интервала, где рисуем график по оси Ox
//    double b =  1; //Конец интервала, где рисуем график по оси Ox
//    double h = 0.01; //Шаг, с которым будем пробегать по оси Ox

//    int N=(b-a)/h + 2; //Вычисляем количество точек, которые будем отрисовывать
//    QVector<double> x(N), y(N); //Массивы координат точек

//    //Вычисляем наши данные
//    int i=0;
//    for (double X=a; X<=b; X+=h)//Пробегаем по всем точкам
//    {
//        x[i] = X;
//        y[i] = X*X;//Формула нашей функции
//        i++;
//    }

//    ui->widget->clearGraphs();//Если нужно, но очищаем все графики
//    //Добавляем один график в widget
//    ui->widget->addGraph();
//    //Говорим, что отрисовать нужно график по нашим двум массивам x и y
//    ui->widget->graph(0)->setData(x, y);

//    //Подписываем оси Ox и Oy
//    ui->widget->xAxis->setLabel("x");
//    ui->widget->yAxis->setLabel("y");

//    //Установим область, которая будет показываться на графике
//    ui->widget->xAxis->setRange(a, b);//Для оси Ox

//    //Для показа границ по оси Oy сложнее, так как надо по правильному
//    //вычислить минимальное и максимальное значение в векторах
//    double minY = y[0], maxY = y[0];
//    for (int i=1; i<N; i++)
//    {
//        if (y[i]<minY) minY = y[i];
//        if (y[i]>maxY) maxY = y[i];
//    }
//    ui->widget->yAxis->setRange(minY, maxY);//Для оси Oy

//    //И перерисуем график на нашем widget
//    ui->widget->replot();
//}

//void MainWindow::drawPieChart()
//{

//}

