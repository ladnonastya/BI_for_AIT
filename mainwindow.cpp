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
#include <QPixmap>


MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    connect(ui->actionOpen,SIGNAL(triggered()),this,SLOT(openFile()));
    connect(ui->actionSave_as, SIGNAL(triggered()),this,SLOT(saveAs()));
    connect(ui->pushButton,SIGNAL(clicked(bool)),this,SLOT(drawPieChart()));
    connect(ui->pushButton_2,SIGNAL(clicked(bool)),this,SLOT(drawCostLine()));
    //ui->widget->hide();
    QPixmap myPixmap("D:/BI_for_AIT/piechart.png");
    ui->label->setPixmap(myPixmap);
    ui->label->hide();
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

void MainWindow::drawCostLine()
{

    QVector<double> x(10), y(10);

    for(int i=0;i<10;i++)
        x[i]=i+1;

    y[0]=73.9;
    y[1]=77.56;
    y[2]=81.42;
    y[3]=87;
    y[4]=89.9;
    y[5]=97;
    y[6]=121.56;
    y[7]=135;
    y[8]=166.25;
    y[9]=307.1;


    ui->widget_2->clearGraphs();
    ui->widget_2->addGraph();
    ui->widget_2->graph(0)->setData(x, y);
    ui->widget_2->xAxis->setLabel("Country");
    ui->widget_2->yAxis->setLabel("Average cost");

    ui->widget_2->xAxis->setRange(0, 12);

//    QVector<double> ticks;
//    QVector<QString> labels;
//    ticks << 1 << 2 << 3 << 4 << 5 << 6 << 7 << 8 << 9 << 10;
//    labels << "ЮАР" << "Австрия" << "Франция" << "Испания" << "Чили" << "Италия" << "Австралия" << "Португалия" << "Германия" << "США";
//    ui->widget_2->xAxis->setAutoTicks(false);
//    ui->widget_2->xAxis->setAutoTickLabels(false);
//    ui->widget_2->xAxis->setTickVector(ticks);
//    ui->widget_2->xAxis->setTickVectorLabels(labels);
//    ui->widget_2->xAxis->setSubTickCount(0);
//    ui->widget_2->xAxis->setTickLength(0, 4);
//    ui->widget_2->xAxis->grid()->setVisible(true);
//    ui->widget_2->xAxis->setRange(0, 8);

    ui->widget_2->yAxis->setRange(0,315);

    //ui->widget_2->graph(0)->setData(ticks, y);

    ui->widget_2->graph(0)->setLineStyle(QCPGraph::lsNone);
    ui->widget_2->graph(0)->setScatterStyle(QCPScatterStyle(QCPScatterStyle::ssCircle, 4));

    ui->widget_2->replot();

    //ui->widget_2->plotLayout()->addElement(0,1,ui->widget_2->legend);
    //ui->widget_2->xAxis->
}

void MainWindow::drawPieChart()
{
    //ui->widget->show();
//    QPixmap image("https://ibb.co/cWcuob");
//    ui->label->setPixmap(image);
    //ui->label->show();
    //ui->label->setPixmap(myPixmap);
    ui->label->show();
}

