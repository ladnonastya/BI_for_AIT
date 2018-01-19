#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QDialog>
#include "xlsxdocument.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    QString fileName;
    QStringList variables;
    ~MainWindow();

signals:
    sendVariables(QStringList var);

private slots:
    void openFile();
    void saveAs();

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
