#-------------------------------------------------
#
# Project created by QtCreator 2018-01-19T21:32:38
#
#-------------------------------------------------

QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets
                                  QT += printsupport

TARGET = BI_for_AIT
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp \
    qcustomplot.cpp \
    piechartwidget.cpp

HEADERS  += mainwindow.h \
    qcustomplot.h \
    piechartwidget.h

FORMS    += mainwindow.ui

include(3rdparty/qtxlsx/src/xlsx/qtxlsx.pri)
#include(3rdparty/qsint-0.2.2-src/qsint.pri)
