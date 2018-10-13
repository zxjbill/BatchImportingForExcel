#include "widget.h"
#include "ui_widget.h"
#include <QApplication>
#include <QAxObject>
#include <QAxWidget>
#include <QFileDialog>
#include <QObject>
#include <QDebug>
#include "excelbase.h"
#include <QString>

using namespace std;


Widget::Widget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Widget)
{
    ui->setupUi(this);
    ui->label->setVisible(false);
    excel.setControl("Excel.Application");
    excel.setProperty("Visible", false); // 在实际使用中可设置为false，不让用户看到底层运行
}

Widget::~Widget()
{
    delete ui;
    excel.dynamicCall("Quit(void)");
}

QStringList Widget::getFileNames(const QString &path)
{
    QDir dir(path);
    QStringList nameFilters;
    nameFilters << "*.txt";
    QStringList files = dir.entryList(nameFilters, QDir::Files|QDir::Readable, QDir::Name);
    return files;
}

void Widget::on_pushButton_clicked()
{
    ui->label->setVisible(false);
    QString path = QDir::currentPath();
    QString file_name = QFileDialog::getExistingDirectory(NULL, QStringLiteral("选择文件夹"), path);
    ui->textEdit->setPlainText(file_name);
    ui->textEdit_2->setPlainText(file_name.split("/").last());
}

void Widget::on_pushButton_2_clicked()
{
    QString path = ui->textEdit->toPlainText();
    path.replace("/", "\\");
    QString after_name = ui->textEdit_2->toPlainText();
    int start_raw = ui->horizontalSlider->value();
    bool is_suc = ImportAllTxt(path, after_name, start_raw);

    ui->label->setVisible(true);

    if(is_suc)
    {
        ui->label->setText("success");
    }
    else
    {
        ui->label->setText("fault");
    }
}

void Widget::on_horizontalSlider_valueChanged(int value)
{
    int data = ui->horizontalSlider->value();
    ui->label_4->setText(QString::number(data));
}

void Widget::on_textEdit_textChanged()
{
    ui->label->setVisible(false);
}

bool Widget::ImportAllTxt(QString &path, QString &after_name, int start_raw)
{
    QAxObject *work_books = excel.querySubObject("WorkBooks");

    work_books->dynamicCall("Add()");
    work_book = excel.querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");
    bool is_new_book = true;

    QStringList txtList = this->getFileNames(path);
    qDebug()<<txtList;

    for(int i = 0; i<txtList.size(); i++)
    {
        QString name = txtList[i].replace(".txt", "");

        if(name == QString("说明"))
        {
            continue;
        }

        QFile file(path + "\\" + txtList[i] + ".txt");

        if(!file.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            qDebug()<<"Can't open the file!"<<endl;
            continue;
        }

        if(!is_new_book)
        {
            work_sheets->dynamicCall("Add()");
        }
        else
        {
            is_new_book = false;
        }

        QAxObject *work_sheet = work_book->querySubObject("Sheets(int)", 1);
        work_sheet->setProperty("Name", name);

        QTextStream in(&file);
        QList<QList<QVariant>> cells;
        QString line = in.readLine();
        int raw = 1;

        while(!line.isNull())
        {
            if(raw < start_raw)
            {
                line = line.trimmed();
                QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", raw, 1);
                cell->setProperty("NumberFormatLocal","@");
                cell->setProperty("Value", line);
            }
            else
            {
                line = line.trimmed();
                line.replace(QRegExp("\\s+"), " ");
                QStringList str_list = line.split(" ");
                QList<QVariant> listVariant;

                for(int j = 0; j<str_list.size();j++)
                {
                    listVariant.append(QVariant(str_list[j]));
                }

                cells.append(listVariant);
            }

            line = in.readLine();
            raw++;
        }

        if(cells.size() <= 0)
            continue;
        if(NULL == work_sheet || work_sheet->isNull())
            continue;
        int row = cells.size() + 6;
        int col = cells.at(0).size();
        QString rangStr;
        excelbase::convertToColName(col-1,rangStr);
        rangStr += QString::number(row);
        rangStr = "A"+ QString::number(start_raw)+ ":" + rangStr;
        qDebug()<<rangStr;
        QAxObject *range = work_sheet->querySubObject("Range(const QString&)",rangStr);
        if(NULL == range || range->isNull())
        {
            continue;
        }
        bool succ = false;
        QVariant var;
        excelbase::castListListVariant2Variant(cells,var);
        succ = range->setProperty("Value", var);
        delete range;
        qDebug()<<name<<":"<<succ<<endl;
    }

    QString afer_path = path + "\\" + after_name + ".xlsx";
    work_book->dynamicCall("SaveAs(const QString&)", afer_path);
    work_book->dynamicCall("Close(Boolean)", false);

    return true;
}
