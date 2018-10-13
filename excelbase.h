#ifndef EXCELBASE_H
#define EXCELBASE_H
#include <QVariant>
#include <QString>


class excelbase
{
public:
    excelbase();
    static void convertToColName(int data, QString &res);
    static void castListListVariant2Variant(const QList<QList<QVariant> > &cells,QVariant& res);
    static QString to26AlphabetString(int data);
};

#endif // EXCELBASE_H
