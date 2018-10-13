#include "excelbase.h"
#include <QVariant>
#include <QString>

excelbase::excelbase()
{

}


///
/// \brief 把列数转换为excel的字母列号
/// \param data 大于0的数
/// \return 字母列号，如0->A 25->Z 26 AA
///
void excelbase::convertToColName(int data, QString &res)
{
    Q_ASSERT(data>-1 && data<65534);
    int tempData = data / 26;
    if(tempData > 0)
    {
        int mode = data % 26;
        convertToColName(mode,res);
        convertToColName(tempData-1,res);
    }
    else
    {
        res=(to26AlphabetString(data)+res);
    }
}


///
/// \brief QList<QList<QVariant> >转换为QVariant
/// \param cells
/// \return
///
void excelbase::castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res)
{
    QVariantList vars;
    const int rows = cells.size();
    for(int i=0;i<rows;++i)
    {
        vars.append(QVariant(cells[i]));
    }
    res = QVariant(vars);
}


///
/// \brief 数字转换为26字母
///
/// 0->A 25->Z
/// \param data
/// \return
///
QString excelbase::to26AlphabetString(int data)
{
    QChar ch = data + 0x41;//A对应0x41
    return QString(ch);
}
