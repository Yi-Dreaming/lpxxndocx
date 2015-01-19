#ifndef TABLE_H
#define TABLE_H
#include "docx_global.h"
#include "shared.h"
#include "blkcntnr.h"
#include "./oxml/oxmltable.h"
#include <QString>

namespace Docx {
class Paragraph;
class Column;
class Row;
class Cell;
class Rows;
class Columns;

class DOCX_EXPORT Table : public Parented
{
public:
    Table();
    Table(const CT_Tbl &table);
    void setAlignment(const QString &align);
    Cell* cell(int rowIdx, int colIdx);
    Row* addRow();
    QVector<Cell*> rowCells(int rowIdx);
    Rows* rows();
    Columns* columns();
    virtual ~Table();

private:
    Column* addColumn();


private:
    CT_Tbl m_table;
};

class Cell
{
public:
    Cell();
    Paragraph *addParagraph(const QString &text = "", const QString &style = "");
    Table *addTable(int rows, int cols);
    virtual ~Cell();

private:

};

class Columns
{
public:
    Columns();

    virtual ~Columns();

private:

};

class Column : public Parented
{
public:
    Column();
    Column(const CT_TblGridCol &gridCol, Table* parent);

    virtual ~Column();

private:
    CT_TblGridCol m_gridCol;
    Table *m_table;
};

class Rows
{
public:
    Rows();

    virtual ~Rows();

private:

};

class Row
{
public:
    Row();

    virtual ~Row();

private:


};
}
#endif // TABLE_H
