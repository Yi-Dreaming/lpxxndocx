//#include "MainWindow.h"
#include <QApplication>
#include <QFile>
#include <QTextCodec>
#include <QDebug>
#include "document.h"
#include "table.h"

int main(int argc, char *argv[])
{
    char ver[10] = "0.3";
    QApplication a(argc, argv);
    QTextCodec::setCodecForLocale(QTextCodec::codecForName("UTF8")); // 用于QT 5.6 版本
//    QTextCodec::setCodecForLocale(QTextCodec::codecForName("GBK")); // 用于QT 5.12 版本
    Docx::Document doc = Docx::Document("D:/build-docx-Desktop_Qt_5_15_2_MinGW_64_bit-Debug/debug/test.docx");
//    doc.addHeading("title1--",1);

//    /**
//     * @brief table 协议总表格式
//     */
//    {
//        int rowCount = 34;
//        doc.addParagraph("表格一，协议总表");
//        auto table = doc.addTable(rowCount,9,"auto");

//        table->setAlignment(Docx::WD_TABLE_ALIGNMENT::CENTER); // 设置表格在文档中居中
////        table->cell(0,0)->addText(QString::fromLocal8Bit("型号"));
//        auto cell1 = table->merge(0,0,1,1);
//        qDebug()<<"cell1:"<<cell1<<"cell1.m_currentpara"<<cell1->m_currentpara;
//        cell1->addText(QString::fromLocal8Bit("型号"),"left");
//        table->cell(0,2)->addText(QString::fromLocal8Bit("版本"),"center","top");

//// todo 实现单元格的对齐方式

//        table->merge(0,3,0,6);
//        table->merge(1,3,1,6);

//        for(int i = 2; i < 7; i++)
//        {
//            table->merge(i, 0, i, 1);
//            table->merge(i, 2, i, 4);
//            table->merge(i, 6, i, 8);
//        }

//        table->merge(7, 0, 7, 8);

//        for(int i = 8; i < 25; i++)
//        {
//            table->merge(i, 1, i, 3);
//            table->merge(i, 4, i, 7);
//        }

//        table->merge(25, 0, rowCount-1, 8);
//    }
//    /**
//     * @brief table 协议每个字的表
//     */
//    {
//        int rowCount = 35;
//        doc.addParagraph("表格二，每个字的表格");
//        auto table = doc.addTable(rowCount,9,"auto");

//        table->setAlignment(Docx::WD_TABLE_ALIGNMENT::CENTER); // 设置表格在文档中居中
//        table->merge(0,0,1,1);

//        table->merge(0,3,0,6);
//        table->merge(1,3,1,6);

//        table->merge(2,0,2,1);
//        table->merge(2,2,2,8);

//        table->merge(8,0,8,1);
//        table->merge(8,2,8,8);


//        for(int i = 3; i < 8; i++)
//        {
//            table->merge(i, 0, i, 1);
//            table->merge(i, 2, i, 4);
//            table->merge(i, 6, i, 8);
//        }

//        table->merge(9, 0, 9, 8);

//        for(int i = 10; i < 27; i++)
//        {
//            table->merge(i, 1, i, 2);
//            table->merge(i, 3, i, 8);
//        }

//        table->merge(27, 0, rowCount-1, 8);
//    }

    // 测试添加图像
    doc.addPicture("C:\\Windows\\Web\\Screen\\img100.jpg");
    auto pic = doc.addPicture("F:\\My Knowledge\\My Personal Knowledge\\d3630f7d_11380317_0.jpg");

    doc.save("D:/build-docx-Desktop_Qt_5_15_2_MinGW_64_bit-Debug/debug/test1.docx");

    return a.exec();
}
