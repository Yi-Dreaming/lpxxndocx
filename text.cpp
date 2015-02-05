#include "text.h"
#include "./oxml/oxmltext.h"
#include "./oxml/oxmlshape.h"
#include "./parts/documentpart.h"
#include "./shape.h"

#include <QDomDocument>
#include <QDebug>

namespace Docx {


Paragraph::Paragraph(DocumentPart *part, const QDomElement &element)
    : m_part(part)
{
    m_dom = part->m_dom;
    m_pEle = new QDomElement(element);
    m_style = new CT_PPr(this);
}

Run *Paragraph::addRun(const QString &text, const QString &style)    
{
    Run *run = new Run(m_part, m_pEle);
    if (!text.isEmpty())
        run->addText(text);
    if (!style.isEmpty())
        setStyle(style);

    m_runs.append(run);
    return run;
}

void Paragraph::addText(const QString &text)
{
    addRun(text);
}

QString Paragraph::text() const
{
    QString str;
    for (const Run *r : m_runs) {
        str += r->text();
    }
    return str;
}

void Paragraph::setStyle(const QString &style)
{
    m_style->setStyle(style);
}

void Paragraph::setAlignment(WD_PARAGRAPH_ALIGNMENT align)
{
    m_style->setAlignment(align);
}

Paragraph *Paragraph::insertParagraphBefore(const QString &text, const QString &style)
{
    QDomElement pEle = m_dom->createElement(QStringLiteral("w:p"));
    Paragraph *p = new Paragraph(m_part, pEle);
    p->addRun(text, style);
    QDomNode parent = m_pEle->parentNode();

    parent.insertBefore(pEle, *m_pEle);
    return p;
}

Paragraph::~Paragraph()
{
    qDeleteAll(m_runs);
    delete m_pEle;
}

// End Paragraph


Run::Run(DocumentPart *part, QDomElement *parent)
    : m_part(part), m_parent(parent)
{
    m_dom = part->m_dom;
    m_rEle = m_dom->createElement(QStringLiteral("w:r"));
    m_parent->appendChild(m_rEle);
    m_style = new CT_RPr(this);
}

/*!
 * \brief 添加 Tab
 */
void Run::addTab()
{
    QDomElement tabEle = m_dom->createElement(QStringLiteral("w:tab"));
    m_rEle.appendChild(tabEle);
}

/*!
 * \brief 添加文字
 * \param text
 */
void Run::addText(const QString &text)
{
    QDomElement tEle = m_dom->createElement(QStringLiteral("w:t"));
    tEle.appendChild(m_dom->createTextNode(text));
    m_rEle.appendChild(tEle);
    //m_rEle
}

QString Run::text() const
{
    return m_rEle.text();
}

/*!
 * \brief 返回 InlineShape
 * \param path
 * \param width
 * \param height
 * \return
 */
InlineShape *Run::addPicture(const QString &path, const Length &width, const Length &height)
{
    InlineShapes *ships = m_part->m_inlineshapes;
    InlineShape *picture = ships->addPicture(path, this);
    return scalePicture(picture, width, height);

}

InlineShape *Run::scalePicture(InlineShape *picture, const Length &width, const Length &height)
{
    if (!width.isEmpty() || !height.isEmpty()) {
        int lwidth = width.emu();
        int lheight = height.emu();

        int native_width = picture->width();
        int native_height = picture->height();
        if (width.isEmpty()) {
            float scaling_factor = float(lheight) / float(native_height);
            lwidth = int(round(native_width * scaling_factor));
        }
        if (height.isEmpty()) {
            float scaling_factor = float(lwidth) / float(native_width);
            lheight = int(round(native_height * scaling_factor));
        }
        picture->setWidth(lwidth);
        picture->setHeight(lheight);
    }
    qDebug() << picture->width() << " height " << picture->height();
    return picture;
}

InlineShape *Run::addPicture(const QImage &img, const Length &width, const Length &height)
{
    InlineShapes *ships = m_part->m_inlineshapes;
    InlineShape *picture = ships->addPicture(img, this);
    return scalePicture(picture, width, height);
}

/*!
 * \brief 设置样式
 * \param style
 */
void Run::setStyle(const QString &style)
{
    m_style->setStyle(style);
}

Run::~Run()
{
    delete m_style;
}

/*!
 * \brief 全部大写
 * \param isallcaps
 */
void Run::setAllcaps(bool isallcaps)
{
    m_style->setAllcaps(isallcaps);
}

/*!
 * \brief 小一号的大写
 * \param smallcpas
 */
void Run::setSmallcaps(bool smallcpas)
{
    m_style->setSmallcaps(smallcpas);
}

/*!
 * \brief 加粗
 * \param isbold
 */
void Run::setBold(bool isbold)
{
    m_style->setBold(isbold);
}

/*!
 * \brief 倾斜
 * \param isItalic
 */
void Run::setItalic(bool isItalic)
{
    m_style->setItalic(isItalic);
}

/*!
 * \brief 双划线
 * \param isDoubleStrike
 */
void Run::setDoubleStrike(bool isDoubleStrike)
{
    m_style->setDoubleStrike(isDoubleStrike);
}

/*!
 * \brief 阴影
 * \param shadow
 */
void Run::setShadow(bool shadow)
{
    m_style->setShadow(shadow);
}

/*!
 * \brief 下划线
 * \param underline
 */
void Run::setUnderLine(WD_UNDERLINE underline)
{
    m_style->setUnderLine(underline);
}

void Run::addDrawing(CT_Inline *imline)
{
    QDomElement drawing = m_dom->createElement(QStringLiteral("w:drawing"));

    m_rEle.appendChild(drawing);
    drawing.appendChild(imline->element());
}



}


