#include "text.h"
#include "./oxml/oxmltext.h"

#include <QDomDocument>
#include <QDebug>

namespace Docx {


Paragraph::Paragraph(QDomDocument *domDocument, const QDomElement &element)
{
    m_dom = domDocument;
    m_pEle = new QDomElement(element);
    m_style = new CT_PPr(this);
}

Run *Paragraph::addRun(const QString &text, const QString &style)    
{
    Run *run = new Run(m_dom, m_pEle);
    if (!text.isEmpty())
        run->addText(text);
    if (!style.isEmpty())
        setStyle(style);

    m_runs.append(run);
    return run;
}

void Paragraph::setStyle(const QString &style)
{
    m_style->setStyle(style);
}

void Paragraph::setAlignment(const WD_PARAGRAPH_ALIGNMENT &align)
{
    m_style->setAlignment(align);
}

Paragraph *Paragraph::insertParagraphBefore(const QString &text, const QString &style)
{
    QDomElement pEle = m_dom->createElement(QStringLiteral("w:p"));
    Paragraph *p = new Paragraph(m_dom, pEle);
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


Run::Run(QDomDocument *domDocument, QDomElement *parent)
    : m_dom(domDocument), m_parent(parent)
{
    m_rEle = m_dom->createElement(QStringLiteral("w:r"));
    m_parent->appendChild(m_rEle);
    m_style = new CT_RPr(this);
}

void Run::addTab()
{
    QDomElement tabEle = m_dom->createElement(QStringLiteral("w:tab"));
    m_rEle.appendChild(tabEle);
}

void Run::addText(const QString &text)
{
    QDomElement tEle = m_dom->createElement(QStringLiteral("w:t"));
    tEle.appendChild(m_dom->createTextNode(text));
    m_rEle.appendChild(tEle);
    //m_rEle
}

void Run::setStyle(const QString &style)
{
    m_style->setStyle(style);
}

Run::~Run()
{
    delete m_style;
}

void Run::setAllcaps(bool isallcaps)
{
    m_style->setAllcaps(isallcaps);
}

void Run::setBold(bool isbold)
{
    m_style->setBold(isbold);
}

void Run::setItalic(bool isItalic)
{
    m_style->setItalic(isItalic);
}

void Run::setDoubleStrike(bool isDoubleStrike)
{
    m_style->setDoubleStrike(isDoubleStrike);
}

void Run::setUnderLine(const WD_UNDERLINE &underline)
{
    m_style->setUnderLine(underline);
}



Text::Text()
{

}

Text::Text(QDomDocument *domDocument, QDomElement *parent)
    : m_dom(domDocument), m_parent(parent)
{

}

Text::~Text()
{

}



}


