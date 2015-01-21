#include "contenttypemap.h"

#include <QDomDocument>

namespace Docx {

ContentTypeMap::ContentTypeMap()
{

}

ContentTypeMap *ContentTypeMap::fromXml(const QByteArray &contentTypesXml)
{
    QDomDocument doc;
    doc.setContent(contentTypesXml);

    QDomNodeList defaults =  doc.elementsByTagName("Default");
    QDomNodeList overrides =  doc.elementsByTagName("Override");

    ContentTypeMap *types = new ContentTypeMap();
    for (int i = 0, sz = defaults.count(); i < sz; i++) {
        QDomNode n = defaults.at(i);
        QString ext = n.toElement().attribute("Extension");
        QString cty = n.toElement().attribute("ContentType");
        types->addDefault(ext, cty);
    }

    for (int i = 0; i < overrides.count(); i++) {
        QDomNode n = overrides.at(i);
        QString ext = n.toElement().attribute("PartName");
        QString cty = n.toElement().attribute("ContentType");
        types->addOverride(ext, cty);
    }

    return types;
}

void ContentTypeMap::addDefault(const QString &extension, const QString &contentType)
{
    m_defaults.insert(extension, contentType);
}

void ContentTypeMap::addOverride(const QString &partName, const QString &contentType)
{
    m_overrides.insert(partName, contentType);
}

QString ContentTypeMap::contentType(const PackURI &partname) const
{
    if (m_overrides.contains(partname.fullURI()))
        return m_overrides.value(partname.fullURI());
    if (m_defaults.contains(partname.ext()))
        return m_defaults.value(partname.ext());
    return QStringLiteral("");
}

ContentTypeMap::~ContentTypeMap()
{

}
}
