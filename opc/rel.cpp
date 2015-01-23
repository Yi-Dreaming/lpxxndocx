#include "rel.h"

namespace Docx {


Relationship::Relationship(const QString &rId, const QString &reltype, Part *target, const QString &baseURI, bool external)
    : m_rId(rId), m_reltype(reltype), m_baseURI(baseURI), m_isexternal(external), m_target(target)
{

}

Relationship::~Relationship()
{
    delete m_target;
}


Relationships::Relationships(const QString &baseURI)
    : m_baseURI(baseURI)
{

}

Relationship *Relationships::addRelationship(const QString &reltype, Part *target, const QString &rId, bool external)
{
    Relationship *rel = new Relationship(rId, reltype, target, m_baseURI, external);
    m_rels[rId] = rel;
    if (!external)
        m_targetPartsByrId[rId] = target;
    return rel;
}

Part *Relationships::partWithReltype(const QString &reltype)
{
    for (Relationship *s : m_rels.values()) {
        if (s->relType() == reltype)
            return s->target();
    }
    return nullptr;
}

Relationships::~Relationships()
{
    m_targetPartsByrId.clear();
    m_rels.clear();
}


}

