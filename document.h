#ifndef DOCUMENT_H
#define DOCUMENT_H

#include "docx_global.h"

class QIODevice;

namespace Docx {

class DOCX_EXPORT Document
{
public:
    Document();
    explicit Document(const QString& name);
    explicit Document(QIODevice* device);
    ~Document();

    void save(const QString& path);

};

}


#endif // DOCUMENT_H
