#ifndef IMAGEPART_H
#define IMAGEPART_H

#include "../opc/part.h"
#include "../image/image.h"

namespace Docx {
class ImagePart : public Part
{
public:
    ImagePart(const PackURI &partName, const QString &contentType, const QByteArray &blob = QByteArray(), Image *image = nullptr);
    static ImagePart *load(const PackURI &partName, const QString &contentType, const QByteArray &blob = QByteArray(), Image *image = nullptr);
    void afterUnmarshal();
    Image * image() const;
    QByteArray Hash() const;

    ~ImagePart();   

private:
    Image *m_image;    
    QByteArray m_hash;
};
}

#endif // IMAGEPART_H
