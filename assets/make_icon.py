"""Draw the Unbound icon: a grey, faceted, upside-down triangle.

Run:  .venv\\Scripts\\python assets\\make_icon.py
Writes assets/icon.ico (16-256 px) and assets/icon.png (256 px).

The triangle is split into three facets that meet at its centre, each a
different grey, so it reads as a solid geometric shape lit from the top.
Change the colours or shape here and re-run to update the icon.
"""

import os
import struct

from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QPointF, Qt
from PySide6.QtGui import QColor, QGuiApplication, QImage, QPainter, QPolygonF

HERE = os.path.dirname(os.path.abspath(__file__))
SIZES = [16, 24, 32, 48, 64, 128, 256]

TOP = "#B4B9C1"    # top facet, lightest (faces the light)
LEFT = "#868C96"   # left facet
RIGHT = "#5C626C"  # right facet, darkest
EDGE = "#3E434B"   # thin outline so it holds up on light backgrounds


def render(size):
    img = QImage(size, size, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    p.scale(size / 256, size / 256)

    # Equilateral triangle, point down, centred in a 256 grid
    a = QPointF(20, 42)     # top left
    b = QPointF(236, 42)    # top right
    d = QPointF(128, 229)   # bottom point
    c = QPointF(128, 104)   # centre where the facets meet

    p.setPen(Qt.NoPen)
    for pts, colour in (((a, b, c), TOP), ((a, c, d), LEFT), ((b, d, c), RIGHT)):
        p.setBrush(QColor(colour))
        p.drawPolygon(QPolygonF(list(pts)))

    # Outline at every size keeps small icons crisp on light taskbars
    pen = p.pen()
    pen.setStyle(Qt.SolidLine)
    pen.setColor(QColor(EDGE))
    pen.setWidthF(max(1.0, 256 / size * 0.9))
    pen.setJoinStyle(Qt.MiterJoin)
    p.setPen(pen)
    p.setBrush(Qt.NoBrush)
    p.drawPolygon(QPolygonF([a, b, d]))
    p.end()
    return img


def png_bytes(img):
    data = QByteArray()
    buf = QBuffer(data)
    buf.open(QIODevice.WriteOnly)
    img.save(buf, "PNG")
    return bytes(data)


def write_ico(path, images):
    """ICO file with PNG-compressed entries (supported since Windows Vista)."""
    pngs = [png_bytes(i) for i in images]
    out = struct.pack("<HHH", 0, 1, len(pngs))
    offset = 6 + 16 * len(pngs)
    for img, data in zip(images, pngs):
        w = img.width() if img.width() < 256 else 0  # 0 means 256
        out += struct.pack("<BBBBHHII", w, w, 0, 0, 1, 32, len(data), offset)
        offset += len(data)
    with open(path, "wb") as fh:
        fh.write(out + b"".join(pngs))


if __name__ == "__main__":
    app = QGuiApplication([])
    images = [render(s) for s in SIZES]
    write_ico(os.path.join(HERE, "icon.ico"), images)
    images[-1].save(os.path.join(HERE, "icon.png"))
    print("Wrote icon.ico and icon.png")
