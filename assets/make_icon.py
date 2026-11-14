"""Draw the Unbound icon: a grey, faceted envelope seen from the back.

Run:  .venv\\Scripts\\python assets\\make_icon.py
Writes assets/icon.ico (16-256 px), assets/icon.png (256 px) and
assets/social-preview.png (1280 x 640, for the GitHub repository's
Settings -> Social preview).

The envelope is split into four flaps, each a different grey, so it reads as
a solid shape lit from the top left. The top flap is drawn last, over the
others, and points down like the triangle of the earlier icon. Change the
colours or shape here and re-run to update all three images.
"""

import os
import struct

from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QPointF, Qt
from PySide6.QtGui import QColor, QGuiApplication, QImage, QPainter, QPolygonF

HERE = os.path.dirname(os.path.abspath(__file__))
SIZES = [16, 24, 32, 48, 64, 128, 256]

TOP = "#B4B9C1"     # top flap, lightest (faces the light)
LEFT = "#868C96"    # left flap
BOTTOM = "#71777F"  # bottom flap
RIGHT = "#5C626C"   # right flap, darkest
EDGE = "#3E434B"    # thin outline so it holds up on light backgrounds

PREVIEW_SIZE = (1280, 640)
PREVIEW_BACKGROUND = "#161B22"  # GitHub's dark background


def draw(p, pen_width):
    """Draw the envelope in a 256 x 256 grid on painter p."""
    tl, tr = QPointF(16, 52), QPointF(240, 52)     # top corners
    bl, br = QPointF(16, 204), QPointF(240, 204)   # bottom corners
    meet = QPointF(128, 122)  # where the side and bottom flaps meet, under the top flap
    tip = QPointF(128, 150)   # point of the top flap

    p.setPen(Qt.NoPen)
    for pts, colour in (((tl, bl, meet), LEFT), ((tr, br, meet), RIGHT),
                        ((bl, br, meet), BOTTOM), ((tl, tr, tip), TOP)):
        p.setBrush(QColor(colour))
        p.drawPolygon(QPolygonF(list(pts)))

    pen = p.pen()
    pen.setStyle(Qt.SolidLine)
    pen.setColor(QColor(EDGE))
    pen.setWidthF(pen_width)
    pen.setJoinStyle(Qt.MiterJoin)
    p.setPen(pen)
    p.setBrush(Qt.NoBrush)
    # The top flap's lower edges, so the envelope still reads at 16 px
    p.drawPolyline(QPolygonF([tl, tip, tr]))
    p.drawPolygon(QPolygonF([tl, tr, br, bl]))


def render(size):
    img = QImage(size, size, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    p.scale(size / 256, size / 256)
    # Outline at every size keeps small icons crisp on light taskbars
    draw(p, max(1.0, 256 / size * 0.9))
    p.end()
    return img


def render_preview():
    w, h = PREVIEW_SIZE
    img = QImage(w, h, QImage.Format_RGB32)
    img.fill(QColor(PREVIEW_BACKGROUND))
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    scale = 2.8  # envelope about 630 x 425 px, centred
    p.translate(w / 2 - 128 * scale, h / 2 - 128 * scale)
    p.scale(scale, scale)
    draw(p, 0.9)
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
    render_preview().save(os.path.join(HERE, "social-preview.png"))
    print("Wrote icon.ico, icon.png and social-preview.png")
