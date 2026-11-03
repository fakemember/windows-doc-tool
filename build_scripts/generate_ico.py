import struct
from pathlib import Path
from PySide6.QtGui import QIcon, QPixmap, QGuiApplication
from PySide6.QtCore import QSize, QByteArray, QBuffer, QIODevice
import sys

app = QGuiApplication.instance() or QGuiApplication(sys.argv)

svg_path = Path("src/wdt/ui/assets/tflow_logo.svg")
ico_path = Path("src/wdt/ui/assets/tflow_logo.ico")

icon = QIcon(str(svg_path))
sizes = [16, 24, 32, 48, 64, 128, 256]

png_buffers = []
for s in sizes:
    pixmap = icon.pixmap(QSize(s, s))
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QIODevice.OpenModeFlag.WriteOnly)
    pixmap.save(buf, "PNG")
    buf.close()
    png_buffers.append((s, bytes(ba.data())))

num_images = len(png_buffers)
header = struct.pack("<HHH", 0, 1, num_images)
offset = 6 + 16 * num_images
entries = []
image_data_list = []

for s, data in png_buffers:
    w = 0 if s == 256 else s
    h = 0 if s == 256 else s
    size_bytes = len(data)
    entry = struct.pack("<BBBBHHII", w, h, 0, 0, 1, 32, size_bytes, offset)
    entries.append(entry)
    image_data_list.append(data)
    offset += size_bytes

with open(ico_path, "wb") as f:
    f.write(header)
    for entry in entries:
        f.write(entry)
    for data in image_data_list:
        f.write(data)

print(f"Generated ICO: {ico_path} (size={ico_path.stat().st_size} bytes)")
