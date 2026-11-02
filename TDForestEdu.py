#!/usr/bin/env python3

# -*- coding: utf-8 -*-

"""

================================================================================

 TDForest.py - PHẦN MỀM VẼ TRẮC ĐỒ NGANG - DỌC RỪNG (2D/3D)

================================================================================

Mô phỏng trắc đồ (mặt cắt) cây rừng trong ô tiêu chuẩn, thay thế cách vẽ tay

truyền thống bằng bản đồ trực quan 2D và mô hình 3D có thể xoay.



CÀI ĐẶT THƯ VIỆN (chạy trong terminal / PyCharm Terminal):

    pip install PySide6 pandas numpy matplotlib openpyxl



CHẠY CHƯƠNG TRÌNH:

    python TDForest.py



ĐỊNH DẠNG DỮ LIỆU ĐẦU VÀO (file Excel .xlsx/.xls), mỗi dòng là 1 cây:

    TT | Tên loài cây | D1.3 (cm) | Hvn (m) | Hdc (m) |

    Dtán Đông (m) | Dtán Tây (m) | Dtán Nam (m) | Dtán Bắc (m) | X (m) | Y (m)



Tên cột trong file Excel của bạn không cần trùng khớp tuyệt đối - dùng chức

năng "Ánh xạ cột" để chọn cột tương ứng.

================================================================================

"""



import sys

import os

import re

import uuid

import base64

import unicodedata
import random
import math
import json
import copy
from datetime import datetime



import numpy as np

import pandas as pd



from PySide6.QtCore import (

    Qt, QTimer, QModelIndex, QAbstractTableModel, QPointF, QRectF,

    QThread, Signal, QSettings, QUrl, QSize,

)

from PySide6.QtGui import (

    QColor, QIcon, QAction, QPixmap, QFont, QPainter, QPainterPath, QBrush, QPen,

    QDesktopServices, QCursor, QLinearGradient,

)

from PySide6.QtWidgets import (

    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,

    QLabel, QPushButton, QComboBox, QFileDialog, QTableView, QTabWidget,

    QGroupBox, QSlider, QCheckBox, QColorDialog, QMessageBox, QDialog,

    QDialogButtonBox, QListWidget, QListWidgetItem, QSplitter, QDoubleSpinBox,

    QLineEdit, QToolBar, QStatusBar, QFrame, QHeaderView, QSizePolicy,

    QScrollArea, QSplashScreen

)





# ==============================================================================

# 0-bis. LOGO & BỘ ICON THÔNG BÁO

# Logo phần mềm dùng ảnh "Forest" của Icons8 (icons8-forest-65.png), nhúng

# thẳng dạng base64 ngay trong file .py để không phải kèm theo file ảnh rời

# khi đóng gói phần mềm. Sinh bằng: base64.b64encode(open("icons8-forest-65.png","rb").read()).

# LƯU Ý BẢN QUYỀN: icon thuộc bộ icon có bản quyền của Icons8 (icons8.com) -

# dùng theo gói miễn phí BẮT BUỘC ghi rõ nguồn "Icon by Icons8" ở nơi hiển thị

# hợp lý trong phần mềm (xem dòng ghi nguồn trong AboutDialog bên dưới). Nếu

# đã mua gói bản quyền trả phí của Icons8 thì có thể bỏ dòng ghi nguồn đó.

# ==============================================================================



_APP_LOGO_B64 = (
"""iVBORw0KGgoAAAANSUhEUgAAAEEAAABBCAYAAACO98lFAAAACXBIWXMAAAsTAAALEwEAmpwYAAAL30lEQVR4nO1baXBUVRbusijLsubH1PzQcujXr5eEQNBC
oUSdGcdCy30cUaliRC3LkdEZGUcdZwYVyds6CSEbSwgkyqKIaJRdFHAJQQ0C/e593QlJSGASAgmEJSzZk+4+U+eGjv1679CdwTJf1SnSr999fe/3zj3nu+de
DIZRjGIUo/gJ4brMI9dzEnnU8HMGL5EDFpl4U+2VY+Ntmy5U/8IglF9jSBrKyq7mJMeSGzJd45P1E7a8/ZxVoWBRKPAi2er/HZ/heH2scGBGUCOAqzhJ/cCq
0H6TRADNZqfdvKjmJryDxgz6HP4AL5LPDEkCPvuWXBcsqjgBKZnEw94soqzsapOkejlJbdM1ALjKLNOjSNpT61ohY0cnvP1FBzy++hjwEgGzTCsT28OysquR
XW5ele1yHmMUHC9xEn0/8PqUEse1qXbqWfF9G3T0umFclgYm2bHc9z0v09d4RXvAvw16gFmmIH/ZBe/sA539e9t55hUmmc41XGngJLXNJKpegwBj/K+bZFIy
LpPChR43IKQdxyElk/bh2w73LJwC6AGBBPjsodJGsMj0dFwd/PV89R6TRJ0mkdTxgmOWIQnAmHKDQO8LvJ6aSXvtu46DD60X+tGdw75JnCr4pnEKhCPhH5vO
gtVOPTF3LkUg6RiVf7vkEEwrbgAeO2CnDxtGAJzkeAsH3HK+D/wx59NGSLXT9pCNhPJrcN5jDAhHwpyNZzDIxk6CSSYl47NcULLPyx4wtaAOeFH92jACSLXT
8y9vaIJA1JzsHgxwGdrMUO0wCzyx+lhYEqYtawCLTI7H3BFTBhGsMoUF3/TCom8HAAnhRLLekGRg2sOBVp/ohlCYsaYerLLaGKotBmpsi0EwkIAXyk4BL1Ew
ieozsfdGgDFmWTvJ0qDEcnWXMb/yVyE7Lu29DTtgEsl8Xij/pe+6UXBN52VShPM41ixitZOjM9+rh3Aob7gAvEwgRXDcHqo9pkHsMwbBlzedZVMAPQDHwEmO
tYa4AXAVL5InTQr9C6bEULeYZHU3/kCq3ckCl0WmHk50KGaFnsfOpGW52HWzTMAo0l2Rfs44n9yKA6w4fBEi4e5lNcDL6v5wz0HSMQtYFerBGGBW6LH4PCAO
cCJZigP8z2cXmMstq/TAQ6VNLB/ftqgOsr7uYddLfvDC3zdiZMZcr+4O9zyzRNRpRTXgjUgBwMf0DKCSREVp+H/AKLimmwSaxwv0VatCO2as0Qei5Xs98PQH
rVD0vTtoXr62pX3QKzK0mcYMms1lkDd90wfXBjiwT51no1AAMODxwuQ8V5CUTjpsCvmNRaEd+JaxsyxnSwRe/OR02GgcaKX7AMZlOsH/GZiGjYpjjklSt6NE
7ndH84NBLPvuJKTY6Y9SOtng5lXZLAp13764DuxfdbMBrdjrhVe3tEN+RX/MJKAhaa9sbmeE4PR5bFUzxhGvTSHe4u9OhhxwZ58bjpzp1V0LJaWTCl5Sv5+Y
UwXLKoNd/HIN48XUwjrmHTfluOBAc2cQCVur22HSQhd09nl011FKo7IcERKsdtqH8zzRBPgbEvzY6maYkK1Be/eAbrB9bi/cWlAFJZVtuuvRpHRCYVM095wN
Z5JKgs8rJi6ogvXkTJA3IAEYDHsGPEFSOsWunkt+PJBp96MrjyadBLQ7Fh+CuduOgturD5Bd/R7AwPne/lNxSelIy/WYwEukAn8AixR3LK4bERJ+v7SeKcLJ
eVVAjunjw+KKE3B7YXVQBmFSWiGN8S7Xo4ITHSstMoV/bT03tIgaKUON8cd3j0J6tlMXHy72uuHGHCcTS/FI6XDL9aiwKVr3rLUtIzr4wPiQnlMF69TTugEv
/LoF7lxykAmmeKR03DBJ6hqcBsLO8AWKkTCsZeR+06obLHoGZpDNVe2662Xa2cRJaZOdPoyu9ddPY1eCybDF3w2gIoTtB88FZQp553G4p7gG/J1hSEpLZFsC
vIDU4hu4nAGgmhR3dsHC8t5hPyN3dx+WykHxK7H5cKqjH8ZlarCj9lxypLRVoQMvfnJq2J3P390Hk3KrByu7EoHHVzUP+1lvbB8MeIdP9wQRMW97MzxYUqtb
caKUTkuElLYo1ItrguF2/OHSJnj03UOsSqwe62TpNVLdL5rdkncwSCkisP6InoKZIXCqpNgjV6WjwqbQnmfWnYi5k4u+HWB1fozm+Hlyfg2s/OFHQXNXUQ3M
LmsbNgl3Lq0PCo4+vL65CaavPJR4KW0UHV/iYqm40hO1g1jLsygac/uphbVQuKefvbk1fqrunuJa+PPHoUnA+wv29EeMLWlZTtjk0mcCHxrP9uIKFPY2dSRW
So8vqLjBohDPfSuOsE5EIuG2wlp4Y1szc82HSutg0sJqSM92RSUBl9F/WtvC5jum4qfXhdYjT33QCjfmuNhcDwcc8FNrG+KW0tG9QXBNR5fCQmUkEu4qaoBZ
axuYzsdl7hOr65lXRCKh9AeA6auacQkMO+vOs/yOZCwNUYmakl8DeeWhp4IPh071sPqlI2AJHk1KxwSrnXqikYA1RHTX2R8fYXq+d8ADc7c26zrkTwLK7z+8
28QiOAZNBK4YkfBQtYo7l4SPB/54/qMjzOKR0jHBZte6ZsVQQ0AiJmS7mEsGLnP9ScDA+UBJI1sPuFq7hoolSMDzIQInTsXxWU74RItec6THu5j7V116rk5K
S+TAZW2Lpy9wRY0LaHm7++CmnGoWqQPnL5Lw7PqTcO+KIzAxx8nmKwJlLyMgTNDEIg7K40jxwB/PfngYSgNSKU41i0yGL6Vtefs5rNnfX/LfofQXkYiKQZH0
YEkdW+35k5Bi12BSrgsaLokeXBSZJRqySIuehb+JBO2qOw+xAuNSYIk2IVJ6rHBgBhKBgxMi7Pb6p7y0bKdu5Xff8lqYnF/F0hkCNQQOEHeGQuqConq4d3kN
aC161x4uihMhpbl5VTZepu6Z7x+PSdwE6oT9RzvgxMX+oQ4hAVhtDtV28bcDkJKp6RZN9ad6WP3AF0fixZCUlsiK+EYuwBhe1m7xfbTZNfffNpweFgn+ixuz
QuH1redCtnvr84swpaCGBTNfrQDJQNk9Jf8g+9dficaDYUlpXiQOzPdj5b2p+Nlip20Y2YdLAqa5QQKCd4vRcC8DozsKn7NdP1aSkBAUTHjPSxvOwIRsp275
HCt8Uhp3vGImwZihvmzE1HKpJoebsthJPACVsaMDXtnUDm9tvxgTCfnlrWCza/Dm56HvR+1wV1E9PPm+XvUh0hc4h4izf9nNhJjv+E68SEhV2qg45lgU2je4
VT9YfPVtuPrbzXnVsHqfXjE++2HoBRnuWz7yThPTDr7AmSwSfFIaX6ghESdMDQKMschaCwax5z9qA3FXJ7zx2QXAUyImicLcbc1BYumfW84xl8aSHXoFDn5c
lhMmZDnh7e3HWNAMtNRMDV7ZPBhD5F1djITC3a0h743Fbl7ounwprYMAY4wS2YJ7lH4HObptmcSz2m86YFF0akEt+96qUK/Pk1Dv2+zUk5al9U3I0npDGbZ5
4ZKWmP9FJyMh0v2xmNVODxuSgetztet8eTgti/b7x4SJORpq+AGT4JiN3+NWPJ5RjOW5Zpl6A6eDeZ7KG650pAWQgMJnuIrtiiDBJDhm+84xxppeRpIEXnIs
MIrqHkOywONJMJnC3cWH4f4VjYMnvySy6UoiwSiRFpNEPHFvr8UKi6K5n/FLbXjmCAVHoAbnJPUQJ5KN8ZKAbxH3CCNp+mgkYNuknVuaUuK4FiP4/B0dulUi
U5GZjmn64/5qHyeptXGTIDq+wrcYaY4nOibgqTo8bxXXwYzHVzWzeiB24rn1J9mxmqCjfAGuGNd0iOLGiSdB7eUktScu6WyWCVOAt2Kel/GYvFr4U84OWGrD
g6dxNeIE7XcoiDhR3cFJ6iOxtLmSSRgxpI2SYPj5kWAU1T14DEZnEmFb6uOzNGYYV0yiCkH3xWKXDn1ijdJ2aadrWM8R1eSsF3ynXE0SKdCZSDaaZLrc95mT
6Ab8vxNB98ViItnKi2Qp+1smhZyofj6c58Qaz0YxilGMAhn4H+JR69W6FGnCAAAAAElFTkSuQmCC"""
)



def make_tree_pixmap(size=64, bg_color=None):

    """Trả về logo phần mềm (icon "Forest" của Icons8) dạng QPixmap vuông

    kích thước `size`, scale mượt (SmoothTransformation) từ ảnh gốc 65x65.

    Giữ nguyên tên hàm/API cũ (make_tree_pixmap/make_tree_icon, tham số

    bg_color) để KHÔNG phải sửa hàng chục nơi gọi hàm này trong toàn bộ file."""

    raw = base64.b64decode(_APP_LOGO_B64)

    src = QPixmap()

    src.loadFromData(raw, "PNG")

    if src.isNull():

        # An toàn: nếu vì lý do gì đó không giải mã được ảnh, trả về pixmap

        # rỗng trong suốt thay vì làm crash toàn bộ phần mềm.

        pix = QPixmap(size, size)

        pix.fill(Qt.transparent)

        return pix



    scaled = src.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)



    pix = QPixmap(size, size)

    pix.fill(Qt.transparent)

    p = QPainter(pix)

    p.setRenderHint(QPainter.Antialiasing)

    p.setRenderHint(QPainter.SmoothPixmapTransform)



    if bg_color is not None:

        p.setBrush(QBrush(QColor(bg_color)))

        p.setPen(Qt.NoPen)

        p.drawEllipse(0, 0, size, size)



    x = (size - scaled.width()) / 2.0

    y = (size - scaled.height()) / 2.0

    p.drawPixmap(int(round(x)), int(round(y)), scaled)



    p.end()

    return pix





def make_tree_icon(size=64):

    return QIcon(make_tree_pixmap(size))




# ==============================================================================

# 0-ter. MÀN HÌNH KHỞI ĐỘNG (SPLASH SCREEN)

# Phần mềm mất khoảng 5-10s để dựng xong toàn bộ giao diện + vẽ dữ liệu mẫu

# lần đầu (matplotlib 2D/3D khá nặng) - hiển thị màn hình chờ có logo + dòng

# trạng thái để người dùng biết phần mềm đang khởi động, tránh cảm giác

# "đứng máy"/treo trong lúc chờ. Cập nhật trạng thái qua build_startup_splash()

# trả về, gọi .showMessage(...) tại các mốc trong MainWindow.__init__.

# ==============================================================================



def build_startup_splash():

    """Tạo QSplashScreen nền gradient đồng bộ theme (navy - xanh ngọc) với

    logo phần mềm + tên app ở giữa, sẵn sàng để gọi .showMessage(...) hiển

    thị dòng trạng thái đang tải bên dưới."""

    w, h = 420, 260

    pix = QPixmap(w, h)

    pix.fill(Qt.transparent)

    p = QPainter(pix)

    p.setRenderHint(QPainter.Antialiasing)

    p.setRenderHint(QPainter.SmoothPixmapTransform)



    # Nền gradient bo góc, cùng tông với header các hộp thoại "Thông tin" /

    # "Kích hoạt" để cảm giác xuyên suốt ngay từ giây đầu mở phần mềm.
    grad = QPainterPath()

    grad.addRoundedRect(QRectF(0, 0, w, h), 18, 18)

    lg = QLinearGradient(0, 0, w, h)

    lg.setColorAt(0.0, QColor("#16213e"))

    lg.setColorAt(0.5, QColor("#1f4287"))

    lg.setColorAt(1.0, QColor("#278ea5"))

    p.setBrush(QBrush(lg))

    p.setPen(Qt.NoPen)

    p.drawPath(grad)



    # Logo phần mềm, khung tròn trắng mờ làm nền cho nổi bật.

    logo_size = 84

    cx = w / 2.0

    logo_cy = 92

    p.setBrush(QBrush(QColor(255, 255, 255, 46)))

    p.setPen(Qt.NoPen)

    p.drawEllipse(QPointF(cx, logo_cy), logo_size / 2 + 10, logo_size / 2 + 10)

    logo_pix = make_tree_pixmap(logo_size)

    p.drawPixmap(int(cx - logo_size / 2), int(logo_cy - logo_size / 2), logo_pix)



    # Tên phần mềm + mô tả ngắn.

    p.setPen(QPen(QColor("#ffffff")))

    font_title = QFont("Segoe UI", 20, QFont.Bold)

    p.setFont(font_title)

    p.drawText(QRectF(0, 150, w, 30), Qt.AlignCenter, "TDForest")



    p.setPen(QPen(QColor("#dff3f7")))

    font_sub = QFont("Segoe UI", 9)

    font_sub.setItalic(True)

    p.setFont(font_sub)

    p.drawText(QRectF(20, 180, w - 40, 22), Qt.AlignCenter,

               "Phần mềm vẽ trắc đồ ngang - dọc rừng (2D/3D)")



    p.end()



    splash = QSplashScreen(pix)

    splash.setWindowFlag(Qt.WindowStaysOnTopHint, True)

    splash.showMessage("  Đang khởi động phần mềm...", Qt.AlignBottom | Qt.AlignHCenter,

                        QColor("#ffffff"))

    return splash




# Icon "Ánh xạ cột" - 2 mắt xích liên kết + dấu tích xanh, nhúng base64
# (nguồn ảnh do người dùng cung cấp: Link.png, 48x48). Dùng cho QAction
# "Ánh xạ cột" trên toolbar và nút "Ánh xạ cột..." trong panel dữ liệu.
_MAPPING_ICON_B64 = (
"""iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAAsTAAALEwEAmpwYAAAEbklEQVR4nO1Y7WtbVRw+VTcVFJXa5pys4ofVD06pStl8
2dDm3u4eBWGyLb78A8MPtjKbc9qCElb84FyZhdEPguzesqTVfhC/Kw6/DNRtzTkJdmta50JZFWGzJpnvi/zSpN7cntwuL82tkAcOhHt/uTzPye88vycXoSaa
aKIxCJ++zT/0XeuOcGIr+r+gs2/udhKSfYSJbwgXNwiXOViYiRjhcrgtnLgLbVbggcQOzOVckbRqYS4XfQOxJ9FmAx6K78RcXnMjbxNxnQzJbrRZ4A8lHi9P
XiwrrzO5sCnOBnnr7P2Yi9TaXRaRbVx0QA0ePt+GmTihEHHIa/4Iczm2lpgYUdYycdLRSp8jzx2HiXQJKSZOlKv38ZlHHEJ/RF4CD4pnHYSudBw+c2fv9NV7
9kR/uc9ZDxZKSsX+hrwEYeI1R99/2hPJ9mvRzJ9aJPNXYDLbZ6/3h+TTjvqUx/0f328n1PXBYlKLZnO29SvK5VpWqnMt0POkVMDHngqw9/Sjoyk78cLKJIvk
CZfjzsPuY/EXkdcAP1eRD0QzmZ7I9d3tw9JHuJxWWOhXIKxxbjMY2+PnsVfJoNgHkaF4b9f4T584yfdE0rmH3k3+gLk8h5n4QzHcloozogFDSo45rbLg4XM7
x3+eUpHvHJlziRJiCfP4wxtO3j8Ye0I1YVd7/tjatlmPPOZipi0U3+55MFORD0Qz2a7RVJgw+bU9Rhf8fpZw+Qa0ImpQJFaQF0uYyc8eO744r9r5ruOpd+yt
187FU34ef6bj8Mw21Cg8GL50hyrPQzyACatFs2+uIX8qndt+JAntcQZ5DcJlv2Lnj8C958xr9+YnrIJ8IUb8Aztf7tnB6eBWfcJ4RTeNSd0yLmgmzcCCz3Ct
16IvQ02NAsS3jmxzqnhPn1huDUQyf9t9vnMkeblkMPGZXarn6qZxQLPogm7RnOsyjfnek3R/dewPnd2S30UboQfYOb+9BFoIiAeimQUYUuDz9nrIOvb64HTw
Vt2io+sSt5xC6LFwOHxLRfwLE9PW9/KqW71vINYOadL+Hedwqoq8VVzG0cqnbWnv34BXIurqXAtmcsoh+GIJedM4UI6cE+XqNNN4qSIRhMu4Y9qOqckr/h5y
2V9yYE1jvnYB9PuKDjZm4m3nr4CZ+AhSZyubvZuE4rsJF18qbDYGFry6++A2Lu3hhFutZj4fvGkBeZJMXCmfY9TZpmMo0Wl/jm4ZU3UTYNEIqgTgJPn3NjdJ
XhXMNIvOuRFeD47DfAFVk4UIk5dcyTNxulxM0E2arpsAk6ZRNQBXwky+jrn8At4eYC5+x0xcxlxOYi5ecPtuPQVoprGMGg3NNC7WTYBFZxsuYCXv1OcQ66ax
GmcaBghm9XOhvQcbLqD7w+4tukmTNQ8yiy7UnFCrBaTKmn+BCboPeQlIla79bbn1Pn0PeQ2IxJplvF+5AONoxXF6IwGp0u1M6P/tetLztnE72BDMNItGwNvz
w86kafi8cm3vQajxmmcTTTSBNh7/Aj+4H8ncvmbWAAAAAElFTkSuQmCC"""
)



def make_mapping_pixmap(size=32):

    """Trả về icon "Ánh xạ cột" dạng QPixmap vuông kích thước `size`, scale
    mượt (SmoothTransformation) từ ảnh gốc 48x48."""

    raw = base64.b64decode(_MAPPING_ICON_B64)

    src = QPixmap()

    src.loadFromData(raw, "PNG")

    if src.isNull():

        pix = QPixmap(size, size)

        pix.fill(Qt.transparent)

        return pix

    return src.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)



def make_mapping_icon(size=32):

    return QIcon(make_mapping_pixmap(size))




# Icon "Dùng dữ liệu Demo" - người dùng + tài liệu trong khung tròn xanh,
# nhúng base64 (nguồn ảnh do người dùng cung cấp: Demo.png, 24x24). Dùng
# cho nút "Dùng dữ liệu Demo (không cần file)" trong panel dữ liệu.
_DEMO_ICON_B64 = (
"""iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAACXBIWXMAAAsTAAALEwEAmpwYAAAByUlEQVR4nO2N30tTcRjGD/TP5FVd79CO0tBCEFZJRKDL
ublfoEMWXUhQioTaRXkjLgdOclN0QTg3KsSLY2TNK2EJsUndBGHgrVcfefddd32PeOalD3zg5X2f530M41LnVe8qV2IbeCMlUpEiS5Eie9FNatEif6ObnDRQ
c01uDU+JlGQk6/g8VOB2uMDP8DtwQ6jAYXidW9qCgRzHwTy0wkCOY23Bo0W4CAyd+ubhIjB0Gs7Cw9nWGMk6FNh1SC3Cgxl3PM7CTt2hYP8PCB+/Q2YbJtdg
ZB6GXkNAnrxQyCy7ZFp5xPvpQGUFbcGbLShXofLbHeUqyA9tgf8JCINT8OwtzH2A/C6UqrBVA/uXQmbZyU084g1NqaygLeh/Dj3J1giMOxTkvkFwHLpj7ghO
QL7iUPD+BwiZHZhcgeQr6B+D+6PgT0B3WCGz7AJjyiNeyfzLawsmliFtw9qBO9I2yA9twc1eEO7FITENT7PwsghzNmT2YGlfIbPs5CYe8d6NqaygLfDd4bCj
B1rB56euLfD6aPN28cXbBW640clnq5OrxlnyWFw324l7LBbMdr56LGqmxZFpcdLkqLnbbXriZgfXznx8KeM/OgXaPauNqxhcGQAAAABJRU5ErkJggg=="""
)



def make_demo_pixmap(size=24):

    """Trả về icon "Dùng dữ liệu Demo" dạng QPixmap vuông kích thước `size`,
    scale mượt (SmoothTransformation) từ ảnh gốc 24x24."""

    raw = base64.b64decode(_DEMO_ICON_B64)

    src = QPixmap()

    src.loadFromData(raw, "PNG")

    if src.isNull():

        pix = QPixmap(size, size)

        pix.fill(Qt.transparent)

        return pix

    return src.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)



def make_demo_icon(size=24):

    return QIcon(make_demo_pixmap(size))



# Icon "Giải thích cột" (1.png, do người dùng cung cấp, 100x100), nhúng base64.
_EXPLAIN_ICON_B64 = (
"""iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAYAAABw4pVUAAAACXBIWXMAAAsTAAALEwEAmpwYAAAGPElEQVR4nO2d/1MUZRzHF5IZGX9u+iPCElNnEjU1k3FG
iRGMr4qooSYcD1QzoKIZ0RfLwrLUoAwFR2RUSis8D9cZZgBLmSsrUwRxdOQmM28fciYD+TSfmz1a7g644569XdnPe+b1G8c+z+d1t8+zO3f7kSQKhUKhUCgU
CoVCoVDGGwCI5pw/xzn/kHN+QVGU24qiDHDOwQooijKAc8a5c853cc7nAUCUITLcbvdiRVE6jC4KNx8X+/r6FkVMBAA8pijKuyaYOJgZRVH2A8AkvWVM4px/
Y/Rk+SOCy+VqwjewbkI457uNniR/xHA6nXWSJEXpsmaMdNC21p+hrLQKli/dCgsTGMx/1mYJFiYwSFm2FbaVVkN766URpVRWVhYJlaLupi76HujePTe8V1Fr
eGHmm4Sdb9d5auJbp56enp7Y2NgZwoT09fXND2SeZNgCSglUq+Li4h2SJD0hRAjn/KNApyntQJJy90PeoRtQ6BgAJoMlKHQMeOactHr/sFqcb/M/fTkcjlOS
JC0RcurinP/oewBcM7QyjC4OM5gX1/5fj22bq/2EdHd3X5MkKUOSpMfDFqIoSq/vAXAB9w4gr/aG4QVhBpN3qGeoHrjQ+9bL5XL9pQqJFyHE73aIdjdlpdMU
G4FCR/9QPbA2vvVyu90PVSGJYQsJtEhpz5lGF4OZBG1NAtVMFZJMQmRTCUknIbKphCAkhJGQyGI7Y/zGwtJCCs8OQtrn1yCh6HuYml0LT2bUQHzuEXh+Wwus
O3aPhESSfPu/MO/1Mx4JgYjLOgipn/5OQiI28ZLmEWUMkVkDmQcie8FqSSFZX90cW4bK9HUNUBDBi1ZLClmw+VzQQpDsmpskRE+eyTsWkpCkXT+RED2ZvrYh
JCFLd3aQED2Zzb4LSciKvVdJiJ4kV14KWkZc1iHYeOo+CdH7GmT62qNBCVlScT6iY7OkECYDrK7/A55edXhUGXOKmyK65bW0ECYDvHzCDXNfs/uJmLqy1rOQ
25ofRnxMlhbCVNZ/zSGtqgtSP7viueYosPcbNhYSIpsLEiKbCxIimwsSIpsLywjJrb8DSR84IfHNNkjcMQbl7ZC8+1fY1PQPCRGNzTEAi8paQrpV4iV+Tb1H
JAkRSGJ527hkeJmx/rhHKgkRwIaTf0NcxsGwhCCr6npJiAjSqrrCloFkVF8nISJYsa9TiJB0EiKGNQ13SYiZhDAZYM6rp0mImYRs/PY+zHqlkYSYRQjzfEW0
33Ohh5+WWRtOwIy84xCfW09CzHTrJL36OgkhIUBCGAkhIYyEkJCQQ0KAhDASQkIYCSEh4woJARLCSIh5hRTY+yG75ha8tK8TUvdegcU7WoMWgn+Lr8HXZh+8
BQVn9PtmoyWELP/k8phfrA4F/F8pe/T5he6EF7L849+EifAlZc9lEhIK+U0P4KmVdboJwU9K/ukHJCRYMg/c0E2Gl8wvxf6OfUILWbH3qu5CcKEnIUGSU9er
u5Ccwy4SEiyFzYMwc8MJ3WTM2tgI7OwgCQmF3Po7Qre82gV9zdE/hY93wgth+FvCRsXzOA387WC4IqaurIOFW85BXqOiy1gtIYRpwG0qbofHheAtLgmRzQ8J
kc0FCZHNBQmRLSqEHjUOIh81LkTIbd8D0MP4YcSH8acmlfkJcblcd0WesvzaVWCbH+8AsH+G0QVhBrMsZ99QPbZv+cJPSFdX1zWRjxrf5XsA7LmkPWeiFHyX
WK+hS88wGQtmF8IP7b/4CbHb7SdFPox/XqBFCtv7aKUQNnj/ncAtjxhj21Uhi8MWgu1E1faiww6CDbBIim2YjEBNwbC7TnR0dKYqZJokIthONJB5BHsu4ZqC
i5nV2ualJpV5WhwFOk15KSkpKdcs6OG3PBqtORjBR8XhcHjXDjHddbTBNqLYTtToSfJHBKfTeWHy5MlZGiFi2ub5Suno6Dhg9GS5yWlpaWmeMmVKtkZG+M3A
RklURUXFps7OzitGT5ybjO7u7q7S0tK3NCKQubr0w/WVEhMTMw07WGLTRNxJYGs4vEVgdFF4hMC54pxx7rhWFBUVvaHZTXmZGQkZ2uCu4QWfQRCSZwEXv2aE
KCZeHUiyeosgwyKkq3NOVGsgbmtLoVAoFAqFQqFIEy7/AXm09MDPI3XBAAAAAElFTkSuQmCC"""
)



def make_explain_pixmap(size=24):

    """Trả về icon "Giải thích cột" dạng QPixmap vuông kích thước `size`,
    scale mượt (SmoothTransformation) từ ảnh gốc 100x100."""

    raw = base64.b64decode(_EXPLAIN_ICON_B64)

    src = QPixmap()

    src.loadFromData(raw, "PNG")

    if src.isNull():

        pix = QPixmap(size, size)

        pix.fill(Qt.transparent)

        return pix

    return src.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)



def make_explain_icon(size=24):

    return QIcon(make_explain_pixmap(size))




# Icon "Xuất Demo → Excel" (2.png, do người dùng cung cấp, 64x64), nhúng base64.
_EXPORTDEMO_ICON_B64 = (
"""iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAACXBIWXMAAAsTAAALEwEAmpwYAAAK3ElEQVR4nO1bC2xT1xm+bOtL6jZNG1s3adqmlY4Wykan
rVu3Sl3Vdepj7bqRlqmle/SpqsDWbqIojECpCCQk9zqhhBASQiCJ73ESnPeDkNAS8vA5zsNxbMfBvk5iXxrYmCCBhIf6T//1g9jxKybOY+WXfsk6j+v7f+c/
/+ucy3E3aeGRXJ4AoRhcFX4c2M9yn/Ixt1DpJgDlNwGABQWAinTexxP6Bk9oNk9Yg0DoAE/oOTeza4LILvAiM/OEHRdElsOL7LU0wu4nhHx+wQKQrqa/QIEF
kY4IhEEszBd3fXe6AFyx5sF/WzYojL9nFYDsbHZLOmGv8iK1+glC2KAgsoMCoW8LhP5aKOlcsqfQ8JUkYrw1jbTekVmk/1a6pmOZSs1+pRLZPwRC1agpof4n
HACROC4AAMAiFdGt4Qmz+4QWqUMQ2baMEraUm2GaVwDsKtZ9WyC07rrgrFMlsudD7d//KwAEDUvgCTuvCC/SEZ6wP6M2cHGmOQcAABbxhG3mRfqpZ+U1uKe5
WaI5BSApCT7HE5bvMW5XBJG+yc0yzRkA4F75LPeq0zFBTZ/k5oBiFX74yKobA4AnLNkrvEqje4ibI4pV+N5Dz8QOgEDYc+49Ty+rNOzx6cy1Wv/9JZPk+qNZ
knPMkktndrjOWByuK8j4W2mT5H39knO12Xz2i5GeN1mIG2EuWtpdov+OQNh/PAbvrWjnmWzOeywOOddsd17UGSzQ2NIOlQ3NUFJRB+ojVSBqq5Tf2NbY0gG0
1wI41izJ+/tODS+ZNwAIhB7z+PjiaMYPDQ3dYXG4dlkk11Vm7IcjNUehuKwyKsaxOEfREMmVIknS7XMKgKBhCZ6VP7urkH0t0nhcOYvk6i1kDFI+boay7h7o
sQ+DcWAQusw2oAYLtHf1+TG2YR+OwbE4J+XjJjhMGQLRZnKc+SY3F6Sqsd6mhLRKoKN/KdJ4q+RcaXK4ziQfb4LleQLcl8vDjw5kQEu/DV7UquFxdS48V1IA
a7RqyOvoUBh/Yxv24ZjWfhusPJChzF2Wy0PisVros7uGLbbhFdxsEy/q/uoWnhnR/0daeRT+9apS5eW9vKGhGg7qdH5tyFubjyoc2J5PdfDPhiq/tjery8Ak
ycNGaeSuWRMeABYJhBkUADQsIdxY3Kdmh9yd1NTg9+KoBSctNnhCzIsagCfFA9DSf8qnQV7GZ1skmaF9mRUAeA19xLP3hyIlNmjwOgbskFB62O/FV5Uehjqj
aYqQ4QBArjea4A+lh/yATCgtBPwPs+TaOjsAiPRDT7j7fiRX131q8GpSUz3U9PZBW78N9ra1wjt1lbC/ox22BGhFNADgnJz2Nni3vgqy29ug3WpTnr35WD10
nRoci/tWAAx5RTaMAGA5K9xY9PMfHG/0vfyjxTnwdo0WPmw9CSa7E3rtTqADDkUTCvVMMXy7W09CtaFPYfyNbdiHY3AszumTnEofPuvRohzf89HAmh2uPXEF
IEPd+UOv+keM8OzOiw8d2jtlFR84mKm4tPvzBPh5wR54muTDy+UivFpVAutryxUXh7yuVqu0oTfAMTgW53TbhmBlfuaU5/7ycBYCOxZNxOgldN8q0rk4agAE
Ql/xqH9+uHEY3hbp9UHV+FlNgaK2wfoibQHkWqMJfqvJD9on6rvAYne+EK08PKEiciz7/2/hxmFs/97RmqAv+VaNFva1t8UMANoAdH/B+hIba9EjZEcji0pD
n/XVJyN4Mx8p5WnCIF2te4wLQ5jE/H6StZ7MGxtrlWguVgBSTzTDhobg4CaUFSoRIheF6vOEfTKpQHtWKGr7RqR5HNbolQklnUvCAyCffbhw6v5HRsOIniFW
ANDvbwvR/0jhPjSEIxHlIFScUmoXmTYyAISewcGZZR1fDTfO4pAv44tiuBrI2m4DYD4QrA9Z7OxSOFQ/zj3SbQjah27S7JAnolb9QA7cCul4iCHSDV4WRHbV
bQNoorctWAEEAUg9cVxxTYFcaegFTVdX0D5kFA45VD/OrTYYg/bh9ggHAC5cgOoHsv9W2DV1r4SfMGkLoDoGVfGmhpBBUPgtkA73Zu2EV8QCWJWXBd/ftkHh
u7dvhHt2bYGle3ZE3AJouwTCdiDzhKZ68pmr3jbkKUUd1XRUxgeAS4fhbzABMQlC7ZgOAPdmp8LdOzfBj1PSYC3pgndKZdiovaDwWjIECXnN8EBqGvwkYzv0
Dji6uCgIF87j0U5HHiwyMh2jgWUsdEnBBEQXhqFwtAAszdoJSz5IhBcOnIDEinHYVDERlBMrLsEbxQxO6k3GqABQ65Z7AOjlZtpt9EvO1RiUBBMQo7pQiVAg
ALjyKPxrReaQggfyjtpRRxKBWyPJhKruKeQe42Y6cDCOjNyJYWkwV4hFjR4bhsIq+NkhdyiM4S6GvRj+Tg6FH8sR4KWCtqiFv64NE5mR5FGMutugC1EB4Asd
o/GX7m2wHy1zYDKU2dqiJDQG+zDorJIS2h5m7mQI+7zJUFJ9NawXK2BTGLX3clJlYNv4tU3ll5eFez88afYYwT9x0ZKKdC5OLe3+ejRjsRLUYx+8gqkqxv6t
/XbI8qbDuo7w6XBTg2LZn889HlH4bTXj0DV0EaoNl/y1oHw8PUJRZ0gBQK1bzsWLLJIrJVhBBENWLG6EswFLkjfCOjIUVvj3a8ahe2gM5HOjYHaNKWD4+svH
Qxq3NEJ/6j2qj+uhrSRJt2NsHlgSw7QWNQLLXKE0AH38e9oLfgJvrpyA5Lrgwu+on7INLoR6L0Gk26e9/2MlrNKYHPJQYAaHCc1BSkNqgBuA837Cf9R/EQY+
GQVV0zh0DbqF7z89CjsbptqJxPLx88HeB2+f8IS6lKROwx7mZoMstuEVWL0NVhZHD/AbMQ8we1xTLvrK4g9mJMP6EqdPoK1V46B3uIV2nRsNs/LhtwAaPY83
M8zGnQUfWa3yYrPD9REmMxizY7yPnqDPNgydysFIP3T0mH2HIrVNJyCjyuYn1JaqCaD28Cs/CYA0LoCwiMuLtMej/i9zs01Wq/U2rN5aJNcYM1pBW9sY8ijs
YHEJpO4j8K8AN4iaoKaXfLYgOI9f26idmFKzxDsLnujPjluBmysySiN3YQHTbHeO4cEnHoDiQaimog5EbbXCmso6yCkogtTS6KPASauvChbNeg9yeTX7XVQv
Ksf54iFGjFjDszhce80OVztmcphKI+Nvw4Cju0Xf15tcc8ExDQCOvp4Nt0y9wEHLPHv/aNQrJccZgGgJY3sMb5UIL4za48oHCo8kEPauZ9+fyyDt31twAHgJ
w1uM8NDCJ1aMjyLjbzR4wfY8Eq/RPYF3lpRLHBr2NDcdkucZANMlvGEqEHbJE/Nvm/YD5AUMAF7U4gkd9aS8u2N6iDxfLx9HIEGkf1dum7td3v6YAx55rm9e
TpM8ZS6NJ9H5lCd00w1Fe/ICAQAjPIHo/+Lz84SdTyd01Q0/WJ7nAKDg6YS96DuwcQtfjRe1Z+QP5HkKgKDRrxAITREIdV6vU9IBgehWz+gfyXEEYDcx3ikQ
1i8QVsSLbB0WKfHjiB2EfRkvWHtZ+axGo3tCRehavI7nTWUnVafNmOElNTd/YUaFjzcAPNE9E/MnM4Sd5gnL5DXsQS6eJMfx8jHeLsMVx09p8IMojNF5QvvQ
kF3/cIqeEwg18YTWCiLbq3x2Q3Q/iKvQc375eD4Rm+27t/ON2E0AnvpsA8B9hul/W7WaIXO/tV4AAAAASUVORK5CYII="""
)



def make_exportdemo_pixmap(size=22):

    """Trả về icon "Xuất Demo → Excel" dạng QPixmap vuông kích thước `size`,
    scale mượt (SmoothTransformation) từ ảnh gốc 64x64."""

    raw = base64.b64decode(_EXPORTDEMO_ICON_B64)

    src = QPixmap()

    src.loadFromData(raw, "PNG")

    if src.isNull():

        pix = QPixmap(size, size)

        pix.fill(Qt.transparent)

        return pix

    return src.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)



def make_exportdemo_icon(size=22):

    return QIcon(make_exportdemo_pixmap(size))




# Icon "Lưu cấu hình" (3.png, do người dùng cung cấp, 100x100), nhúng base64.
_SAVECFG_ICON_B64 = (
"""iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAYAAABw4pVUAAAACXBIWXMAAAsTAAALEwEAmpwYAAAKzklEQVR4nO2dd3AU1x3H10n+SpnEqTOJgzASBFCQERIg
wBSBKA7dIDoG0WzRJYoASTbVNIEtwGVojiihONgY0UwJzWADKrcrIQi4IDqWdPsQxA7C+Jv5vTuMhtxJu7d7u3un/c68kUa3t7f7Pvfe7/3KPgmCLVu2bNmy
ZcuWLVu2bNmyZcuWLVtBrtBateXQkNqwUHsYGhLSTqipok7Q8zilKsH5XzBIUTIKXpQhJTBI4xmkFPqc8Hr1yjNWpi6h1+iY2yj8uVBTZAQQYPuPXxk3eHKj
8AYsKjLi2w/3ry5hkOCp0eccPbMNTZs8h4MnN1V+7aoMaR9BYyhoSucUglH+AgLgRzLEDgxSFoPEWsRE4ejZrbyzW8ZEoyog9NMLlMpNZpDek+FoT58lBIv0
BsJQ+GsG6VX6Rnvq6Cd/9wZEIRR3E68wSOkM0tNCoEsvIHch/YFBWsIgllfX0UqBqIPCwdBnL7oLx++FmgoEOPITJ8RJDOKdqjor1Ecg6qHwdo9sDV2bUJOA
3IGjOYPoUNJJoRqA+AiFWh4tAIRgBwLgKRoVMsT7SjsnVCMQX6HIkB7IEGcHjOFXC+QOin4jQ9qr4lsKanyVdWYbjpzeqmiVpScUFxgxmxYcQrAAqVMr5H7u
pf1hMqQLamEwSNj9r/UcBLU9R97zeMzX93PRoG7das+lAUqRE4W1hGAAEhHe4NTrGdNYSUWeahhMQaPzZqxIw+AhvRQdf/jTzYgIb6j6c2SIN5yQIoRABlKG
/PCiqwed1Fn1w8L8EsOqHxaGwUN648L1w4o7t7rprQoopeUobCgEIpBSiM8wiMX+GBVMY/MViLtdc+JciBBIQNwG3CebwawPhNsUyxl67w4fnmKQPjK705kf
gbih7KF7FawOxBVVNb/TmZ+BuKBIyYJV5M0DlyFV1CAgFQxStGBFIJRnYBDzze5sZiAQd8u1RI7lSSAyxCSzO5qZA4SixeMEKwFxhdCrjtoGNxBJLkfe7ywD
xJXPML+jmXlAyJ4ssAQQV6bPc3KpJgFhrhniaQsA4WlX0zuZmQ6EQ0kzFQjlCtx5adM7mVkDSLFp+RP3TcWZ3bnMUkDIlhTEmgkkK1CBbP5npp/OL643EwgL
VCCdO7fFxKQE3PrPGb3PL5viKKoZ9tkH16F79zjdcyL1w8L4een8aoHc/jYHU1JGI65Da+T9e7feUKItC4Q6q3WrZjjwyUae3dPzxksq8vDxiY14vmUz7Dqk
HErla//7lmWIjozAlp0rdbsuGeJ0ywLp1i2Ow9D5G4jKbf/xDejRo6NPQKidPb8LsW1b4LX5SSj7zqEHkD2WBULTir/y6azSSKHP8RUItSvsFAYN7okhQ3vj
WvlnGq9JvGIoDHokQCkQfy4xL944wqeaeYumIKJhA8xekITNOzLxVdknPl0TjY602RPQqVMbFF05qGWEfG/ooxD07IWZQC7dPoqpM8bwuX/s+KF44610rN+8
FG++/RrGJA5CZMRf0bt3F/43T6uo6q5pddZCNI+O5LVgvkMpbGwYEBmOPmYBKbpykBfPzV88DVfvfOrxmNIH+dh3LAujxgxEk8YRmLswGZfLTqq6JrJLVMO1
Y++7PgKRehsIREowC0ifvi/gnXULFB9//tohzEgby8EsWpbCR4zSa8q9sButWjTFmo2LVV+nE46XDAPCIE0wA0jO+Wy0aR0D50P1KyECQ44gVUDWfbaO4vcV
XT2EDrGtuJ2ybNKKQZxlBpAde9/lNkLLOY7nbkf79s9j2PC+fFGg5D1flp7gDuj0ma+o+TLMDHogR05v4cZa63nIxlD5KdmIrG3LFL3n5r0z6D+gBx9lCn2V
mUE/ZZVU5KFZVGPkXdQn1EE2ghzXEaP6o1h+bPS9NSrqHj4yHi+PHcyhWmbKMtOor9mwiI+S2//N0Q3ynNeTFC8UCErCiHiMHD2gSofXUKNu5rLX+dDBfY+E
Uf1w65uzup5bDURaUlMr8zJSDF32mu0YllTkYWJyAje0aqre9Ww0ZQ1PiMeEycM9GnpDHUMKC1ghdPLWmrncoyYn0AwoFMLv1787kqaOgvy9aF7ohGQFIMz9
VFRM8yjuJ/jin2ht5GiSs5o2Z0Klv4vFhsKwEhAGCZ9/fYx/U4e+9CKP2hoNhaLDXbt2wJI3Z1k//G4EEEaR2gf5SJ87iec1pC8/NhwKOY/k0VNgUoY4zXAg
dZ+tU+6PR9RCdWjhf6mHE3nvGw6FwizNmzVBwsh+UwwH4qp25xu56HZDzocO7gn37NkJSzNT0aVzO1XJovez30biuKE8rUtBQTNGyrHc7eV1QmqfE8wQ7aqj
581krEhDfL9uPzh96XMm8tSs0qXtPz5YgQmThvPfk6ePQdbWDMOBMIhrBbNEWxzpOQdHRTbCpZuPA360lHy0tP1gX/V5CQqBrN3kCpWT9535zmzDgTghtTUN
iJ6lpJSKTRw31GuE9oUusejZoyMHQ+GLyq/fuHcar86bzI95FM6gahLyDwweHZdN34rDvd+U5puh5NHSzFSvr9NoofD7gIE9eIp24KCeGDYiHrFtW6LJc434
VFU5g3i6cCevuTIYSKpgtpzI+aUeFYzLV6Vj/qKpio6lIgZKsW7ZkYnOndphUvJIjwDJYSz4yijDLt6Rkf8rwQpikBZrvSFaGfXq1Vn1+yjA2K5tC74N4JOv
zUhNxLKVaYYAkSHNE6wi2olN6yghx65FTBSfatS+l0LnKbMS/+/vji/28QWBXqH6KpqTNksQrCTX7nDabmztpsXoG99VdUzqowNreDje02sUkaUcin+BOBIF
q8ntKOZpGiXfObjRppJONe+juixv78n/fC9P1VKRg5+A5FjisWhPom3xtG4cUCyf5HkO8tiv363eS6eVFRnvzwo+9HoM+SNUJqp3NNi9M16UYGXJKJiq9UZv
fXOWG2Sa/6kzKZrr6bhzxQc4vOpWZzTyqECBgo86A5ksWF3uzWd26nHDZ8/vQvK00WjcKBytWjbFqJcH8rwHdSyF2mkqIk9eybloJJHjSA5k5WSSBhjZltp8
pirRo9KuLYz0+SaWPcjH/CXTeMXJwowUDiH78DpFU1rlRrmSPn3+xnPh3spQFbZCy/gcSlWK/D9RKEHPKWLxGzN5BFhLZ1JohSrc6SEiWp35cI6rlt9/0Zto
OzzaFk9PKDNSE/kzHFoN9KFTm9Axro2qul0ZUkk5ChoIgSy6AT23+nM+dPAVU1VxL6WNjH31BW8/tOtOiI2EYFAJcv/IIEl6Qfmi5Dhf6vr7UbnHI0M8V4aC
PwvBJPd/P9ilVycdOrWJp0u9LYl1bDsDzoArlS9bjbMqGj20Q2Wd/hkV0gPXxvwBsrTV6tHTTmx6LIe7d4/Dhu3L9QaSY3kPXG9RZo0KkrWuwvIu7uYOok4l
pTKNYMvGpowQ7cRGm39p2ZFu+ap0nkHU4IEzGdL8cuT81uz+sIzIcDKIab4skWnqotKhdZuXqARBnyWmUtbT7Pu3rMD/KVhBrLvESHHdF4XZqWql8HK1KVsn
g7hORkE70wsSAk1w5Viiae8Q1/8eqXr0rFo9hxc+P/LiqQqd3kO1tlTeSYa6RtsHf+gmHD+T4YikB2KcEIfRo2Pu3bRTSityx8c0a3IhISF+DR1zHTk/Nft6
a7xCnwkNM62U05YtW7Zs2bJly5YtW7Zs2bJly5Ytoebof3PDAuMdx1baAAAAAElFTkSuQmCC"""
)



def make_savecfg_pixmap(size=24):

    """Trả về icon "Lưu cấu hình" dạng QPixmap vuông kích thước `size`,
    scale mượt (SmoothTransformation) từ ảnh gốc 100x100."""

    raw = base64.b64decode(_SAVECFG_ICON_B64)

    src = QPixmap()

    src.loadFromData(raw, "PNG")

    if src.isNull():

        pix = QPixmap(size, size)

        pix.fill(Qt.transparent)

        return pix

    return src.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)



def make_savecfg_icon(size=24):

    return QIcon(make_savecfg_pixmap(size))




# Icon "Tải cấu hình" (4.png, do người dùng cung cấp, 65x65), nhúng base64.
_LOADCFG_ICON_B64 = (
"""iVBORw0KGgoAAAANSUhEUgAAAEEAAABBCAYAAACO98lFAAAACXBIWXMAAAsTAAALEwEAmpwYAAAG+ElEQVR4nO1aeWxURRh/IIp9rweHIEcUEMKhxiOQCBpv
/5AAGjxIFBVNiEEU/8EARk0TElRAaPbNtqXlhobSclVoASUltgW0ULkKBVpKO/N2e+3VY4/en5nXzu5bLNt93deT90u+P3Z3dma+35v55vt98zhOhw4dOnT0
YWwiYWHxZPzQmJKJXIKZ5+4LpMIDvCjNFRDZySN8UxBxq4AIMONFUs0jki2IZGWkAU/mBhSiYTAv4s8FhLHS6YAm4lYekdRII5nC9XcMFUsnCSI+p3RwRJwE
845UQfS5aoi7XAtb8+tgXW4NfJ1pgxm7yu4igzTyIvmJAxjE9Uc8jExzeIRtzKHJO8wQf6UOHPUtEAgXKxvgkxNWCPffKin9Lm6EITxLQNjJnPjmtA2cja2g
BrnlDfDkbr+VkU7jCtcfIMTdHi0gUkYnTp9m4tU66Cos7hZ482Clggi8geszSMh7UDBIrwsIr22L9iRVEEl8uBF/JSBylE065t9aCBXV9S3wXJJyRUiv9b7z
iKxgTzqQfZhuAXUb4N4ocjTBI/ESOznO91qg5DeXjuVFknO3sxO2mWD6LrMc9dl3I+MkuFPTBFrix7MOb/9hRvxujxMQbigbxSNcyCYxbadZ3uuVrmbvJFta
Ac6a62HtP9VwCntAa9BtMXpLG9E8wvt6lgGAQbyIsxgBS05awdWk1UJXh48yLO0kEAfdmj3GAY/IF4yAxSesmu3zrmBPgdOXOxjwzJ4jQSQFdNCJ20xQ19Cb
FAD8XVbviwsIv9ODSQ+RB91woQZ6GyU1TUoSlvUUCcvYoDftjb3NAUh1zUpdsbL7CTBIs9mRGGEk0EjDfy/jUlWDkgSPgHAujzAKM0ovaOp8uFj6koBInjIX
mLTdDH0BN2yNAeQ3uUAfXGjeR8NgQSS/KosddAW8mloBvxe7NXGCHq2Hi1xg8wRWkoGQcccN63KrZVH2SkqFn9oUEG4REF5PfekSAbxIUpQZ35ocB5Q7fclQ
qHA3tcJb7WKIaoFQiFDC7GyGVdkOGK7IWHmEk7hoGKKKA14km1gHzyeVwzWrtkGwudWX6DCjK0zLpKvA1ghP71EILRFvDJoAQSx9g20BSkCVW7unT0HdXPqn
zacrmBBqT760jLdWTws8u7fMV6ITpbnBbQNEbtA/UZV2qxuOwe/P+IQPLZIcKXLBrH3l3u9W5zg0Ha/Q3qTUGPmdxgdeLJ3HJvPzee2TIZpgsf5pxnmo0AXZ
kgfSi90wdafZ+1vs5dDrDkrQwOlLqsjCTkjAybThqHgJ7BoFKoakG05v5B6XKMG+AqdMALOUW04Ynyh5T6G029qcQBS1DS3e+gOP8N57MwAwiBVH3jtaBVri
2B03RMW2ETAq3gQ78uv8CGC2Lb/OGyPoiUS1gVb44BhTnLjqnhxEbiIj2JJZr6EuyJLqZYdov8NiJTBcrOmQAGYbL1TLK4G2f2yrSd7TWoD6xPyLiikZ1iEJ
4WLJdNZo93WnJgNftTTCuESTd4n/khuYAGars+3eCT+1u8yvWNNVUJ9Yn+GxJTM6JCEixjRNSxKKq5vk9JpVmX844wiKAGafHvflES+nVKgu0QcigfraMQnI
NJI1+i2vJuSsTXlH8O1puyoCqGVJHpif5iuvv3/MAk0hxOqNeUFsB66tbGahjeiNTyi1v9nJvnOfZoZqCWCWid0wZ7+vrxWn7aGX4kRi7yRbJGm04eNbTdBA
c1uVoGmv8nJkQVql/ESzQ7CMYjdMU9xJbu7C3YWnuRUeTWBHJNkfdP3wYKFL9WCL0v31wKTtJpiy3QxjtkgQFSvBhK1tnzszub2ReNuPbXeAxReqHNWAJl/e
eRnI4oAkcAlmnlZumbJTUzihqtBfxnaffXnKFvS8qPZhAVpApISLvvZQMApyFRuMXnCoQfJNF3x20irHFKVFGn030W8fruzUItrHf2Zv+f/6Wp5pk0+eYEC3
9Py0KqWkXsIFhehrD/GIXNfiApWB3UZ9fNwaVBwYHtvW/rtsR0ip8gI/AsgBTg0ikGkqjaKsg5VZ9pDO6Z4kgcbzA4UuP0EmIHJxhKEwkuvSSxUitrKO6FKm
aoyWw9QaTZdlEjLUkbDwqCWo/g8VuUC8VCu/3aLY/8z+GLneEqGaAOWKoLe+WgU0mj8EQ4IWAZa+FRMukuXavMgRLRdalvIivhLqxGgJLRgSWGBUbSLx8Aj/
Re9HouLwcK47EEFjhQHPVGsCIg1diQm8EScHO0YYwuO4vgwBkfoukYCIgRsoEHQSOJ0ECp0ETidBhk4Cp5MgQyeB00mQoZPA+Uh4cX8FrMmxd2pRxgFMgqBa
CQ4gEnhEMtUrQtwqGPEibsAgGoYMjcNPqDHeWDKmt6etQ4cOHdz9iP8AKH548WEStmMAAAAASUVORK5CYII="""
)



def make_loadcfg_pixmap(size=24):

    """Trả về icon "Tải cấu hình" dạng QPixmap vuông kích thước `size`,
    scale mượt (SmoothTransformation) từ ảnh gốc 65x65."""

    raw = base64.b64decode(_LOADCFG_ICON_B64)

    src = QPixmap()

    src.loadFromData(raw, "PNG")

    if src.isNull():

        pix = QPixmap(size, size)

        pix.fill(Qt.transparent)

        return pix

    return src.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)



def make_loadcfg_icon(size=24):

    return QIcon(make_loadcfg_pixmap(size))





def make_notice_pixmap(kind="info", size=48):

    """Icon tròn cho hộp thoại thông báo, tông màu thống nhất với theme:

    info = xanh lá (đồng bộ logo), warning = vàng cam, error = đỏ, success = xanh đậm."""

    colors = {

        "info": ("#43a047", "i"),

        "success": ("#1b5e20", "✓"),

        "warning": ("#f9a825", "!"),

        "error": ("#c62828", "✕"),

    }

    color, glyph = colors.get(kind, colors["info"])

    pix = QPixmap(size, size)

    pix.fill(Qt.transparent)

    p = QPainter(pix)

    p.setRenderHint(QPainter.Antialiasing)

    p.setBrush(QBrush(QColor(color)))

    p.setPen(Qt.NoPen)

    p.drawEllipse(1, 1, size - 2, size - 2)

    p.setPen(QPen(Qt.white))

    font = QFont("Segoe UI", int(size * 0.42), QFont.Bold)

    p.setFont(font)

    p.drawText(QRectF(0, 0, size, size), Qt.AlignCenter, glyph)

    p.end()

    return pix





def make_exit_icon(size=40):

    """Vẽ icon nút "Thoát phần mềm" bằng QPainter (vector, không dùng emoji) theo

    kiểu icon "log-out/exit" chuẩn: khung cửa hé mở (rounded rect, chỉ 3 cạnh) +

    tay nắm cửa + mũi tên chỉ ra ngoài kèm vạch ngang, tông màu đỏ cam để nổi bật

    và dễ nhận biết là hành động "rời khỏi", tách biệt hẳn với tông xanh lá chủ

    đạo dùng cho các nút chức năng khác của TDForest."""

    pix = QPixmap(size, size)

    pix.fill(Qt.transparent)

    p = QPainter(pix)

    p.setRenderHint(QPainter.Antialiasing)

    p.setRenderHint(QPainter.SmoothPixmapTransform)



    color = QColor("#c0392b")

    pen = QPen(color)

    pen.setWidthF(size * 0.075)

    pen.setCapStyle(Qt.RoundCap)

    pen.setJoinStyle(Qt.RoundJoin)

    p.setPen(pen)

    p.setBrush(Qt.NoBrush)



    # Khung cửa hé mở: chỉ vẽ 3 cạnh (trái, trên, dưới), bỏ trống cạnh phải để

    # mũi tên "thoát ra" đi xuyên qua, tạo cảm giác cửa đang mở.

    door = QRectF(size * 0.17, size * 0.14, size * 0.38, size * 0.72)

    path_door = [

        QPointF(door.right(), door.top()),

        QPointF(door.left() + size * 0.08, door.top()),

        QPointF(door.left(), door.top() + size * 0.08),

        QPointF(door.left(), door.bottom() - size * 0.08),

        QPointF(door.left() + size * 0.08, door.bottom()),

        QPointF(door.right(), door.bottom()),

    ]

    for p1, p2 in zip(path_door[:-1], path_door[1:]):

        p.drawLine(p1, p2)



    # Tay nắm cửa - chấm tròn nhỏ

    p.setBrush(QBrush(color))

    p.setPen(Qt.NoPen)

    p.drawEllipse(QPointF(door.right() - size * 0.09, size / 2.0), size * 0.032, size * 0.032)



    # Mũi tên chỉ ra ngoài (phải): thân mũi tên + đầu mũi tên hình chữ V

    p.setPen(pen)

    p.setBrush(Qt.NoBrush)

    arrow_y = size / 2.0

    shaft_x0 = door.right() + size * 0.06

    shaft_x1 = size * 0.90

    p.drawLine(QPointF(shaft_x0, arrow_y), QPointF(shaft_x1, arrow_y))

    head = size * 0.16

    p.drawLine(QPointF(shaft_x1, arrow_y), QPointF(shaft_x1 - head, arrow_y - head))

    p.drawLine(QPointF(shaft_x1, arrow_y), QPointF(shaft_x1 - head, arrow_y + head))



    p.end()

    return pix





def make_exit_qicon(size=40):

    return QIcon(make_exit_icon(size))





def show_notice(parent, kind, title, text):

    """Thay cho QMessageBox.information/warning/critical mặc định của Qt -

    dùng icon tự vẽ đồng bộ theme thay vì icon hệ điều hành."""

    box = QMessageBox(parent)

    box.setWindowTitle(title)

    box.setText(text)

    box.setIconPixmap(make_notice_pixmap(kind, 48))

    box.setWindowIcon(make_tree_icon(32))

    box.exec()



# ==============================================================================

# 0-ter. CẤU HÌNH LƯU LẠI (QSettings) - dùng để nhớ đã hiện thông báo dùng thử

# hay chưa, giống cơ chế bên ULTS.py.

# ==============================================================================

CONFIG_ORG_NAME = "TDForest_Software"

CONFIG_APP_NAME = "TDForest_TracDoRung"





def get_settings() -> QSettings:

    """Trả về đối tượng QSettings dùng chung cho toàn bộ ứng dụng."""

    return QSettings(CONFIG_ORG_NAME, CONFIG_APP_NAME)





# ==============================================================================

# 0-quat. THÔNG TIN ADMIN + THÔNG BÁO HỆ THỐNG (bảng thongbaotracdo trên Postgres:
#   id, thongbao, phienbanmoinhat). Đã BỎ hoàn toàn key/mã khóa máy.

# ==============================================================================

MK_CONN_STR = (

    'postgres://avnadmin:AVNS_JZGTySXHq8YJ7UHuJJu'

    '@nicfi2025-gisplaines-79c8.k.aivencloud.com:18167'

    '/defaultdb?sslmode=require'

)

_ADMIN_EMAIL = "tvklinh@gmail.com"

_ADMIN_ZALO = "0364400660"

_ADMIN_INFO = f"  📧  Email : {_ADMIN_EMAIL}\n  📱  Zalo  : {_ADMIN_ZALO}"

# ------------------------------------------------------------------------------
# THỜI HẠN BẢN PHÁT HÀNH + PHẠM VI SỬ DỤNG (không còn key/mã khóa máy)
#   - Bản phát hành này dùng được 2 năm, tới _RELEASE_VALID_UNTIL thì phần mềm
#     chỉ nhắc "đã có phiên bản mới, liên hệ Admin" (không nhắc chữ hết hạn).
#   - Muốn gia hạn cho bản build sau: chỉ cần đổi ngày bên dưới rồi build lại.
# ------------------------------------------------------------------------------
_RELEASE_VALID_UNTIL = (2028, 11, 1)   # (năm, tháng, ngày) - 01/11/2028

_SCOPE_NOTICE = (
    "🌲  TDForest được phát triển dành riêng cho Sinh viên và Giảng viên "
    "Khoa Lâm nghiệp - Trường Đại học Nông Lâm TP. Hồ Chí Minh, nhằm phục vụ "
    "công tác học tập, giảng dạy và nghiên cứu khoa học.\n\n"
    "⚠️  Phần mềm không được phép sử dụng vào mục đích thương mại hay bất kỳ "
    "mục đích nào khác ngoài phạm vi nêu trên. Mọi hành vi sao chép, chia sẻ, "
    "phân phối lại hoặc chỉnh sửa phần mềm khi chưa có sự đồng ý của tác giả "
    "đều không được chấp thuận.\n\n"
    "💚  Trân trọng cảm ơn quý Thầy Cô và các bạn Sinh viên đã tin dùng TDForest!"
)

_UPDATE_REQUIRED_MSG = (
    "🎉  TDForest đã có phiên bản mới với nhiều cải tiến!\n\n"
    "Để tiếp tục sử dụng đầy đủ các chức năng, bạn vui lòng liên hệ Admin "
    "để nhận bản cập nhật:\n" + _ADMIN_INFO + "\n\n"
    "──────────────────────────────\n" + _SCOPE_NOTICE
)


def _release_needs_update() -> bool:
    """True nếu đã tới ngày _RELEASE_VALID_UNTIL -> cần nhận bản cập nhật mới.
    Chống chỉnh lùi đồng hồ: lưu lại ngày lớn nhất từng thấy trên máy (QSettings),
    nên dù người dùng chỉnh ngày hệ thống lùi về trước cũng không mở lại được."""
    from datetime import date
    today = date.today()
    settings = get_settings()
    try:
        last_seen = date.fromisoformat(str(settings.value("app/last_seen", "")))
    except ValueError:
        last_seen = today
    if last_seen > today:
        today = last_seen
    else:
        settings.setValue("app/last_seen", today.isoformat())
        settings.sync()
    return today >= date(*_RELEASE_VALID_UNTIL)



# Tên (khớp chính xác, không phân biệt hoa/thường) của bảng thông báo hệ thống +

# phiên bản mới dành riêng cho TDForest (khác "thongbaopmthongke" bên ULTS.py).

NOTICE_TABLE = "thongbaotracdo"



# Phiên bản hiện tại của phần mềm (tự tay tăng số này mỗi khi build bản mới). Dùng

# để so sánh với cột 'phienbanmoinhat' trong bảng thongbaotracdo nhằm tự động báo

# cho người dùng biết khi nào có bản cập nhật.

APP_VERSION = "1.0.0"





def _mk_parse_conn_str(conn_str: str) -> dict:

    """Phân tích chuỗi kết nối dạng postgres://user:pass@host:port/db?sslmode=..."""

    m = re.match(

        r'postgres://([^:]+):([^@]+)@([^:]+):(\d+)/([^?]+)\??(.*)',

        conn_str

    )

    if not m:

        raise ValueError("Chuỗi kết nối không hợp lệ.")

    user, password, host, port, dbname, opts = m.groups()

    params = dict(host=host, port=int(port),

                  dbname=dbname, user=user, password=password)

    for opt in opts.split('&'):

        if '=' in opt:

            k, v = opt.split('=', 1)

            if k == 'sslmode':

                params['sslmode'] = v

    return params





# ==============================================================================

#  THÔNG BÁO HỆ THỐNG & KIỂM TRA PHIÊN BẢN MỚI

#  (bảng: thongbaotracdo - dùng chung Postgres với bản quyền)

#

#  Cấu trúc bảng đang dùng (giống hệt thongbaopmthongke bên ULTS.py):

#

#     CREATE TABLE thongbaotracdo (

#         id               serial PRIMARY KEY,

#         thongbao         text,   -- để trống/NULL thì không hiện thông báo gì

#         phienbanmoinhat  text    -- vd '1.0.1' -- để trống thì không kiểm tra bản mới

#     );

#

#  Chỉ cần 1 dòng dữ liệu trong bảng (id lớn nhất sẽ được lấy).

#  Muốn ngừng thông báo -> để cột 'thongbao' rỗng ('' hoặc NULL) trên Postgres.

#  Muốn báo có bản mới -> điền 'phienbanmoinhat' khác với APP_VERSION hiện tại.

#  Việc dò tên bảng/cột dùng information_schema (không phân biệt hoa/thường) nên

#  đặt tên bảng kiểu gì trên Postgres cũng không lo lệch tên.

# ==============================================================================

class NotificationResult:

    __slots__ = ("ok", "thongbao", "phienbanmoinhat", "error_msg", "debug_info")



    def __init__(self):

        self.ok = False

        self.thongbao = ""

        self.phienbanmoinhat = ""

        self.error_msg = ""

        self.debug_info = ""





def _load_notification() -> NotificationResult:

    """Đọc thông báo hệ thống + phiên bản mới nhất từ bảng thongbaotracdo. Tự tra

    information_schema để tìm đúng tên bảng thật trên Postgres (không đoán mò hoa/

    thường), tránh lỗi 'relation does not exist'."""

    res = NotificationResult()

    try:

        import psycopg2

        from psycopg2 import sql

        params = _mk_parse_conn_str(MK_CONN_STR)

        conn = psycopg2.connect(connect_timeout=8, **params)

        conn.autocommit = True

        cur = conn.cursor()



        # 1) Tìm tên bảng thật (không phân biệt hoa/thường) đang có trên DB. Ưu tiên

        #    khớp CHÍNH XÁC NOTICE_TABLE ("thongbaotracdo") trước, để tránh lấy nhầm

        #    bảng khác nếu có nhiều bảng cùng bắt đầu bằng 'thongbao'.

        cur.execute(

            "SELECT table_name FROM information_schema.tables "

            "WHERE table_schema NOT IN ('pg_catalog', 'information_schema') "

            "AND table_name ILIKE 'thongbao%' "

            "ORDER BY (lower(table_name) = %s) DESC, table_name",

            (NOTICE_TABLE.lower(),)

        )

        rows = cur.fetchall()

        if not rows:

            cur.close(); conn.close()

            res.error_msg = (

                "Không tìm thấy bảng nào có tên bắt đầu bằng 'thongbao' "

                "trên Postgres (đã kiểm tra mọi schema)."

            )

            return res

        table_name = rows[0][0]  # tên thật (đúng hoa/thường) lấy từ information_schema

        res.debug_info = f"bảng khớp: {[r[0] for r in rows]}, chọn: {table_name}"



        # 2) Tìm đúng tên cột thật (không phân biệt hoa/thường) trong bảng đó

        cur.execute(

            "SELECT column_name FROM information_schema.columns "

            "WHERE table_name = %s",

            (table_name,)

        )

        cols = {c[0].lower(): c[0] for c in cur.fetchall()}

        col_thongbao = cols.get("thongbao")

        col_phienban = cols.get("phienbanmoinhat")

        col_id = cols.get("id")



        if not col_thongbao and not col_phienban:

            cur.close(); conn.close()

            res.error_msg = (

                f"Đã tìm thấy bảng '{table_name}' nhưng không có cột "

                f"'thongbao' hay 'phienbanmoinhat'. Các cột hiện có: "

                f"{', '.join(cols.values())}"

            )

            return res



        order_clause = sql.SQL("ORDER BY {} DESC".format(

            sql.Identifier(col_id).as_string(conn))) if col_id else sql.SQL("")



        if col_thongbao:

            q = sql.SQL("SELECT {} FROM {} WHERE {} IS NOT NULL AND {} <> '' {} LIMIT 1").format(

                sql.Identifier(col_thongbao), sql.Identifier(table_name),

                sql.Identifier(col_thongbao), sql.Identifier(col_thongbao), order_clause)

            cur.execute(q)

            r = cur.fetchone()

            res.thongbao = (r[0] or "").strip() if r else ""



        if col_phienban:

            q = sql.SQL("SELECT {} FROM {} WHERE {} IS NOT NULL AND {} <> '' {} LIMIT 1").format(

                sql.Identifier(col_phienban), sql.Identifier(table_name),

                sql.Identifier(col_phienban), sql.Identifier(col_phienban), order_clause)

            cur.execute(q)

            r = cur.fetchone()

            raw_val = r[0] if r else None

            res.phienbanmoinhat = (raw_val or "").strip()

            res.debug_info += f" | phienbanmoinhat thô = {raw_val!r}"



        cur.close()

        conn.close()

        res.ok = True

    except Exception as e:

        res.error_msg = str(e)

    return res





class NotificationWorker(QThread):

    """Chạy nền khi khởi động app để lấy thông báo + kiểm tra phiên bản mới, không

    chặn UI và không hiện lỗi nếu mất mạng / bảng chưa có dữ liệu."""

    finished_notice = Signal(object)



    def run(self):

        self.finished_notice.emit(_load_notification())





import matplotlib

matplotlib.use("QtAgg")

import matplotlib.colors as mcolors

import matplotlib.patches as mpatches

from matplotlib.patches import Polygon, Circle, Rectangle, Ellipse

from matplotlib.path import Path as MplPath

from matplotlib.figure import Figure

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas

from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar

# --- Import "cứng" các backend xuất file (PDF/SVG/PNG/PS) ---
# Khi gọi fig.savefig("xxx.pdf"), matplotlib tự động nạp động (lazy-import)
# module matplotlib.backends.backend_pdf theo phần đuôi file. Trên PyCharm/khi
# chạy bằng python thuần thì việc lazy-import này không vấn đề gì vì toàn bộ
# thư viện matplotlib đều nằm sẵn trên đĩa. Nhưng khi đóng gói bằng PyInstaller,
# công cụ phân tích import KHÔNG "nhìn thấy" các import kiểu lazy-import này
# (chúng nằm sâu trong matplotlib, không phải import tường minh ở đầu file),
# nên module backend_pdf (và backend_svg, backend_ps) bị bỏ sót, không được
# đóng gói vào file .exe -> khi bấm "Xuất PDF" sẽ báo lỗi:
#   No module named 'matplotlib.backends.backend_pdf'
# Cách khắc phục triệt để nhất: import tường minh các backend này ngay tại
# đây để PyInstaller "nhìn thấy" và tự động đóng gói theo.
import matplotlib.backends.backend_pdf   # noqa: F401  (bắt buộc để xuất PDF khi đóng gói exe)
import matplotlib.backends.backend_svg   # noqa: F401  (bắt buộc để xuất SVG khi đóng gói exe)
import matplotlib.backends.backend_ps    # noqa: F401  (dự phòng, một số máy cần khi xuất EPS/PS)

from mpl_toolkits.mplot3d import Axes3D  # noqa: F401  (kích hoạt projection '3d')





# ==============================================================================

# 1. HẰNG SỐ / CẤU HÌNH TRƯỜNG DỮ LIỆU

# ==============================================================================



# key nội bộ, nhãn hiển thị, từ khoá gợi ý để tự động ánh xạ cột

FIELD_DEFS = [

    ("TT",     "TT (Số thứ tự)",            ["tt", "stt", "sothutu", "so"]),

    ("Loai",   "Tên loài cây",              ["loai", "tenloai", "loaicay", "species", "ten"]),

    ("D13",    "D1.3 (cm)",                 ["d1.3", "d13", "d1,3", "duongkinh1.3", "dbh"]),

    ("Hvn",    "Hvn - Chiều cao vút ngọn (m)", ["hvn", "vutngon", "chieucaovutngon"]),

    ("Hdc",    "Hdc - Chiều cao dưới cành (m)", ["hdc", "duoicanh", "chieucaoduoicanh"]),

    ("DtanD",  "Dtán hướng Đông (m)",       ["dong", "dtand", "dtandong", "east"]),

    ("DtanT",  "Dtán hướng Tây (m)",        ["tay", "dtant", "dtantay", "west"]),

    ("DtanN",  "Dtán hướng Nam (m)",        ["nam", "dtann", "dtannam", "south"]),

    ("DtanB",  "Dtán hướng Bắc (m)",        ["bac", "dtanb", "dtanbac", "north"]),

    ("X",      "X - Toạ độ ngang (m)",      ["x"]),

    ("Y",      "Y - Toạ độ dọc (m)",        ["y"]),

]

FIELD_KEYS_NUMERIC = ["D13", "Hvn", "Hdc", "DtanD", "DtanT", "DtanN", "DtanB", "X", "Y"]

FIELD_KEYS_CRITICAL = ["X", "Y", "Hvn", "Hdc"]



# Giải thích chi tiết ý nghĩa từng cột (biến) dữ liệu - hiển thị trong hộp

# thoại "Giải thích các cột" để người dùng không chuyên (hoặc mới dùng phần

# mềm) hiểu rõ mỗi cột trong file Excel đầu vào tương ứng với đại lượng đo

# đếm nào ngoài thực địa, đơn vị gì và dùng để làm gì trong bản vẽ trắc đồ.

FIELD_EXPLANATIONS = {

    "TT": "Số thứ tự của cây trong ô tiêu chuẩn (OTC), dùng để đối chiếu "

          "ngược lại với biên bản đo đếm ngoài thực địa. Không bắt buộc - "

          "nếu bỏ trống, phần mềm sẽ tự đánh số lại theo thứ tự dòng.",

    "Loai": "Tên loài cây (tên phổ thông hoặc tên khoa học tuỳ bạn ghi). "

            "Dùng để tô màu phân biệt các loài trên trắc đồ và lập bảng "

            "chú dẫn (legend) số lượng cây theo loài.",

    "D13": "Đường kính thân cây đo tại vị trí ngang ngực (1,3 m tính từ mặt "

           "đất), đơn vị cm - chỉ tiêu lâm học cơ bản (D1.3 hay DBH). Hiện "

           "chưa dùng để vẽ hình dạng trên trắc đồ nhưng vẫn được lưu và "

           "hiển thị trong bảng dữ liệu để tham khảo, tính toán thêm.",

    "Hvn": "Chiều cao vút ngọn - khoảng cách từ gốc cây đến đỉnh ngọn cao "

           "nhất, đơn vị mét. Quyết định độ cao đỉnh tán trên trắc đồ dọc "

           "và mô hình 3D. Đây là cột BẮT BUỘC.",

    "Hdc": "Chiều cao dưới cành - khoảng cách từ gốc cây đến điểm bắt đầu "

           "có cành/tán lá (đáy tán), đơn vị mét. Cùng với Hvn xác định "

           "chiều cao và độ dày của khối tán. Đây là cột BẮT BUỘC.",

    "DtanD": "Bán kính (hoặc 1/2 đường kính) tán lá đo theo hướng Đông tính "

             "từ tâm thân cây, đơn vị mét - dùng để vẽ hình chiếu tán trên "

             "mặt bằng (trắc đồ ngang) theo 4 hướng thực đo, không giả định "

             "tán tròn đều.",

    "DtanT": "Bán kính tán lá đo theo hướng Tây tính từ tâm thân cây, đơn vị "

             "mét. Cùng nhóm với Dtán Đông/Nam/Bắc để dựng hình tán thực tế "

             "(có thể lệch tâm) trên trắc đồ ngang.",

    "DtanN": "Bán kính tán lá đo theo hướng Nam tính từ tâm thân cây, đơn vị "

             "mét. Cùng nhóm với Dtán Đông/Tây/Bắc để dựng hình tán thực tế "

             "trên trắc đồ ngang.",

    "DtanB": "Bán kính tán lá đo theo hướng Bắc tính từ tâm thân cây, đơn vị "

             "mét. Cùng nhóm với Dtán Đông/Tây/Nam để dựng hình tán thực tế "

             "trên trắc đồ ngang.",

    "X": "Toạ độ vị trí cây theo trục ngang (trục X) trong ô tiêu chuẩn, "

         "đơn vị mét - thường đo từ 1 góc mốc cố định của OTC. Quyết định "

         "vị trí cây trên trắc đồ ngang và mô hình 3D. Đây là cột BẮT BUỘC.",

    "Y": "Toạ độ vị trí cây theo trục dọc (trục Y) trong ô tiêu chuẩn, đơn "

         "vị mét - thường đo từ cùng góc mốc với X. Quyết định vị trí cây "

         "trên trắc đồ ngang, đồng thời dùng làm trục chiếu cho trắc đồ dọc "

         "(mặt cắt đứng). Đây là cột BẮT BUỘC.",

}



# ==============================================================================

# 1-bis. DỮ LIỆU DEMO TÍCH HỢP SẴN (không cần file Excel ngoài)

# ------------------------------------------------------------------------------

# Trích 1 ô tiêu chuẩn (OTC) rừng mẫu - 40 cây / 8 loài - để người dùng dùng

# thử ngay toàn bộ chức năng phần mềm (bảng dữ liệu, trắc đồ ngang, trắc đồ

# dọc, mô phỏng 3D) mà KHÔNG cần chuẩn bị file Excel trước. Tên cột đặt giống

# hệt định dạng chuẩn nên được ánh xạ tự động 1-1, không cần qua hộp thoại

# "Ánh xạ cột".

# ==============================================================================

DEMO_COLUMNS = {

    "TT": "TT", "Loai": "Tên loài cây", "D13": "D1,3 (cm)",

    "Hvn": "Hvn (m)", "Hdc": "Hdc (m)",

    "DtanD": "Dtán Đ (m)", "DtanT": "Dtán T (m)",

    "DtanN": "Dtán N (m)", "DtanB": "Dtán B (m)",

    "X": "X (m)", "Y": "Y (m)",

}



# TT, Loài, D1.3, Hvn, Hdc, DtánĐ, DtánT, DtánN, DtánB, X, Y

DEMO_ROWS = [

    (1, "Dầu rái", 24.8, 16.6, 8.8, 1.66, 1.19, 1.50, 1.50, 13.4, 19.8),

    (2, "Sao đen", 10.5, 6.0, 3.5, 0.70, 0.50, 1.02, 0.85, 4.9, 13.7),

    (3, "Trâm", 28.7, 13.6, 7.3, 2.38, 1.55, 0.99, 1.41, 7.5, 20.6),

    (4, "Chò chỉ", 8.0, 6.0, 2.5, 0.50, 0.54, 0.50, 1.38, 9.2, 17.8),

    (5, "Sao đen", 22.9, 15.8, 7.2, 2.43, 1.64, 1.32, 2.29, 18.7, 16.9),

    (6, "Dầu rái", 30.4, 18.3, 7.5, 1.12, 1.63, 1.01, 0.50, 14.6, 12.2),

    (7, "Dầu rái", 15.0, 6.0, 2.6, 0.50, 1.40, 0.69, 0.84, 9.3, 19.2),

    (8, "Gõ đỏ", 42.4, 25.4, 14.8, 2.71, 2.07, 3.05, 3.14, 3.3, 16.3),

    (9, "Dầu rái", 39.1, 20.5, 9.4, 3.23, 3.18, 3.23, 2.95, 9.3, 1.8),

    (10, "Bằng lăng", 27.2, 17.3, 10.1, 1.65, 1.60, 1.44, 1.94, 4.8, 14.3),

    (11, "Dầu rái", 32.8, 18.8, 6.7, 2.60, 3.23, 2.36, 2.42, 13.1, 19.9),

    (12, "Bằng lăng", 31.0, 18.8, 6.8, 1.62, 1.44, 1.36, 2.31, 3.1, 8.5),

    (13, "Bằng lăng", 43.3, 23.4, 13.2, 3.10, 2.74, 2.84, 2.78, 7.4, 10.5),

    (14, "Dầu rái", 30.1, 22.0, 9.4, 2.17, 2.66, 2.25, 2.01, 7.7, 5.5),

    (15, "Trâm", 15.0, 7.0, 3.2, 0.50, 0.83, 1.19, 1.35, 14.8, 2.0),

    (16, "Chò chỉ", 45.1, 25.1, 10.6, 3.42, 3.33, 2.57, 3.10, 19.0, 23.4),

    (17, "Sao đen", 23.1, 12.9, 7.7, 2.02, 1.30, 1.81, 1.15, 11.8, 2.7),

    (18, "Gõ đỏ", 19.9, 6.2, 2.4, 1.18, 1.02, 0.50, 1.07, 15.4, 4.0),

    (19, "Gõ đỏ", 11.3, 7.5, 3.6, 1.09, 0.70, 1.26, 0.91, 6.4, 18.6),

    (20, "Bằng lăng", 20.5, 13.2, 4.8, 0.50, 1.44, 1.57, 1.25, 11.9, 8.6),

    (21, "Trâm", 20.3, 13.9, 6.1, 1.39, 1.19, 1.40, 0.90, 5.2, 23.7),

    (22, "Sao đen", 26.3, 13.5, 7.7, 1.43, 1.01, 1.01, 1.26, 18.2, 9.1),

    (23, "Chò chỉ", 19.0, 10.6, 4.0, 0.96, 0.64, 1.06, 1.33, 18.3, 11.8),

    (24, "Dầu rái", 20.6, 12.9, 7.4, 0.96, 1.27, 1.26, 0.66, 17.0, 21.4),

    (25, "Chò chỉ", 25.1, 12.6, 6.0, 2.02, 2.31, 1.55, 2.53, 4.2, 11.9),

    (26, "Dầu rái", 22.1, 13.5, 6.5, 1.01, 0.80, 1.19, 1.45, 13.0, 18.3),

    (27, "Dầu rái", 29.6, 13.9, 5.8, 2.39, 2.15, 2.48, 2.16, 15.8, 12.4),

    (28, "Dầu rái", 34.4, 20.5, 10.8, 2.08, 2.14, 1.39, 2.67, 7.3, 11.5),

    (29, "Dầu rái", 27.9, 15.0, 6.3, 1.94, 2.17, 1.54, 1.44, 9.4, 23.8),

    (30, "Sao đen", 36.9, 21.9, 9.6, 3.15, 2.41, 3.08, 2.17, 5.7, 22.8),

    (31, "Bằng lăng", 24.2, 10.2, 5.9, 1.98, 0.93, 0.83, 1.16, 11.3, 19.8),

    (32, "Trâm", 16.0, 8.7, 4.3, 0.50, 0.92, 1.06, 0.89, 17.9, 23.2),

    (33, "Trâm", 25.8, 15.0, 6.4, 1.77, 1.13, 2.26, 2.10, 16.1, 3.8),

    (34, "Dầu rái", 25.6, 11.6, 5.7, 1.46, 0.73, 1.27, 1.25, 8.3, 16.0),

    (35, "Sao đen", 20.2, 8.7, 3.5, 1.69, 2.05, 1.09, 2.44, 7.5, 12.2),

    (36, "Trâm", 47.8, 27.4, 10.5, 2.15, 3.10, 2.99, 2.75, 1.9, 6.1),

    (37, "Chò chỉ", 39.4, 23.8, 8.8, 1.79, 2.13, 2.31, 1.88, 5.7, 14.5),

    (38, "Chò chỉ", 8.0, 7.1, 2.7, 0.79, 0.50, 0.50, 1.27, 18.8, 12.1),

    (39, "Sao đen", 25.7, 9.4, 4.4, 1.45, 1.68, 1.47, 1.58, 18.2, 21.2),

    (40, "Sao đen", 35.9, 20.2, 8.7, 1.80, 1.98, 2.40, 2.22, 18.1, 11.7),

]



DEMO_MARKER = "__TDFOREST_DEMO__"

# Định danh giả cho mục "dữ liệu vừa TẢI TỪ FILE CẤU HÌNH (.geojson)" trong
# combo "File Excel đang xem" - dùng chung cơ chế với DEMO_MARKER ở trên,
# nhưng đại diện cho dữ liệu KHÔNG gắn với file Excel nào trên đĩa (đã được
# nhúng sẵn trong file .geojson nên vẫn xem lại được dù file Excel gốc mất).
CONFIG_MARKER = "__TDFOREST_CONFIG__"

# Số hiệu định dạng file cấu hình .geojson - tăng lên nếu sau này thay đổi
# cấu trúc lưu (để có thể xử lý tương thích ngược khi cần).
TDFOREST_CONFIG_VERSION = 1


def _json_safe_default(obj):
    """Hàm default cho json.dump/json.dumps - xử lý các kiểu dữ liệu numpy
    (numpy.float64, numpy.int64...) hay lọt vào self.style/DataFrame mà
    bản thân module json không tự serialize được."""
    if isinstance(obj, np.integer):
        return int(obj)
    if isinstance(obj, np.floating):
        v = float(obj)
        return None if np.isnan(v) else v
    if isinstance(obj, np.bool_):
        return bool(obj)
    if isinstance(obj, np.ndarray):
        return obj.tolist()
    if isinstance(obj, (set, tuple)):
        return list(obj)
    return str(obj)







def get_demo_dataframe():

    """Trả về DataFrame dữ liệu Demo, tên cột giống hệt định dạng chuẩn."""

    cols = list(DEMO_COLUMNS.values())

    return pd.DataFrame(DEMO_ROWS, columns=cols)





# Hình chiếu tán trên mặt bằng (trắc đồ ngang) - vẽ theo 4 bán kính đo thực địa

PLAN_SHAPE_OPTIONS = [

    "Elip 4 hướng (mượt, Đông-Tây-Nam-Bắc)",

    "Tán tròn đều (Circle)",

]



# Dạng khối tán theo mặt cắt đứng / 3D - ứng với từng trạng thái rừng khác

# nhau (rừng lá rộng, rừng lá kim, rừng trồng thuần loài, rừng già tầng trên,

# tán loe...). Mỗi dạng dùng chung 1 "hàm bán kính chuẩn hoá" r(u) với u chạy

# từ 0 (đáy tán, ngang Hdc) đến 1 (đỉnh tán, ngang Hvn) - áp dụng cho cả vẽ

# mặt cắt dọc 2D lẫn khối tròn xoay 3D.

CROWN_FORM_OPTIONS = [

    "Tán tròn/elip (Oval - phổ biến, cây lá rộng)",

    "Hình cầu tròn đều (Spherical - tán rậm hình cầu, cây gỗ lớn đơn lẻ)",

    "Hình nón (Conifer - thông, sa mộc, lá kim)",

    "Hình trụ (Columnar - rừng trồng thuần loài, mật độ cao)",

    "Hình dù/bán cầu (Dome - cây tầng trên, rừng già)",

    "Hình tán bàn/xoè phẳng đỉnh (Flat-top - cây vượt tán, cổ thụ)",

    "Hình trứng ngược (Obovate - tán loe rộng phía trên)",

    "Hình phễu ngược (Inverted funnel - đáy hẹp, đỉnh loe thẳng)",

    "Hình trứng xuôi (Ovoid - phình dưới, thuôn nhỏ dần lên đỉnh)",

    "Hình thoi/con quay (Diamond - phình giữa, nhọn 2 đầu)",

    "Hình sao nhiều mũi (Star - mặt cắt thuỳ, sống động khi xoay 3D)",

    "Tán tự nhiên gợn sóng (Irregular - bề mặt không đều, chân thực)",

    "Hình đa tầng (Layered - tán phân tầng như thông cảnh quan)",

    "Hình xoắn ốc (Spiral - vân xoắn quanh thân, ấn tượng khi xoay 3D)",

    "Hình cầu gai nhiều thuỳ (Spiky sphere - thuỳ lồi lõm sâu, sống động)",

    "Hình pha lê đa giác (Faceted/Low-poly - khối nhiều mặt, phong cách hiện đại)",

]



# giữ tên cũ để không phá vỡ các chỗ khác lỡ còn tham chiếu

SHAPE_OPTIONS = PLAN_SHAPE_OPTIONS



COLOR_MODES = [

    "Theo loài cây",

    "Một màu đồng nhất",

    "Theo chiều cao Hvn (gradient)",

]



# Tỷ lệ bản vẽ so với thực tế - dùng để hiển thị chú thích tỷ lệ trên trắc đồ

# và tính kích thước hình khi xuất ảnh (in đúng tỷ lệ 1:N ở chế độ in "Actual

# size / 100%", không "Fit to page").

SCALE_OPTIONS = ["Tự động (vừa khung)", "1:100", "1:200", "1:500", "1:1000", "1:2000", "Tuỳ chỉnh..."]

# Khổ giấy tiêu chuẩn (mm) theo chiều ĐỨNG (portrait): (rộng, cao). Dùng khi
# xuất trắc đồ ra file in đúng khổ - đúng tỷ lệ 1:N (xem export_paper_size).
PAPER_SIZES_MM = {
    "A2": (420.0, 594.0),
    "A3": (297.0, 420.0),
    "A4": (210.0, 297.0),
    "A5": (148.0, 210.0),
}

# Các mẫu số tỷ lệ "đẹp" thường dùng trong bản đồ/trắc đồ lâm nghiệp - dùng
# để tự động chọn tỷ lệ vừa khít khổ giấy khi người dùng chọn "Tự động".
STANDARD_SCALE_DENOMS = [50, 75, 100, 150, 200, 250, 500, 750, 1000, 1500, 2000, 2500, 5000]

# Lề dành RIÊNG cho SỐ + TÊN TRỤC X/Y khi xuất trắc đồ ra khổ giấy chuẩn (cm).
# Nhãn trục (tick label + axis label như "X (m)", "Y (m)", "Hvn (m)") luôn được
# matplotlib vẽ ra NGOÀI khung vẽ (axes), KHÔNG nằm trong khung nội dung/khung
# tỷ lệ. Nếu không trừ riêng phần lề này, ở chế độ "Phóng/thu vừa khổ giấy"
# (không theo tỷ lệ chuẩn) khung nội dung sẽ được đặt khít sát mép vùng in,
# đẩy số/tên trục ra ngoài khổ giấy -> bị cắt/che khi in (xem export_paper_size
# và _render_paper_export).
AXIS_LABEL_MARGIN_LEFT_CM = 1.5
AXIS_LABEL_MARGIN_BOTTOM_CM = 1.2
AXIS_LABEL_MARGIN_RIGHT_CM = 0.35
AXIS_LABEL_MARGIN_TOP_CM = 0.35




FONT_FAMILY_OPTIONS = ["Segoe UI", "Arial", "Times New Roman", "Calibri", "DejaVu Sans", "Verdana", "Tahoma"]



# Danh sách các biểu đồ con trong tab "Phân tích cấu trúc rừng" cho phép

# người dùng tự đặt Tiêu đề / nhãn trục X-Y / Màu sắc / Kiểu chữ riêng (xem

# _build_struct_chart_style_group). has_xlabel=False / has_color=False nghĩa

# là biểu đồ đó không có trục X mang ý nghĩa dữ liệu / không có 1 màu chủ

# đạo duy nhất (VD: sơ đồ phân tầng tô theo 3 màu tầng cố định) nên ẩn field

# tương ứng đi cho đỡ rối.

STRUCT_CHART_OPTIONS = [

    # key,     nhãn hiển thị,                                có trục X?, có màu chủ đạo?, có chú dẫn/hộp thông tin?

    ("nd",     "Phân bố N/D (số cây theo cấp kính)",      True,  True,  False),

    ("ndg",    "Phân bố N/Dg (cấp kính tương đối)",        True,  True,  False),

    ("g",      "Phân bố tiết diện ngang (G)",              True,  True,  False),

    ("cover",  "Độ che phủ tán",                           True,  True,  True),

    ("nh",     "Phân bố N/H (số cây theo cấp cao)",        True,  True,  False),

    ("layers", "Sơ đồ phân tầng đứng",                     False, False, True),

]



# Bảng màu tán lá - mở rộng ~40 sắc thái phân biệt rõ (xanh lá/rêu/lam/nâu/

# tím/cam/hồng...) để đủ dùng cho ô tiêu chuẩn có nhiều loài cây khác nhau.

# Nếu số loài vượt quá bảng này, generate_extra_color() sẽ tự sinh thêm màu

# phân biệt bằng cách xoay đều vòng tròn màu (hue) theo tỉ lệ vàng.

SPECIES_PALETTE = [

    "#2e7d32", "#6a994e", "#1b6b52", "#3d5a80", "#588157",

    "#7d5ba6", "#a5714e", "#457b9d", "#606c38", "#8d6a9f",

    "#40916c", "#bc6c25", "#2f6690", "#9d4edd", "#3a5a40",

    "#c9184a", "#ee9b00", "#0a9396", "#5f0f40", "#3a86ff",

    "#ff595e", "#8ac926", "#1982c4", "#6a4c93", "#ffca3a",

    "#005f73", "#9b2226", "#4a5759", "#e56b6f", "#355070",

    "#b56576", "#6d597a", "#84a98c", "#f4a261", "#264653",

    "#e76f51", "#2a9d8f", "#e9c46a", "#7209b7", "#606c38",

]



# Giao diện sáng - chủ đạo xanh rừng + nền kem nhạt

LIGHT_QSS = """

QMainWindow, QDialog { background-color: #f3f6f1; }

QWidget { font-family: 'Segoe UI', 'Calibri', Arial; font-size: 10pt; color: #2b3a2f; }

QGroupBox {

    background-color: #ffffff;

    border: 1px solid #dfe7dc;

    border-radius: 10px;

    margin-top: 16px;

    padding: 12px 10px 10px 10px;

    font-weight: 600;

    color: #1b5e20;

}

QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; }

QPushButton {

    background-color: #2e7d32; color: white; border: none;

    border-radius: 7px; padding: 7px 14px; font-weight: 500;

}

QPushButton:hover { background-color: #388e3c; }

QPushButton:pressed { background-color: #1b5e20; }

QPushButton:disabled { background-color: #b7c4b4; color: #eef2ec; }

QPushButton#secondary {

    background-color: #ffffff; color: #2e7d32; border: 1.4px solid #2e7d32;

}

QPushButton#secondary:hover { background-color: #eaf5ea; }

QPushButton#danger {

    background-color: #fff0ef; color: #c62828; border: 1.4px solid #c62828;
    font-weight: 600;

}

QPushButton#danger:hover { background-color: #ffd9d6; border-color: #b71c1c; }

QPushButton#danger:pressed { background-color: #f4b8b3; }

QComboBox, QDoubleSpinBox, QLineEdit {

    background-color: #ffffff; border: 1px solid #cfd9cc;

    border-radius: 6px; padding: 4px 8px; min-height: 22px;

}

QComboBox::drop-down { border: none; width: 22px; }

QTabWidget::pane { border: 1px solid #dfe7dc; border-radius: 10px; background: #ffffff; top: -1px; }

QTabBar::tab {

    background: #e7efe4; padding: 9px 20px; margin-right: 3px;

    border-top-left-radius: 9px; border-top-right-radius: 9px; color: #4b5c4d;

}

QTabBar::tab:selected { background: #2e7d32; color: white; font-weight: 600; }

QTabBar::tab:hover:!selected { background: #d3e4cf; }

QTableView {

    background: white; gridline-color: #eef1ee; alternate-background-color: #f7faf6;

    selection-background-color: #c8e6c9; selection-color: #1b3a1e;

    border: 1px solid #dfe7dc; border-radius: 6px;

}

QHeaderView::section {

    background-color: #eaf3e8; padding: 6px; border: none;

    border-bottom: 2px solid #2e7d32; font-weight: 600; color: #23522a;

}

QListWidget { background: white; border: 1px solid #dfe7dc; border-radius: 6px; }

QListWidget::item { padding: 3px 4px; }

QCheckBox, QLabel { color: #33422f; }

QSlider::groove:horizontal { height: 5px; background: #dfe8de; border-radius: 2px; }

QSlider::handle:horizontal { background: #2e7d32; width: 15px; margin: -6px 0; border-radius: 8px; }

QSlider::handle:horizontal:hover { background: #1b5e20; }

QStatusBar { background: #eaf3e8; color: #23522a; border-top: 1px solid #dfe7dc; }

QToolBar { background: #ffffff; border-bottom: 1px solid #dfe7dc; spacing: 8px; padding: 6px; }

QToolBar QToolButton {

    background: #ffffff; color: #1b5e20; border: 1px solid #cfe0cc;

    border-radius: 7px; padding: 6px 10px; font-weight: 600;

}

QToolBar QToolButton:hover { background: #eaf5ea; }

QScrollArea { border: none; }

QSplitter::handle { background: #e3ece0; }
QSpinBox, QDateEdit {
    background-color: #ffffff; border: 1px solid #cfd9cc;
    border-radius: 6px; padding: 4px 8px; min-height: 22px;
}
QComboBox:hover, QDoubleSpinBox:hover, QSpinBox:hover, QDateEdit:hover, QLineEdit:hover { border-color: #9fc29b; }
QComboBox:focus, QDoubleSpinBox:focus, QSpinBox:focus, QDateEdit:focus, QLineEdit:focus,
QPlainTextEdit:focus, QTextEdit:focus { border: 1.5px solid #2e7d32; }
QPlainTextEdit, QTextEdit {
    background: #ffffff; border: 1px solid #cfd9cc; border-radius: 6px; padding: 4px;
}
QComboBox QAbstractItemView {
    background: #ffffff; border: 1px solid #cfd9cc; outline: 0;
    selection-background-color: #c8e6c9; selection-color: #1b3a1e;
}
QScrollBar:vertical { background: transparent; width: 11px; margin: 2px; }
QScrollBar::handle:vertical { background: #c9d6c6; border-radius: 4px; min-height: 32px; }
QScrollBar:horizontal { background: transparent; height: 11px; margin: 2px; }
QScrollBar::handle:horizontal { background: #c9d6c6; border-radius: 4px; min-width: 32px; }
QScrollBar::handle:hover { background: #9fb79b; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
QScrollBar::add-page, QScrollBar::sub-page { background: none; }
QToolTip {
    background-color: #1b3a1e; color: #ffffff; border: none;
    padding: 6px 9px; border-radius: 5px; font-size: 9pt;
}
QMenu { background: #ffffff; border: 1px solid #dfe7dc; padding: 4px; }
QMenu::item { padding: 6px 22px; border-radius: 4px; }
QMenu::item:selected { background: #eaf5ea; color: #1b5e20; }
QToolBar::separator { background: #e3ece0; width: 1px; margin: 4px 6px; }
QToolBar QToolButton#primaryAction {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2e7d32, stop:1 #1b5e20);
    color: #ffffff; border: none;
}
QToolBar QToolButton#primaryAction:hover { background: #388e3c; }
QTabBar::tab:disabled { color: #a9b5a7; background: #eef1ed; }
QTabWidget#subTabs::pane { border: 1px solid #e3ebe0; border-radius: 8px; top: -1px; }
QTabWidget#subTabs QTabBar::tab {
    background: transparent; color: #4b5c4d; padding: 7px 14px; margin-right: 2px;
    border: none; border-bottom: 3px solid transparent; border-radius: 0;
}
QTabWidget#subTabs QTabBar::tab:selected {
    color: #1b5e20; border-bottom: 3px solid #2e7d32; font-weight: 600; background: transparent;
}
QTabWidget#subTabs QTabBar::tab:hover:!selected { background: #eef5ec; }
QFrame#kpiCard {
    background: #ffffff; border: 1px solid #dfe7dc; border-left: 4px solid #2e7d32; border-radius: 8px;
}
QLabel#kpiValue { font-size: 16pt; font-weight: 700; color: #1b5e20; background: transparent; }
QLabel#kpiLabel { font-size: 8.5pt; color: #66756a; background: transparent; }
QLabel#formulaBox {
    background: #e8f5e9; border: 1px solid #cfe5cb; border-left: 4px solid #2e7d32;
    border-radius: 8px; padding: 10px 14px; font-size: 10.5pt; color: #1b5e20;
}
QLabel#hintLabel { color: #66756a; font-size: 8.8pt; }
QFrame#dialogHeader {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0f3d1e, stop:1 #2e7d32);
}
QLabel#dialogHeaderTitle { color: #ffffff; font-size: 14pt; font-weight: 700; background: transparent; }
QLabel#dialogHeaderSub { color: #cfe8cf; font-size: 9pt; background: transparent; }
QPushButton#primaryBig { padding: 9px 22px; font-weight: 700; font-size: 10.5pt; }
QStatusBar::item { border: none; }
"""





# ==============================================================================

# 2. TIỆN ÍCH CHUNG

# ==============================================================================



def strip_accents(text):

    """Bỏ dấu tiếng Việt để so khớp gần đúng tên cột."""

    text = str(text)

    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))





def normalize_key(text):

    return strip_accents(text).lower().replace(" ", "").replace("_", "").replace(".", "").replace(",", "")





def generate_extra_color(i):

    """Sinh thêm 1 màu phân biệt (ngoài SPECIES_PALETTE) bằng cách xoay hue

    theo tỉ lệ vàng - đảm bảo các màu sinh thêm luôn tách biệt rõ với nhau

    dù số loài rất lớn."""

    golden = 0.6180339887498949

    hue = (i * golden) % 1.0

    sat = 0.55 + 0.20 * (((i * 3) % 4) / 4.0)

    val = 0.62 + 0.28 * (((i * 5) % 3) / 3.0)

    return mcolors.to_hex(mcolors.hsv_to_rgb((hue, sat, val)))





def build_species_colors(species_list, overrides=None):

    """Gán màu ổn định cho từng loài cây theo thứ tự bảng chữ cái, ưu tiên

    dùng bảng màu dựng sẵn (đa dạng), sinh thêm màu nếu vượt quá bảng, và áp

    dụng màu người dùng tự chọn riêng (overrides) nếu có."""

    unique = sorted(set(str(s) for s in species_list))

    colors = {}

    for i, sp in enumerate(unique):

        if i < len(SPECIES_PALETTE):

            colors[sp] = SPECIES_PALETTE[i]

        else:

            colors[sp] = generate_extra_color(i - len(SPECIES_PALETTE))

    if overrides:

        for sp, hexcol in overrides.items():

            if sp in colors:

                colors[sp] = hexcol

    return colors





def darker(hex_color, factor=0.68):

    rgb = np.array(mcolors.to_rgb(hex_color))

    return mcolors.to_hex(np.clip(rgb * factor, 0, 1))





def lighter(hex_color, factor=0.5):

    """Pha màu về phía trắng theo hệ số factor (0 = giữ nguyên, 1 = trắng

    hoàn toàn) - dùng cho phần đỉnh tán khi tô gradient theo độ cao."""

    rgb = np.array(mcolors.to_rgb(hex_color))

    return mcolors.to_hex(np.clip(rgb + (1.0 - rgb) * factor, 0, 1))





def crown_facecolors(color, U, alpha, top_lighten=0.28, bottom_darken=0.45):

    """Sinh mảng RGBA cho từng ô lưới của mặt tán 3D theo độ cao chuẩn hoá U

    (0 = đáy tán/Hdc, 1 = đỉnh tán/Hvn). Blend theo không gian HSV thay vì

    RGB: chỉ thay đổi độ sáng (Value) theo chiều cao, GIỮ NGUYÊN sắc (Hue)

    và tăng nhẹ độ bão hoà (Saturation) ở đáy - tán vẫn có khối/chiều sâu

    (đỉnh sáng như có nắng chiếu, đáy tối như vùng khuất) nhưng không bị

    "bạc màu"/ngả xám như cách pha thẳng về trắng trước đây."""

    h, s, v = mcolors.rgb_to_hsv(np.array(mcolors.to_rgb(color)))

    t = np.clip(U, 0.0, 1.0)

    v_arr = v * (1.0 - bottom_darken) + v * bottom_darken * t + top_lighten * t * (1.0 - v)

    v_arr = np.clip(v_arr, 0.0, 1.0)

    s_arr = np.clip(s * (1.05 - 0.15 * t), 0.0, 1.0)

    h_arr = np.full_like(v_arr, h)

    hsv = np.stack([h_arr, s_arr, v_arr], axis=-1)

    rgb = mcolors.hsv_to_rgb(hsv)

    rgba = np.concatenate([rgb, np.full(rgb.shape[:-1] + (1,), alpha)], axis=-1)

    return rgba





def get_cmap(name):

    try:

        return matplotlib.colormaps.get_cmap(name)

    except Exception:

        return matplotlib.cm.get_cmap(name)





def crown_polygon(cx, cy, rE, rN, rW, rS, n=28):

    """

    Tạo đa giác tán cây MƯỢT (đóng kín) từ 4 bán kính đo theo 4 hướng

    Đông / Bắc / Tây / Nam, dùng nội suy Catmull-Rom tuần hoàn qua 4 điểm

    bán kính đó (theo goc cực). Khác với cách nối bằng 4 cung 1/4 ellipse

    riêng biệt (vốn bị "gãy" độ cong tại đúng 4 hướng chính), Catmull-Rom

    tuần hoàn cho đường cong liên tục cả về vị trí lẫn độ dốc (C1) khi đi

    qua đường Đông -> Bắc -> Tây -> Nam -> Đông, nhìn tự nhiên hơn nhiều.

    """

    rE = max(float(rE), 0.05); rN = max(float(rN), 0.05)

    rW = max(float(rW), 0.05); rS = max(float(rS), 0.05)

    radii = [rE, rN, rW, rS]          # tại góc 0°, 90°, 180°, 270°

    m = 4

    pts = []

    for i in range(m):

        p0 = radii[(i - 1) % m]

        p1 = radii[i]

        p2 = radii[(i + 1) % m]

        p3 = radii[(i + 2) % m]

        for t in np.linspace(0.0, 1.0, n, endpoint=False):

            t2 = t * t; t3 = t2 * t

            r = 0.5 * ((2 * p1) + (-p0 + p2) * t

                       + (2 * p0 - 5 * p1 + 4 * p2 - p3) * t2

                       + (-p0 + 3 * p1 - 3 * p2 + p3) * t3)

            r = max(r, 0.02)

            theta = np.radians(90.0 * i + 90.0 * t)

            pts.append((cx + r * np.cos(theta), cy + r * np.sin(theta)))

    return np.array(pts)





# ---- Dạng khối tán (mặt cắt dọc / 3D) - hàm bán kính chuẩn hoá r(u) ----

# u: 0 = đáy tán (ngang Hdc)  ->  1 = đỉnh tán (ngang Hvn). Trả về hệ số

# nhân với bán kính lớn nhất của tán (0..1).



def _crown_radius_profile(u, form):

    u = np.clip(np.asarray(u, dtype=float), 0.0, 1.0)

    if form.startswith("Hình nón"):

        return 1.0 - u

    if form.startswith("Hình trụ"):

        edge = 0.14

        r = np.ones_like(u)

        lo = u < edge

        hi = u > (1 - edge)

        r[lo] = np.sin(np.clip(u[lo] / edge, 0, 1) * np.pi / 2.0)

        r[hi] = np.sin(np.clip((1 - u[hi]) / edge, 0, 1) * np.pi / 2.0)

        return r

    if form.startswith("Hình dù"):

        return np.sqrt(np.clip(1.0 - u ** 2, 0.0, 1.0))

    if form.startswith("Hình trứng ngược"):

        return np.sqrt(u)

    if form.startswith("Hình cầu"):

        # Cung tròn (không phải hình sin) -> tán bầu bo tròn đều, "đầy đặn"

        # hơn dạng Oval một chút quanh phần giữa tán.

        return np.sqrt(np.clip(1.0 - (2.0 * u - 1.0) ** 2, 0.0, 1.0))

    if form.startswith("Hình tán bàn"):

        # Giữ bán kính tối đa (mặt bàn phẳng) tới ~72% chiều cao tán, rồi

        # bo tròn xuống 0 ở đỉnh -> mô phỏng tán cây tầng vượt tán xoè rộng,

        # phẳng đỉnh, thường gặp ở cây cổ thụ / cây tầng trên rừng già.

        plateau = 0.72

        r = np.ones_like(u)

        lo = u < 0.10

        r[lo] = np.sin(np.clip(u[lo] / 0.10, 0.0, 1.0) * np.pi / 2.0)

        hi = u > plateau

        r[hi] = np.cos(np.clip((u[hi] - plateau) / (1.0 - plateau), 0.0, 1.0) * np.pi / 2.0)

        return r

    if form.startswith("Hình phễu ngược"):

        # Bán kính tăng tuyến tính từ đáy (hẹp) lên đỉnh (loe rộng).

        return np.clip(u, 0.02, 1.0)

    if form.startswith("Hình trứng xuôi"):

        # Ngược với "trứng ngược": phình to ở đáy tán, thuôn nhỏ dần lên đỉnh.

        return np.sqrt(np.clip(1.0 - u, 0.0, 1.0))

    if form.startswith("Hình thoi"):

        # Bán kính tăng tuyến tính tới giữa tán rồi giảm tuyến tính về đỉnh

        # -> hình thoi/con quay, phình bụng, nhọn 2 đầu, dáng góc cạnh rõ.

        return 1.0 - np.abs(2.0 * u - 1.0)

    if form.startswith("Hình đa tầng"):

        # Bao hình thu nhỏ dần lên đỉnh (như cây thông) + gợn sóng tuần hoàn

        # theo u để tạo các "tầng" tán tách bạch rõ ràng (mỗi tầng phình ra

        # rồi thu hẹp lại trước khi sang tầng kế tiếp).

        n_tang = 4

        envelope = 1.0 - 0.55 * u

        pha = (u * n_tang) % 1.0

        tang = np.sin(pha * np.pi) ** 0.8

        r = envelope * (0.32 + 0.68 * tang)

        r[u < 0.03] = r[u < 0.03] * (u[u < 0.03] / 0.03)

        return np.clip(r, 0.0, 1.0)

    if form.startswith("Hình cầu gai"):

        # Bao hình cầu tròn đều làm nền cho các thuỳ gai (thuỳ được tạo ở

        # _crown_azimuthal_modulation), giữ mặt cắt dọc đầy đặn như hình cầu.

        return np.sqrt(np.clip(1.0 - (2.0 * u - 1.0) ** 2, 0.0, 1.0))

    # "Hình sao nhiều mũi", "Tán tự nhiên gợn sóng", "Hình xoắn ốc" và

    # "Hình pha lê đa giác"

    # "Hình cầu gai nhiều thuỳ" và "Hình pha lê đa giác" dùng chung biên dạng

    # chiều cao mặc định (Oval) - nét đặc trưng của 2 dạng này nằm ở việc

    # gợn sóng/thuỳ THEO PHƯƠNG VỊ (xem _crown_azimuthal_modulation), chỉ

    # thể hiện rõ trên khối 3D chứ không ảnh hưởng tới mặt cắt dọc 2D.

    # mặc định: Tán tròn/elip (Oval)

    return np.sin(u * np.pi)





def get_radii(row, radius_mode):

    """Trả về (rE, rN, rW, rS) - bán kính tán theo 4 hướng.

    Nếu radius_mode=False, dữ liệu Dtán được hiểu là đường kính -> chia đôi."""

    factor = 1.0 if radius_mode else 0.5

    rE = float(row["DtanD"]) * factor

    rN = float(row["DtanB"]) * factor

    rW = float(row["DtanT"]) * factor

    rS = float(row["DtanN"]) * factor

    return rE, rN, rW, rS





def resolve_crown_form(row, style):

    """Trả về dạng khối tán (mặt cắt dọc / 3D) áp dụng cho 1 cây cụ thể.

    Ưu tiên dạng tán RIÊNG đã gán cho loài cây đó (style["crown_form_map"],

    thiết lập ở hộp thoại "Dạng tán theo loài..."); nếu loài chưa được gán

    riêng thì dùng dạng tán CHUNG (style["crown_form"]) - nhờ vậy có thể mô

    phỏng linh động nhiều dạng tán khác nhau cùng lúc trong 1 ô tiêu chuẩn

    (VD: thông lá kim hình nón xen với cây lá rộng tán tròn/dù)."""

    form_map = style.get("crown_form_map") or {}

    species = str(row["Loai"])

    return form_map.get(species) or style.get("crown_form", CROWN_FORM_OPTIONS[0])





def resolve_color(row, color_map, style):

    mode = style["color_mode"]

    if mode.startswith("Một"):

        return style["uniform_color"]

    if mode.startswith("Theo chiều cao"):

        hvn = float(row["Hvn"])

        hmin, hmax = style["h_range"]

        t = 0.5 if hmax <= hmin else (hvn - hmin) / (hmax - hmin)

        cmap = get_cmap("YlGn")

        return mcolors.to_hex(cmap(0.30 + 0.65 * t))

    return color_map.get(str(row["Loai"]), "#558b2f")





def safe_hdc(hdc, hvn):

    """Đảm bảo Hdc luôn thấp hơn Hvn một khoảng hợp lý để tán cây không suy biến."""

    hdc = max(float(hdc), 0.0)

    hvn = max(float(hvn), 0.1)

    return min(hdc, hvn - max(0.05 * hvn, 0.1))





def frustum_surface(cx, cy, z0, z1, r0, r1, n=16):

    """Mặt khối chóp cụt tròn xoay (dùng cho thân cây thon dần và cho khối

    hình nón khi r1=0)."""

    theta = np.linspace(0, 2 * np.pi, n)

    u = np.array([0.0, 1.0])

    T, U = np.meshgrid(theta, u)

    R = r0 + (r1 - r0) * U

    X = cx + R * np.cos(T)

    Y = cy + R * np.sin(T)

    Z = z0 + U * (z1 - z0)

    return X, Y, Z





def cone_surface(cx, cy, z0, z1, r, n=16):

    return frustum_surface(cx, cy, z0, z1, r, 0.0, n=n)





def _crown_azimuthal_modulation(theta, form, seed=0, u=None):

    """Hệ số nhân bán kính THEO PHƯƠNG VỊ (góc quanh trục thân cây) - mặc

    định = 1 (khối tròn xoay đối xứng hoàn hảo, giống mọi dạng tán trước

    đây). Hai dạng tán "đẹp mắt" bổ sung phá vỡ tính đối xứng tròn xoay để

    khối 3D trông sống động, tự nhiên hơn khi xoay xem:

    - "Hình sao nhiều mũi": 5 thuỳ đều nhau quanh tán, như ngôi sao khi

      nhìn từ trên xuống.

    - "Tán tự nhiên gợn sóng": tổng hợp vài sóng hài biên độ nhỏ, pha

      ngẫu nhiên NHƯNG cố định theo `seed` (thường là số TT/ chỉ số cây)

      để mỗi cây có hình dạng gợn sóng riêng nhưng ổn định qua các lần vẽ

      lại (không "nhấp nháy" đổi hình khi xoay/refresh)."""

    if form.startswith("Hình sao"):

        k = 5

        return 1.0 + 0.16 * np.cos(k * theta)

    if form.startswith("Tán tự nhiên"):

        rng = np.random.RandomState(int(seed) % (2 ** 31))

        mod = np.ones_like(theta)

        for k, amp in ((3, 0.08), (5, 0.05), (7, 0.035)):

            phase = rng.uniform(0, 2 * np.pi)

            mod = mod + amp * np.cos(k * theta + phase)

        return mod

    if form.startswith("Hình xoắn ốc"):

        # Pha của gợn sóng dịch chuyển dần theo độ cao u -> tạo vân xoắn

        # chạy dọc thân tán, nhìn rất động khi xoay khối 3D.

        so_vong_xoan = 1.6

        k = 6

        uu = np.zeros_like(theta) if u is None else u

        return 1.0 + 0.16 * np.cos(k * theta - so_vong_xoan * 2 * np.pi * uu)

    if form.startswith("Hình cầu gai"):

        k = 9

        return 1.0 + 0.24 * np.cos(k * theta)

    return np.ones_like(theta)





def crown_surface(cx, cy, z0, z1, r_max, form, n=28, m=14, seed=0):

    """Khối tán 3D dạng mặt tròn xoay, sinh ra từ hàm bán kính chuẩn hoá

    r(u) của dạng tán đã chọn - thống nhất với hình vẽ mặt cắt dọc 2D - có

    thể điều biến thêm theo phương vị (xem _crown_azimuthal_modulation) để

    tạo các kiểu tán 3D lởm chởm/nhiều thuỳ tự nhiên hơn."""

    if form.startswith("Hình pha lê"):

        # Giảm mạnh số cạnh/vòng lưới -> mặt tán gồm các mảng phẳng rõ rệt,

        # tạo hiệu ứng khối đa giác (low-poly/pha lê) thay vì mặt cong mượt.

        n, m = 9, 5

    theta = np.linspace(0, 2 * np.pi, n)

    u = np.linspace(0.0, 1.0, m)

    T, U = np.meshgrid(theta, u)

    base_R = r_max * _crown_radius_profile(U[:, 0], form)[:, None]

    R = base_R * _crown_azimuthal_modulation(T, form, seed=seed, u=U)

    X = cx + R * np.cos(T)

    Y = cy + R * np.sin(T)

    Z = z0 + U * (z1 - z0)

    return X, Y, Z, U





def ellipsoid_surface(cx, cy, zc, r_xy, rz, n=18, m=12):

    theta = np.linspace(0, 2 * np.pi, n)

    phi = np.linspace(0, np.pi, m)

    T, P = np.meshgrid(theta, phi)

    X = cx + r_xy * np.sin(P) * np.cos(T)

    Y = cy + r_xy * np.sin(P) * np.sin(T)

    Z = zc + rz * np.cos(P)

    return X, Y, Z





# ==============================================================================

# 3. HÀM VẼ TRẮC ĐỒ (2D NGANG / 2D DỌC / 3D)

# ==============================================================================



def _apply_labels(ax, style, kind, default_title=None, default_xlabel=None, default_ylabel=None):

    """Áp dụng Tiêu đề / nhãn trục X-Y / màu chữ / kiểu chữ theo tuỳ chỉnh

    của người dùng (mục ⑤ Nhãn & phông chữ biểu đồ), có fallback về nhãn

    mặc định nếu người dùng để trống."""

    font_color = style.get("font_color", "#1b5e20")

    font_family = style.get("font_family", "Segoe UI")

    title_size = style.get("title_size", 11.5)

    label_size = style.get("label_size", 10.0)



    if kind == "plan":

        title = style.get("title_plan") or default_title

        xlabel = style.get("xlabel_plan") or default_xlabel or "X (m)"

        ylabel = style.get("ylabel_plan") or default_ylabel or "Y (m)"

    elif kind == "profile":

        title = style.get("title_profile") or default_title

        xlabel = style.get("xlabel_profile") or default_xlabel

        ylabel = style.get("ylabel_profile") or default_ylabel or "Chiều cao (m)"

    else:

        title, xlabel, ylabel = default_title, default_xlabel, default_ylabel



    ax.set_xlabel(xlabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.set_ylabel(ylabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.set_title(title, fontsize=title_size, fontweight="bold", color=font_color,

                 fontfamily=font_family, pad=10)

    ax.tick_params(colors=font_color)





def _struct_style(style, chart_key, default_title, default_xlabel, default_ylabel, default_color):

    """Lấy Tiêu đề / nhãn trục X / nhãn trục Y / Màu sắc / màu chữ / kiểu chữ

    / cỡ chữ tiêu đề / cỡ chữ trục / cỡ chữ chú dẫn (chỉ dùng cho "cover" và

    "layers" - hộp thông tin độ che phủ tán và chú dẫn tầng rừng) do người

    dùng tự đặt riêng cho 1 biểu đồ con trong tab "Phân tích cấu trúc rừng"

    (lưu ở style["struct_chart_style"][chart_key]). Trường nào bị để trống

    (hoặc bằng 0 với 3 cỡ chữ) sẽ tự rơi về giá trị mặc định (title/xlabel/

    ylabel/color) hoặc về cấu hình chữ CHUNG của toàn phần mềm (font_color/

    font_family/title_size/label_size) - riêng legend_size rơi về label_size

    (không phải mục ⑤ chung) vì chú dẫn thường muốn theo cỡ trục của biểu đồ."""

    cs = (style.get("struct_chart_style") or {}).get(chart_key, {})

    title = cs.get("title") or default_title

    xlabel = cs.get("xlabel") or default_xlabel

    ylabel = cs.get("ylabel") or default_ylabel

    color = cs.get("color") or default_color

    font_color = cs.get("font_color") or style.get("font_color", "#1b5e20")

    font_family = cs.get("font_family") or style.get("font_family", "Segoe UI")

    title_size = cs.get("title_size") or style.get("title_size", 11.5)

    label_size = cs.get("label_size") or style.get("label_size", 9.0)

    legend_size = cs.get("legend_size") or label_size

    return title, xlabel, ylabel, color, font_color, font_family, title_size, label_size, legend_size





_DEFAULT_SCALENOTE_POS = (0.012, 0.985)  # vị trí mặc định chú thích tỷ lệ

                                          # (1:N), góc trên-trái, phân số trục





def _real_size_for_scale(style, kind):

    """Trả về (chiều rộng thực, chiều cao thực) tính bằng mét, dùng để quy

    đổi tỷ lệ 1:N - LUÔN dựa vào ô nhập ② (không phụ thuộc việc có tick

    hiển thị khung viền hay không).

    - "plan" (trắc đồ ngang): đúng bằng chiều rộng x chiều dài ô.

    - "profile" (trắc đồ dọc): chiều ngang là cạnh ô theo trục chiếu (X

      hoặc Y), chiều đứng là chiều cao cây lớn nhất (không phải chiều

      dài/rộng ô - 2 đại lượng khác bản chất nên không thể dùng chung

      plot_h). Hàm này ở mức module (không phải phương thức của

      MainWindow) để cả draw_plan_view/draw_profile_view (qua

      _draw_scale_note) lẫn MainWindow._real_size_for_scale đều dùng

      chung 1 chỗ, tránh trùng lặp logic."""

    field_w = style.get("field_w")

    field_h = style.get("field_h")

    if kind == "profile":

        axis_key = style.get("profile_axis", "X")

        real_w = field_w if axis_key == "X" else field_h

        hmin, hmax = style.get("h_range", (0.0, 30.0))

        real_h = max(hmax * 1.12, 1.0)

        return real_w, real_h

    return field_w, field_h





def _effective_scale_denom(ax, style, kind):

    """Ước lượng mẫu số tỷ lệ (N trong 1:N) ĐANG hiển thị trên khung vẽ khi

    người dùng chọn "Tự động (vừa khung)" (style["scale_denom"] rỗng) - lúc

    này không có 1 tỷ lệ cố định do người dùng chọn, nhưng vẫn có thể suy ra

    tỷ lệ "tại thời điểm hiện tại" từ kích thước vật lý hiện tại của khung

    trục (ax) so với kích thước thực (mét) của ô đo/trắc đồ, để người dùng

    vẫn biết tương đối 1 cm trên màn hình/giấy tương ứng bao nhiêu mét

    ngoài thực địa. Giá trị này CHỈ mang tính tham khảo (đổi theo kích

    thước cửa sổ/khi export), nên _draw_scale_note sẽ hiển thị kèm dấu "≈"."""

    real_w, real_h = _real_size_for_scale(style, kind)

    if not real_w or not real_h:

        return None

    fig = ax.figure

    if fig is None:

        return None

    fig_w_in, fig_h_in = fig.get_size_inches()

    bbox = ax.get_position()

    axes_w_cm = bbox.width * fig_w_in * 2.54

    axes_h_cm = bbox.height * fig_h_in * 2.54

    if axes_w_cm <= 0 or axes_h_cm <= 0:

        return None

    denom_w = (real_w * 100.0) / axes_w_cm

    denom_h = (real_h * 100.0) / axes_h_cm

    return max(denom_w, denom_h)





def _draw_scale_note(ax, style, key):

    """Vẽ chú thích tỷ lệ bản vẽ dạng "Tỷ lệ 1:N / 1 cm trên sơ đồ = X m

    ngoài thực tế" ở góc biểu đồ. Nếu người dùng đã chọn 1 tỷ lệ cụ thể ở

    mục ② thì hiển thị đúng tỷ lệ đó; nếu đang để "Tự động (vừa khung)" thì

    tự ước lượng tỷ lệ hiện tại từ kích thước khung vẽ (xem

    _effective_scale_denom) và hiển thị kèm dấu "≈" vì giá trị này sẽ đổi

    theo kích thước cửa sổ/khi export. Kích thước (hệ số nhân,

    style["scale_note_size"]), màu chữ (style["scale_note_color"]) và font

    chữ (style["scale_note_font_family"]) có thể chỉnh riêng, mặc định dùng

    theo màu chữ / font chữ chung của toàn phần mềm (mục ⑤). Dùng

    ax.annotate (không phải ax.text) để tận dụng cơ chế KÉO THẢ có sẵn của

    matplotlib (Annotation.draggable) - vị trí sau khi kéo được ghi nhớ vào

    style["<key>_scale_note_pos"] để giữ nguyên ở những lần vẽ lại sau (đổi

    màu, đổi tab, xuất ảnh...), tương tự cơ chế đã dùng cho hộp thông tin độ

    che phủ tán (canopy_info_pos)."""

    ax._scale_note_ann = None

    if not style.get("show_scale_note", True):

        return

    denom = style.get("scale_denom")

    is_auto = not denom

    if is_auto:

        denom = _effective_scale_denom(ax, style, key)

        if not denom:

            return

    denom = float(denom)

    size_mult = float(style.get("scale_note_size", 1.0) or 1.0)

    color = style.get("scale_note_color") or style.get("font_color", "#1b5e20")

    font_family = style.get("scale_note_font_family") or style.get("font_family", "Segoe UI")



    real_m = denom / 100.0

    if abs(real_m - round(real_m)) < 1e-6:

        real_str = f"{real_m:g}"

    else:

        real_str = f"{real_m:.2f}".rstrip("0").rstrip(".")

    denom_str = f"{'≈' if is_auto else ''}{int(round(denom))}"

    text = f"Tỷ lệ 1:{denom_str}\n1 cm trên sơ đồ = {real_str} m ngoài thực tế"



    pos = style.get(f"{key}_scale_note_pos") or _DEFAULT_SCALENOTE_POS

    ann = ax.annotate(text, xy=pos, xycoords="axes fraction", ha="left", va="top",

                       multialignment="center",

                       fontsize=8.8 * size_mult, fontweight="600", color=color,

                       fontfamily=font_family,

                       bbox=dict(boxstyle="round,pad=0.32", facecolor="none",

                                 edgecolor="none", alpha=0.0),

                       annotation_clip=False, zorder=99999)

    ann.set_clip_on(False)

    ann.draggable(True, use_blit=False)

    ax._scale_note_ann = ann



    canvas = ax.figure.canvas

    old_cid = getattr(ax, "_scale_note_drag_cid", None)

    if old_cid is not None:

        try:

            canvas.mpl_disconnect(old_cid)

        except Exception:

            pass



    def _on_note_release(event, ann=ann, style=style, key=key):

        try:

            style[f"{key}_scale_note_pos"] = (float(ann.xyann[0]), float(ann.xyann[1]))

        except Exception:

            pass



    ax._scale_note_drag_cid = canvas.mpl_connect("button_release_event", _on_note_release)







def _draw_species_legend_box(ax, df, color_map, style, view_key="plan"):

    """Vẽ hộp chú dẫn màu tán theo TÊN LOÀI CÂY ngay trên trắc đồ (góc trên

    bên phải), để xem/in trắc đồ mà không cần đối chiếu riêng bảng chú giải

    ở khung điều khiển bên trái. Chỉ hiển thị khi đang tô màu "Theo loài

    cây" (các chế độ tô 1 màu chung / gradient theo chiều cao không có ý

    nghĩa chú giải theo loài) và người dùng chưa tắt tuỳ chọn này. Dùng toạ

    độ trục (ax.transAxes) nên không ảnh hưởng tới vùng dữ liệu / tỷ lệ bản

    vẽ thực tế khi xuất ảnh đúng tỷ lệ 1:N.

    `view_key` ("plan" = trắc đồ ngang, "profile" = trắc đồ dọc) - vị trí đã
    kéo được lưu RIÊNG cho từng trắc đồ (style[f"{view_key}_legend_bbox_..."])
    giống cơ chế của la bàn/thước tỷ lệ (_draw_north_arrow/_draw_scale_bar),
    vì vùng dữ liệu thật (ax.dataLim) của trắc đồ ngang (X-Y mặt bằng) và
    trắc đồ dọc (X-Hvn) khác hẳn nhau về tỷ lệ - nếu dùng chung 1 toạ độ,
    chú dẫn sẽ "nhảy" sang vị trí khác hẳn khi chuyển từ trắc đồ này sang
    trắc đồ kia (kể cả sau khi lưu/tải lại cấu hình)."""

    if not style.get("show_species_legend", True):

        return

    if not str(style.get("color_mode", "")).startswith("Theo loài"):

        return

    if df is None or df.empty or not color_map:

        return



    counts = df["Loai"].astype(str).value_counts()

    species_list = [s for s in color_map.keys() if s in counts.index]

    if not species_list:

        return

    max_items = 12

    truncated = len(species_list) > max_items

    shown = species_list[:max_items]



    # Màu chữ / cỡ chữ CHÚ DẪN có thể tách riêng khỏi màu chữ / cỡ chữ chung

    # của tiêu đề-trục (mục ⑤); nếu người dùng chưa chỉnh riêng thì dùng lại

    # giá trị chung để tương thích ngược.

    font_color = style.get("legend_font_color") or style.get("font_color", "#1b5e20")

    font_family = style.get("font_family", "Segoe UI")

    font_size = float(style.get("legend_font_size", 7.6))

    title_size = font_size + 0.7



    handles = [mpatches.Patch(facecolor=color_map[s], edgecolor=darker(color_map[s]),

                               linewidth=0.8, label=f"{s}  ({int(counts[s])})")

               for s in shown]

    if truncated:

        handles.append(mpatches.Patch(facecolor="none", edgecolor="none",

                                       label=f"… và {len(species_list) - max_items} loài khác"))



    # Vị trí chú dẫn: neo theo TỶ LỆ PHẦN TRĂM so với VÙNG DỮ LIỆU THẬT của
    # trục (ax.dataLim - hộp bao các cây, hoàn toàn không đổi dù xem trên màn
    # hình hay in ra khổ giấy nào), KHÔNG neo theo giới hạn hiển thị của trục
    # (ax.get_xlim()/get_ylim()) hay theo phân số ax.transAxes như trước đây.
    # Lý do: để giữ tỷ lệ khung hình vuông (aspect="equal"), matplotlib tự
    # "phồng" thêm giới hạn hiển thị (view limits) ra ngoài vùng dữ liệu thật
    # một khoảng đệm - và khoảng đệm này KHÁC NHAU giữa khung xem trước (hình
    # dạng cửa sổ tuỳ ý) và khung khi in (đúng tỷ lệ khổ giấy). Nếu neo theo
    # view limits (hoặc theo phân số trục), vị trí đã kéo sẽ bị lệch/rơi ra
    # ngoài khung khi in vì khoảng đệm đó thay đổi. Neo theo dataLim (không
    # bị ảnh hưởng bởi khoảng đệm này) thì vị trí đã kéo LUÔN giữ đúng chỗ so
    # với các cây, ở bất kỳ khổ giấy/hướng giấy/tỷ lệ nào.
    def _data_point_from_datalim_frac(frac_xy):
        dl = ax.dataLim
        w = (dl.x1 - dl.x0) or 1.0
        h = (dl.y1 - dl.y0) or 1.0
        dx = dl.x0 + frac_xy[0] * w
        dy = dl.y0 + frac_xy[1] * h
        return dx, dy

    try:
        ax.autoscale_view()
    except Exception:
        pass

    legend_bbox_datalim_frac = style.get(f"{view_key}_legend_bbox_datalim_frac")
    if legend_bbox_datalim_frac is None:
        # Tương thích ngược với cấu hình cũ (trước khi tách riêng theo từng
        # trắc đồ) - chỉ dùng làm giá trị khởi tạo 1 lần, sau đó mỗi trắc đồ
        # sẽ tự lưu vị trí riêng khi người dùng kéo thả.
        legend_bbox_datalim_frac = style.get("legend_bbox_datalim_frac")
    if legend_bbox_datalim_frac is None:
        # Mặc định: góc trên-phải, lùi vào một chút NGOÀI rìa vùng dữ liệu
        # thật (frac > 1 nghĩa là ra ngoài rìa cây một chút, vẫn còn nằm
        # trong phần đệm quanh trục).
        # (1.0, 1.0) = đúng góc trên-phải rìa dữ liệu; nhích thêm 1 chút ra
        # ngoài rìa nhưng vẫn nằm AN TOÀN bên trong khoảng đệm 0.14 (14%) ở
        # trên - để không bị hụt/cắt khi in đúng tỷ lệ (không còn đệm dư ra
        # do ép aspect="equal" như khi xem trên màn hình).
        legend_bbox_datalim_frac = (1.03, 1.03)
    # SỬA LỖI vị trí chú dẫn "xem 1 nơi, in ra 1 nẻo" ở TRẮC ĐỒ NGANG: trước
    # đây toạ độ neo được quy đổi qua "phân số trục" (ax.transAxes) TẠI THỜI
    # ĐIỂM VẼ - nhưng trắc đồ ngang dùng ax.set_aspect("equal",
    # adjustable="datalim"), nên lượng "đệm" mà matplotlib tự thêm vào giới
    # hạn hiển thị (view limits) để giữ khung hình vuông là KHÁC NHAU giữa
    # khung xem trước (khung Qt có tỷ lệ tuỳ ý) và khung khi xuất khổ giấy
    # (tỷ lệ đúng bằng đúng tỷ lệ vẽ 1:N nên gần như không cần đệm thêm) ->
    # cùng 1 "phân số trục" lại rơi vào 2 vị trí thực tế khác nhau. Neo trực
    # tiếp bằng TOẠ ĐỘ DỮ LIỆU (bbox_transform=ax.transData) bỏ qua hẳn bước
    # quy đổi qua phân số trục này, nên chú dẫn luôn nằm đúng 1 vị trí thực
    # so với các cây, bất kể xem trên màn hình hay xuất ra khổ giấy nào.
    dx, dy = _data_point_from_datalim_frac(legend_bbox_datalim_frac)
    bbox_anchor = (float(dx), float(dy))



    leg = ax.legend(handles=handles, title="Chú dẫn loài cây", loc="upper right",

                     bbox_to_anchor=bbox_anchor, bbox_transform=ax.transData,

                     fontsize=font_size, title_fontsize=title_size, frameon=True, framealpha=0.92,

                     facecolor="white", edgecolor="#c8d3c4", borderpad=0.7,

                     labelspacing=0.45, handlelength=1.3, handleheight=1.1,

                     handletextpad=0.6, borderaxespad=0.4)

    leg.set_zorder(99998)

    leg.get_frame().set_linewidth(1.0)

    for text in leg.get_texts():

        text.set_color(font_color)

        text.set_fontfamily(font_family)

    if leg.get_title() is not None:

        leg.get_title().set_color(font_color)

        leg.get_title().set_fontfamily(font_family)

        leg.get_title().set_fontweight("bold")

    ax.add_artist(leg)



    # ---- Cho phép KÉO chú dẫn tới bất kỳ vị trí nào trên biểu đồ bằng

    # chuột (tính năng có sẵn của matplotlib), đồng thời GHI NHỚ lại vị trí

    # mới vào style["legend_bbox"] để giữ nguyên ở những lần vẽ lại sau

    # (đổi màu, đổi tab, xoay 3D, xuất ảnh...). Mỗi lần vẽ lại legend là một

    # đối tượng hoàn toàn mới nên phải huỷ kết nối sự kiện của lần vẽ trước

    # (lưu trên chính ax) để tránh chồng nhiều callback.

    leg.set_draggable(True, use_blit=False)

    canvas = ax.figure.canvas

    old_cid = getattr(ax, "_legend_drag_cid", None)

    if old_cid is not None:

        try:

            canvas.mpl_disconnect(old_cid)

        except Exception:

            pass



    def _on_legend_release(event, ax=ax, leg=leg, style=style, view_key=view_key):

        try:

            frame_bbox = leg.get_window_extent()

            inv = ax.transAxes.inverted()

            x1, y1 = inv.transform((frame_bbox.x1, frame_bbox.y1))  # góc trên-phải

            # Lưu RIÊNG cho từng trắc đồ (plan/profile) - xem ghi chú view_key
            # ở đầu hàm _draw_species_legend_box.
            style[f"{view_key}_legend_bbox"] = (float(x1), float(y1))  # giữ lại để tương thích ngược

            # Quy đổi thêm sang TỶ LỆ PHẦN TRĂM so với VÙNG DỮ LIỆU THẬT
            # (ax.dataLim) - đây mới là toạ độ được ưu tiên dùng khi vẽ lại
            # (xem _draw_species_legend_box), vì không bị ảnh hưởng bởi khoảng
            # đệm "phồng thêm" để giữ tỷ lệ khung hình vuông, vốn khác nhau
            # giữa màn hình xem trước và khi xuất khổ giấy A4/A3, đổi hướng
            # giấy... nên chú dẫn giữ ĐÚNG vị trí đã kéo khi in.
            dx, dy = ax.transData.inverted().transform((frame_bbox.x1, frame_bbox.y1))
            dl = ax.dataLim
            w = (dl.x1 - dl.x0) or 1.0
            h = (dl.y1 - dl.y0) or 1.0
            style[f"{view_key}_legend_bbox_datalim_frac"] = (float((dx - dl.x0) / w), float((dy - dl.y0) / h))

        except Exception:

            pass



    ax._legend_drag_cid = canvas.mpl_connect("button_release_event", _on_legend_release)





_DEFAULT_NORTH_POS = (0.078, 0.865)      # vị trí mặc định la bàn (góc trên-trái, phân số trục,

                                          # đặt thấp hơn 1 chút để không chồng lên chú thích tỷ lệ 1:N)

_DEFAULT_SCALEBAR_POS = (0.045, 0.045)   # vị trí mặc định thước tỷ lệ (góc dưới-trái, phân số trục)





def _point_in_bbox(fx, fy, bbox, pad=0.012):

    """Kiểm tra điểm (fx, fy) - toạ độ phân số trục (0..1) - có nằm trong

    hộp bbox=(x0, y0, x1, y1) (mở rộng thêm biên `pad` để dễ bấm trúng)."""

    if bbox is None:

        return False

    x0, y0, x1, y1 = bbox

    return (x0 - pad) <= fx <= (x1 + pad) and (y0 - pad) <= fy <= (y1 + pad)





def _nice_round_length(x):

    """Làm tròn 1 khoảng cách (m) về số 'đẹp' gần nhất theo dãy 1-2-5 x10^n,

    dùng để chọn tổng chiều dài thước tỷ lệ tự nhiên, dễ đọc (VD: 5m, 10m,

    20m, 50m... thay vì những số lẻ như 7.3m)."""

    if x <= 0 or not np.isfinite(x):

        return 1.0

    exp = np.floor(np.log10(x))

    base = x / (10 ** exp)

    if base < 1.5:

        nice = 1

    elif base < 3.5:

        nice = 2

    elif base < 7.5:

        nice = 5

    else:

        nice = 10

    return float(nice * (10 ** exp))





def _draw_north_arrow(ax, style, key):

    """Vẽ la bàn hướng Bắc dạng "hoa gió" (compass rose) 8 cánh đen-trắng

    xen kẽ kiểu bản đồ giấy cổ điển - N/E/S/W ở 4 cánh chính (dài), 4 cánh

    phụ Đông Bắc/Đông Nam/Tây Nam/Tây Bắc ngắn hơn xen giữa, mỗi cánh chia

    đôi theo trục tâm thành 2 nửa tam giác tô màu đối lập (đậm/trắng) tạo

    hiệu ứng "cối xay gió" đặc trưng của la bàn bản đồ (theo mẫu người dùng

    cung cấp). Dùng toạ độ phân số theo hệ trục (ax.transAxes) nên luôn ở

    đúng vị trí tương đối trên khung nhìn dù dữ liệu/tỷ lệ trục thay đổi.

    Kích thước (style["north_arrow_size"], hệ số nhân, mặc định 1.0) và màu

    cánh đậm (style["north_arrow_color"]) có thể chỉnh riêng, giống cơ chế

    cỡ chữ/màu chữ của chú dẫn loài cây. Có thể KÉO THẢ tới vị trí bất kỳ

    bằng chuột (xem MainWindow._on_overlay_press/_motion), vị trí sau khi

    kéo được ghi nhớ vào style["<key>_north_arrow_pos"] để giữ nguyên ở

    những lần vẽ lại sau (đổi màu, đổi tab, xoay 3D, xuất ảnh...)."""

    ax._north_arrow_artists = []

    ax._north_arrow_bbox = None

    if not style.get("show_north_arrow", True):

        return



    fx, fy = style.get(f"{key}_north_arrow_pos") or _DEFAULT_NORTH_POS

    size_mult = float(style.get("north_arrow_size", 1.0) or 1.0)

    dark = style.get("north_arrow_color") or "#111111"

    font_family = style.get("font_family", "Segoe UI")

    r_long = 0.062 * size_mult    # bán kính cánh chính (N/E/S/W)

    r_short = 0.036 * size_mult   # bán kính cánh phụ (chéo)

    bbox_ax = ax.get_window_extent()

    aspect = bbox_ax.width / max(bbox_ax.height, 1)

    artists = []



    def _tip(angle_deg, radius):

        """Điểm đầu cánh theo góc (0° = Bắc, tăng dần theo chiều kim đồng

        hồ), hiệu chỉnh theo tỷ lệ khung để hình sao không bị méo khi trục

        X/Y có tỷ lệ hiển thị khác nhau."""

        th = np.radians(angle_deg)

        return (fx + radius * np.sin(th) / max(aspect, 1e-6), fy + radius * np.cos(th))



    # 8 hướng: chính (dài) xen kẽ phụ (ngắn) - mỗi cánh vẽ thành 2 tam giác

    # (tâm -> đỉnh -> 1 trong 2 điểm "chân cánh" sát tâm) tô màu đối lập,

    # cùng chiều quanh cả vòng để tạo hiệu ứng cối xay gió như ảnh mẫu.

    base_w = r_long * 0.16

    dirs = [0, 45, 90, 135, 180, 225, 270, 315]

    center = (fx, fy)

    for ang in dirs:

        is_main = (ang % 90 == 0)

        radius = r_long if is_main else r_short

        tip = _tip(ang, radius)

        perp = ang + 90

        b1 = (fx + base_w * np.sin(np.radians(perp)) / max(aspect, 1e-6),

              fy + base_w * np.cos(np.radians(perp)))

        b2 = (fx - base_w * np.sin(np.radians(perp)) / max(aspect, 1e-6),

              fy - base_w * np.cos(np.radians(perp)))

        tri_dark = mpatches.Polygon([center, tip, b1], closed=True, transform=ax.transAxes,

                                     facecolor=dark, edgecolor=dark, linewidth=0.5, zorder=99997)

        tri_light = mpatches.Polygon([center, tip, b2], closed=True, transform=ax.transAxes,

                                      facecolor="#ffffff", edgecolor=dark, linewidth=0.5, zorder=99997)

        ax.add_patch(tri_dark); artists.append(tri_dark)

        ax.add_patch(tri_light); artists.append(tri_light)



    # Chấm tròn nhỏ ở tâm, giống mẫu la bàn giấy cổ điển.

    hub = mpatches.Circle((fx, fy), base_w * 0.55, transform=ax.transAxes, facecolor=dark,

                           edgecolor=dark, linewidth=0, zorder=99998)

    ax.add_patch(hub); artists.append(hub)



    # Nhãn N/E/S/W đậm ở 4 đầu cánh chính, đặt cách đầu cánh 1 khoảng nhỏ.

    label_r = r_long * 1.34

    for ang, txt in ((0, "N"), (90, "E"), (180, "S"), (270, "W")):

        lx, ly = _tip(ang, label_r)

        lbl = ax.text(lx, ly, txt, transform=ax.transAxes, ha="center", va="center",

                       fontsize=9.6 * size_mult, fontweight="bold", color=dark,

                       fontfamily=font_family, zorder=99998)

        artists.append(lbl)



    ax._north_arrow_artists = artists

    ext = label_r * 1.22

    ax._north_arrow_bbox = (fx - ext, fy - ext, fx + ext, fy + ext)





def _draw_scale_bar(ax, style, key):

    """Vẽ thước tỷ lệ đồ hoạ (scale bar) dạng vạch đen-trắng xen kẽ ở góc

    biểu đồ - cách trực quan phổ biến trên bản đồ/trắc đồ giấy để người xem

    ước lượng khoảng cách thực mà KHÔNG cần đọc số ghi tỷ lệ. Chiều dài thực

    (m) được tự động chọn về 1 số 'đẹp' (1-2-5 x10^n) dựa theo khoảng hiển

    thị hiện tại của trục X, nên luôn dễ đọc dù phóng to/thu nhỏ. Kích thước

    (style["scale_bar_size"], hệ số nhân, mặc định 1.0) và màu vạch đậm

    (style["scale_bar_color"], None = dùng theo màu chữ chung) có thể chỉnh

    riêng, giống cơ chế của chú dẫn loài cây. Có thể KÉO THẢ tới vị trí bất

    kỳ bằng chuột, vị trí lưu vào style["<key>_scalebar_pos"]."""

    ax._scale_bar_artists = []

    ax._scale_bar_bbox = None

    if not style.get("show_scale_bar", True):

        return



    xlim = ax.get_xlim()

    span = abs(xlim[1] - xlim[0])

    if span <= 1e-9 or not np.isfinite(span):

        return

    size_mult = float(style.get("scale_bar_size", 1.0) or 1.0)

    total_m = _nice_round_length(span * 0.28)

    bar_w = (total_m / span) * size_mult

    if bar_w <= 0.02:

        return

    bar_w = min(bar_w, 0.55)



    fx, fy = style.get(f"{key}_scalebar_pos") or _DEFAULT_SCALEBAR_POS

    font_color = style.get("scale_bar_color") or style.get("font_color", "#1b5e20")

    font_family = style.get("font_family", "Segoe UI")

    bar_h = 0.016 * size_mult

    pad = 0.014 * size_mult

    n_seg = 4

    seg_w = bar_w / n_seg

    artists = []



    backing = mpatches.FancyBboxPatch(

        (fx - pad, fy - pad * 1.43), bar_w + pad * 2, bar_h + pad * 3.7,

        boxstyle=f"round,pad=0.006,rounding_size={0.007 * size_mult:.4f}", transform=ax.transAxes,

        facecolor="none", edgecolor="none", linewidth=0, alpha=0.0, zorder=99996)

    ax.add_patch(backing); artists.append(backing)



    for i in range(n_seg):

        seg = mpatches.Rectangle((fx + i * seg_w, fy), seg_w, bar_h, transform=ax.transAxes,

                                  facecolor=(font_color if i % 2 == 0 else "white"),

                                  edgecolor=font_color, linewidth=0.9, zorder=99997)

        ax.add_patch(seg); artists.append(seg)



    lbl0 = ax.text(fx, fy + bar_h + 0.010 * size_mult, "0", transform=ax.transAxes, ha="center",

                    va="bottom", fontsize=7.4 * size_mult, color=font_color, fontfamily=font_family,

                    zorder=99998)

    lblend = ax.text(fx + bar_w, fy + bar_h + 0.010 * size_mult, f"{total_m:g} m",

                      transform=ax.transAxes, ha="center", va="bottom", fontsize=7.8 * size_mult,

                      fontweight="bold", color=font_color, fontfamily=font_family, zorder=99998)

    artists += [lbl0, lblend]



    ax._scale_bar_artists = artists

    ax._scale_bar_bbox = (fx - pad, fy - pad * 1.43, fx + bar_w + pad, fy + bar_h + pad * 2.3)





def figsize_for_scale(real_w_m, real_h_m, scale_denom, margin_cm=2.0):

    """Tính kích thước hình (inch) để khi in ở chế độ 'Actual size / 100%'

    (KHÔNG dùng 'Fit to page'), bản vẽ đúng tỷ lệ 1:scale_denom so với thực

    tế: 1 cm trên giấy = scale_denom cm ngoài thực địa."""

    w_cm = (real_w_m * 100.0 / scale_denom) + margin_cm

    h_cm = (real_h_m * 100.0 / scale_denom) + margin_cm

    return max(w_cm / 2.54, 2.0), max(h_cm / 2.54, 2.0)





def _crown_boundary_points(x, y, rE, rN, rW, rS, r_avg, plan_shape, n_circle=72):

    """Trả về mảng điểm (Nx2) khép kín theo viền tán - dùng chung cho việc tô

    màu (fill) lẫn tách đoạn nét đứt/nét liền theo vùng bị che khuất."""

    if plan_shape.startswith("Elip"):

        return crown_polygon(x, y, rE, rN, rW, rS)

    theta = np.linspace(0.0, 2 * np.pi, n_circle, endpoint=False)

    return np.column_stack([x + r_avg * np.cos(theta), y + r_avg * np.sin(theta)])





def _boundary_segments_by_coverage(pts, covered_mask):

    """Tách viền tán (vòng khép kín) thành các đoạn polyline liên tục theo

    trạng thái covered_mask (bị che / không bị che), để có thể vẽ mỗi đoạn

    bằng nét đứt (bị che) hoặc nét liền (không bị che) - thay vì cả viền chỉ

    dùng 1 kiểu nét duy nhất."""

    n = len(pts)

    if n == 0:

        return []

    if covered_mask.all():

        return [(True, np.vstack([pts, pts[:1]]))]

    if not covered_mask.any():

        return [(False, np.vstack([pts, pts[:1]]))]



    # xoay điểm bắt đầu tới đúng 1 điểm chuyển trạng thái, để không cắt dở

    # đoạn nét đang liên tục ở chỗ nối giữa cuối mảng và đầu mảng.

    start = 0

    for k in range(n):

        if covered_mask[k] != covered_mask[k - 1]:

            start = k

            break

    order = [(start + k) % n for k in range(n)]



    segments = []

    cur_state = bool(covered_mask[order[0]])

    cur_pts = [pts[order[0]]]

    for idx in order[1:]:

        state = bool(covered_mask[idx])

        cur_pts.append(pts[idx])  # điểm nối để đoạn liền mạch với đoạn kế tiếp

        if state != cur_state:

            segments.append((cur_state, np.array(cur_pts)))

            cur_state = state

            cur_pts = [pts[idx]]

    cur_pts.append(pts[order[0]])  # khép vòng lại điểm bắt đầu

    segments.append((cur_state, np.array(cur_pts)))

    return segments





def draw_plan_view(ax, df, color_map, style):

    """Trắc đồ NGANG - mặt bằng phân bố cây theo toạ độ (X, Y).



    Phần viền tán nằm trong vùng chiếu của (các) cây CAO HƠN được vẽ bằng

    NÉT ĐỨT (bị che khuất khi nhìn từ trên xuống); phần viền còn lại - không

    bị tán cây nào cao hơn che - vẽ NÉT LIỀN, nên 1 tán có thể vừa có đoạn

    nét đứt vừa có đoạn nét liền, đúng với thực tế quan sát trên hiện trường.

    Cây cao hơn luôn được vẽ đè lên (zorder cao hơn) cây thấp hơn.

    """

    ax.clear()

    ax.set_facecolor(style["bg_color"])

    if style.get("plot_w") and style.get("plot_h"):

        ax.add_patch(Rectangle((0, 0), style["plot_w"], style["plot_h"], fill=False,

                                edgecolor="#8fa38a", linewidth=1.6, linestyle="--", zorder=1))



    n = len(df)

    if n == 0:

        _apply_labels(ax, style, "plan",

                      default_title="TRẮC ĐỒ NGANG - Mặt bằng phân bố cây",

                      default_xlabel="X (m)", default_ylabel="Y (m)")

        for spine in ax.spines.values():

            spine.set_color("#c8d3c4")

        return



    xs = df["X"].astype(float).to_numpy()

    ys = df["Y"].astype(float).to_numpy()

    hvns = df["Hvn"].astype(float).to_numpy()

    r_avgs = np.empty(n)

    radii_list = []

    boundary_list = []

    for i in range(n):

        rE, rN, rW, rS = get_radii(df.iloc[i], style["radius_mode"])

        radii_list.append((rE, rN, rW, rS))

        r_avgs[i] = (rE + rN + rW + rS) / 4.0

        boundary_list.append(_crown_boundary_points(

            xs[i], ys[i], rE, rN, rW, rS, r_avgs[i], style["plan_shape"]))



    # Path của từng tán, dùng để kiểm tra 1 điểm biên có nằm trong vùng tán

    # của cây khác hay không (point-in-polygon), phục vụ tách nét đứt/liền.

    paths_list = [MplPath(b) for b in boundary_list]



    # Vẽ cây thấp trước, cây cao sau => cây cao "phủ lên" cây thấp giống

    # góc nhìn từ trên xuống (bird's-eye view).

    order = np.argsort(hvns, kind="stable")

    for rank, i in enumerate(order):

        row = df.iloc[i]

        x, y = xs[i], ys[i]

        pts = boundary_list[i]

        color = resolve_color(row, color_map, style)

        edge = style["crown_edge_color"] if style.get("crown_edge_color") else darker(color)

        edge_w = style.get("crown_edge_width", 1.0)

        z = 2 + rank



        # Phần viền tán nằm bên trong vùng tán của bất kỳ cây nào CAO HƠN sẽ

        # bị đánh dấu "che khuất" -> vẽ nét đứt; phần còn lại (không nằm

        # trong tán cây cao hơn nào) vẽ nét liền - cho phép 1 tán vừa có

        # đoạn nét đứt vừa có đoạn nét liền tuỳ theo góc nhìn từ trên xuống.

        covered_mask = np.zeros(len(pts), dtype=bool)

        for j in range(n):

            if j == i or not (hvns[j] > hvns[i]):

                continue

            inside = paths_list[j].contains_points(pts)

            if inside.any():

                covered_mask |= inside



        # Tô màu tán (chỉ phần fill, không viền ở đây - viền vẽ riêng bên

        # dưới theo từng đoạn nét đứt/liền). picker=True để có thể CLICK vào

        # tán cây xem thông tin chi tiết (xem MainWindow.on_tree_picked).

        crown_fill = Polygon(pts, closed=True, facecolor=color, edgecolor="none",

                              alpha=style["alpha"], zorder=z, picker=True)

        crown_fill._tree_row = row

        ax.add_patch(crown_fill)

        for covered, seg in _boundary_segments_by_coverage(pts, covered_mask):

            ax.plot(seg[:, 0], seg[:, 1], color=edge, linewidth=edge_w,

                    linestyle="--" if covered else "-", zorder=z + 0.5)



        d13_m = float(row["D13"]) / 100.0

        ax.plot(x, y, marker="o", markersize=max(3.0, min(9.0, d13_m * 14)), color="#5d4037",

                markeredgecolor="#3e2723", zorder=5000 + rank)

        if style["show_labels"]:

            label = str(row["TT"]) if style["label_mode"] == "TT" else str(row["Loai"])

            ax.annotate(label, (x, y), fontsize=7.3, color=style.get("label_color", "#2b3a2f"),

                        fontweight="medium",

                        xytext=(3, 3), textcoords="offset points", zorder=6000 + rank)



    _apply_labels(ax, style, "plan",

                  default_title="TRẮC ĐỒ NGANG - Mặt bằng phân bố cây",

                  default_xlabel="X (m)", default_ylabel="Y (m)")

    ax.set_aspect("equal", adjustable="datalim")

    if style["show_grid"]:

        ax.grid(True, linestyle=":", linewidth=0.6, alpha=0.55)

    # LƯU Ý: đây là khoảng đệm THẬT DUY NHẤT quanh dữ liệu mà chú dẫn loài
    # cây (mặc định neo ở góc trên-phải, hơi ra ngoài rìa dữ liệu - xem
    # _draw_species_legend_box) có thể "mượn" để không đè lên tán cây, và
    # là khoảng đệm GIỐNG NHAU trên màn hình lẫn khi xuất khổ giấy (khi in
    # đúng tỷ lệ 1:N, khung vẽ được đặt đúng bằng tỷ lệ khung hình thật của
    # ô đo đếm nên KHÔNG có thêm đệm "phồng" do ép aspect="equal" như trên
    # màn hình xem trước). Vì vậy KHÔNG dùng cách kéo giãn Xmax/Ymax bằng
    # nút "Edit axis" của thanh công cụ matplotlib để "tạo chỗ" cho chú dẫn
    # - thao tác đó chỉ đổi GIỚI HẠN HIỂN THỊ tạm thời trên màn hình, không
    # được lưu vào cấu hình và không tồn tại khi in, nên chú dẫn sẽ "nhảy"
    # trở lại đúng vị trí thực (theo margins ở đây) sau khi tải lại/xuất
    # PDF. Muốn chừa chỗ cho chú dẫn, hãy KÉO chú dẫn (bằng chuột) vào bên
    # trong khoảng đệm này, hoặc tăng khoảng đệm ở đây nếu vẫn chưa đủ.
    ax.margins(float(style.get("legend_margin_frac", 0.14)))

    for spine in ax.spines.values():

        spine.set_color("#c8d3c4")

    _draw_scale_note(ax, style, "plan")

    _draw_species_legend_box(ax, df, color_map, style, "plan")

    _draw_north_arrow(ax, style, "plan")

    _draw_scale_bar(ax, style, "plan")





def draw_profile_view(ax, df, color_map, style):

    """Trắc đồ DỌC - mặt cắt đứng chiếu theo trục X hoặc Y, thể hiện chiều cao.



    Áp dụng phối cảnh chiều sâu (painter's algorithm): cây có toạ độ theo

    trục vuông góc (chiều sâu) LỚN HƠN được xem là ở PHÍA SAU nên vẽ trước;

    cây ở PHÍA TRƯỚC (toạ độ chiều sâu nhỏ hơn) vẽ sau cùng nên sẽ đè lên,

    che khuất một phần cây phía sau - đúng như cách vẽ trắc đồ dọc bằng tay

    truyền thống. Giảm "Độ trong tán lá" (alpha) ở khung ③ để nhìn xuyên qua

    tán trước và thấy được cây bị che ở phía sau.



    Giống như trắc đồ ngang: phần viền tán của 1 cây nằm TRONG vùng tán của

    (các) cây PHÍA TRƯỚC nó (che khuất khi nhìn từ mặt cắt) được vẽ bằng NÉT

    ĐỨT; phần viền còn lại - không bị tán cây phía trước nào che - vẽ NÉT

    LIỀN, nên 1 tán có thể vừa có đoạn nét đứt vừa có đoạn nét liền tuỳ vị

    trí. Nhờ vậy vẫn nhận ra được ranh giới cây phía sau ngay cả khi tán

    phía trước gần như không trong suốt.

    """

    # (đã bỏ chú thích cố định trên canvas theo yêu cầu người dùng)

    ax.clear()

    ax.set_facecolor(style["bg_color"])

    axis_key = style["profile_axis"]

    depth_key = "Y" if axis_key == "X" else "X"

    ax.axhline(0, color="#6d4c41", linewidth=2.2, zorder=1)



    if df.empty:

        _apply_labels(ax, style, "profile",

                      default_title=f"TRẮC ĐỒ DỌC - Mặt cắt đứng chiếu theo trục {axis_key}",

                      default_xlabel=f"{axis_key} (m)")

        for spine in ax.spines.values():

            spine.set_color("#c8d3c4")

        return



    # Sắp xếp cây theo chiều sâu giảm dần: cây xa nhất (phía sau) vẽ trước.

    rows_sorted = df.assign(_depth=df[depth_key].astype(float)).sort_values(

        "_depth", ascending=False, kind="stable")



    # ---- Bước 1: dựng trước hình học (thân + tán) của TỪNG cây theo đúng

    # thứ tự vẽ (xa -> gần), để ở bước 2 có thể kiểm tra phần viền tán nào

    # rơi vào vùng tán của (các) cây phía trước (vẽ sau) -> tách nét đứt.

    items = []

    for _, row in rows_sorted.iterrows():

        pos = float(row[axis_key])

        hvn = max(float(row["Hvn"]), 0.1)

        hdc = safe_hdc(row["Hdc"], hvn)

        rE, rN, rW, rS = get_radii(row, style["radius_mode"])

        if axis_key == "X":

            r1, r2 = rE, rW

        else:

            r1, r2 = rN, rS

        crown_w = max(r1 + r2, 0.3)

        color = resolve_color(row, color_map, style)

        edge = style["crown_edge_color"] if style.get("crown_edge_color") else darker(color)



        # ---- Thân cây: hình chóp cụt thon dần (gốc lớn, ngọn nhỏ dần),

        # khớp liền với đáy tán ở độ cao Hdc thay vì hình chữ nhật đều. ----

        trunk_w_top = max(float(row["D13"]) / 100.0 * style["trunk_exagg"], 0.05)

        trunk_w_base = trunk_w_top * 1.55  # phình gốc (bạnh vè) tự nhiên hơn

        trunk_poly = np.array([

            [pos - trunk_w_base / 2, 0], [pos + trunk_w_base / 2, 0],

            [pos + trunk_w_top / 2, hdc], [pos - trunk_w_top / 2, hdc],

        ])



        # ---- Tán cây: dựng từ hàm bán kính chuẩn hoá r(u) của dạng tán đã

        # chọn, đối xứng 2 bên quanh trục thân - khớp liền vào đỉnh thân (Hdc)

        # thay vì hình elip/tam giác cố định. ----

        n_pts = 26

        u = np.linspace(0.0, 1.0, n_pts)

        rel_r = _crown_radius_profile(u, resolve_crown_form(row, style))

        heights = hdc + u * max(hvn - hdc, 0.05)

        right_side = np.column_stack([pos + rel_r * (crown_w / 2.0), heights])

        left_side = np.column_stack([pos - rel_r * (crown_w / 2.0), heights])[::-1]

        crown_poly = np.vstack([right_side, left_side])



        items.append(dict(row=row, pos=pos, hvn=hvn, trunk_poly=trunk_poly,

                           crown_poly=crown_poly, color=color, edge=edge))



    # Path của từng tán, dùng để kiểm tra 1 điểm viền có nằm trong vùng tán

    # của cây khác hay không (point-in-polygon), phục vụ tách nét đứt/liền.

    paths_list = [MplPath(it["crown_poly"]) for it in items]



    # ---- Bước 2: vẽ theo thứ tự xa -> gần (painter's algorithm) ----

    for rank, it in enumerate(items):

        row = it["row"]

        pos, hvn = it["pos"], it["hvn"]

        color, edge = it["color"], it["edge"]

        edge_w = style.get("crown_edge_width", 0.9)

        z = 10 + rank * 3  # cây phía trước (vẽ sau) có zorder cao hơn => che cây phía sau



        ax.add_patch(Polygon(it["trunk_poly"], closed=True, facecolor=style["trunk_color"],

                              edgecolor=darker(style["trunk_color"]), linewidth=0.6, zorder=z))



        # Phần viền tán nằm bên trong vùng tán của bất kỳ cây nào PHÍA TRƯỚC

        # (rank lớn hơn, vẽ sau) sẽ bị đánh dấu "che khuất" -> vẽ nét đứt;

        # phần còn lại vẽ nét liền.

        crown_poly = it["crown_poly"]

        covered_mask = np.zeros(len(crown_poly), dtype=bool)

        for j in range(rank + 1, len(items)):

            inside = paths_list[j].contains_points(crown_poly)

            if inside.any():

                covered_mask |= inside



        # Tô màu tán (chỉ phần fill, không viền ở đây - viền vẽ riêng bên

        # dưới theo từng đoạn nét đứt/liền, giống trắc đồ ngang). picker=True

        # để có thể CLICK vào tán cây xem thông tin chi tiết.

        crown_fill = Polygon(crown_poly, closed=True, facecolor=color, edgecolor="none",

                              alpha=style["alpha"], zorder=z + 1, picker=True)

        crown_fill._tree_row = row

        ax.add_patch(crown_fill)

        for covered, seg in _boundary_segments_by_coverage(crown_poly, covered_mask):

            ax.plot(seg[:, 0], seg[:, 1], color=edge, linewidth=edge_w,

                    linestyle="--" if covered else "-", zorder=z + 1.5)



        if style["show_labels"]:

            label = str(row["TT"]) if style["label_mode"] == "TT" else str(row["Loai"])

            ax.annotate(label, (pos, hvn), fontsize=7.3, color=style.get("label_color", "#2b3a2f"),

                        xytext=(0, 4), textcoords="offset points", ha="center", zorder=z + 2)



    _apply_labels(ax, style, "profile",

                  default_title=f"TRẮC ĐỒ DỌC - Mặt cắt đứng chiếu theo trục {axis_key}",

                  default_xlabel=f"{axis_key} (m)")

    if style["show_grid"]:

        ax.grid(True, linestyle=":", linewidth=0.6, alpha=0.55)

    # LƯU Ý: gọi ax.set_ylim(bottom=0) mà không truyền "top" sẽ tự động TẮT

    # autoscale trục Y trong matplotlib (auto=False theo mặc định), khiến

    # toàn bộ tán cây phía trên bị cắt mất khỏi khung nhìn. Do đó phải tự

    # tính chiều cao lớn nhất từ dữ liệu để đặt cả 2 cận rõ ràng.

    hvn_max = float(df["Hvn"].max()) if not df.empty else 1.0

    ax.set_ylim(0, max(hvn_max * float(style.get("profile_margin_y_top", 1.12)), 1.0))

    ax.margins(x=float(style.get("profile_margin_x", 0.08)))

    for spine in ax.spines.values():

        spine.set_color("#c8d3c4")

    _draw_scale_note(ax, style, "profile")

    _draw_species_legend_box(ax, df, color_map, style, "profile")

    _draw_north_arrow(ax, style, "profile")

    _draw_scale_bar(ax, style, "profile")





def draw_3d_view(ax, df, color_map, style):

    """Mô phỏng 3D toàn bộ ô tiêu chuẩn - thân cây + tán cây, có thể xoay."""

    ax.clear()

    ax.set_facecolor(style["bg_color"])

    for i, row in df.iterrows():

        x, y = float(row["X"]), float(row["Y"])

        hvn = max(float(row["Hvn"]), 0.1)

        hdc = safe_hdc(row["Hdc"], hvn)

        rE, rN, rW, rS = get_radii(row, style["radius_mode"])

        r_avg = (rE + rN + rW + rS) / 4.0

        color = resolve_color(row, color_map, style)

        # Thân cây 3D: khối chóp cụt thon dần, gốc lớn - ngọn nhỏ, khớp với

        # đáy tán (giống cách xử lý ở trắc đồ dọc 2D) thay vì 1 nét thẳng đều.

        trunk_r_top = max(float(row["D13"]) / 100.0 * style["trunk_exagg"] / 2.0, 0.03)

        trunk_r_base = trunk_r_top * 1.55

        Xt, Yt, Zt = frustum_surface(x, y, 0, hdc, trunk_r_base, trunk_r_top, n=10)

        ax.plot_surface(Xt, Yt, Zt, color=style["trunk_color"], linewidth=0,

                         shade=True, antialiased=True, zorder=1)

        X, Y, Z, U = crown_surface(x, y, hdc, hvn, r_avg, resolve_crown_form(row, style), seed=i)

        facecolors = crown_facecolors(color, U, style["alpha"])

        ax.plot_surface(X, Y, Z, facecolors=facecolors, linewidth=0,

                         shade=True, antialiased=True)

    if style.get("plot_w") and style.get("plot_h"):

        w, h = style["plot_w"], style["plot_h"]

        xs = [0, w, w, 0, 0]; ys = [0, 0, h, h, 0]; zs = [0, 0, 0, 0, 0]

        ax.plot3D(xs, ys, zs, color="#8fa38a", linestyle="--", linewidth=1.3)

    font_color = style.get("font_color", "#1b5e20")

    font_family = style.get("font_family", "Segoe UI")

    ax.set_xlabel(style.get("xlabel_plan") or "X (m)", color=font_color, fontfamily=font_family)

    ax.set_ylabel(style.get("ylabel_plan") or "Y (m)", color=font_color, fontfamily=font_family)

    ax.set_zlabel(style.get("ylabel_profile") or "Chiều cao (m)", color=font_color, fontfamily=font_family)

    ax.set_title(style.get("title_3d") or "MÔ PHỎNG 3D - Trắc đồ rừng",

                 fontsize=style.get("title_size", 11.5), fontweight="bold",

                 color=font_color, fontfamily=font_family, pad=6)

    try:

        ax.set_box_aspect((1, 1, 0.55))

    except Exception:

        pass

    ax.xaxis.pane.set_facecolor("#eef3ea")

    ax.yaxis.pane.set_facecolor("#eef3ea")

    ax.zaxis.pane.set_facecolor("#e6ede2")





# ==============================================================================

# 3-bis. PHÂN TÍCH CẤU TRÚC RỪNG (mật độ, D1.3, G, phân bố N/D, N/Dg,

# phân bố tiết diện ngang, độ che phủ tán, chỉ số chồng lấn tán...)

# ==============================================================================

#

# Toàn bộ hàm trong khối này KHÔNG đụng tới Qt/GUI - chỉ nhận vào DataFrame

# + style (dict cấu hình đang dùng để vẽ trắc đồ ngang) và trả về số liệu /

# vẽ lên 1 trục matplotlib (ax) do nơi gọi cung cấp sẵn, đúng kiểu draw_*

# đã dùng ở các phần trên - để dễ tái sử dụng lại get_radii(), _crown_boundary_points()...



def polygon_area(pts):

    """Diện tích đa giác khép kín (m²) theo công thức Shoelace - dùng để

    tính diện tích tán của từng cây từ các điểm biên viền tán đã có."""

    pts = np.asarray(pts, dtype=float)

    if len(pts) < 3:

        return 0.0

    x = pts[:, 0]

    y = pts[:, 1]

    return 0.5 * abs(np.dot(x, np.roll(y, -1)) - np.dot(y, np.roll(x, -1)))





def compute_crown_areas(df, style):

    """Trả về mảng diện tích tán (m²) của từng cây, tính theo đúng hình

    dạng tán (Elip 4 hướng hoặc Tròn bán kính trung bình) đang chọn ở

    mục ③ Kiểu hiển thị cây - đồng bộ hoàn toàn với trắc đồ ngang."""

    n = len(df)

    areas = np.zeros(n)

    if n == 0:

        return areas

    xs = df["X"].astype(float).to_numpy()

    ys = df["Y"].astype(float).to_numpy()

    for i in range(n):

        rE, rN, rW, rS = get_radii(df.iloc[i], style["radius_mode"])

        r_avg = (rE + rN + rW + rS) / 4.0

        pts = _crown_boundary_points(xs[i], ys[i], rE, rN, rW, rS, r_avg, style["plan_shape"])

        areas[i] = polygon_area(pts)

    return areas





def compute_structure_summary(df, style):

    """Tính các chỉ tiêu cấu trúc rừng tổng hợp: mật độ/ha, D1.3 tb, G tb,

    G/ha, diện tích tán tb/tổng - KHÔNG bao gồm độ che phủ & chỉ số chồng

    lấn tán (2 chỉ tiêu đó cần rasterize nên tính riêng, xem

    compute_canopy_coverage, vì tốn nhiều thời gian hơn với OTC nhiều cây)."""

    n = len(df)

    field_w = style.get("field_w")

    field_h = style.get("field_h")

    area_ha = (field_w * field_h / 10000.0) if (field_w and field_h) else None



    if n == 0:

        return {

            "n": 0, "area_ha": area_ha, "density_ha": None, "d13_mean": None,

            "g_mean_cm2": None, "g_per_ha": None, "dg": None,

            "crown_area_mean": None, "crown_area_total": None, "crown_area_per_ha": None,

        }



    d13 = df["D13"].astype(float).to_numpy()

    g_i_m2 = np.pi / 4.0 * (d13 / 100.0) ** 2  # tiết diện ngang từng cây (m²)

    dg = float(np.sqrt(np.mean(d13 ** 2))) if n > 0 else None  # đường kính bình quân theo tiết diện ngang



    crown_areas = compute_crown_areas(df, style)



    return {

        "n": n,

        "area_ha": area_ha,

        "density_ha": (n / area_ha) if area_ha else None,

        "d13_mean": float(np.mean(d13)),

        "g_mean_cm2": float(np.mean(g_i_m2) * 10000.0),   # đổi ra cm² cho dễ đọc (số nhỏ)

        "g_per_ha": (float(np.sum(g_i_m2)) / area_ha) if area_ha else None,

        "dg": dg,

        "crown_area_mean": float(np.mean(crown_areas)),

        "crown_area_total": float(np.sum(crown_areas)),

        "crown_area_per_ha": (float(np.sum(crown_areas)) / area_ha) if area_ha else None,

        "_g_i_m2": g_i_m2,

        "_crown_areas": crown_areas,

        "_d13": d13,

    }





def compute_nd_distribution(d13, bin_width):

    """Phân bố N/D - số cây (N) theo cấp kính D1.3, cấp đều nhau, độ rộng

    cấp (bin_width, cm) do người dùng chọn."""

    if len(d13) == 0:

        return np.array([0.0, bin_width]), np.array([0])

    dmin, dmax = float(np.min(d13)), float(np.max(d13))

    start = np.floor(dmin / bin_width) * bin_width

    end = np.ceil(dmax / bin_width) * bin_width

    if end <= start:

        end = start + bin_width

    edges = np.arange(start, end + bin_width / 2.0, bin_width)

    counts, edges = np.histogram(d13, bins=edges)

    return edges, counts





def compute_g_distribution(d13, g_i_m2, bin_width):

    """Phân bố diện tích tiết diện ngang (G) theo cấp kính D1.3 - dùng

    CHUNG cấp kính với phân bố N/D, nhưng mỗi cấp cộng dồn G (m²) của các

    cây thuộc cấp đó thay vì đếm số cây, cho thấy cấp kính nào đóng góp

    nhiều tiết diện ngang (trữ lượng) nhất."""

    edges, _ = compute_nd_distribution(d13, bin_width)

    if len(d13) == 0:

        return edges, np.array([0.0])

    g_sum, edges = np.histogram(d13, bins=edges, weights=g_i_m2)

    return edges, g_sum





DG_RATIO_BOUNDS = [0.0, 0.5, 0.8, 1.0, 1.2, 1.5, 2.0, np.inf]

DG_RATIO_LABELS = ["<0,5 Dg", "0,5–0,8 Dg", "0,8–1,0 Dg", "1,0–1,2 Dg",

                    "1,2–1,5 Dg", "1,5–2,0 Dg", ">2,0 Dg"]





def compute_ndg_distribution(d13):

    """Phân bố N/Dg - số cây theo cấp kính tương đối so với Dg (đường kính

    bình quân theo tiết diện ngang, quadratic mean diameter): mỗi cây được

    xếp vào 1 trong các lớp tỉ lệ D1.3/Dg, cho biết cấu trúc lâm phần lệch

    trái/phải bao nhiêu so với đường kính bình quân G - chỉ tiêu cấu trúc

    kinh điển khác với phân bố N/D thông thường (vốn dùng cấp kính tuyệt

    đối, không so với Dg)."""

    n = len(d13)

    if n == 0:

        return DG_RATIO_LABELS, np.zeros(len(DG_RATIO_LABELS), dtype=int), None

    dg = float(np.sqrt(np.mean(np.asarray(d13, dtype=float) ** 2)))

    ratio = np.asarray(d13, dtype=float) / dg if dg > 0 else np.zeros(n)

    counts, _ = np.histogram(ratio, bins=DG_RATIO_BOUNDS)

    return DG_RATIO_LABELS, counts, dg





def compute_canopy_coverage(df, style, grid_res=140):

    """Độ che phủ tán & chỉ số chồng lấn tán - dùng phương pháp raster hoá:

    phủ 1 lưới ô vuông đều lên toàn bộ ô tiêu chuẩn, với mỗi ô kiểm tra có

    nằm trong tán của ÍT NHẤT 1 cây hay không (point-in-polygon, tái dùng

    đúng biên dạng tán như trắc đồ ngang). Từ đó suy ra:

      - Độ che phủ tán (%) = diện tích mặt đất có tán che / diện tích OTC

      - Chỉ số chồng lấn tán CI = (Σ diện tích tán từng cây − diện tích che

        phủ thực) / diện tích che phủ thực - CI = 0: tán không chồng lấn

        nhau; CI càng lớn: tán chồng lấn lên nhau càng nhiều."""

    n = len(df)

    field_w = style.get("field_w")

    field_h = style.get("field_h")

    if n == 0 or not field_w or not field_h or field_w <= 0 or field_h <= 0:

        return None



    xs = df["X"].astype(float).to_numpy()

    ys = df["Y"].astype(float).to_numpy()

    boundary_list = []

    for i in range(n):

        rE, rN, rW, rS = get_radii(df.iloc[i], style["radius_mode"])

        r_avg = (rE + rN + rW + rS) / 4.0

        boundary_list.append(_crown_boundary_points(xs[i], ys[i], rE, rN, rW, rS, r_avg, style["plan_shape"]))

    crown_areas = np.array([polygon_area(p) for p in boundary_list])



    grid_res = max(40, min(400, int(grid_res)))

    gx = np.linspace(0.0, field_w, grid_res)

    gy = np.linspace(0.0, field_h, grid_res)

    GX, GY = np.meshgrid(gx, gy)

    grid_pts = np.column_stack([GX.ravel(), GY.ravel()])

    covered = np.zeros(grid_pts.shape[0], dtype=bool)

    for pts in boundary_list:

        remaining = ~covered

        if not remaining.any():

            break

        inside = MplPath(pts).contains_points(grid_pts[remaining])

        idx = np.where(remaining)[0]

        covered[idx[inside]] = True



    plot_area = field_w * field_h

    cover_frac = float(covered.mean())

    covered_area = cover_frac * plot_area

    total_crown_area = float(crown_areas.sum())

    overlap_index = ((total_crown_area - covered_area) / covered_area) if covered_area > 0 else 0.0



    return {

        "cover_pct": min(100.0, cover_frac * 100.0),

        "covered_area": covered_area,

        "total_crown_area": total_crown_area,

        "overlap_index": max(0.0, overlap_index),

        "plot_area": plot_area,

        "grid_x": gx, "grid_y": gy,

        "covered_mask": covered.reshape(GX.shape),

    }





def draw_nd_distribution(ax, df, style, bin_width):

    ax.clear()

    ax.set_facecolor(style["bg_color"])

    title, xlabel, ylabel, bar_color, font_color, font_family, title_size, label_size, _legend_size = _struct_style(

        style, "nd", "PHÂN BỐ SỐ CÂY THEO CẤP KÍNH (N/D1.3)",

        f"Cấp kính D1.3 (cm) - cấp {bin_width:g} cm", "Số cây (N)", "#43a047")

    d13 = df["D13"].astype(float).to_numpy() if len(df) else np.array([])

    edges, counts = compute_nd_distribution(d13, bin_width)

    centers = (edges[:-1] + edges[1:]) / 2.0

    ax.bar(centers, counts, width=bin_width * 0.86, color=bar_color,

           edgecolor=darker(bar_color), linewidth=0.8, zorder=3)

    for cx, cnt in zip(centers, counts):

        if cnt > 0:

            ax.annotate(str(int(cnt)), (cx, cnt), ha="center", va="bottom",

                        fontsize=8.2, color=font_color, xytext=(0, 2), textcoords="offset points")

    ax.set_title(title, fontsize=title_size,

                 fontweight="bold", color=font_color, fontfamily=font_family, pad=6)

    ax.set_xlabel(xlabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.set_ylabel(ylabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.tick_params(colors=font_color, labelsize=label_size * 0.9)

    ax.grid(True, axis="y", linestyle=":", linewidth=0.6, alpha=0.55)

    for spine in ax.spines.values():

        spine.set_color("#c8d3c4")





def draw_ndg_distribution(ax, df, style):

    ax.clear()

    ax.set_facecolor(style["bg_color"])

    d13 = df["D13"].astype(float).to_numpy() if len(df) else np.array([])

    labels, counts, dg = compute_ndg_distribution(d13)

    default_title = "PHÂN BỐ SỐ CÂY THEO CẤP KÍNH TƯƠNG ĐỐI (N/Dg)"

    if dg is not None:

        default_title += f"  -  Dg = {dg:.1f} cm"

    title, xlabel, ylabel, bar_color, font_color, font_family, title_size, label_size, _legend_size = _struct_style(

        style, "ndg", default_title,

        "Cấp kính so với Dg (đường kính bình quân theo G)", "Số cây (N)", "#7cb342")

    x = np.arange(len(labels))

    ax.bar(x, counts, width=0.68, color=bar_color, edgecolor=darker(bar_color), linewidth=0.8, zorder=3)

    for xi, cnt in zip(x, counts):

        if cnt > 0:

            ax.annotate(str(int(cnt)), (xi, cnt), ha="center", va="bottom",

                        fontsize=8.2, color=font_color, xytext=(0, 2), textcoords="offset points")

    ax.set_xticks(x)

    ax.set_xticklabels(labels, rotation=20, ha="right", fontsize=label_size * 0.9, color=font_color)

    ax.set_title(title, fontsize=title_size, fontweight="bold",

                 color=font_color, fontfamily=font_family, pad=6)

    ax.set_xlabel(xlabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.set_ylabel(ylabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.tick_params(axis="y", colors=font_color, labelsize=label_size * 0.9)

    ax.grid(True, axis="y", linestyle=":", linewidth=0.6, alpha=0.55)

    for spine in ax.spines.values():

        spine.set_color("#c8d3c4")





def draw_g_distribution(ax, df, style, bin_width):

    ax.clear()

    ax.set_facecolor(style["bg_color"])

    title, xlabel, ylabel, bar_color, font_color, font_family, title_size, label_size, _legend_size = _struct_style(

        style, "g", "PHÂN BỐ TIẾT DIỆN NGANG (G) THEO CẤP KÍNH",

        f"Cấp kính D1.3 (cm) - cấp {bin_width:g} cm", "Tổng tiết diện ngang G (m²)", "#ef6c00")

    d13 = df["D13"].astype(float).to_numpy() if len(df) else np.array([])

    g_i_m2 = np.pi / 4.0 * (d13 / 100.0) ** 2

    edges, g_sum = compute_g_distribution(d13, g_i_m2, bin_width)

    centers = (edges[:-1] + edges[1:]) / 2.0

    ax.bar(centers, g_sum, width=bin_width * 0.86, color=bar_color,

           edgecolor=darker(bar_color), linewidth=0.8, zorder=3)

    for cx, gv in zip(centers, g_sum):

        if gv > 0:

            ax.annotate(f"{gv:.3f}", (cx, gv), ha="center", va="bottom",

                        fontsize=7.6, color=font_color, xytext=(0, 2), textcoords="offset points")

    ax.set_title(title, fontsize=title_size,

                 fontweight="bold", color=font_color, fontfamily=font_family, pad=6)

    ax.set_xlabel(xlabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.set_ylabel(ylabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.tick_params(colors=font_color, labelsize=label_size * 0.9)

    ax.grid(True, axis="y", linestyle=":", linewidth=0.6, alpha=0.55)

    for spine in ax.spines.values():

        spine.set_color("#c8d3c4")





def draw_canopy_coverage(ax, df, style, coverage):

    ax.clear()

    ax.set_facecolor(style["bg_color"])

    title, xlabel, ylabel, cover_color, font_color, font_family, title_size, label_size, legend_size = _struct_style(

        style, "cover", "ĐỘ CHE PHỦ TÁN & CHỒNG LẤN TÁN (nhìn từ trên xuống)",

        "X (m)", "Y (m)", "#66bb6a")

    field_w = style.get("field_w")

    field_h = style.get("field_h")



    if coverage is None:

        ax.text(0.5, 0.5, "Bấm \"Tính độ che phủ tán\" để xem kết quả",

                ha="center", va="center", fontsize=10.5, color=font_color,

                transform=ax.transAxes)

        ax.set_xticks([]); ax.set_yticks([])

        for spine in ax.spines.values():

            spine.set_color("#c8d3c4")

        return



    gx, gy = coverage["grid_x"], coverage["grid_y"]

    mask = coverage["covered_mask"]

    cmap = mcolors.ListedColormap(["#f3f7ef", cover_color])

    ax.imshow(mask.astype(int), origin="lower", extent=[gx[0], gx[-1], gy[0], gy[-1]],

              cmap=cmap, aspect="equal", zorder=1, alpha=0.9)



    if field_w and field_h:

        ax.add_patch(Rectangle((0, 0), field_w, field_h, fill=False,

                                edgecolor=darker(cover_color), linewidth=1.4, zorder=3))

    if len(df):

        # picker=True + pickradius để CLICK vào chấm điểm 1 cây xem thông tin
        # chi tiết (xem MainWindow.on_tree_picked) - _tree_rows_df lưu kèm
        # DataFrame theo ĐÚNG thứ tự các điểm đã vẽ (df.iloc[i] <-> điểm i),
        # để tra ngược từ event.ind (matplotlib) ra đúng dòng dữ liệu cây.
        tree_pts, = ax.plot(df["X"].astype(float), df["Y"].astype(float), marker="o", linestyle="none",

                markersize=3.2, color="#3e2723", zorder=4, picker=True, pickradius=6)

        tree_pts._tree_rows_df = df

        # Nhãn từng cây (tắt/mở qua ô "Hiển thị nhãn cây trên biểu đồ" ở
        # khung 🎨 riêng của biểu đồ này) - dùng chung "Nội dung nhãn"
        # (Số TT / Tên loài) với trắc đồ ngang/dọc, giống hệt cơ chế nhãn
        # ở draw_plan_view/draw_profile_view.
        cover_cs = style.get("struct_chart_style", {}).get("cover", {})
        if cover_cs.get("show_labels", False):
            label_mode = style.get("label_mode", "TT")
            xs_lbl = df["X"].astype(float).to_numpy()
            ys_lbl = df["Y"].astype(float).to_numpy()
            for i in range(len(df)):
                row = df.iloc[i]
                lbl = str(row["TT"]) if label_mode == "TT" else str(row["Loai"])
                ax.annotate(lbl, (xs_lbl[i], ys_lbl[i]), fontsize=6.6, color=font_color,
                            fontfamily=font_family, xytext=(3, 3), textcoords="offset points",
                            zorder=5)



    ax.set_title(title, fontsize=title_size, fontweight="bold",

                 color=font_color, fontfamily=font_family, pad=6)

    ax.set_xlabel(xlabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.set_ylabel(ylabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.tick_params(colors=font_color, labelsize=label_size * 0.9)

    ax.set_aspect("equal", adjustable="box")

    for spine in ax.spines.values():

        spine.set_color("#c8d3c4")



    info = (f"Độ che phủ tán: {coverage['cover_pct']:.1f}%\n"

            f"Chỉ số chồng lấn tán (CI): {coverage['overlap_index']:.2f}\n"

            f"Σ diện tích tán: {coverage['total_crown_area']:.1f} m²\n"

            f"Diện tích che phủ thực: {coverage['covered_area']:.1f} m²")



    # ---- Hộp thông tin có thể KÉO THẢ tới bất kỳ vị trí nào bằng chuột -

    # kể cả kéo ra NGOÀI khung biểu đồ (clip_on=False) - vị trí sau khi kéo

    # được ghi nhớ vào style["canopy_info_pos"] (toạ độ phân số theo hệ trục

    # ax.transAxes) để giữ nguyên ở những lần vẽ lại sau (đổi tham số, đổi

    # tab, xuất ảnh...), tương tự cơ chế của chú dẫn loài cây

    # (_draw_species_legend_box). Dùng ax.annotate thay vì ax.text vì Text

    # thường không hỗ trợ kéo thả, còn Annotation thì có.

    pos = style.get("canopy_info_pos") or (0.02, 0.98)

    ann = ax.annotate(info, xy=pos, xycoords="axes fraction", ha="left", va="top",

                       fontsize=legend_size, color="#1b3d1e", fontfamily=font_family,

                       bbox=dict(boxstyle="round,pad=0.4", facecolor="white",

                                 edgecolor="#8fa38a", alpha=0.88),

                       annotation_clip=False, zorder=5)

    ann.set_clip_on(False)

    ann.draggable(True, use_blit=False)



    canvas = ax.figure.canvas

    old_cid = getattr(ax, "_canopy_info_drag_cid", None)

    if old_cid is not None:

        try:

            canvas.mpl_disconnect(old_cid)

        except Exception:

            pass



    def _on_info_release(event, ann=ann, style=style):

        try:

            style["canopy_info_pos"] = (float(ann.xyann[0]), float(ann.xyann[1]))

        except Exception:

            pass



    ax._canopy_info_drag_cid = canvas.mpl_connect("button_release_event", _on_info_release)





# ------------------------------------------------------------------------------

# 3b. CẤU TRÚC ĐỨNG (theo chiều cao vút ngọn Hvn): Hvn tb/max/min, phân bố

#     N/H, phân tầng rừng (tầng trội / tầng ưu thế / tầng dưới tán)

# ------------------------------------------------------------------------------



# Tên hiển thị + màu riêng cho từng tầng rừng, dùng chung cho sơ đồ phân

# tầng và các đường phân cách trên biểu đồ phân bố N/H.

LAYER_INFO = {

    "A1": ("Tầng trội (vượt tán)", "#ef6c00"),

    "A2": ("Tầng ưu thế (tán chính)", "#2e7d32"),

    "A3": ("Tầng dưới tán", "#6d4c41"),

}





def compute_vertical_structure(df, k_std=1.0):

    """Phân tầng rừng theo chiều cao vút ngọn (Hvn) bằng phương pháp thống

    kê kinh điển Htb ± k·σ (k mặc định = 1 lần độ lệch chuẩn):

      - Tầng trội   (A1 - vượt tán) : Hvn > Htb + k·σ

      - Tầng ưu thế (A2 - tán chính): Htb − k·σ ≤ Hvn ≤ Htb + k·σ

      - Tầng dưới tán (A3)          : Hvn < Htb − k·σ

    "Chiều cao tầng rừng" lấy bằng Hvn bình quân của tầng ưu thế (A2) - vì

    đây là tầng tán chính, đại diện chiều cao chung của lâm phần."""

    n = len(df)

    if n == 0:

        return None

    hvn = df["Hvn"].astype(float).to_numpy()

    h_mean = float(np.mean(hvn))

    h_std = float(np.std(hvn))

    h_max = float(np.max(hvn))

    h_min = float(np.min(hvn))

    upper = h_mean + k_std * h_std

    lower = h_mean - k_std * h_std



    layer = np.where(hvn > upper, "A1", np.where(hvn < lower, "A3", "A2"))



    def _layer_stats(code):

        mask = layer == code

        cnt = int(mask.sum())

        if cnt == 0:

            return {"count": 0, "pct": 0.0, "h_min": None, "h_max": None, "h_mean": None}

        vals = hvn[mask]

        return {"count": cnt, "pct": cnt / n * 100.0,

                "h_min": float(vals.min()), "h_max": float(vals.max()), "h_mean": float(vals.mean())}



    a2 = _layer_stats("A2")

    return {

        "n": n, "h_mean": h_mean, "h_std": h_std, "h_max": h_max, "h_min": h_min,

        "upper": upper, "lower": lower,

        "A1": _layer_stats("A1"), "A2": a2, "A3": _layer_stats("A3"),

        "canopy_height": a2["h_mean"] if a2["h_mean"] is not None else h_mean,

        "_hvn": hvn, "_layer": layer,

    }





def compute_nh_distribution(hvn, bin_width):

    """Phân bố N/H - số cây (N) theo cấp chiều cao vút ngọn Hvn, cùng cách

    chia cấp đều nhau như compute_nd_distribution (dùng cho D1.3)."""

    if len(hvn) == 0:

        return np.array([0.0, bin_width]), np.array([0])

    hmin, hmax = float(np.min(hvn)), float(np.max(hvn))

    start = np.floor(hmin / bin_width) * bin_width

    end = np.ceil(hmax / bin_width) * bin_width

    if end <= start:

        end = start + bin_width

    edges = np.arange(start, end + bin_width / 2.0, bin_width)

    counts, edges = np.histogram(hvn, bins=edges)

    return edges, counts





def draw_nh_distribution(ax, df, style, bin_width, vstruct=None):

    ax.clear()

    ax.set_facecolor(style["bg_color"])

    title, xlabel, ylabel, bar_color, font_color, font_family, title_size, label_size, _legend_size = _struct_style(

        style, "nh", "PHÂN BỐ SỐ CÂY THEO CẤP CHIỀU CAO (N/Hvn)",

        f"Cấp chiều cao Hvn (m) - cấp {bin_width:g} m", "Số cây (N)", "#00838f")

    hvn = df["Hvn"].astype(float).to_numpy() if len(df) else np.array([])

    edges, counts = compute_nh_distribution(hvn, bin_width)

    centers = (edges[:-1] + edges[1:]) / 2.0

    ax.bar(centers, counts, width=bin_width * 0.86, color=bar_color,

           edgecolor=darker(bar_color), linewidth=0.8, zorder=3)

    for cx, cnt in zip(centers, counts):

        if cnt > 0:

            ax.annotate(str(int(cnt)), (cx, cnt), ha="center", va="bottom",

                        fontsize=8.2, color=font_color, xytext=(0, 2), textcoords="offset points")



    # Kẻ 2 đường ranh giới tầng (Htb ± σ) lên biểu đồ phân bố, cho thấy trực

    # quan cấp chiều cao nào thuộc tầng nào.

    if vstruct is not None:

        for val in (vstruct["lower"], vstruct["upper"]):

            ax.axvline(val, color="#c62828", linestyle="--", linewidth=1.1, zorder=4)

        ax.axvline(vstruct["h_mean"], color="#37474f", linestyle=":", linewidth=1.0, zorder=4)



    ax.set_title(title, fontsize=title_size,

                 fontweight="bold", color=font_color, fontfamily=font_family, pad=6)

    ax.set_xlabel(xlabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.set_ylabel(ylabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.tick_params(colors=font_color, labelsize=label_size * 0.9)

    ax.grid(True, axis="y", linestyle=":", linewidth=0.6, alpha=0.55)

    for spine in ax.spines.values():

        spine.set_color("#c8d3c4")





def draw_forest_layers(ax, df, style, vstruct):

    """Sơ đồ phân tầng đứng: mỗi cây được vẽ thành 1 điểm ở đúng chiều cao

    Hvn của nó (trục hoành chỉ để rải cây ra cho dễ nhìn, KHÔNG mang ý

    nghĩa toạ độ), tô màu theo tầng (trội / ưu thế / dưới tán), cùng 2 dải

    nền phân cách ranh giới tầng và hộp chú dẫn có thể kéo thả."""

    ax.clear()

    ax.set_facecolor(style["bg_color"])

    title, _xlabel, ylabel, _color, font_color, font_family, title_size, label_size, legend_size = _struct_style(

        style, "layers", "SƠ ĐỒ PHÂN TẦNG RỪNG THEO CHIỀU CAO VÚT NGỌN (Hvn)",

        "", "Hvn (m)", "#2e7d32")



    if vstruct is None or len(df) == 0:

        ax.text(0.5, 0.5, "Chưa có dữ liệu Hvn để phân tầng rừng",

                ha="center", va="center", fontsize=10.5, color=font_color,

                transform=ax.transAxes)

        ax.set_xticks([]); ax.set_yticks([])

        for spine in ax.spines.values():

            spine.set_color("#c8d3c4")

        return



    hvn = vstruct["_hvn"]

    layer = vstruct["_layer"]

    n = len(hvn)

    rng = np.random.default_rng(42)  # seed cố định để vị trí cây không "nhảy" mỗi lần vẽ lại

    xs = rng.uniform(0.06, 0.94, size=n)



    h_top = max(vstruct["h_max"] * 1.08, vstruct["upper"] + 1.0, 1.0)

    bands = [

        (vstruct["upper"], h_top, LAYER_INFO["A1"][1]),

        (vstruct["lower"], vstruct["upper"], LAYER_INFO["A2"][1]),

        (0.0, max(vstruct["lower"], 0.0), LAYER_INFO["A3"][1]),

    ]

    for y0, y1, color in bands:

        if y1 > y0:

            ax.axhspan(y0, y1, color=color, alpha=0.12, zorder=1)



    for code in ("A1", "A2", "A3"):

        mask = layer == code

        if not mask.any():

            continue

        label, color = LAYER_INFO[code]

        stt = vstruct[code]

        # picker=True + pickradius để CLICK vào 1 chấm cây xem thông tin chi
        # tiết (xem MainWindow.on_tree_picked). `mask` được lập từ `hvn`/
        # `layer` (vstruct["_hvn"]/["_layer"]) vốn LẤY THEO ĐÚNG THỨ TỰ dòng
        # của `df` (xem compute_vertical_structure) - nên df.iloc[mask] cũng
        # giữ đúng thứ tự tương ứng với các điểm xs[mask]/hvn[mask] đã vẽ,
        # cho phép tra ngược từ event.ind ra đúng dòng dữ liệu cây.
        scat = ax.scatter(xs[mask], hvn[mask], s=36, color=color, edgecolor="white",

                   linewidth=0.6, zorder=4, picker=True, pickradius=6,

                   label=f"{label}: {stt['count']} cây ({stt['pct']:.1f}%)")

        scat._tree_rows_df = df.iloc[np.where(mask)[0]]

    # Nhãn từng cây (tắt/mở qua ô "Hiển thị nhãn cây trên biểu đồ" ở khung
    # 🎨 riêng của biểu đồ này) - dùng chung "Nội dung nhãn" (Số TT / Tên
    # loài) với trắc đồ ngang/dọc, giống hệt cơ chế nhãn ở
    # draw_plan_view/draw_profile_view/draw_canopy_coverage.
    layers_cs = style.get("struct_chart_style", {}).get("layers", {})
    if layers_cs.get("show_labels", False):
        label_mode = style.get("label_mode", "TT")
        for idx in range(n):
            row = df.iloc[idx]
            lbl = str(row["TT"]) if label_mode == "TT" else str(row["Loai"])
            ax.annotate(lbl, (xs[idx], hvn[idx]), fontsize=6.6, color=font_color,
                        fontfamily=font_family, xytext=(0, 5), textcoords="offset points",
                        ha="center", zorder=5)



    for val in (vstruct["lower"], vstruct["upper"]):

        if val > 0:

            ax.axhline(val, color="#37474f", linestyle="--", linewidth=1.0, zorder=3)



    ax.set_xlim(0, 1)

    ax.set_ylim(0, h_top)

    ax.set_xticks([])

    ax.set_ylabel(ylabel, fontsize=label_size, color=font_color, fontfamily=font_family)

    ax.tick_params(axis="y", colors=font_color, labelsize=label_size * 0.9)

    ax.set_title(title, fontsize=title_size, fontweight="bold",

                 color=font_color, fontfamily=font_family, pad=6)

    for spine in ax.spines.values():

        spine.set_color("#c8d3c4")



    leg = ax.legend(loc="upper right", fontsize=legend_size, framealpha=0.92, facecolor="white",

                     edgecolor="#c8d3c4", borderpad=0.6, handletextpad=0.5)

    leg.set_zorder(99998)

    for text in leg.get_texts():

        text.set_color(font_color)

        text.set_fontfamily(font_family)



    # ---- Chú dẫn tầng rừng cũng có thể KÉO THẢ tới vị trí bất kỳ, ghi nhớ

    # vị trí vào style["forest_layers_bbox"] - cùng cơ chế đã dùng cho chú

    # dẫn loài cây (_draw_species_legend_box) và hộp thông tin độ che phủ

    # tán (draw_canopy_coverage).

    bbox_anchor = style.get("forest_layers_bbox")

    if bbox_anchor:

        leg.set_bbox_to_anchor(bbox_anchor, transform=ax.transAxes)

    leg.set_draggable(True, use_blit=False)

    canvas = ax.figure.canvas

    old_cid = getattr(ax, "_layers_legend_drag_cid", None)

    if old_cid is not None:

        try:

            canvas.mpl_disconnect(old_cid)

        except Exception:

            pass



    def _on_layers_legend_release(event, ax=ax, leg=leg, style=style):

        try:

            frame_bbox = leg.get_window_extent()

            inv = ax.transAxes.inverted()

            x1, y1 = inv.transform((frame_bbox.x1, frame_bbox.y1))

            style["forest_layers_bbox"] = (float(x1), float(y1))

        except Exception:

            pass



    ax._layers_legend_drag_cid = canvas.mpl_connect("button_release_event", _on_layers_legend_release)





# ==============================================================================

# 4. MODEL BẢNG DỮ LIỆU (QTableView <- DataFrame)

# ==============================================================================



DISPLAY_COLUMNS = [

    ("TT", "TT"), ("Loai", "Tên loài"), ("D13", "D1.3 (cm)"),

    ("Hvn", "Hvn (m)"), ("Hdc", "Hdc (m)"),

    ("DtanD", "Dtán Đông"), ("DtanT", "Dtán Tây"), ("DtanN", "Dtán Nam"), ("DtanB", "Dtán Bắc"),

    ("X", "X (m)"), ("Y", "Y (m)"),

]





class PandasModel(QAbstractTableModel):

    def __init__(self, df=None):

        super().__init__()

        self._df = df if df is not None else pd.DataFrame(columns=[k for k, _ in DISPLAY_COLUMNS])



    def setDataFrame(self, df):

        self.beginResetModel()

        self._df = df[[k for k, _ in DISPLAY_COLUMNS]].copy()

        self.endResetModel()



    def rowCount(self, parent=QModelIndex()):

        return len(self._df.index)



    def columnCount(self, parent=QModelIndex()):

        return len(self._df.columns)



    def data(self, index, role=Qt.DisplayRole):

        if not index.isValid() or role not in (Qt.DisplayRole, Qt.EditRole):

            return None

        val = self._df.iat[index.row(), index.column()]

        if isinstance(val, float):

            return f"{val:.2f}"

        return str(val)



    def headerData(self, section, orientation, role=Qt.DisplayRole):

        if role != Qt.DisplayRole:

            return None

        if orientation == Qt.Horizontal:

            return DISPLAY_COLUMNS[section][1]

        return str(section + 1)





# ==============================================================================

# 5. HỘP THOẠI ÁNH XẠ CỘT

# ==============================================================================



class ColumnMappingDialog(QDialog):

    """Cho phép người dùng ánh xạ cột trong file Excel thực tế vào các trường

    dữ liệu chuẩn của chương trình. Có gợi ý tự động dựa theo tên cột."""



    def __init__(self, columns, current_mapping=None, parent=None):

        super().__init__(parent)

        self.setWindowTitle("Ánh xạ cột dữ liệu Excel")

        self.setWindowIcon(make_mapping_icon(32))

        screen = QApplication.primaryScreen()

        avail = screen.availableGeometry() if screen else None

        h = min(480, int(avail.height() * 0.85)) if avail else 480

        self.resize(460, h)

        self.columns = [str(c) for c in columns]

        self.combos = {}



        layout = QVBoxLayout(self)

        info = QLabel("Chọn cột trong file Excel tương ứng với từng trường dữ liệu. "

                       "Chương trình đã tự động gợi ý theo tên cột, hãy kiểm tra lại.")

        info.setWordWrap(True)

        info.setStyleSheet("color:#4b5c4d; margin-bottom:6px;")

        layout.addWidget(info)



        scroll = QScrollArea(); scroll.setWidgetResizable(True)

        inner = QWidget(); form = QFormLayout(inner)

        form.setSpacing(8)

        for key, label, keywords in FIELD_DEFS:

            combo = QComboBox()

            combo.addItem("- Không sử dụng -", None)

            for c in self.columns:

                combo.addItem(c, c)

            guess = (current_mapping or {}).get(key)

            if guess not in self.columns:

                guess = self._guess(self.columns, keywords)

            if guess in self.columns:

                combo.setCurrentIndex(self.columns.index(guess) + 1)

            self.combos[key] = combo

            lab = QLabel(label + ":")

            if key in ("X", "Y", "Hvn", "Hdc"):

                lab.setStyleSheet("font-weight:600; color:#1b5e20;")

            form.addRow(lab, combo)

        scroll.setWidget(inner)

        layout.addWidget(scroll)



        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)

        btns.button(QDialogButtonBox.Ok).setText("Xác nhận")

        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")

        btns.accepted.connect(self.accept)

        btns.rejected.connect(self.reject)

        layout.addWidget(btns)



    @staticmethod

    def _guess(columns, keywords):

        for c in columns:

            norm = normalize_key(c)

            for kw in keywords:

                if normalize_key(kw) in norm:

                    return c

        return None



    def get_mapping(self):

        return {key: combo.currentData() for key, combo in self.combos.items()}





class ColumnExplainDialog(QDialog):

    """Hộp thoại giải thích ý nghĩa từng cột (biến) dữ liệu - giúp người

    dùng không chuyên hiểu rõ mỗi cột trong file Excel đầu vào tương ứng

    với đại lượng đo đếm nào ngoài thực địa, đơn vị gì, có bắt buộc hay

    không. Không phụ thuộc dữ liệu đã nạp hay chưa - luôn dùng chung 1

    danh sách FIELD_DEFS / FIELD_EXPLANATIONS của toàn phần mềm."""



    def __init__(self, parent=None):

        super().__init__(parent)

        self.setWindowTitle("Giải thích các cột (biến) dữ liệu")

        self.setWindowIcon(make_tree_icon(32))

        screen = QApplication.primaryScreen()

        avail = screen.availableGeometry() if screen else None

        h = min(560, int(avail.height() * 0.85)) if avail else 560

        self.resize(520, h)



        layout = QVBoxLayout(self)

        info = QLabel(

            "Mỗi dòng dữ liệu trong file Excel ứng với 1 cây trong ô tiêu "

            "chuẩn (OTC), gồm các cột (biến) dưới đây. Tên cột trong file "

            "Excel của bạn KHÔNG cần trùng khớp tuyệt đối - dùng chức năng "

            "\"Ánh xạ cột\" để chọn cột tương ứng cho từng trường.")

        info.setWordWrap(True)

        info.setStyleSheet("color:#4b5c4d; margin-bottom:6px;")

        layout.addWidget(info)



        scroll = QScrollArea()

        scroll.setWidgetResizable(True)

        scroll.setFrameShape(QFrame.NoFrame)

        inner = QWidget()

        inner_lay = QVBoxLayout(inner)

        inner_lay.setSpacing(10)

        for key, label, _ in FIELD_DEFS:

            box = QFrame()

            box.setFrameShape(QFrame.StyledPanel)

            box.setStyleSheet(

                "QFrame { background-color:#f7faf6; border:1px solid #dfe7dc; "

                "border-radius:8px; padding:8px; }")

            box_lay = QVBoxLayout(box)

            box_lay.setSpacing(3)

            box_lay.setContentsMargins(8, 6, 8, 8)



            required = key in FIELD_KEYS_CRITICAL

            head = QHBoxLayout()

            title = QLabel(label)

            title.setStyleSheet("font-weight:700; color:#1b5e20; font-size:10.5pt;")

            head.addWidget(title)

            head.addStretch(1)

            tag = QLabel("BẮT BUỘC" if required else "Tuỳ chọn")

            tag.setStyleSheet(

                ("color:#c62828; border:1px solid #c62828;" if required

                 else "color:#6b7d6d; border:1px solid #cfd9cc;") +

                " border-radius:8px; padding:1px 8px; font-size:8.2pt; font-weight:600;")

            head.addWidget(tag)

            box_lay.addLayout(head)



            desc = QLabel(FIELD_EXPLANATIONS.get(key, ""))

            desc.setWordWrap(True)

            desc.setStyleSheet("color:#2b3a2f;")

            box_lay.addWidget(desc)



            inner_lay.addWidget(box)

        inner_lay.addStretch(1)

        scroll.setWidget(inner)

        layout.addWidget(scroll)



        btns = QDialogButtonBox(QDialogButtonBox.Ok)

        btns.button(QDialogButtonBox.Ok).setText("Đóng")

        btns.accepted.connect(self.accept)

        layout.addWidget(btns)





class SpeciesCrownFormDialog(QDialog):

    """Cho phép gán DẠNG TÁN RIÊNG cho từng loài cây (mặt cắt dọc / 3D), thay

    vì dùng chung 1 dạng tán cho toàn bộ ô tiêu chuẩn - mô phỏng linh động

    hơn khi lô có nhiều loài với hình thái tán rất khác nhau (VD: thông lá

    kim hình nón mọc xen cây lá rộng tán tròn/dù)."""



    DEFAULT_LABEL = "- Dùng dạng tán chung -"



    def __init__(self, species_list, current_map=None, parent=None):

        super().__init__(parent)

        self.setWindowTitle("Dạng tán riêng theo loài")

        screen = QApplication.primaryScreen()

        avail = screen.availableGeometry() if screen else None

        h = min(480, int(avail.height() * 0.8)) if avail else 420

        self.resize(440, h)

        self.species_list = species_list

        self.combos = {}



        layout = QVBoxLayout(self)

        info = QLabel("Chọn dạng khối tán riêng cho từng loài. Loài để "

                       f"\"{self.DEFAULT_LABEL}\" sẽ dùng dạng tán chung ở "

                       "mục ③ Kiểu hiển thị cây.")

        info.setWordWrap(True)

        info.setStyleSheet("color:#4b5c4d; margin-bottom:6px;")

        layout.addWidget(info)



        scroll = QScrollArea(); scroll.setWidgetResizable(True)

        inner = QWidget(); form = QFormLayout(inner)

        form.setSpacing(8)

        current_map = current_map or {}

        for sp in self.species_list:

            combo = QComboBox()

            combo.addItem(self.DEFAULT_LABEL)

            combo.addItems(CROWN_FORM_OPTIONS)

            saved = current_map.get(sp)

            if saved in CROWN_FORM_OPTIONS:

                combo.setCurrentIndex(CROWN_FORM_OPTIONS.index(saved) + 1)

            self.combos[sp] = combo

            form.addRow(QLabel(sp + ":"), combo)

        scroll.setWidget(inner)

        layout.addWidget(scroll)



        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)

        btns.button(QDialogButtonBox.Ok).setText("Xác nhận")

        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")

        btns.accepted.connect(self.accept)

        btns.rejected.connect(self.reject)

        layout.addWidget(btns)



    def get_map(self):

        result = {}

        for sp, combo in self.combos.items():

            text = combo.currentText()

            if text != self.DEFAULT_LABEL:

                result[sp] = text

        return result





class TreeInfoDialog(QDialog):

    """Hộp thoại hiển thị đầy đủ thông tin 1 cây, bật lên khi người dùng

    CLICK vào tán cây trên Trắc đồ ngang (2D) hoặc Trắc đồ dọc (2D)."""



    def __init__(self, row, parent=None):

        super().__init__(parent)

        self.setWindowTitle(f"Thông tin cây - TT {row.get('TT', '')}")

        self.setMinimumWidth(360)



        layout = QVBoxLayout(self)

        layout.setSpacing(8)



        header = QLabel(f"🌳  {row.get('Loai', 'Không rõ')}")

        header.setWordWrap(True)

        header.setStyleSheet("font-size:13pt; font-weight:700; color:#1b5e20;")

        layout.addWidget(header)



        sub = QLabel(f"Số thứ tự (TT) trong ô tiêu chuẩn: {row.get('TT', '')}")

        sub.setStyleSheet("color:#5c6b5d; margin-bottom:4px;")

        layout.addWidget(sub)



        line = QFrame(); line.setFrameShape(QFrame.HLine); line.setStyleSheet("color:#d9e2d6;")

        layout.addWidget(line)



        def fnum(v, unit="", nd=1):

            try:

                if v is None or (isinstance(v, float) and not np.isfinite(v)):

                    return "-"

                return f"{float(v):.{nd}f} {unit}".strip()

            except Exception:

                return "-"



        rows = [

            ("D1.3 - Đường kính ngang ngực", fnum(row.get("D13"), "cm")),

            ("Hvn - Chiều cao vút ngọn", fnum(row.get("Hvn"), "m")),

            ("Hdc - Chiều cao dưới cành", fnum(row.get("Hdc"), "m")),

            ("Đường kính tán - hướng Đông", fnum(row.get("DtanD"), "m")),

            ("Đường kính tán - hướng Tây", fnum(row.get("DtanT"), "m")),

            ("Đường kính tán - hướng Nam", fnum(row.get("DtanN"), "m")),

            ("Đường kính tán - hướng Bắc", fnum(row.get("DtanB"), "m")),

            ("Toạ độ X", fnum(row.get("X"), "m", 2)),

            ("Toạ độ Y", fnum(row.get("Y"), "m", 2)),

        ]

        form = QFormLayout()

        form.setSpacing(7)

        form.setLabelAlignment(Qt.AlignLeft)

        for label, value in rows:

            lab = QLabel(label + ":")

            lab.setStyleSheet("color:#3d4a3e;")

            val = QLabel(value)

            val.setStyleSheet("font-weight:700; color:#1b3d1e;")

            form.addRow(lab, val)

        layout.addLayout(form)



        btns = QDialogButtonBox(QDialogButtonBox.Ok)

        btns.button(QDialogButtonBox.Ok).setText("Đóng")

        btns.accepted.connect(self.accept)

        layout.addWidget(btns)





# ==============================================================================

# 6. CANVAS MATPLOTLIB NHÚNG VÀO QT

# ==============================================================================



class MplCanvas(FigureCanvas):

    def __init__(self, three_d=False):

        self.fig = Figure(figsize=(6, 5), dpi=100, facecolor="#ffffff")

        self.fig.subplots_adjust(left=0.09, right=0.97, top=0.92, bottom=0.10)

        self.ax = self.fig.add_subplot(111, projection="3d") if three_d else self.fig.add_subplot(111)

        super().__init__(self.fig)

        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)





# ==============================================================================

# 7. CỬA SỔ CHÍNH

# ==============================================================================



class AboutDialog(QDialog):

    """Hộp thoại "Thông tin" - mở từ nút ℹ️ trên toolbar của cửa sổ chính. Các dòng

    liên hệ (Email/Zalo/Facebook) là nút bấm được, bấm vào là mở ứng dụng/trình

    duyệt tương ứng qua QDesktopServices (Facebook mở thẳng ra trình duyệt mặc

    định của máy)."""



    def __init__(self, parent=None, auto_close_ms: int = None):

        """`auto_close_ms`: nếu được truyền vào (vd 3000), hộp thoại sẽ TỰ ĐỘNG đóng

        sau đúng số mili-giây đó (dùng khi hiển thị lúc khởi động phần mềm). Để None

        (mặc định) khi mở từ nút "ℹ️ Thông tin" trên toolbar - lúc đó người dùng tự

        đóng bằng nút "❌ Đóng"."""

        super().__init__(parent)

        self.setWindowTitle("Thông tin phần mềm")

        self.setWindowIcon(make_tree_icon(32))

        self.setMinimumWidth(460)

        self.setModal(True)

        # Nền trắng rõ ràng cho toàn hộp thoại, tránh kế thừa màu nền tối/khác

        # từ theme hệ điều hành khiến chữ khó đọc.

        self.setStyleSheet("QDialog { background:#ffffff; }")



        outer = QVBoxLayout(self)

        outer.setContentsMargins(0, 0, 0, 0)

        outer.setSpacing(0)



        # ---- Khối tiêu đề: nền gradient xanh lá, bố cục NGANG (logo bên trái

        # trong khung tròn nổi + tên/mô tả bên phải) - tránh khoảng trống chết

        # phía trên như kiểu xếp dọc trước đây, trông cân đối và hiện đại hơn.

        header = QWidget()

        header.setFixedHeight(86)

        header.setStyleSheet(

            "background: qlineargradient(x1:0,y1:0,x2:1,y2:1,"

            "stop:0 #16213e, stop:0.5 #1f4287, stop:1 #278ea5);"

            "border-top-left-radius:10px; border-top-right-radius:10px;"

        )

        header_layout = QHBoxLayout(header)

        header_layout.setContentsMargins(22, 12, 20, 12)

        header_layout.setSpacing(14)



        # Khung tròn trắng mờ làm nền cho logo, tạo điểm nhấn thay vì để

        # logo "trôi nổi" giữa nền gradient trơn.

        badge = QLabel()

        badge.setFixedSize(58, 58)

        badge.setStyleSheet(

            "background:rgba(255,255,255,0.18); border-radius:29px;"

        )

        badge_layout = QVBoxLayout(badge)

        badge_layout.setContentsMargins(0, 0, 0, 0)

        lbl_icon = QLabel()

        lbl_icon.setPixmap(make_tree_pixmap(38))

        lbl_icon.setAlignment(Qt.AlignCenter)

        lbl_icon.setStyleSheet("background:transparent;")

        badge_layout.addWidget(lbl_icon)

        header_layout.addWidget(badge, 0, Qt.AlignVCenter)



        text_col = QVBoxLayout()

        text_col.setSpacing(3)

        lbl_name = QLabel("TDForest")

        lbl_name.setStyleSheet("font-size:19px; font-weight:800; color:#ffffff; background:transparent;")

        text_col.addWidget(lbl_name)



        lbl_sub = QLabel("Phần mềm vẽ trắc đồ ngang - dọc rừng (2D/3D)")

        lbl_sub.setWordWrap(True)

        lbl_sub.setStyleSheet("color:#dff3f7; font-size:11px; font-style:italic; background:transparent;")

        text_col.addWidget(lbl_sub)



        header_layout.addLayout(text_col, 1)



        # Huy hiệu số phiên bản - ghi rõ ngay trên nền header thay vì chỉ hiện

        # mờ nhạt ở phần thân, giúp người dùng nhận ra bản đang chạy ngay từ

        # cái nhìn đầu tiên.

        lbl_ver_badge = QLabel(f"PHIÊN BẢN V{APP_VERSION}")

        lbl_ver_badge.setAlignment(Qt.AlignCenter)

        lbl_ver_badge.setStyleSheet(

            "color:#16213e; background:#ffffff; font-size:11px; font-weight:800;"

            "padding:5px 12px; border-radius:11px;"

        )

        header_layout.addWidget(lbl_ver_badge, 0, Qt.AlignVCenter)



        outer.addWidget(header)



        # ---- Phần thân: nền trắng, chứa khối liên hệ + phiên bản ----

        body = QWidget()

        body.setStyleSheet("background:#ffffff;")

        layout = QVBoxLayout(body)

        layout.setSpacing(14)

        layout.setContentsMargins(26, 20, 26, 20)

        outer.addWidget(body)



        scope_box = QGroupBox("📜  Phạm vi sử dụng")
        scope_box.setStyleSheet(
            "QGroupBox { font-weight:700; color:#1b5e20; border:1px solid #d5dbdb;"
            "border-radius:10px; margin-top:10px; padding-top:12px; background:#ffffff; }"
            "QGroupBox::title { subcontrol-origin: margin; left:12px; padding:0 6px; background:#ffffff; }"
        )
        scope_layout = QVBoxLayout()
        lbl_scope = QLabel(_SCOPE_NOTICE)
        lbl_scope.setWordWrap(True)
        lbl_scope.setStyleSheet("font-weight:400; color:#37474F; font-size:12.5px;")
        scope_layout.addWidget(lbl_scope)
        scope_box.setLayout(scope_layout)
        layout.addWidget(scope_box)

        box = QGroupBox("📇  Liên hệ")

        box.setStyleSheet(

            "QGroupBox { font-weight:700; color:#1b5e20; border:1px solid #d5dbdb;"

            "border-radius:10px; margin-top:10px; padding-top:12px; background:#ffffff; }"

            "QGroupBox::title { subcontrol-origin: margin; left:12px; padding:0 6px; background:#ffffff; }"

        )

        box_layout = QVBoxLayout()

        box_layout.setSpacing(8)



        def make_link_row(icon_text, label_text, url, tooltip):

            btn = QPushButton(f"{icon_text}   {label_text}")

            btn.setCursor(QCursor(Qt.PointingHandCursor))

            btn.setToolTip(tooltip)

            btn.setStyleSheet(

                "QPushButton { text-align:left; padding:10px 14px; border:1.4px solid #d5dbdb;"

                "border-radius:9px; background:#ffffff; color:#1b3a20; font-size:13.5px; font-weight:600; }"

                "QPushButton:hover { background:#eaf5ea; border-color:#66bb6a; color:#1b5e20; }"

                "QPushButton:pressed { background:#c8e6c9; }"

            )

            btn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(url)))

            return btn



        box_layout.addWidget(make_link_row(

            "✉️", _ADMIN_EMAIL, f"mailto:{_ADMIN_EMAIL}", "Bấm để gửi email"))

        box_layout.addWidget(make_link_row(

            "💬", f"Zalo: {_ADMIN_ZALO}", f"https://zalo.me/{_ADMIN_ZALO}", "Bấm để mở Zalo"))

        box_layout.addWidget(make_link_row(

            "📘", "Facebook: facebook.com/nlskl.gis", "https://www.facebook.com/nlskl.gis",

            "Bấm để mở Facebook trên trình duyệt"))

        box.setLayout(box_layout)

        layout.addWidget(box)



        lbl_ver = QLabel(f"⚙️  Phiên bản phần mềm hiện tại: {APP_VERSION}")

        lbl_ver.setAlignment(Qt.AlignHCenter)

        lbl_ver.setStyleSheet("color:#607d68; font-size:11.5px; padding-top:2px; background:transparent;")

        layout.addWidget(lbl_ver)



        # Ghi nguồn icon logo theo yêu cầu bản quyền của Icons8 (gói miễn phí -

        # xem ghi chú tại hàm make_tree_pixmap). Link bấm được, mở thẳng trang

        # Icons8 tương ứng qua QDesktopServices.

        lbl_icon_credit = QLabel(
            '<a href="https://icons8.com/icon/hSXAV3TAsu3z/forest" style="color:#278ea5; text-decoration:none;">'
            'Icon "Forest" by Icons8</a>'
        )

        lbl_icon_credit.setTextFormat(Qt.RichText)

        lbl_icon_credit.setOpenExternalLinks(True)

        lbl_icon_credit.setAlignment(Qt.AlignHCenter)

        lbl_icon_credit.setStyleSheet("color:#9AA7B4; font-size:10.5px; padding-top:1px; background:transparent;")

        layout.addWidget(lbl_icon_credit)



        if auto_close_ms:

            lbl_auto = QLabel(f"(Tự động đóng sau {max(1, auto_close_ms // 1000)} giây...)")

            lbl_auto.setAlignment(Qt.AlignHCenter)

            lbl_auto.setStyleSheet("color:#9AA7B4; font-size:11px; font-style: italic; background:transparent;")

            layout.addWidget(lbl_auto)

            QTimer.singleShot(auto_close_ms, self.close)



        btn_close = QPushButton("❌   Đóng")

        btn_close.setCursor(QCursor(Qt.PointingHandCursor))

        btn_close.setStyleSheet(

            "QPushButton { padding:9px 26px; border-radius:8px; background:#1b5e20; color:white;"

            "font-weight:600; }"

            "QPushButton:hover { background:#144a18; }"

            "QPushButton:pressed { background:#0d3510; }"

        )

        btn_close.clicked.connect(self.accept)

        row = QHBoxLayout()

        row.addStretch(1)

        row.addWidget(btn_close)

        row.addStretch(1)

        layout.addLayout(row)





# ==============================================================================
# 6-bis. HIỆU ỨNG TUNG HOA KHI KHỞI ĐỘNG (CONFETTI / FLOWER OVERLAY)
# Một lớp QWidget trong suốt phủ toàn bộ cửa sổ chính, vẽ các cánh hoa/mảnh
# confetti nhiều màu rơi từ trên xuống kèm xoay + đung đưa nhẹ, mờ dần rồi
# tự hủy sau vài giây - tạo cảm giác chào mừng khi mở phần mềm. Không chặn
# thao tác chuột (WA_TransparentForMouseEvents) nên không ảnh hưởng thao
# tác của người dùng ngay cả khi hiệu ứng đang chạy.
# ==============================================================================

_CONFETTI_COLORS = [
    "#ff6b6b", "#ffd93d", "#6bcb77", "#4d96ff", "#ff8fab",
    "#c77dff", "#ff9f1c", "#2ec4b6", "#f15bb5", "#00bbf9",
    "#ffb703", "#ef476f",
]


class _ConfettiPiece:
    """Một mảnh hoa/confetti đơn lẻ trong hiệu ứng - tự lưu vị trí, tốc độ,
    góc xoay và hình dạng (cánh hoa / tròn / vuông bo góc) để paintEvent chỉ
    việc vẽ theo trạng thái hiện tại."""

    __slots__ = ("x", "y", "vx", "vy", "size", "color", "angle", "spin",
                 "shape", "sway", "phase")

    def __init__(self, w, h):
        self.x = random.uniform(0, max(w, 1))
        self.y = random.uniform(-h, 0)
        self._randomize_motion()

    def _randomize_motion(self):
        self.vy = random.uniform(90, 230)
        self.vx = random.uniform(-25, 25)
        self.size = random.uniform(7, 15)
        self.color = QColor(random.choice(_CONFETTI_COLORS))
        self.angle = random.uniform(0, 360)
        self.spin = random.uniform(-170, 170)
        self.shape = random.choice(("petal", "petal", "circle", "rect"))
        self.sway = random.uniform(0.6, 1.9)
        self.phase = random.uniform(0.0, 2 * math.pi)

    def reset_above(self, w, h):
        self.x = random.uniform(0, max(w, 1))
        self.y = random.uniform(-h * 0.35, -10)
        self._randomize_motion()


class ConfettiOverlay(QWidget):
    """Hiệu ứng "tung hoa" chào mừng hiển thị đè lên cửa sổ chính khi vừa mở
    phần mềm. Tự chạy trong `duration_ms`, mờ dần ở đoạn cuối rồi tự đóng và
    tự hủy (deleteLater) - không cần dọn dẹp thủ công ở nơi gọi."""

    def __init__(self, parent, duration_ms=3600, count=110):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WA_NoSystemBackground, True)
        self.setAutoFillBackground(False)

        self._duration_ms = duration_ms
        self._elapsed = 0
        self._fade_start = int(duration_ms * 0.7)

        w = max(parent.width() if parent else 800, 400)
        h = max(parent.height() if parent else 600, 300)
        self._pieces = [_ConfettiPiece(w, h) for _ in range(count)]

        if parent is not None:
            self.setGeometry(parent.rect())

        self._timer = QTimer(self)
        self._timer.setInterval(16)
        self._timer.timeout.connect(self._tick)
        self._timer.start()

        self.show()
        self.raise_()

    def _tick(self):
        dt = 0.016
        self._elapsed += 16
        w, h = max(self.width(), 1), max(self.height(), 1)
        t = self._elapsed / 1000.0
        for p in self._pieces:
            p.y += p.vy * dt
            p.x += p.vx * dt + math.sin(t * p.sway + p.phase) * 0.7
            p.angle += p.spin * dt
            if p.y > h + 20:
                p.reset_above(w, h)

        if self._elapsed >= self._duration_ms:
            self._timer.stop()
            self.hide()
            self.deleteLater()
            return
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        opacity = 1.0
        if self._elapsed > self._fade_start:
            span = max(1, self._duration_ms - self._fade_start)
            opacity = max(0.0, 1.0 - (self._elapsed - self._fade_start) / span)
        painter.setOpacity(opacity)
        painter.setPen(Qt.NoPen)

        for p in self._pieces:
            painter.save()
            painter.translate(p.x, p.y)
            painter.rotate(p.angle)
            painter.setBrush(QBrush(p.color))
            s = p.size
            if p.shape == "circle":
                painter.drawEllipse(QRectF(-s / 2, -s / 2, s, s))
            elif p.shape == "rect":
                painter.drawRoundedRect(QRectF(-s / 2, -s / 3, s, s * 0.66), 2, 2)
            else:  # "petal" - cánh hoa nhỏ vẽ bằng 2 đường cong Bézier
                path = QPainterPath()
                path.moveTo(0, -s / 2)
                path.cubicTo(s / 2, -s / 2, s / 2, s / 2, 0, s / 2)
                path.cubicTo(-s / 2, s / 2, -s / 2, -s / 2, 0, -s / 2)
                painter.drawPath(path)
            painter.restore()

        painter.end()


class ExportPaperDialog(QDialog):
    """Hộp thoại chọn nội dung (trắc đồ ngang/dọc), khổ giấy tiêu chuẩn
    (A2/A3/A4/A5), hướng giấy (đứng/ngang) và tỷ lệ 1:N khi xuất bản vẽ ra
    file in đúng khổ - đúng tỷ lệ, phục vụ hồ sơ điều tra rừng."""

    def __init__(self, kind_options, default_kind, default_denom, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Xuất trắc đồ theo khổ giấy")
        self.resize(440, 0)

        layout = QVBoxLayout(self)

        info = QLabel(
            "Chọn nội dung, khổ giấy, hướng giấy và tỷ lệ. Với tỷ lệ 1:N (chuẩn hoặc "
            "tuỳ chỉnh), khi in hãy chọn chế độ \"Actual size / 100%\" (KHÔNG dùng "
            "\"Fit to page\") để giữ đúng tỷ lệ. Riêng chế độ \"Phóng/thu vừa khổ giấy\" "
            "sẽ tự co giãn bản vẽ cho lấp đầy khổ giấy đã chọn, KHÔNG theo tỷ lệ 1:N chuẩn.")
        info.setWordWrap(True)
        info.setStyleSheet("color:#4b5c4d; margin-bottom:6px;")
        layout.addWidget(info)

        form = QFormLayout()
        form.setSpacing(8)

        self.kind_combo = QComboBox()
        for key, label in kind_options:
            self.kind_combo.addItem(label, key)
        i = self.kind_combo.findData(default_kind)
        if i >= 0:
            self.kind_combo.setCurrentIndex(i)
        form.addRow("Nội dung:", self.kind_combo)

        self.paper_combo = QComboBox()
        for name in ("A2", "A3", "A4", "A5"):
            w_mm, h_mm = PAPER_SIZES_MM[name]
            self.paper_combo.addItem(f"{name}   ({int(w_mm)} × {int(h_mm)} mm)", name)
        self.paper_combo.setCurrentIndex(2)  # A4 mặc định
        form.addRow("Khổ giấy:", self.paper_combo)

        self.orient_combo = QComboBox()
        self.orient_combo.addItem("Đứng (Portrait)", "portrait")
        self.orient_combo.addItem("Ngang (Landscape)", "landscape")
        form.addRow("Hướng giấy:", self.orient_combo)

        self.scale_combo = QComboBox()
        self.scale_combo.addItem("Phóng/thu vừa khổ giấy (không theo tỷ lệ)", "fit")
        self.scale_combo.addItem("Tự động - chọn tỷ lệ chuẩn vừa khít khổ giấy", None)
        for d in STANDARD_SCALE_DENOMS:
            self.scale_combo.addItem(f"1 : {d}", d)
        self.scale_combo.addItem("Tuỳ chỉnh...", "custom")
        if default_denom:
            j = self.scale_combo.findData(int(default_denom) if float(default_denom).is_integer() else default_denom)
            if j >= 0:
                self.scale_combo.setCurrentIndex(j)
            else:
                self.scale_combo.setCurrentIndex(self.scale_combo.count() - 1)
        form.addRow("Tỷ lệ:", self.scale_combo)

        self.custom_scale_spin = QDoubleSpinBox()
        self.custom_scale_spin.setRange(10, 100000)
        self.custom_scale_spin.setDecimals(0)
        self.custom_scale_spin.setSingleStep(50)
        self.custom_scale_spin.setValue(float(default_denom) if default_denom else 200.0)
        self.custom_scale_spin.setPrefix("1 : ")
        self.custom_scale_spin.setVisible(self.scale_combo.currentData() == "custom")
        form.addRow("", self.custom_scale_spin)

        self.scale_note_lbl = QLabel(
            "Bản vẽ sẽ được co giãn lấp đầy khổ giấy đã chọn (không giữ tỷ lệ 1:N "
            "chuẩn) - phù hợp khi chỉ cần bản in xem nhanh, không dùng để đo đạc.")
        self.scale_note_lbl.setWordWrap(True)
        self.scale_note_lbl.setStyleSheet("color:#8a6d3b; font-size:10px;")
        self.scale_note_lbl.setVisible(self.scale_combo.currentData() == "fit")
        form.addRow("", self.scale_note_lbl)

        self.scale_combo.currentIndexChanged.connect(self._on_scale_changed)

        self.format_combo = QComboBox()
        self.format_combo.addItem("PDF (khuyên dùng để in đúng tỷ lệ)", "pdf")
        self.format_combo.addItem("PNG (ảnh, 300 dpi)", "png")
        form.addRow("Định dạng file:", self.format_combo)

        layout.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Xuất file")
        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _on_scale_changed(self):
        data = self.scale_combo.currentData()
        self.custom_scale_spin.setVisible(data == "custom")
        self.scale_note_lbl.setVisible(data == "fit")

    def get_values(self):
        denom = self.scale_combo.currentData()
        if denom == "custom":
            denom = self.custom_scale_spin.value()
        return {
            "kind": self.kind_combo.currentData(),
            "paper": self.paper_combo.currentData(),
            "orientation": self.orient_combo.currentData(),
            "denom": denom,  # "fit" = phóng/thu vừa khổ giấy; None = tự động chọn tỷ lệ chuẩn; số = 1:N cụ thể
            "fmt": self.format_combo.currentData(),
        }


# ==============================================================================
# 9. TỔ THÀNH LOÀI - ĐA DẠNG SINH HỌC - TRỮ LƯỢNG
#    + BÁO CÁO CHUYÊN NGHIỆP (PDF nhiều trang khổ A4 / Excel nhiều sheet)
# ------------------------------------------------------------------------------
#  - IV% (Importance Value) = (N% + G%) / 2 ; loài có IV% >= 5% là loài ưu thế
#    và được đưa vào công thức tổ thành (hệ số = IV% / 10), các loài còn lại
#    gộp thành "Lk" (loài khác).
#  - Đa dạng: Shannon H' = -Σ pi·ln(pi) ; Pielou J' = H'/ln(S) ;
#    Simpson (1 - D) = 1 - Σ pi² ; Margalef d = (S - 1)/ln(N).
#  - Trữ lượng: V = G · Hvn · f (f = hình số thân cây, mặc định 0,45).
#  Toàn bộ phần tính toán + vẽ báo cáo chỉ dùng numpy/pandas/matplotlib nên
#  KHÔNG cần cài thêm thư viện nào (Excel dùng openpyxl - vốn đã cần để đọc
#  file .xlsx đầu vào).
# ==============================================================================
DEFAULT_FORM_FACTOR = 0.45
IV_DOMINANT_THRESHOLD = 5.0
REPORT_FONTS = ["Segoe UI", "Arial", "Tahoma", "DejaVu Sans"]
RPT_GREEN_DARK = "#0f3d1e"
RPT_GREEN = "#1b5e20"
RPT_GREEN_MID = "#2e7d32"
RPT_GREEN_LIGHT = "#e8f5e9"
RPT_ZEBRA = "#f5f9f4"
RPT_LINE = "#cfdccb"
RPT_TEXT = "#26332a"
RPT_MUTED = "#66756a"
RPT_ACCENT = "#c8a24a"
A4_PORTRAIT = (8.27, 11.69)
A4_LANDSCAPE = (11.69, 8.27)

REPORT_SECTIONS = [
    ("cover",       "Trang bìa & chỉ tiêu nổi bật"),
    ("summary",     "Chỉ tiêu tổng hợp & nhận xét"),
    ("composition", "Tổ thành loài & đa dạng sinh học"),
    ("plan",        "Trắc đồ ngang"),
    ("profile",     "Trắc đồ dọc"),
    ("view3d",      "Mô phỏng 3D"),
    ("charts",      "Biểu đồ phân bố & phân tầng"),
    ("appendix",    "Phụ lục bảng số liệu cây"),
]


def _vn(x, nd=1, unit=""):
    """Định dạng số kiểu Việt Nam: 1.234,5 (dấu chấm hàng nghìn, phẩy thập phân)."""
    try:
        if x is None or not np.isfinite(float(x)):
            return "-"
    except (TypeError, ValueError):
        return "-"
    s = f"{float(x):,.{nd}f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{s} {unit}".strip() if unit else s


def compute_species_composition(df):
    """Bảng tổ thành loài: N, N%, G, G%, IV%, D1.3 tb, Hvn tb - sắp theo IV% giảm dần."""
    cols = ["STT", "Loai", "N", "N%", "G", "G%", "IV%", "D13_tb", "Hvn_tb"]
    if df is None or len(df) == 0:
        return pd.DataFrame(columns=cols)
    d = pd.to_numeric(df["D13"], errors="coerce").fillna(0).to_numpy(dtype=float)
    t = pd.DataFrame({
        "Loai": df["Loai"].astype(str).to_numpy(),
        "G": np.pi / 4.0 * (d / 100.0) ** 2,
        "D13": d,
        "Hvn": pd.to_numeric(df["Hvn"], errors="coerce").to_numpy(dtype=float),
    })
    grp = t.groupby("Loai", sort=False).agg(
        N=("G", "size"), G=("G", "sum"), D13_tb=("D13", "mean"), Hvn_tb=("Hvn", "mean")
    ).reset_index()
    n_tot = float(grp["N"].sum())
    g_tot = float(grp["G"].sum())
    grp["N%"] = grp["N"] / n_tot * 100.0 if n_tot else 0.0
    grp["G%"] = grp["G"] / g_tot * 100.0 if g_tot > 0 else 0.0
    grp["IV%"] = (grp["N%"] + grp["G%"]) / 2.0
    grp = grp.sort_values(["IV%", "N"], ascending=False).reset_index(drop=True)
    grp.insert(0, "STT", np.arange(1, len(grp) + 1))
    return grp[cols]


def compute_diversity_indices(comp):
    """Các chỉ số đa dạng sinh học tính trên số cây của từng loài."""
    res = {"S": 0, "N": 0, "H": None, "Hmax": None, "J": None, "simpson": None, "margalef": None}
    if comp is None or comp.empty:
        return res
    n_i = comp["N"].to_numpy(dtype=float)
    N = float(n_i.sum())
    S = int(len(n_i))
    p = n_i / N
    H = float(-np.sum(p * np.log(p)))
    res.update({
        "S": S, "N": int(N), "H": H,
        "Hmax": float(np.log(S)) if S > 1 else 0.0,
        "J": H / float(np.log(S)) if S > 1 else None,
        "simpson": float(1.0 - np.sum(p ** 2)),
        "margalef": (S - 1) / float(np.log(N)) if N > 1 else None,
    })
    return res


def build_composition_formula(comp, threshold=IV_DOMINANT_THRESHOLD):
    """Công thức tổ thành theo IV%: '3,25 Dầu rái + 2,10 Sao đen + ... + 1,05 Lk'."""
    if comp is None or comp.empty:
        return "-"
    major = comp[comp["IV%"] >= threshold]
    minor = comp[comp["IV%"] < threshold]
    parts = [f"{_vn(r['IV%'] / 10.0, 2)} {r['Loai']}" for _, r in major.iterrows()]
    if len(minor):
        parts.append(f"{_vn(minor['IV%'].sum() / 10.0, 2)} Lk")
    return " + ".join(parts) if parts else "-"


def compute_stand_volume(df, area_ha, form_factor=DEFAULT_FORM_FACTOR):
    """Trữ lượng gỗ: V(cây) = G · Hvn · f ; M = ΣV ; M/ha = M / diện tích (ha)."""
    if df is None or len(df) == 0:
        return {"f": form_factor, "v_i": np.array([]), "M": 0.0, "M_ha": None}
    d = pd.to_numeric(df["D13"], errors="coerce").fillna(0).to_numpy(dtype=float)
    h = pd.to_numeric(df["Hvn"], errors="coerce").fillna(0).to_numpy(dtype=float)
    v = np.pi / 4.0 * (d / 100.0) ** 2 * h * float(form_factor)
    M = float(v.sum())
    return {"f": float(form_factor), "v_i": v, "M": M, "M_ha": (M / area_ha) if area_ha else None}


def draw_species_composition(ax, comp, color_map, title=None, max_species=15):
    """Biểu đồ thanh ngang IV% của các loài (tối đa `max_species` loài đầu)."""
    ax.clear()
    if comp is None or comp.empty:
        ax.set_axis_off()
        ax.text(0.5, 0.5, "Chưa có dữ liệu - hãy nạp file Excel cây rừng",
                ha="center", va="center", color=RPT_MUTED, fontsize=11, transform=ax.transAxes)
        return
    ax.set_axis_on()
    top = comp.head(max_species).iloc[::-1]
    y = np.arange(len(top))
    colors = [(color_map or {}).get(str(s), "#43a047") for s in top["Loai"]]
    xmax = max(float(top["IV%"].max()) * 1.55, IV_DOMINANT_THRESHOLD * 2.2)
    ax.barh(y, top["IV%"], color=colors, edgecolor="white", linewidth=1.2, height=0.62, zorder=3)
    for yi, (iv, n_pct, g_pct) in enumerate(zip(top["IV%"], top["N%"], top["G%"])):
        ax.text(iv + xmax * 0.012, yi, f"{_vn(iv, 1)}%",
                va="center", ha="left", fontsize=9, fontweight="bold", color=RPT_GREEN, zorder=4)
        ax.text(iv + xmax * 0.012 + xmax * 0.085, yi, f"(N {_vn(n_pct, 1)}% · G {_vn(g_pct, 1)}%)",
                va="center", ha="left", fontsize=7.6, color=RPT_MUTED, zorder=4)
    ax.set_yticks(y)
    ax.set_yticklabels([str(s) for s in top["Loai"]], fontsize=9, color=RPT_TEXT)
    ax.set_xlim(0, xmax)
    ax.set_ylim(-0.7, len(top) - 0.3)
    ax.axvline(IV_DOMINANT_THRESHOLD, color="#c62828", ls=(0, (4, 3)), lw=1.1, zorder=2)
    ax.text(IV_DOMINANT_THRESHOLD, len(top) - 0.35, "  Ngưỡng loài ưu thế (IV% = 5%)",
            color="#c62828", fontsize=7.6, va="bottom", ha="left")
    ax.set_xlabel("Chỉ số giá trị quan trọng IV% = (N% + G%) / 2", fontsize=9.5, color=RPT_TEXT)
    ax.set_title(title or "TỔ THÀNH LOÀI THEO CHỈ SỐ GIÁ TRỊ QUAN TRỌNG (IV%)",
                 fontsize=11.5, fontweight="bold", color=RPT_GREEN, pad=12)
    ax.grid(axis="x", color="#e3eae1", lw=0.8, zorder=0)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color("#9fb19b")
    ax.tick_params(axis="x", colors=RPT_MUTED, labelsize=8.5)
    ax.tick_params(axis="y", length=0)


# ------------------------------------------------------------------------------
# Thu thập số liệu dùng chung cho báo cáo PDF & Excel
# ------------------------------------------------------------------------------
def _skewness(x):
    x = np.asarray(x, dtype=float)
    if len(x) < 3 or np.std(x) == 0:
        return 0.0
    return float(np.mean((x - x.mean()) ** 3) / np.std(x) ** 3)


def _collect_report_data(df, style, meta, canopy=None, vstruct=None):
    f = float(meta.get("form_factor") or DEFAULT_FORM_FACTOR)
    summary = compute_structure_summary(df, style)
    comp = compute_species_composition(df)
    div = compute_diversity_indices(comp)
    formula = build_composition_formula(comp)
    vol = compute_stand_volume(df, summary.get("area_ha"), f)
    if vstruct is None:
        vstruct = compute_vertical_structure(df, k_std=style.get("vstruct_k_std", 1.0))
    area_ha = summary.get("area_ha")
    g_tot = float(np.sum(summary.get("_g_i_m2", np.array([0.0]))))
    hdc = pd.to_numeric(df["Hdc"], errors="coerce").to_numpy(dtype=float)
    fw, fh = style.get("field_w"), style.get("field_h")

    h_rows = [
        ("Diện tích ô tiêu chuẩn", "S", area_ha, "ha", 4),
        ("Số cây đo đếm trong OTC", "n", summary.get("n"), "cây", 0),
        ("Mật độ", "N/ha", summary.get("density_ha"), "cây/ha", 0),
        ("Đường kính bình quân D1.3", "D̄1.3", summary.get("d13_mean"), "cm", 1),
        ("Đường kính bình quân tiết diện", "Dg", summary.get("dg"), "cm", 1),
        ("Tiết diện ngang bình quân 1 cây", "ḡ", summary.get("g_mean_cm2"), "cm²", 0),
        ("Tổng tiết diện ngang OTC", "ΣG", g_tot, "m²", 3),
        ("Tiết diện ngang trên ha", "G/ha", summary.get("g_per_ha"), "m²/ha", 2),
        ("Hình số thân cây", "f", f, "", 2),
        ("Trữ lượng OTC", "ΣM", vol["M"], "m³", 3),
        ("Trữ lượng trên ha", "M/ha", vol["M_ha"], "m³/ha", 1),
        ("Diện tích tán bình quân", "S̄tán", summary.get("crown_area_mean"), "m²", 1),
        ("Tổng diện tích tán", "ΣStán", summary.get("crown_area_total"), "m²", 1),
        ("Độ che phủ tán", "CP", canopy.get("cover_pct") if canopy else None, "%", 1),
        ("Chỉ số chồng lấn tán", "CI", canopy.get("overlap_index") if canopy else None, "", 2),
    ]
    v_rows = []
    if vstruct is not None:
        v_rows = [
            ("Chiều cao vút ngọn bình quân", "H̄vn", vstruct["h_mean"], "m", 1),
            ("Độ lệch chuẩn Hvn", "σH", vstruct["h_std"], "m", 2),
            ("Hvn lớn nhất", "Hmax", vstruct["h_max"], "m", 1),
            ("Hvn nhỏ nhất", "Hmin", vstruct["h_min"], "m", 1),
            ("Chiều cao dưới cành bình quân", "H̄dc", float(np.nanmean(hdc)) if len(hdc) else None, "m", 1),
            ("Chiều cao tầng rừng (tán chính)", "Htầng", vstruct["canopy_height"], "m", 1),
        ]
        for code, name in (("A1", "Tầng trội (A1)"), ("A2", "Tầng ưu thế (A2)"), ("A3", "Tầng dưới tán (A3)")):
            st = vstruct[code]
            v_rows.append((f"{name} - số cây", code, st["count"], "cây", 0))
            v_rows.append((f"{name} - tỷ lệ", f"{code}%", st["pct"], "%", 1))

    # ---- nhận xét tự động ----
    remarks = []
    n = summary.get("n") or 0
    if n:
        remarks.append(
            f"Ô tiêu chuẩn{(' ' + meta.get('otc')) if meta.get('otc') else ''} "
            f"({_vn(fw, 0)} × {_vn(fh, 0)} m, {_vn(area_ha, 3)} ha) có {n} cây thuộc {div['S']} loài; "
            f"mật độ ước tính {_vn(summary.get('density_ha'), 0)} cây/ha, tổng tiết diện ngang "
            f"{_vn(summary.get('g_per_ha'), 2)} m²/ha và trữ lượng {_vn(vol['M_ha'], 1)} m³/ha (f = {_vn(f, 2)}).")
        sk = _skewness(summary.get("_d13", []))
        if sk > 0.3:
            shape = "lệch trái (số cây tập trung ở các cỡ kính nhỏ)"
        elif sk < -0.3:
            shape = "lệch phải (số cây tập trung ở các cỡ kính lớn)"
        else:
            shape = "tương đối cân đối quanh giá trị bình quân"
        remarks.append(
            f"Đường kính bình quân D1.3 = {_vn(summary.get('d13_mean'), 1)} cm (Dg = {_vn(summary.get('dg'), 1)} cm); "
            f"phân bố số cây theo cỡ kính có dạng {shape}, hệ số lệch Sk = {_vn(sk, 2)}.")
        if vstruct is not None:
            remarks.append(
                f"Chiều cao vút ngọn bình quân {_vn(vstruct['h_mean'], 1)} m, dao động "
                f"{_vn(vstruct['h_min'], 1)} - {_vn(vstruct['h_max'], 1)} m. Tầng ưu thế (A2) chiếm "
                f"{_vn(vstruct['A2']['pct'], 1)}% số cây, tầng trội (A1) {_vn(vstruct['A1']['pct'], 1)}% "
                f"và tầng dưới tán (A3) {_vn(vstruct['A3']['pct'], 1)}%.")
        if not comp.empty:
            k_dom = int((comp["IV%"] >= IV_DOMINANT_THRESHOLD).sum())
            top = comp.iloc[0]
            remarks.append(
                f"Có {k_dom} loài ưu thế (IV% ≥ 5%), trong đó {top['Loai']} chiếm ưu thế nhất với "
                f"IV% = {_vn(top['IV%'], 1)}%. Chỉ số đa dạng Shannon H' = {_vn(div['H'], 2)}, "
                f"độ đồng đều Pielou J' = {_vn(div['J'], 2)}, Simpson (1 - D) = {_vn(div['simpson'], 2)}.")
        if canopy:
            remarks.append(
                f"Độ che phủ tán đạt {_vn(canopy.get('cover_pct'), 1)}% với chỉ số chồng lấn tán "
                f"CI = {_vn(canopy.get('overlap_index'), 2)}.")
    return {
        "summary": summary, "comp": comp, "div": div, "formula": formula, "vol": vol,
        "vstruct": vstruct, "canopy": canopy, "h_rows": h_rows, "v_rows": v_rows,
        "remarks": remarks, "form_factor": f,
    }


# ------------------------------------------------------------------------------
# BÁO CÁO PDF (A4, nhiều trang) - vẽ hoàn toàn bằng matplotlib
# ------------------------------------------------------------------------------
def _wrap(text, width_in, fontsize, factor=0.53):
    import textwrap
    chars = max(12, int(width_in * 72.0 / (fontsize * factor)))
    out = []
    for para in str(text).split("\n"):
        out.extend(textwrap.wrap(para, chars) or [""])
    return out


class _PdfPage:
    """Một trang báo cáo: lớp phủ 'inch' (0..W, 0..H) để vẽ khung, chữ, bảng."""

    def __init__(self, fig, ctx, title=None):
        self.fig = fig
        self.ctx = ctx
        self.W, self.H = fig.get_size_inches()
        self.M = 0.6  # lề trái/phải (inch)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_xlim(0, self.W)
        ax.set_ylim(0, self.H)
        ax.set_axis_off()
        ax.set_zorder(-10)
        self.ax = ax
        self.y = self.H - 0.95
        if title is not None:
            self._chrome(title)

    # --- khung trang: thanh tiêu đề + chân trang ---
    def _chrome(self, title):
        import matplotlib.patches as mp
        W, H, ax = self.W, self.H, self.ax
        ax.add_patch(mp.Rectangle((0, H - 0.42), W, 0.42, color=RPT_GREEN, lw=0))
        ax.add_patch(mp.Rectangle((0, H - 0.46), W, 0.04, color=RPT_ACCENT, lw=0))
        ax.text(self.M, H - 0.21, "TDForest  |  BÁO CÁO CẤU TRÚC LÂM PHẦN", color="white",
                fontsize=8.5, fontweight="bold", va="center")
        right = self.ctx["meta"].get("otc") or ""
        right = f"OTC: {right}" if right else self.ctx["meta"].get("location", "")
        ax.text(W - self.M, H - 0.21, right, color="#dcefdc", fontsize=8, va="center", ha="right")
        # chân trang
        ax.plot([self.M, W - self.M], [0.55, 0.55], color=RPT_LINE, lw=0.8)
        ax.text(self.M, 0.36, "Khoa Lâm nghiệp - Trường Đại học Nông Lâm TP. Hồ Chí Minh",
                fontsize=7.2, color=RPT_MUTED, va="center")
        ax.text(W / 2, 0.36, f"TDForest v{self.ctx['version']}", fontsize=7.2, color=RPT_MUTED,
                va="center", ha="center")
        ax.text(W - self.M, 0.36, f"Trang {self.ctx['page']} / {self.ctx['total']}", fontsize=7.6,
                color=RPT_GREEN, fontweight="bold", va="center", ha="right")
        # tiêu đề mục
        ax.text(self.M, H - 0.95, title, fontsize=14.5, fontweight="bold", color=RPT_GREEN, va="center")
        ax.add_patch(mp.Rectangle((self.M, H - 1.2), 0.9, 0.05, color=RPT_ACCENT, lw=0))
        self.y = H - 1.5

    def heading(self, text, size=10.5):
        import matplotlib.patches as mp
        self.ax.add_patch(mp.Rectangle((self.M, self.y - 0.2), 0.06, 0.24, color=RPT_GREEN_MID, lw=0))
        self.ax.text(self.M + 0.14, self.y - 0.08, text, fontsize=size, fontweight="bold",
                     color=RPT_TEXT, va="center")
        self.y -= 0.42

    def paragraph(self, text, size=9, width=None, color=RPT_TEXT, line_h=None, x=None):
        width = width or (self.W - 2 * self.M)
        x = self.M if x is None else x
        line_h = line_h or size / 72.0 * 1.55
        for line in _wrap(text, width, size):
            self.ax.text(x, self.y, line, fontsize=size, color=color, va="top")
            self.y -= line_h
        return self.y

    def box(self, x, y_top, w, h, fill=RPT_GREEN_LIGHT, edge=None, accent=RPT_GREEN_MID, radius=0.06):
        import matplotlib.patches as mp
        self.ax.add_patch(mp.FancyBboxPatch((x, y_top - h), w, h, boxstyle=f"round,pad=0,rounding_size={radius}",
                                            fc=fill, ec=edge or fill, lw=0.8))
        if accent:
            self.ax.add_patch(mp.Rectangle((x, y_top - h), 0.06, h, color=accent, lw=0))

    def table(self, x, y_top, widths, header, rows, aligns=None, size=8.4, row_h=0.26,
              highlight=None, total_row=False):
        """Bảng kẻ đẹp: header xanh đậm chữ trắng, dòng kẻ mảnh, sọc xen kẽ."""
        import matplotlib.patches as mp
        ax = self.ax
        tw = sum(widths)
        aligns = aligns or (["left"] + ["right"] * (len(widths) - 1))
        hh = row_h * 1.25
        ax.add_patch(mp.Rectangle((x, y_top - hh), tw, hh, color=RPT_GREEN, lw=0))
        cx = x
        for w, txt, al in zip(widths, header, aligns):
            tx = cx + 0.08 if al == "left" else (cx + w - 0.08 if al == "right" else cx + w / 2)
            ax.text(tx, y_top - hh / 2, txt, color="white", fontsize=size, fontweight="bold",
                    ha=al, va="center")
            cx += w
        y = y_top - hh
        for i, row in enumerate(rows):
            is_total = total_row and i == len(rows) - 1
            if is_total:
                ax.add_patch(mp.Rectangle((x, y - row_h), tw, row_h, color="#dcebd8", lw=0))
            elif highlight is not None and highlight(i):
                ax.add_patch(mp.Rectangle((x, y - row_h), tw, row_h, color="#eaf4e6", lw=0))
            elif i % 2 == 1:
                ax.add_patch(mp.Rectangle((x, y - row_h), tw, row_h, color=RPT_ZEBRA, lw=0))
            cx = x
            for w, txt, al in zip(widths, row, aligns):
                tx = cx + 0.08 if al == "left" else (cx + w - 0.08 if al == "right" else cx + w / 2)
                bold = is_total or (highlight is not None and highlight(i) and cx == x + widths[0])
                ax.text(tx, y - row_h / 2, str(txt), color=RPT_TEXT, fontsize=size,
                        fontweight="bold" if bold else "normal", ha=al, va="center")
                cx += w
            y -= row_h
            ax.plot([x, x + tw], [y, y], color=RPT_LINE, lw=0.5)
        ax.plot([x, x + tw], [y, y], color=RPT_GREEN, lw=1.0)
        return y

    def kpi(self, x, y_top, w, h, value, label):
        self.box(x, y_top, w, h, fill="#ffffff", edge=RPT_LINE, accent=RPT_GREEN_MID)
        self.ax.text(x + 0.2, y_top - h * 0.42, value, fontsize=15, fontweight="bold",
                     color=RPT_GREEN, va="center")
        self.ax.text(x + 0.2, y_top - h * 0.78, label, fontsize=7.6, color=RPT_MUTED, va="center")

    def chart_axes(self, x, y_bottom, w, h, projection=None):
        W, H = self.W, self.H
        return self.fig.add_axes([x / W, y_bottom / H, w / W, h / H], projection=projection)

    def caption(self, text, y):
        self.ax.text(self.W / 2, y, text, fontsize=8.6, style="italic", color=RPT_MUTED,
                     ha="center", va="top")


def _rpt_cover(fig, ctx):
    import matplotlib.patches as mp
    from matplotlib.colors import LinearSegmentedColormap
    pg = _PdfPage(fig, ctx, title=None)
    ax, W, H, M = pg.ax, pg.W, pg.H, pg.M
    meta, data = ctx["meta"], ctx["data"]
    band_h = 4.2
    cmap = LinearSegmentedColormap.from_list("rptg", [RPT_GREEN_DARK, RPT_GREEN_MID])
    grad = np.linspace(0, 1, 256).reshape(1, -1)
    ax.imshow(grad, extent=(0, W, H - band_h, H), cmap=cmap, aspect="auto", zorder=0)
    for cx, cy, r, a in ((W - 1.0, H - 1.0, 2.2, 0.06), (W - 0.3, H - 3.6, 1.4, 0.05), (0.6, H - 4.4, 1.0, 0.04)):
        ax.add_patch(mp.Circle((cx, cy), r, color="white", alpha=a, lw=0, zorder=1))
    # biểu tượng cây
    tx, ty = W - 1.75, H - 1.95
    ax.add_patch(mp.Circle((tx, ty), 0.78, fc=(1, 1, 1, 0.12), ec=(1, 1, 1, 0.55), lw=1.4, zorder=2))
    for dy, half in ((0.38, 0.30), (0.12, 0.40), (-0.14, 0.50)):
        ax.add_patch(mp.Polygon([(tx - half, ty + dy - 0.30), (tx + half, ty + dy - 0.30), (tx, ty + dy + 0.18)],
                                closed=True, fc="white", ec="none", alpha=0.92, zorder=3))
    ax.add_patch(mp.Rectangle((tx - 0.06, ty - 0.62), 0.12, 0.2, fc="white", alpha=0.92, zorder=3))
    ax.text(M, H - 0.65, "KHOA LÂM NGHIỆP  ·  TRƯỜNG ĐẠI HỌC NÔNG LÂM TP. HỒ CHÍ MINH",
            color="#d9eed9", fontsize=8.6, fontweight="bold", va="center", zorder=4)
    ax.add_patch(mp.Rectangle((M, H - 0.9), 1.1, 0.045, color=RPT_ACCENT, lw=0, zorder=4))
    title = meta.get("title") or "BÁO CÁO ĐIỀU TRA CẤU TRÚC LÂM PHẦN"
    ty_txt = H - 1.65
    for line in _wrap(title.upper(), W - 2 * M - 2.2, 21, factor=0.68):
        ax.text(M, ty_txt, line, color="white", fontsize=21, fontweight="bold", va="center", zorder=4)
        ty_txt -= 0.48
    if meta.get("otc"):
        ax.text(M, ty_txt - 0.15, f"Ô tiêu chuẩn: {meta['otc']}", color="#f1f8f1", fontsize=12.5,
                va="center", zorder=4)
        ty_txt -= 0.45
    ax.text(M, ty_txt - 0.2, "Cấu trúc ngang  ·  Cấu trúc đứng  ·  Tổ thành loài  ·  Đa dạng sinh học",
            color="#bfe3bf", fontsize=9, va="center", zorder=4)
    # thẻ thông tin
    s = data["summary"]
    fw, fh = ctx["style"].get("field_w"), ctx["style"].get("field_h")
    info = [
        ("Địa điểm", meta.get("location") or "-"),
        ("Trạng thái rừng", meta.get("forest_state") or "-"),
        ("Kích thước / diện tích OTC", f"{_vn(fw, 0)} × {_vn(fh, 0)} m  ({_vn(s.get('area_ha'), 3)} ha)"),
        ("Số cây / số loài", f"{s.get('n', 0)} cây  /  {data['div']['S']} loài"),
        ("Ngày điều tra", meta.get("survey_date") or "-"),
        ("Người thực hiện", meta.get("author") or "-"),
        ("MSSV / Lớp", meta.get("student_id") or "-"),
        ("Giảng viên hướng dẫn", meta.get("supervisor") or "-"),
    ]
    top = H - band_h - 0.45
    row_h = 0.33
    card_h = row_h * len(info) + 0.3
    pg.box(M, top, W - 2 * M, card_h, fill="#ffffff", edge=RPT_LINE, accent=RPT_GREEN_MID)
    y = top - 0.33
    for i, (k, v) in enumerate(info):
        ax.text(M + 0.3, y, k, fontsize=9, color=RPT_MUTED, va="center")
        ax.text(M + 2.55, y, v, fontsize=9.6, color=RPT_TEXT, fontweight="bold", va="center")
        if i < len(info) - 1:
            ax.plot([M + 0.3, W - M - 0.25], [y - row_h / 2, y - row_h / 2], color="#edf2ec", lw=0.7)
        y -= row_h
    # KPI
    y_k = top - card_h - 0.45
    ax.text(M, y_k, "CHỈ TIÊU NỔI BẬT", fontsize=10, fontweight="bold", color=RPT_GREEN, va="center")
    y_k -= 0.25
    vs, cov = data["vstruct"], data["canopy"]
    kpis = [
        (_vn(s.get("density_ha"), 0), "Mật độ (cây/ha)"),
        (_vn(s.get("g_per_ha"), 2), "Tiết diện ngang G (m²/ha)"),
        (_vn(data["vol"]["M_ha"], 1), "Trữ lượng M (m³/ha)"),
        (_vn(s.get("dg"), 1), "Đường kính Dg (cm)"),
        (_vn(vs["h_mean"], 1) if vs else "-", "Hvn bình quân (m)"),
        (_vn(cov.get("cover_pct"), 1) + "%" if cov else "-", "Độ che phủ tán"),
        (str(data["div"]["S"]), "Số loài (S)"),
        (_vn(data["div"]["H"], 2), "Chỉ số Shannon H'"),
    ]
    gap = 0.15
    cw = (W - 2 * M - 3 * gap) / 4
    ch = 0.82
    for i, (val, lab) in enumerate(kpis):
        r, c = divmod(i, 4)
        pg.kpi(M + c * (cw + gap), y_k - r * (ch + gap), cw, ch, val, lab)
    ax.plot([M, W - M], [0.95, 0.95], color=RPT_LINE, lw=0.8)
    ax.text(W / 2, 0.72, f"Báo cáo được lập ngày {datetime.now().strftime('%d/%m/%Y')} bằng phần mềm "
            f"TDForest v{ctx['version']}", fontsize=8, color=RPT_MUTED, ha="center", va="center")
    ax.text(W / 2, 0.48, "Tài liệu phục vụ học tập, giảng dạy và nghiên cứu khoa học - Khoa Lâm nghiệp, "
            "Trường Đại học Nông Lâm TP. Hồ Chí Minh", fontsize=7.2, color=RPT_MUTED, ha="center", va="center")


def _fmt_rows(rows):
    return [(lab, sym, _vn(val, nd) if isinstance(val, (int, float, np.floating, np.integer)) else "-",
             unit) for lab, sym, val, unit, nd in rows]


def _rpt_summary(fig, ctx):
    pg = _PdfPage(fig, ctx, "1. CHỈ TIÊU CẤU TRÚC LÂM PHẦN")
    data, meta = ctx["data"], ctx["meta"]
    M, W = pg.M, pg.W
    full = W - 2 * M
    gap = 0.22
    half = (full - gap) / 2
    widths = [half * 0.56, half * 0.17, half * 0.27]
    def _rows(rows):
        return [(lab, sym, (_vn(val, nd) + (" " + unit if unit else ""))
                 if isinstance(val, (int, float, np.floating, np.integer)) else "-")
                for lab, sym, val, unit, nd in rows]
    y_head = pg.y
    pg.heading("1.1. Cấu trúc ngang")
    y_tab = pg.y
    yl = pg.table(M, y_tab, widths, ["Chỉ tiêu", "Ký hiệu", "Giá trị"], _rows(data["h_rows"]),
                  aligns=["left", "center", "right"], size=7.9, row_h=0.25)
    yr = y_tab
    if data["v_rows"]:
        xr = M + half + gap
        pg.ax.add_patch(__import__("matplotlib.patches", fromlist=["Rectangle"]).Rectangle(
            (xr, y_head - 0.2), 0.06, 0.24, color=RPT_GREEN_MID, lw=0))
        pg.ax.text(xr + 0.14, y_head - 0.08, "1.2. Cấu trúc đứng (theo Hvn)", fontsize=10.5,
                   fontweight="bold", color=RPT_TEXT, va="center")
        yr = pg.table(xr, y_tab, widths, ["Chỉ tiêu", "Ký hiệu", "Giá trị"], _rows(data["v_rows"]),
                      aligns=["left", "center", "right"], size=7.9, row_h=0.25)
    pg.y = min(yl, yr) - 0.3
    vs = data["vstruct"]
    if vs is not None:
        pg.paragraph(f"Phân tầng theo phương pháp Htb ± k·σ (k = {_vn(ctx['style'].get('vstruct_k_std', 1.0), 1)}): "
                     f"A1 khi Hvn > {_vn(vs['upper'], 1)} m; A2 khi {_vn(vs['lower'], 1)} ≤ Hvn ≤ {_vn(vs['upper'], 1)} m; "
                     f"A3 khi Hvn < {_vn(vs['lower'], 1)} m.", size=7.8, color=RPT_MUTED)
        pg.y -= 0.12
    if not data["canopy"]:
        pg.paragraph("* Độ che phủ tán và chỉ số chồng lấn chưa được tính (cần nhập kích thước ô tiêu chuẩn).",
                     size=7.8, color=RPT_MUTED)
        pg.y -= 0.12
    # nhận xét
    remarks = list(data["remarks"])
    if meta.get("note"):
        remarks.append("Ghi chú của người thực hiện: " + meta["note"])
    if remarks:
        pg.heading("1.3. Nhận xét chung")
        lines = []
        for r in remarks:
            lines.extend(["•  " + l if j == 0 else "    " + l for j, l in enumerate(_wrap(r, full - 0.5, 8.8))])
        line_h = 8.8 / 72.0 * 1.6
        h = len(lines) * line_h + 0.3
        pg.box(M, pg.y + 0.05, full, h, fill="#f4f9f2", edge="#dbe8d6")
        yy = pg.y - 0.12
        for l in lines:
            pg.ax.text(M + 0.25, yy, l, fontsize=8.8, color=RPT_TEXT, va="top")
            yy -= line_h
        pg.y = pg.y - h - 0.3
    # ký tên
    if pg.y > 1.9:
        y0 = max(pg.y - 0.1, 1.95)
        for cx, role, name in ((M + full * 0.22, "NGƯỜI THỰC HIỆN", meta.get("author")),
                               (M + full * 0.78, "GIẢNG VIÊN HƯỚNG DẪN", meta.get("supervisor"))):
            pg.ax.text(cx, y0, role, fontsize=9, fontweight="bold", color=RPT_TEXT, ha="center", va="top")
            pg.ax.text(cx, y0 - 0.2, "(Ký, ghi rõ họ tên)", fontsize=7.8, style="italic", color=RPT_MUTED,
                       ha="center", va="top")
            if name:
                pg.ax.text(cx, y0 - 1.05, name, fontsize=9.5, fontweight="bold", color=RPT_GREEN,
                           ha="center", va="top")


def _rpt_composition(fig, ctx):
    pg = _PdfPage(fig, ctx, "2. TỔ THÀNH LOÀI & ĐA DẠNG SINH HỌC")
    data = ctx["data"]
    M, W = pg.M, pg.W
    full = W - 2 * M
    pg.heading("2.1. Công thức tổ thành (theo chỉ số IV%)")
    lines = _wrap(data["formula"], full - 0.5, 11, factor=0.64)
    h = len(lines) * 0.26 + 0.32
    pg.box(M, pg.y + 0.05, full, h, fill=RPT_GREEN_LIGHT, edge="#cfe5cb")
    yy = pg.y - 0.12
    for l in lines:
        pg.ax.text(M + 0.28, yy, l, fontsize=11, fontweight="bold", color=RPT_GREEN, va="top")
        yy -= 0.26
    pg.y -= h + 0.12
    pg.paragraph("IV% = (N% + G%) / 2. Loài có IV% ≥ 5% được xem là loài ưu thế và ghi vào công thức "
                 "với hệ số IV%/10; các loài còn lại gộp chung là Lk (loài khác).", size=7.8, color=RPT_MUTED)
    pg.y -= 0.2
    pg.heading("2.2. Chỉ số đa dạng sinh học")
    d = data["div"]
    cards = [(str(d["S"]), "Số loài S"), (_vn(d["H"], 3), "Shannon H'"), (_vn(d["J"], 3), "Pielou J'"),
             (_vn(d["simpson"], 3), "Simpson (1 - D)"), (_vn(d["margalef"], 3), "Margalef d")]
    gap = 0.14
    cw = (full - gap * 4) / 5
    for i, (v, lab) in enumerate(cards):
        pg.kpi(M + i * (cw + gap), pg.y, cw, 0.85, v, lab)
    pg.y -= 1.2
    pg.heading("2.3. Biểu đồ tổ thành loài")
    chart_h = max(2.6, pg.y - 1.25)
    ax = pg.chart_axes(M + 1.3, pg.y - chart_h + 0.55, full - 1.45, chart_h - 0.75)
    draw_species_composition(ax, data["comp"], ctx["color_map"], max_species=15)
    ctx["fig_no"] += 1
    pg.caption(f"Hình {ctx['fig_no']}. Tổ thành loài theo chỉ số giá trị quan trọng IV%", pg.y - chart_h + 0.05)


def _comp_rows(comp):
    return [[int(r["STT"]), str(r["Loai"]), int(r["N"]), _vn(r["N%"], 2), _vn(r["G"], 4), _vn(r["G%"], 2),
             _vn(r["IV%"], 2), _vn(r["D13_tb"], 1), _vn(r["Hvn_tb"], 1)] for _, r in comp.iterrows()]


def _rpt_comp_table(fig, ctx, chunk, part, last):
    title = "2.4. BẢNG TỔ THÀNH LOÀI" + (f" (tiếp theo)" if part > 1 else "")
    pg = _PdfPage(fig, ctx, title)
    M, W = pg.M, pg.W
    full = W - 2 * M
    comp = ctx["data"]["comp"]
    widths = [0.45, full - 0.45 - 7 * 0.82, 0.62, 0.82, 0.9, 0.82, 0.82, 0.82, 0.82]
    widths[1] = full - sum(widths[:1] + widths[2:])
    rows = _comp_rows(chunk)
    if last:
        rows.append(["", "Tổng cộng", int(comp["N"].sum()), "100,00", _vn(comp["G"].sum(), 4), "100,00",
                     "100,00", "", ""])
    dom = set(chunk.index[chunk["IV%"] >= IV_DOMINANT_THRESHOLD] - chunk.index[0])
    pg.table(M, pg.y, widths, ["STT", "Loài cây", "N (cây)", "N%", "G (m²)", "G%", "IV%", "D1.3 tb", "Hvn tb"],
             rows, aligns=["center", "left"] + ["right"] * 7, size=8.2, row_h=0.25,
             highlight=lambda i: i in dom, total_row=last)


def _rpt_view(fig, ctx, kind):
    titles = {"plan": "3. TRẮC ĐỒ NGANG", "profile": "4. TRẮC ĐỒ DỌC", "view3d": "5. MÔ PHỎNG 3D LÂM PHẦN"}
    caps = {"plan": "Trắc đồ ngang (hình chiếu bằng tán cây) ô tiêu chuẩn",
            "profile": f"Trắc đồ dọc - mặt cắt đứng chiếu theo trục {ctx['style'].get('profile_axis', 'X')}",
            "view3d": "Mô phỏng 3D cấu trúc lâm phần"}
    pg = _PdfPage(fig, ctx, titles[kind])
    M, W = pg.M, pg.W
    top, bottom = pg.y, 1.15
    df, cm, st = ctx["df"], ctx["color_map"], ctx["style"]
    if kind == "view3d":
        ax = pg.chart_axes(M, bottom, W - 2 * M, top - bottom, projection="3d")
        draw_3d_view(ax, df, cm, st)
        elev, azim = ctx["meta"].get("view3d") or (22, -60)
        ax.view_init(elev=elev, azim=azim)
    else:
        ax = pg.chart_axes(M + 0.1, bottom, W - 2 * M - 0.2, top - bottom)
        (draw_plan_view if kind == "plan" else draw_profile_view)(ax, df, cm, st)
    ctx["fig_no"] += 1
    otc = ctx["meta"].get("otc")
    pg.caption(f"Hình {ctx['fig_no']}. {caps[kind]}{(' ' + otc) if otc and kind != 'view3d' else ''}", 0.95)


def _rpt_charts(fig, ctx, items):
    pg = _PdfPage(fig, ctx, "6. BIỂU ĐỒ PHÂN BỐ & PHÂN TẦNG" + (" (tiếp theo)" if ctx.get("_charts_cont") else ""))
    ctx["_charts_cont"] = True
    M, W = pg.M, pg.W
    top, bottom = pg.y, 1.0
    slot = (top - bottom) / 2
    for i, (fn, cap) in enumerate(items):
        y_b = top - (i + 1) * slot + 1.05
        ax = pg.chart_axes(M + 0.45, y_b, W - 2 * M - 0.6, slot - 1.3)
        fn(ax)
        ctx["fig_no"] += 1
        pg.caption(f"Hình {ctx['fig_no']}. {cap}", y_b - 0.72)


def _rpt_appendix(fig, ctx, chunk, part, last, layer_map, v_i):
    title = "PHỤ LỤC. BẢNG SỐ LIỆU ĐO ĐẾM CÂY" + (" (tiếp theo)" if part > 1 else "")
    pg = _PdfPage(fig, ctx, title)
    M, W = pg.M, pg.W
    full = W - 2 * M
    widths = [0.5, 0, 0.62, 0.58, 0.58, 0.62, 0.58, 0.58, 0.72, 0.72, 0.5]
    widths[1] = full - sum(widths)
    rows = []
    for idx, r in chunk.iterrows():
        dt = np.nanmean([r["DtanD"], r["DtanT"], r["DtanN"], r["DtanB"]])
        g = np.pi / 4 * (float(r["D13"]) / 100) ** 2
        rows.append([r["TT"], r["Loai"], _vn(r["D13"], 1), _vn(r["Hvn"], 1), _vn(r["Hdc"], 1), _vn(dt, 2),
                     _vn(r["X"], 1), _vn(r["Y"], 1), _vn(g, 4), _vn(v_i[idx], 4), layer_map.get(idx, "")])
    if last:
        df = ctx["df"]
        g_all = np.pi / 4 * (df["D13"].astype(float) / 100) ** 2
        rows.append(["", f"Tổng: {len(df)} cây", "", "", "", "", "", "", _vn(g_all.sum(), 4),
                     _vn(float(np.sum(v_i)), 4), ""])
    pg.table(M, pg.y, widths, ["TT", "Loài cây", "D1.3", "Hvn", "Hdc", "Dtán", "X", "Y", "G (m²)", "V (m³)", "Tầng"],
             rows, aligns=["center", "left"] + ["right"] * 8 + ["center"], size=7.6, row_h=0.235, total_row=last)
    if part == 1:
        pg.ax.text(M, 0.75, "Đơn vị: D1.3 (cm); Hvn, Hdc, Dtán (đường kính tán bình quân 4 hướng), X, Y (m); "
                   "V = G · Hvn · f.", fontsize=7.2, color=RPT_MUTED, va="center")


def build_pdf_report(path, df, style, color_map, meta, sections, canopy=None, vstruct=None,
                     version="1.0.0"):
    """Xuất báo cáo PDF khổ A4 nhiều trang. `sections`: tập các khoá trong REPORT_SECTIONS."""
    import logging
    import matplotlib
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_pdf import PdfPages
    logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
    data = _collect_report_data(df, style, meta, canopy, vstruct)
    ctx = {"df": df, "style": style, "color_map": color_map, "meta": meta, "data": data,
           "version": version, "fig_no": 0, "page": 0, "total": 0}
    pages = []
    if "cover" in sections:
        pages.append((A4_PORTRAIT, _rpt_cover))
    if "summary" in sections:
        pages.append((A4_PORTRAIT, _rpt_summary))
    if "composition" in sections and not data["comp"].empty:
        pages.append((A4_PORTRAIT, _rpt_composition))
        comp = data["comp"]
        per = 34
        chunks = [comp.iloc[i:i + per] for i in range(0, len(comp), per)]
        for k, ch in enumerate(chunks, 1):
            pages.append((A4_PORTRAIT, lambda f, c, ch=ch, k=k, last=(k == len(chunks)):
                          _rpt_comp_table(f, c, ch, k, last)))
    fw, fh = style.get("field_w") or 1, style.get("field_h") or 1
    if "plan" in sections:
        pages.append((A4_LANDSCAPE if fw > fh * 1.15 else A4_PORTRAIT, lambda f, c: _rpt_view(f, c, "plan")))
    if "profile" in sections:
        pages.append((A4_LANDSCAPE, lambda f, c: _rpt_view(f, c, "profile")))
    if "view3d" in sections:
        pages.append((A4_LANDSCAPE, lambda f, c: _rpt_view(f, c, "view3d")))
    if "charts" in sections:
        bw = style.get("struct_bin_width", 4.0)
        vbw = style.get("vstruct_bin_width", 2.0)
        vs = data["vstruct"]
        items = [
            (lambda ax: draw_nd_distribution(ax, df, style, bw), f"Phân bố số cây theo cỡ kính N/D (cự ly cỡ kính {_vn(bw, 0)} cm)"),
            (lambda ax: draw_ndg_distribution(ax, df, style), "Phân bố số cây theo cấp kính tương đối N/Dg"),
            (lambda ax: draw_g_distribution(ax, df, style, bw), "Phân bố tiết diện ngang G theo cỡ kính"),
            (lambda ax: draw_nh_distribution(ax, df, style, vbw, vs), f"Phân bố số cây theo cấp chiều cao N/H (cự ly {_vn(vbw, 0)} m)"),
        ]
        if canopy:
            items.append((lambda ax: draw_canopy_coverage(ax, df, style, canopy), "Bản đồ độ che phủ & chồng lấn tán"))
        items.append((lambda ax: draw_forest_layers(ax, df, style, vs), "Sơ đồ phân tầng đứng lâm phần"))
        for i in range(0, len(items), 2):
            pages.append((A4_PORTRAIT, lambda f, c, it=items[i:i + 2]: _rpt_charts(f, c, it)))
    if "appendix" in sections:
        vs = data["vstruct"]
        layer_map = {i: str(l) for i, l in enumerate(vs["_layer"])} if vs is not None else {}
        v_i = data["vol"]["v_i"]
        per = 36
        dfr = df.reset_index(drop=True)
        chunks = [dfr.iloc[i:i + per] for i in range(0, len(dfr), per)]
        for k, ch in enumerate(chunks, 1):
            pages.append((A4_PORTRAIT, lambda f, c, ch=ch, k=k, last=(k == len(chunks)):
                          _rpt_appendix(f, c, ch, k, last, layer_map, v_i)))
    if not pages:
        raise ValueError("Chưa chọn nội dung nào để xuất báo cáo.")
    ctx["total"] = len(pages)
    rc = {"font.family": "sans-serif", "font.sans-serif": REPORT_FONTS + list(matplotlib.rcParams["font.sans-serif"]),
          "axes.unicode_minus": False, "pdf.fonttype": 42}
    with matplotlib.rc_context(rc):
        with PdfPages(path) as pdf:
            info = pdf.infodict()
            info["Title"] = meta.get("title") or "Báo cáo điều tra cấu trúc lâm phần"
            info["Author"] = meta.get("author") or ""
            info["Subject"] = f"OTC {meta.get('otc', '')} - Khoa Lâm nghiệp, ĐH Nông Lâm TP.HCM"
            info["Creator"] = f"TDForest v{version}"
            for i, (size, builder) in enumerate(pages, 1):
                ctx["page"] = i
                fig = Figure(figsize=size, facecolor="white")
                builder(fig, ctx)
                pdf.savefig(fig)
    return len(pages)


# ------------------------------------------------------------------------------
# BÁO CÁO EXCEL (nhiều sheet, định dạng sẵn, có biểu đồ Excel gốc)
# ------------------------------------------------------------------------------
def build_excel_report(path, df, style, meta, canopy=None, vstruct=None, version="1.0.0"):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.chart import BarChart, Reference
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise RuntimeError("Thiếu thư viện openpyxl.\nHãy cài đặt: pip install openpyxl")
    data = _collect_report_data(df, style, meta, canopy, vstruct)
    wb = Workbook()
    F_TITLE = Font(name="Calibri", size=16, bold=True, color="1B5E20")
    F_SUB = Font(name="Calibri", size=10, italic=True, color="66756A")
    F_H = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    F_SEC = Font(name="Calibri", size=12, bold=True, color="1B5E20")
    F_B = Font(name="Calibri", size=11, bold=True, color="26332A")
    F_N = Font(name="Calibri", size=11, color="26332A")
    FILL_H = PatternFill("solid", fgColor="2E7D32")
    FILL_Z = PatternFill("solid", fgColor="F3F8F2")
    FILL_T = PatternFill("solid", fgColor="DCEBD8")
    FILL_HL = PatternFill("solid", fgColor="E8F5E9")
    thin = Side(style="thin", color="CFDCCB")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def header_block(ws, title, ncols):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        ws.cell(1, 1, title).font = F_TITLE
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
        sub = "Khoa Lâm nghiệp - Trường Đại học Nông Lâm TP. Hồ Chí Minh"
        if meta.get("otc"):
            sub = f"Ô tiêu chuẩn: {meta['otc']}   |   " + sub
        ws.cell(2, 1, sub).font = F_SUB
        ws.row_dimensions[1].height = 26
        return 4

    def table(ws, r0, headers, rows, fmts=None, total=False, hl=None):
        for j, h in enumerate(headers, 1):
            c = ws.cell(r0, j, h)
            c.font, c.fill, c.alignment, c.border = F_H, FILL_H, CENTER, BORDER
        ws.row_dimensions[r0].height = 22
        for i, row in enumerate(rows, 1):
            is_tot = total and i == len(rows)
            for j, v in enumerate(row, 1):
                if isinstance(v, (np.floating, np.integer)):
                    v = v.item()
                if isinstance(v, float) and not np.isfinite(v):
                    v = None
                c = ws.cell(r0 + i, j, v)
                c.border = BORDER
                c.font = F_B if is_tot or (hl and hl(i - 1) and j <= 2) else F_N
                if is_tot:
                    c.fill = FILL_T
                elif hl and hl(i - 1):
                    c.fill = FILL_HL
                elif i % 2 == 0:
                    c.fill = FILL_Z
                if fmts and j - 1 < len(fmts) and fmts[j - 1] and isinstance(v, (int, float)):
                    c.number_format = fmts[j - 1]
                    c.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    c.alignment = Alignment(vertical="center", wrap_text=True)
        return r0 + len(rows) + 1

    def widths(ws, ws_w):
        for j, w in enumerate(ws_w, 1):
            ws.column_dimensions[get_column_letter(j)].width = w

    def page(ws, landscape=False):
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "landscape" if landscape else "portrait"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_options.horizontalCentered = True
        ws.oddFooter.center.text = "TDForest - Trang &P / &N"
        ws.sheet_view.showGridLines = False

    nd_fmt = lambda nd: "#,##0" if nd == 0 else "#,##0." + "0" * nd

    # ---- Sheet 1: Tổng quan ----
    ws = wb.active
    ws.title = "Tổng quan"
    r = header_block(ws, (meta.get("title") or "BÁO CÁO ĐIỀU TRA CẤU TRÚC LÂM PHẦN").upper(), 4)
    s = data["summary"]
    info = [("Địa điểm", meta.get("location")), ("Trạng thái rừng", meta.get("forest_state")),
            ("Kích thước OTC (m)", f"{_vn(style.get('field_w'), 0)} × {_vn(style.get('field_h'), 0)}"),
            ("Ngày điều tra", meta.get("survey_date")), ("Người thực hiện", meta.get("author")),
            ("MSSV / Lớp", meta.get("student_id")), ("Giảng viên hướng dẫn", meta.get("supervisor")),
            ("Ngày lập báo cáo", datetime.now().strftime("%d/%m/%Y"))]
    for k, v in info:
        ws.cell(r, 1, k).font = F_B
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(r, 2, v or "-").font = F_N
        r += 1
    r += 1
    ws.cell(r, 1, "CẤU TRÚC NGANG").font = F_SEC
    r = table(ws, r + 1, ["Chỉ tiêu", "Ký hiệu", "Giá trị", "Đơn vị"],
              [(a, b, c, d) for a, b, c, d, _ in data["h_rows"]],
              fmts=[None, None, None, None]) + 1
    for i, (_, _, _, _, nd) in enumerate(data["h_rows"]):
        ws.cell(r - 1 - len(data["h_rows"]) + i, 3).number_format = nd_fmt(nd)
    if data["v_rows"]:
        ws.cell(r, 1, "CẤU TRÚC ĐỨNG").font = F_SEC
        start = r + 2
        r = table(ws, r + 1, ["Chỉ tiêu", "Ký hiệu", "Giá trị", "Đơn vị"],
                  [(a, b, c, d) for a, b, c, d, _ in data["v_rows"]]) + 1
        for i, (_, _, _, _, nd) in enumerate(data["v_rows"]):
            ws.cell(start + i, 3).number_format = nd_fmt(nd)
    for row in ws.iter_rows(min_row=1, max_row=r, min_col=3, max_col=3):
        for c in row:
            if isinstance(c.value, (int, float)):
                c.alignment = Alignment(horizontal="right")
    remarks = list(data["remarks"]) + (["Ghi chú: " + meta["note"]] if meta.get("note") else [])
    if remarks:
        ws.cell(r, 1, "NHẬN XÉT").font = F_SEC
        r += 1
        for rem in remarks:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            c = ws.cell(r, 1, "• " + rem)
            c.font, c.alignment = F_N, Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 15 * (len(rem) // 95 + 1) + 4
            r += 1
    widths(ws, [42, 12, 16, 12])
    page(ws)

    # ---- Sheet 2: Tổ thành loài ----
    ws = wb.create_sheet("Tổ thành loài")
    r = header_block(ws, "TỔ THÀNH LOÀI & ĐA DẠNG SINH HỌC", 9)
    ws.cell(r, 1, "Công thức tổ thành:").font = F_B
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
    c = ws.cell(r, 2, data["formula"])
    c.font = Font(name="Calibri", size=12, bold=True, color="1B5E20")
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 18 * (len(data["formula"]) // 90 + 1) + 6
    r += 2
    comp = data["comp"]
    rows = [[int(x["STT"]), x["Loai"], int(x["N"]), x["N%"], x["G"], x["G%"], x["IV%"], x["D13_tb"], x["Hvn_tb"]]
            for _, x in comp.iterrows()]
    t0 = r
    rows.append(["", "Tổng cộng", int(comp["N"].sum()), 100.0, float(comp["G"].sum()), 100.0, 100.0, None, None])
    r = table(ws, r, ["STT", "Loài cây", "N (cây)", "N%", "G (m²)", "G%", "IV%", "D1.3 tb (cm)", "Hvn tb (m)"],
              rows, fmts=["0", None, "0", "0.00", "0.0000", "0.00", "0.00", "0.0", "0.0"], total=True,
              hl=lambda i: i < len(comp) and comp.iloc[i]["IV%"] >= IV_DOMINANT_THRESHOLD)
    r += 1
    ws.cell(r, 1, "CHỈ SỐ ĐA DẠNG").font = F_SEC
    d = data["div"]
    r = table(ws, r + 1, ["Chỉ số", "Giá trị"], [
        ("Số loài (S)", d["S"]), ("Shannon H'", d["H"]), ("H' max = ln(S)", d["Hmax"]),
        ("Pielou J'", d["J"]), ("Simpson (1 - D)", d["simpson"]), ("Margalef d", d["margalef"])],
        fmts=[None, "0.000"])
    if len(comp):
        ch = BarChart()
        ch.type = "bar"
        ch.style = 10
        ch.title = "Tổ thành loài theo IV%"
        ch.y_axis.title = "IV%"
        ch.legend = None
        ch.add_data(Reference(ws, min_col=7, min_row=t0, max_row=t0 + len(comp)), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=2, min_row=t0 + 1, max_row=t0 + len(comp)))
        ch.series[0].graphicalProperties.solidFill = "2E7D32"
        ch.height, ch.width = max(7, 0.55 * len(comp) + 3), 16
        ws.add_chart(ch, f"K{t0}")
    widths(ws, [7, 26, 10, 10, 12, 10, 10, 13, 12])
    ws.freeze_panes = ws.cell(t0 + 1, 1)
    page(ws, landscape=True)

    # ---- Sheet 3, 4: phân bố N/D và N/H ----
    s_d13 = pd.to_numeric(df["D13"], errors="coerce").to_numpy(dtype=float)
    g_i = np.pi / 4 * (s_d13 / 100) ** 2
    bw = style.get("struct_bin_width", 4.0)
    edges, counts = compute_nd_distribution(s_d13, bw)
    _, g_sum = compute_g_distribution(s_d13, g_i, bw)
    n_tot = max(1, int(np.sum(counts)))
    g_tot = float(np.sum(g_sum)) or 1.0
    for name, title, unit, edges_, counts_, extra in (
            ("Phân bố N-D", f"PHÂN BỐ SỐ CÂY THEO CỠ KÍNH (N/D) - cự ly {_vn(bw, 0)} cm", "cm", edges, counts, g_sum),
            ("Phân bố N-H", None, "m", None, None, None)):
        if name == "Phân bố N-H":
            hv = pd.to_numeric(df["Hvn"], errors="coerce").to_numpy(dtype=float)
            vbw = style.get("vstruct_bin_width", 2.0)
            edges_, counts_ = compute_nh_distribution(hv, vbw)
            title = f"PHÂN BỐ SỐ CÂY THEO CẤP CHIỀU CAO (N/H) - cự ly {_vn(vbw, 0)} m"
        ws = wb.create_sheet(name)
        r = header_block(ws, title, 7)
        hdr = ["Cấp", f"Giới hạn ({unit})", f"Giữa cấp ({unit})", "N (cây)", "N%"]
        if extra is not None:
            hdr += ["G (m²)", "G%"]
        rows = []
        for i in range(len(counts_)):
            lo, hi = float(edges_[i]), float(edges_[i + 1])
            row = [i + 1, f"{_vn(lo, 0)} - {_vn(hi, 0)}", (lo + hi) / 2, int(counts_[i]), counts_[i] / n_tot * 100]
            if extra is not None:
                row += [float(extra[i]), float(extra[i]) / g_tot * 100]
            rows.append(row)
        tot = ["", "Tổng", None, int(np.sum(counts_)), 100.0] + ([float(np.sum(extra)), 100.0] if extra is not None else [])
        rows.append(tot)
        t0 = r
        r = table(ws, r, hdr, rows, fmts=["0", None, "0.0", "0", "0.00", "0.0000", "0.00"], total=True)
        ch = BarChart()
        ch.type = "col"
        ch.style = 10
        ch.title = title.split(" - ")[0].title()
        ch.y_axis.title = "Số cây (N)"
        ch.x_axis.title = f"Giữa cấp ({unit})"
        ch.legend = None
        ch.gapWidth = 30
        ch.add_data(Reference(ws, min_col=4, min_row=t0, max_row=t0 + len(counts_)), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=3, min_row=t0 + 1, max_row=t0 + len(counts_)))
        ch.series[0].graphicalProperties.solidFill = "43A047" if extra is not None else "00838F"
        ch.height, ch.width = 9, 16
        ws.add_chart(ch, f"I{t0}")
        widths(ws, [7, 16, 14, 10, 10, 12, 10])
        page(ws, landscape=True)

    # ---- Sheet 5: dữ liệu cây ----
    ws = wb.create_sheet("Dữ liệu cây")
    r = header_block(ws, "BẢNG SỐ LIỆU ĐO ĐẾM CÂY", 14)
    vs = data["vstruct"]
    layer = list(vs["_layer"]) if vs is not None else [""] * len(df)
    v_i = data["vol"]["v_i"]
    rows = []
    for i, (_, x) in enumerate(df.reset_index(drop=True).iterrows()):
        dt = float(np.nanmean([x["DtanD"], x["DtanT"], x["DtanN"], x["DtanB"]]))
        rows.append([x["TT"], x["Loai"], x["D13"], x["Hvn"], x["Hdc"], x["DtanD"], x["DtanT"], x["DtanN"],
                     x["DtanB"], dt, x["X"], x["Y"], float(g_i[i]), float(v_i[i]), str(layer[i])])
    rows.append(["", f"Tổng: {len(df)} cây", None, None, None, None, None, None, None, None, None, None,
                 float(np.nansum(g_i)), float(np.sum(v_i)), ""])
    t0 = r
    table(ws, r, ["TT", "Loài cây", "D1.3 (cm)", "Hvn (m)", "Hdc (m)", "Dt Đông", "Dt Tây", "Dt Nam", "Dt Bắc",
                  "Dt tb (m)", "X (m)", "Y (m)", "G (m²)", "V (m³)", "Tầng"], rows,
          fmts=[None, None, "0.0", "0.0", "0.0", "0.00", "0.00", "0.00", "0.00", "0.00", "0.0", "0.0",
                "0.0000", "0.0000", None], total=True)
    ws.freeze_panes = ws.cell(t0 + 1, 3)
    ws.auto_filter.ref = f"A{t0}:O{t0 + len(rows) - 1}"
    widths(ws, [7, 22, 10, 9, 9, 9, 9, 9, 9, 10, 9, 9, 11, 11, 8])
    page(ws, landscape=True)
    wb.properties.title = meta.get("title") or "Báo cáo cấu trúc lâm phần"
    wb.properties.creator = meta.get("author") or f"TDForest v{version}"
    wb.save(path)
    return path


from PySide6.QtCore import QDate
from PySide6.QtWidgets import (
    QGridLayout, QDateEdit, QPlainTextEdit, QRadioButton, QButtonGroup,
    QTableWidget, QTableWidgetItem, QAbstractItemView,
)


class ReportDialog(QDialog):
    """Hộp thoại "📑 Xuất báo cáo": nhập thông tin OTC / người thực hiện, chọn
    nội dung và định dạng (PDF / Excel / cả hai). Các thông tin đã nhập được
    ghi nhớ (QSettings) cho lần xuất sau."""

    _FIELDS = [
        ("title", "Tiêu đề báo cáo", "BÁO CÁO ĐIỀU TRA CẤU TRÚC LÂM PHẦN"),
        ("otc", "Tên / số hiệu OTC", "OTC-01"),
        ("location", "Địa điểm", ""),
        ("forest_state", "Trạng thái rừng", ""),
        ("author", "Người thực hiện", ""),
        ("student_id", "MSSV / Lớp", ""),
        ("supervisor", "Giảng viên hướng dẫn", ""),
    ]

    def __init__(self, parent=None, form_factor=DEFAULT_FORM_FACTOR):
        super().__init__(parent)
        self.setWindowTitle("Xuất báo cáo chuyên nghiệp")
        self.setWindowIcon(make_tree_icon(32))
        self.setMinimumWidth(640)
        s = get_settings()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header = QFrame()
        header.setObjectName("dialogHeader")
        hl = QHBoxLayout(header)
        hl.setContentsMargins(22, 14, 22, 14)
        icon = QLabel()
        icon.setPixmap(make_tree_pixmap(40))
        hl.addWidget(icon)
        tcol = QVBoxLayout()
        t1 = QLabel("Xuất báo cáo chuyên nghiệp")
        t1.setObjectName("dialogHeaderTitle")
        t2 = QLabel("PDF khổ A4 có trang bìa, bảng chỉ tiêu, tổ thành loài, trắc đồ, biểu đồ  ·  "
                    "Excel nhiều sheet có biểu đồ")
        t2.setObjectName("dialogHeaderSub")
        t2.setWordWrap(True)
        tcol.addWidget(t1)
        tcol.addWidget(t2)
        hl.addLayout(tcol, 1)
        outer.addWidget(header)

        body = QWidget()
        bl = QVBoxLayout(body)
        bl.setContentsMargins(20, 16, 20, 12)
        bl.setSpacing(12)

        info_box = QGroupBox("Thông tin báo cáo")
        form = QFormLayout(info_box)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(8)
        self.edits = {}
        for key, label, default in self._FIELDS:
            e = QLineEdit(str(s.value(f"report/{key}", default)))
            e.setPlaceholderText(label)
            self.edits[key] = e
            form.addRow(label + ":", e)
        row = QHBoxLayout()
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("dd/MM/yyyy")
        saved = QDate.fromString(str(s.value("report/survey_date", "")), "dd/MM/yyyy")
        self.date_edit.setDate(saved if saved.isValid() else QDate.currentDate())
        row.addWidget(self.date_edit)
        row.addSpacing(18)
        row.addWidget(QLabel("Hình số f:"))
        self.ff_spin = QDoubleSpinBox()
        self.ff_spin.setRange(0.30, 0.80)
        self.ff_spin.setSingleStep(0.01)
        self.ff_spin.setDecimals(2)
        self.ff_spin.setValue(float(form_factor or DEFAULT_FORM_FACTOR))
        self.ff_spin.setToolTip("Hình số thân cây dùng tính trữ lượng V = G · Hvn · f (mặc định 0,45)")
        row.addWidget(self.ff_spin)
        row.addStretch(1)
        form.addRow("Ngày điều tra:", row)
        self.note_edit = QPlainTextEdit(str(s.value("report/note", "")))
        self.note_edit.setPlaceholderText("Ghi chú / nhận xét thêm của người thực hiện (không bắt buộc)")
        self.note_edit.setFixedHeight(64)
        form.addRow("Ghi chú:", self.note_edit)
        bl.addWidget(info_box)

        sec_box = QGroupBox("Nội dung báo cáo")
        grid = QGridLayout(sec_box)
        grid.setHorizontalSpacing(20)
        self.sec_checks = {}
        saved_secs = s.value("report/sections", "")
        saved_secs = set(str(saved_secs).split(",")) if saved_secs else None
        for i, (key, label) in enumerate(REPORT_SECTIONS):
            cb = QCheckBox(label)
            cb.setChecked(True if saved_secs is None else key in saved_secs)
            self.sec_checks[key] = cb
            grid.addWidget(cb, i // 2, i % 2)
        bl.addWidget(sec_box)

        fmt_box = QGroupBox("Định dạng xuất")
        fl = QGridLayout(fmt_box)
        self.fmt_group = QButtonGroup(self)
        self.fmt_radios = {}
        saved_fmt = str(s.value("report/format", "both"))
        fmt_opts = (("pdf", "📄  PDF (in ấn, nộp bài)"), ("word", "📝  Word (chỉnh sửa, nộp bài)"),
                    ("excel", "📊  Excel (số liệu, biểu đồ)"), ("both", "📦  PDF + Excel"),
                    ("all", "🗂️  Cả 3 định dạng"))
        for idx, (key, label) in enumerate(fmt_opts):
            rb = QRadioButton(label)
            rb.setChecked(key == saved_fmt)
            self.fmt_group.addButton(rb)
            self.fmt_radios[key] = rb
            fl.addWidget(rb, idx // 2, idx % 2)
        bl.addWidget(fmt_box)

        scope = QLabel("🌲  " + "Báo cáo phục vụ học tập, giảng dạy và nghiên cứu - Khoa Lâm nghiệp, "
                       "Trường Đại học Nông Lâm TP. Hồ Chí Minh.")
        scope.setWordWrap(True)
        scope.setObjectName("hintLabel")
        bl.addWidget(scope)

        btns = QHBoxLayout()
        btns.addStretch(1)
        b_cancel = QPushButton("Hủy")
        b_cancel.setObjectName("secondary")
        b_cancel.clicked.connect(self.reject)
        b_ok = QPushButton("📑  Xuất báo cáo")
        b_ok.setObjectName("primaryBig")
        b_ok.setDefault(True)
        b_ok.clicked.connect(self._accept)
        btns.addWidget(b_cancel)
        btns.addWidget(b_ok)
        # Nội dung cuộn được; phần nút bấm luôn cố định ở đáy hộp thoại.
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(body)
        outer.addWidget(scroll, 1)
        footer = QFrame()
        footer.setStyleSheet("QFrame { background:#f7faf6; border-top:1px solid #dfe7dc; }")
        fl2 = QHBoxLayout(footer)
        fl2.setContentsMargins(20, 10, 20, 10)
        fl2.addStretch(1)
        fl2.addWidget(b_cancel)
        fl2.addWidget(b_ok)
        outer.addWidget(footer)
        # Không cao quá 90% màn hình để khỏi tràn ra ngoài.
        avail = QApplication.primaryScreen().availableGeometry()
        max_h = int(avail.height() * 0.9)
        self.setMaximumHeight(max_h)
        self.resize(680, min(self.sizeHint().height(), max_h))

    def _accept(self):
        if not any(cb.isChecked() for cb in self.sec_checks.values()):
            show_notice(self, "warning", "Xuất báo cáo", "Vui lòng chọn ít nhất một nội dung cho báo cáo.")
            return
        s = get_settings()
        for key, e in self.edits.items():
            s.setValue(f"report/{key}", e.text().strip())
        s.setValue("report/survey_date", self.date_edit.date().toString("dd/MM/yyyy"))
        s.setValue("report/note", self.note_edit.toPlainText().strip())
        s.setValue("report/sections", ",".join(k for k, cb in self.sec_checks.items() if cb.isChecked()))
        s.setValue("report/format", self.get_format())
        s.sync()
        self.accept()

    def get_format(self):
        for key, rb in self.fmt_radios.items():
            if rb.isChecked():
                return key
        return "pdf"

    def get_values(self):
        meta = {k: e.text().strip() for k, e in self.edits.items()}
        meta["survey_date"] = self.date_edit.date().toString("dd/MM/yyyy")
        meta["note"] = self.note_edit.toPlainText().strip()
        meta["form_factor"] = float(self.ff_spin.value())
        sections = {k for k, cb in self.sec_checks.items() if cb.isChecked()}
        return meta, sections, self.get_format()



# ==============================================================================
# 10. BÁO CÁO WORD (.docx) - chỉnh sửa được, nộp bài được
#     Cần thư viện python-docx (pip install python-docx)
# ==============================================================================
def build_word_report(path, df, style, color_map, meta, sections, canopy=None, vstruct=None,
                      version="1.0.0"):
    try:
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
    except ImportError:
        raise RuntimeError("Thiếu thư viện python-docx.\nHãy cài đặt: pip install python-docx")
    import io
    import matplotlib
    from matplotlib.figure import Figure

    data = _collect_report_data(df, style, meta, canopy, vstruct)
    GREEN = RGBColor(0x1B, 0x5E, 0x20)
    MUTED = RGBColor(0x66, 0x75, 0x6A)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    AL = {"left": WD_ALIGN_PARAGRAPH.LEFT, "right": WD_ALIGN_PARAGRAPH.RIGHT,
          "center": WD_ALIGN_PARAGRAPH.CENTER}
    otc = meta.get("otc") or ""
    fw, fh = style.get("field_w"), style.get("field_h")
    vs = data["vstruct"]
    fig_no = [0]
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin = sec.bottom_margin = Cm(2)
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)

    def shade(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        tcPr.append(shd)

    def heading(text, level=1):
        h = doc.add_heading(text, level=level)
        for r in h.runs:
            r.font.color.rgb = GREEN
            r.font.name = "Calibri"
        return h

    def para(text="", size=10, bold=False, italic=False, color=None, align=None, space_after=4):
        p = doc.add_paragraph()
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        if color is not None:
            r.font.color.rgb = color
        if align is not None:
            p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        return p

    def caption(text):
        fig_no[0] += 1
        return para(f"Hình {fig_no[0]}. {text}", size=9, italic=True, color=MUTED,
                    align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    def picture(draw, w_in, h_in, cap, three_d=False):
        fig = Figure(figsize=(w_in, h_in), facecolor="white")
        ax = fig.add_subplot(111, projection="3d") if three_d else fig.add_subplot(111)
        draw(ax)
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
        buf.seek(0)
        doc.add_picture(buf, width=Cm(16.5))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption(cap)

    def table(headers, rows, widths, aligns, size=8.5, total=False):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"
        t.autofit = False
        hdr = t.rows[0]
        trPr = hdr._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        trPr.append(th)
        for j, h in enumerate(headers):
            c = hdr.cells[j]
            c.text = ""
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(str(h))
            r.bold = True
            r.font.size = Pt(size)
            r.font.color.rgb = WHITE
            shade(c, "2E7D32")
        for i, row in enumerate(rows):
            cells = t.add_row().cells
            is_tot = total and i == len(rows) - 1
            for j, v in enumerate(row):
                c = cells[j]
                c.text = ""
                p = c.paragraphs[0]
                p.alignment = AL[aligns[j]]
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run("" if v is None else str(v))
                r.font.size = Pt(size)
                r.bold = is_tot
                if is_tot:
                    shade(c, "DCEBD8")
                elif i % 2 == 1:
                    shade(c, "F5F9F4")
        for row in t.rows:
            for j, c in enumerate(row.cells):
                c.width = Cm(widths[j])
        doc.add_paragraph().paragraph_format.space_after = Pt(2)
        return t

    def kv_table(pairs, widths=(5.5, 11.5)):
        t = doc.add_table(rows=0, cols=2)
        t.style = "Table Grid"
        t.autofit = False
        for k, v in pairs:
            cells = t.add_row().cells
            for j, txt in enumerate((k, v)):
                cells[j].text = ""
                p = cells[j].paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(str(txt or "-"))
                r.font.size = Pt(9.5)
                r.bold = (j == 0)
            shade(cells[0], "EAF3E8")
        for row in t.rows:
            for j, c in enumerate(row.cells):
                c.width = Cm(widths[j])
        return t

    def fmt_num(val, nd, unit=""):
        if isinstance(val, (int, float, np.floating, np.integer)):
            return _vn(val, nd, unit)
        return "-"

    def _render():
        # ---------------- Trang bìa ----------------
        para("KHOA LÂM NGHIỆP · TRƯỜNG ĐẠI HỌC NÔNG LÂM TP. HỒ CHÍ MINH", size=9, bold=True, color=GREEN)
        for _ in range(4):
            doc.add_paragraph()
        para((meta.get("title") or "BÁO CÁO ĐIỀU TRA CẤU TRÚC LÂM PHẦN").upper(), size=24, bold=True,
             color=GREEN, space_after=10)
        if otc:
            para(f"Ô tiêu chuẩn: {otc}", size=13, color=MUTED, space_after=16)
        s = data["summary"]
        area_txt = (f"{_vn(fw, 0)} × {_vn(fh, 0)} m ({_vn(s.get('area_ha'), 3)} ha)"
                    if fw and fh else "-")
        kv_table([
            ("Địa điểm", meta.get("location")),
            ("Trạng thái rừng", meta.get("forest_state")),
            ("Kích thước / diện tích OTC", area_txt),
            ("Số cây / số loài", f"{s.get('n', 0)} cây / {data['div']['S']} loài"),
            ("Ngày điều tra", meta.get("survey_date")),
            ("Người thực hiện", meta.get("author")),
            ("MSSV / Lớp", meta.get("student_id")),
            ("Giảng viên hướng dẫn", meta.get("supervisor")),
        ])
        doc.add_paragraph()
        para(f"Báo cáo được lập ngày {datetime.now().strftime('%d/%m/%Y')} bằng TDForest v{version}",
             size=9, color=MUTED)
        para("Tài liệu phục vụ học tập, giảng dạy và nghiên cứu khoa học - Khoa Lâm nghiệp, "
             "Trường Đại học Nông Lâm TP. Hồ Chí Minh", size=8.5, italic=True, color=MUTED)
        doc.add_page_break()

        # ---------------- 1. Chỉ tiêu ----------------
        if "summary" in sections:
            heading("1. Chỉ tiêu cấu trúc lâm phần", 1)
            heading("1.1. Cấu trúc ngang", 2)
            table(["Chỉ tiêu", "Ký hiệu", "Giá trị", "Đơn vị"],
                  [[lab, sym, fmt_num(val, nd), unit] for lab, sym, val, unit, nd in data["h_rows"]],
                  [7.5, 2.5, 3.5, 3.5], ["left", "center", "right", "left"])
            if data["v_rows"]:
                heading("1.2. Cấu trúc đứng (theo chiều cao vút ngọn)", 2)
                table(["Chỉ tiêu", "Ký hiệu", "Giá trị", "Đơn vị"],
                      [[lab, sym, fmt_num(val, nd), unit] for lab, sym, val, unit, nd in data["v_rows"]],
                      [7.5, 2.5, 3.5, 3.5], ["left", "center", "right", "left"])
            if vs is not None:
                para(f"Phân tầng theo Htb ± k·σ (k = {_vn(style.get('vstruct_k_std', 1.0), 1)}): "
                     f"A1 khi Hvn > {_vn(vs['upper'], 1)} m; A2 khi {_vn(vs['lower'], 1)} ≤ Hvn ≤ "
                     f"{_vn(vs['upper'], 1)} m; A3 khi Hvn < {_vn(vs['lower'], 1)} m.",
                     size=8.5, italic=True, color=MUTED)
            if not canopy:
                para("* Độ che phủ tán và chỉ số chồng lấn chưa được tính (cần nhập kích thước ô tiêu chuẩn).",
                     size=8.5, italic=True, color=MUTED)
            heading("1.3. Nhận xét", 2)
            remarks = list(data["remarks"])
            if meta.get("note"):
                remarks.append("Ghi chú của người thực hiện: " + meta["note"])
            for r_txt in remarks:
                doc.add_paragraph(r_txt, style="List Bullet")
            doc.add_paragraph()
            sig = doc.add_table(rows=2, cols=2)
            sig.autofit = False
            for j, role in enumerate(("NGƯỜI THỰC HIỆN", "GIẢNG VIÊN HƯỚNG DẪN")):
                c = sig.rows[0].cells[j]
                c.text = ""
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(role)
                r.bold = True
                r.font.size = Pt(9.5)
                p2 = c.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r2 = p2.add_run("(Ký, ghi rõ họ tên)")
                r2.italic = True
                r2.font.size = Pt(8.5)
                name = meta.get("author") if j == 0 else meta.get("supervisor")
                c2 = sig.rows[1].cells[j]
                c2.text = ""
                p3 = c2.paragraphs[0]
                p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p3.paragraph_format.space_before = Pt(40)
                r3 = p3.add_run(name or "")
                r3.bold = True
                r3.font.color.rgb = GREEN
            doc.add_page_break()

        # ---------------- 2. Tổ thành ----------------
        comp = data["comp"]
        if "composition" in sections and not comp.empty:
            d = data["div"]
            heading("2. Tổ thành loài & đa dạng sinh học", 1)
            heading("2.1. Công thức tổ thành (theo IV%)", 2)
            para(data["formula"], size=11, bold=True, color=GREEN)
            para("IV% = (N% + G%) / 2. Loài có IV% ≥ 5% là loài ưu thế; các loài còn lại gộp chung là Lk.",
                 size=8.5, italic=True, color=MUTED)
            heading("2.2. Bảng tổ thành loài", 2)
            rows = [[int(r["STT"]), str(r["Loai"]), int(r["N"]), _vn(r["N%"], 2), _vn(r["G"], 4),
                     _vn(r["G%"], 2), _vn(r["IV%"], 2), _vn(r["D13_tb"], 1), _vn(r["Hvn_tb"], 1)]
                    for _, r in comp.iterrows()]
            rows.append(["", "Tổng cộng", int(comp["N"].sum()), "100,00", _vn(comp["G"].sum(), 4),
                         "100,00", "100,00", "", ""])
            table(["STT", "Loài cây", "N (cây)", "N%", "G (m²)", "G%", "IV%", "D1.3 tb (cm)", "Hvn tb (m)"],
                  rows, [1.1, 3.6, 1.6, 1.6, 1.9, 1.6, 1.6, 2.1, 1.9],
                  ["center", "left"] + ["right"] * 7, total=True)
            heading("2.3. Chỉ số đa dạng sinh học", 2)
            table(["Chỉ số", "Giá trị"],
                  [["Số loài (S)", d["S"]], ["Shannon H'", _vn(d["H"], 3)], ["Pielou J'", _vn(d["J"], 3)],
                   ["Simpson (1 - D)", _vn(d["simpson"], 3)], ["Margalef d", _vn(d["margalef"], 3)],
                   ["Trữ lượng M/ha (m³/ha)", _vn(data["vol"]["M_ha"], 1)]],
                  [8, 4], ["left", "right"])
            picture(lambda ax: draw_species_composition(ax, comp, color_map), 6.5, 4.6,
                    "Tổ thành loài theo chỉ số giá trị quan trọng IV%")
            doc.add_page_break()

        # ---------------- 3-5. Trắc đồ & 3D ----------------
        if "plan" in sections:
            heading("3. Trắc đồ ngang", 1)
            picture(lambda ax: draw_plan_view(ax, df, color_map, style), 6.5, 6.2,
                    f"Trắc đồ ngang (hình chiếu bằng tán cây) ô tiêu chuẩn {otc}".strip())
            doc.add_page_break()
        if "profile" in sections:
            heading("4. Trắc đồ dọc", 1)
            picture(lambda ax: draw_profile_view(ax, df, color_map, style), 6.5, 4.6,
                    f"Trắc đồ dọc - mặt cắt đứng chiếu theo trục {style.get('profile_axis', 'X')}")
            doc.add_page_break()
        if "view3d" in sections:
            heading("5. Mô phỏng 3D lâm phần", 1)
            elev, azim = meta.get("view3d") or (22, -60)

            def _d3(ax):
                draw_3d_view(ax, df, color_map, style)
                ax.view_init(elev=elev, azim=azim)
            picture(_d3, 6.5, 5.6, "Mô phỏng 3D cấu trúc lâm phần", three_d=True)
            doc.add_page_break()

        # ---------------- 6. Biểu đồ ----------------
        if "charts" in sections:
            bw = style.get("struct_bin_width", 4.0)
            vbw = style.get("vstruct_bin_width", 2.0)
            heading("6. Biểu đồ phân bố & phân tầng", 1)
            picture(lambda ax: draw_nd_distribution(ax, df, style, bw), 6.5, 3.6,
                    f"Phân bố số cây theo cỡ kính N/D (cự ly {_vn(bw, 0)} cm)")
            picture(lambda ax: draw_ndg_distribution(ax, df, style), 6.5, 3.6,
                    "Phân bố số cây theo cấp kính tương đối N/Dg")
            picture(lambda ax: draw_g_distribution(ax, df, style, bw), 6.5, 3.6,
                    "Phân bố tiết diện ngang G theo cỡ kính")
            picture(lambda ax: draw_nh_distribution(ax, df, style, vbw, vs), 6.5, 3.6,
                    f"Phân bố số cây theo cấp chiều cao N/H (cự ly {_vn(vbw, 0)} m)")
            if canopy:
                picture(lambda ax: draw_canopy_coverage(ax, df, style, canopy), 6.5, 5.6,
                        "Bản đồ độ che phủ & chồng lấn tán")
            picture(lambda ax: draw_forest_layers(ax, df, style, vs), 6.5, 5.0,
                    "Sơ đồ phân tầng đứng lâm phần")
            doc.add_page_break()

        # ---------------- Phụ lục ----------------
        if "appendix" in sections:
            heading("Phụ lục. Bảng số liệu đo đếm cây", 1)
            v_i = data["vol"]["v_i"]
            layer = list(vs["_layer"]) if vs is not None else [""] * len(df)
            dfr = df.reset_index(drop=True)
            rows = []
            g_sum = 0.0
            for i, (_, x) in enumerate(dfr.iterrows()):
                dt = float(np.nanmean([x["DtanD"], x["DtanT"], x["DtanN"], x["DtanB"]]))
                g = float(np.pi / 4 * (float(x["D13"]) / 100) ** 2)
                g_sum += g
                rows.append([x["TT"], x["Loai"], _vn(x["D13"], 1), _vn(x["Hvn"], 1), _vn(x["Hdc"], 1),
                             _vn(dt, 2), _vn(x["X"], 1), _vn(x["Y"], 1), _vn(g, 4), _vn(v_i[i], 4),
                             str(layer[i])])
            rows.append(["", f"Tổng: {len(dfr)} cây", "", "", "", "", "", "", _vn(g_sum, 4),
                         _vn(float(np.sum(v_i)), 4), ""])
            table(["TT", "Loài cây", "D1.3", "Hvn", "Hdc", "Dtán tb", "X", "Y", "G (m²)", "V (m³)", "Tầng"],
                  rows, [0.9, 4.4, 1.2, 1.1, 1.1, 1.3, 1.0, 1.0, 1.5, 1.5, 1.0],
                  ["center", "left"] + ["right"] * 8 + ["center"], size=7.5, total=True)
            para("Đơn vị: D1.3 (cm); Hvn, Hdc, Dtán (m); X, Y (m); G (m²); V = G · Hvn · f (m³).",
                 size=8, italic=True, color=MUTED)

        # ---------------- chân trang & thuộc tính file ----------------
        fp = sec.footer.paragraphs[0]
        fp.text = f"TDForest v{version}  ·  Khoa Lâm nghiệp - ĐH Nông Lâm TP. Hồ Chí Minh"
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in fp.runs:
            r.font.size = Pt(8)
            r.font.color.rgb = MUTED
        doc.core_properties.title = meta.get("title") or "Báo cáo cấu trúc lâm phần"
        doc.core_properties.author = meta.get("author") or ""
        doc.core_properties.subject = f"OTC {otc} - Khoa Lâm nghiệp, ĐH Nông Lâm TP.HCM"

    rc = {"font.family": "sans-serif",
          "font.sans-serif": REPORT_FONTS + list(matplotlib.rcParams["font.sans-serif"]),
          "axes.unicode_minus": False}
    with matplotlib.rc_context(rc):
        _render()
    doc.save(path)
    return path


class MainWindow(QMainWindow):

    def __init__(self, splash=None):

        super().__init__()

        # Khởi tạo TRƯỚC khi có bất kỳ resize/show nào xảy ra (kể cả gián tiếp
        # qua _fit_to_screen() bên dưới), vì resizeEvent() có thể được Qt gọi
        # ngay trong quá trình dựng cửa sổ trên một số máy (driver/DPI khác
        # nhau) - nếu để _confetti khởi tạo muộn sẽ gây AttributeError khi
        # đóng gói exe và chạy trên máy khác (đã từng xảy ra trên Win 11).
        self._confetti = None

        # Màn hình chờ (nếu có truyền vào từ main()) - dùng để cập nhật dòng
        # trạng thái tại vài mốc dựng giao diện tốn thời gian nhất, giúp
        # người dùng thấy phần mềm đang thực sự tải chứ không bị "đứng".
        self._splash = splash

        self.setWindowTitle("TDForest - Phần mềm vẽ trắc đồ ngang - dọc rừng")

        self.setWindowIcon(make_tree_icon(64))

        self.setMinimumSize(860, 560)

        self._fit_to_screen()



        # ---- trạng thái dữ liệu ----

        self.file_paths = []

        self.raw_df = None

        self.raw_columns = []

        self.column_mapping = {}

        self.df = None

        self.color_map = {}

        self.color_overrides = {}  # màu do người dùng tự chọn riêng cho từng loài (species -> hex)

        self._crown_edge_last_color = "#1b5e20"  # ghi nhớ màu viền tán vừa chọn khi tick/bỏ tick



        self.style = {

            "plan_shape": PLAN_SHAPE_OPTIONS[0],

            "crown_form": CROWN_FORM_OPTIONS[0],

            "color_mode": COLOR_MODES[0],

            "uniform_color": "#43a047",

            "trunk_color": "#795548",

            "alpha": 0.85,

            "radius_mode": True,

            "show_labels": False,

            "label_mode": "TT",

            "label_color": "#2b3a2f",

            "show_grid": True,

            "bg_color": "#fbfdf9",

            "plot_w": 20.0,

            "plot_h": 25.0,

            # kích thước thực của ô tiêu chuẩn dùng để TÍNH TỶ LỆ - luôn lấy

            # theo giá trị người dùng nhập ở khung ②, KHÔNG phụ thuộc việc có

            # tick "Hiển thị khung ô tiêu chuẩn" hay không (đó chỉ là bật/tắt

            # vẽ đường viền nét đứt của ô trên trắc đồ).

            "field_w": 20.0,

            "field_h": 25.0,

            "profile_axis": "X",

            "trunk_exagg": 1.0,

            "h_range": (0.0, 30.0),

            # viền tán riêng + độ nét

            "crown_edge_color": None,        # None = tự động (tối màu hơn màu tán)

            "crown_edge_width": 1.2,

            # tỷ lệ bản vẽ

            "scale_denom": None,             # None = tự động vừa khung

            # nhãn & phông chữ biểu đồ (để trống -> dùng mặc định)

            "title_plan": "", "title_profile": "", "title_3d": "",

            "xlabel_plan": "", "ylabel_plan": "",

            "xlabel_profile": "", "ylabel_profile": "",

            "font_color": "#1b5e20",

            "font_family": FONT_FAMILY_OPTIONS[0],

            "title_size": 11.5,

            "label_size": 10.0,

            # chú dẫn màu loài cây vẽ ngay trên biểu đồ (góc trên bên phải)

            # - có thể KÉO THẢ tới vị trí bất kỳ bằng chuột (vị trí sau khi

            # kéo lưu vào "legend_bbox", None = mặc định góc trên-phải), và

            # tự chỉnh cỡ chữ / màu chữ riêng (None màu = dùng chung font_color).

            "show_species_legend": True,

            "legend_font_size": 7.6,

            "legend_font_color": None,

            "legend_bbox": None,                    # cũ - chỉ còn giữ để tương thích ngược

            "legend_bbox_datalim_frac": None,        # cũ - chỉ còn giữ để tương thích ngược

            # Vị trí chú dẫn loài cây sau khi kéo thả - LƯU RIÊNG cho từng
            # trắc đồ (ngang/dọc), giống la bàn & thước tỷ lệ bên dưới, vì
            # vùng dữ liệu thật của 2 trắc đồ khác tỷ lệ nhau nên không thể
            # dùng chung 1 toạ độ (nếu không chú dẫn sẽ "nhảy" chỗ khi
            # chuyển trắc đồ hoặc sau khi lưu/tải lại cấu hình).
            "plan_legend_bbox": None,

            "plan_legend_bbox_datalim_frac": None,

            "profile_legend_bbox": None,

            "profile_legend_bbox_datalim_frac": None,

            # la bàn (hướng Bắc) & thước tỷ lệ đồ hoạ vẽ ngay trên trắc đồ

            # ngang/dọc - có thể KÉO THẢ tới vị trí bất kỳ bằng chuột giống

            # chú dẫn loài cây; vị trí sau khi kéo lưu riêng cho từng trắc

            # đồ (None = dùng vị trí mặc định).

            "show_north_arrow": True,

            "show_scale_bar": True,

            "plan_north_arrow_pos": None,

            "profile_north_arrow_pos": None,

            "plan_scalebar_pos": None,

            "profile_scalebar_pos": None,

            # kích thước (hệ số nhân, mặc định 1.0) và màu riêng của la bàn

            # & thước tỷ lệ - None ở màu thước tỷ lệ = dùng theo màu chữ

            # chung (font_color); la bàn luôn có màu riêng (mặc định đen).

            "north_arrow_size": 1.0,

            "north_arrow_color": "#111111",

            "scale_bar_size": 1.0,

            "scale_bar_color": None,

            # chú thích "Tỷ lệ 1:N (1 cm trên sơ đồ = X m ngoài thực tế)" -

            # chỉ hiển thị khi đã chọn tỷ lệ cụ thể ở mục ② (không phải "Tự

            # động"); kích thước/màu/font riêng, None = dùng theo mục ⑤ chung.

            "show_scale_note": True,

            "scale_note_size": 1.0,

            "scale_note_color": None,

            "scale_note_font_family": None,

            "plan_scale_note_pos": None,

            "profile_scale_note_pos": None,

            # dạng tán RIÊNG gán cho từng loài (species -> tên dạng tán trong

            # CROWN_FORM_OPTIONS); loài không có trong đây dùng "crown_form"

            # chung ở trên - cho phép mô phỏng nhiều dạng tán linh động cùng

            # lúc trong 1 ô tiêu chuẩn.

            "crown_form_map": {},

            # --- Phân tích cấu trúc rừng (cấu trúc ngang) ---

            "struct_bin_width": 4.0,   # độ rộng cấp kính D1.3 (cm) dùng cho phân bố N/D và G

            "struct_grid_res": 140,    # độ phân giải lưới raster để tính độ che phủ tán

            # vị trí hộp thông tin độ che phủ tán sau khi kéo thả (axes fraction, góc trên-trái)

            "canopy_info_pos": None,

            # --- Phân tích cấu trúc rừng (cấu trúc đứng - theo Hvn) ---

            "vstruct_bin_width": 2.0,  # độ rộng cấp chiều cao Hvn (m) dùng cho phân bố N/H

            "vstruct_k_std": 1.0,      # hệ số k trong ngưỡng phân tầng Htb ± k·σ

            # vị trí chú dẫn sơ đồ phân tầng đứng sau khi kéo thả (axes fraction, góc trên-phải)

            "forest_layers_bbox": None,

            # --- Tuỳ biến Tiêu đề / nhãn trục X-Y / Màu sắc / Kiểu chữ riêng

            # cho TỪNG biểu đồ con trong tab "Phân tích cấu trúc rừng" (để

            # trống -> dùng tiêu đề/nhãn mặc định; font để trống -> dùng theo

            # mục ⑤ chung của toàn phần mềm) ----

            "struct_chart_style": {

                "nd":     {"title": "", "xlabel": "", "ylabel": "", "color": "#43a047", "font_color": "", "font_family": ""},

                "ndg":    {"title": "", "xlabel": "", "ylabel": "", "color": "#7cb342", "font_color": "", "font_family": ""},

                "g":      {"title": "", "xlabel": "", "ylabel": "", "color": "#ef6c00", "font_color": "", "font_family": ""},

                "cover":  {"title": "", "xlabel": "", "ylabel": "", "color": "#66bb6a", "font_color": "", "font_family": ""},

                "nh":     {"title": "", "xlabel": "", "ylabel": "", "color": "#00838f", "font_color": "", "font_family": ""},

                "layers": {"title": "", "xlabel": "", "ylabel": "", "color": "#2e7d32", "font_color": "", "font_family": ""},

            },

        }

        self.crown_form_map = {}  # bản làm việc, đồng bộ vào self.style khi vẽ lại

        self._canopy_result = None  # kết quả tính độ che phủ tán (tính theo yêu cầu, không tự động mỗi lần refresh)

        self._vstruct_result = None  # kết quả phân tầng đứng lần tính gần nhất (Hvn tb/max/min, A1/A2/A3...)

        # trạng thái đang KÉO THẢ la bàn / thước tỷ lệ trên từng trắc đồ

        # (key "plan"/"profile" -> {"target": None|"north"|"scale", "offset": (dx, dy)})

        self._overlay_drag = {

            "plan": {"target": None, "offset": (0.0, 0.0)},

            "profile": {"target": None, "offset": (0.0, 0.0)},

        }



        # ── Trạng thái bản phát hành / thông báo ──────────────────────────────
        self._notification_loader = None
        self._update_notice_shown = False



        self._splash_msg("  Đang tạo thanh công cụ...")

        self._build_toolbar()

        self._splash_msg("  Đang dựng giao diện...")

        self._build_central()

        self.setStatusBar(QStatusBar())

        self.statusBar().showMessage("Sẵn sàng. Hãy mở file Excel dữ liệu cây rừng để bắt đầu.")



        self.rotate_timer = QTimer(self)

        self.rotate_timer.setInterval(45)

        self.rotate_timer.timeout.connect(self._auto_rotate_tick)



        # ── Thời hạn bản phát hành (không cần mạng, không cần key) ─────────────
        if _release_needs_update():
            self._lock_all_features()
        # Thông báo hiện sau khi hộp thoại "Thông tin" (tự đóng sau 3s) đã đóng.
        QTimer.singleShot(3600, self._show_startup_notices)
        # Kiểm tra lại mỗi giờ, phòng khi phần mềm mở liên tục qua ngày tới hạn.
        self.release_check_timer = QTimer(self)
        self.release_check_timer.timeout.connect(self._periodic_release_check)
        self.release_check_timer.start(60 * 60 * 1000)



        # ── Thông báo hệ thống + kiểm tra phiên bản mới (bảng thongbaotracdo) ──

        self._notification_loader = NotificationWorker()

        self._notification_loader.finished_notice.connect(self._on_notification_loaded)

        self._notification_loader.start()

        # ── Hiệu ứng tung hoa chào mừng khi mở phần mềm ─────────────────────
        # (self._confetti đã được khởi tạo = None ngay đầu __init__ ở trên,
        #  trước _fit_to_screen(), để tránh AttributeError khi resizeEvent
        #  bị gọi sớm trên một số máy Windows.)

    # ---------------------------------------------------------- hiệu ứng mở app
    def show_confetti_effect(self):
        """Phát hiệu ứng tung hoa/confetti phủ toàn bộ cửa sổ chính, gọi sau
        khi cửa sổ đã show() và có kích thước thật để overlay phủ đúng vùng."""
        try:
            self._confetti = ConfettiOverlay(self)
        except Exception:
            # Hiệu ứng chỉ mang tính trang trí - nếu vì lý do gì đó (driver
            # đồ họa, môi trường headless...) không khởi tạo được thì bỏ qua
            # âm thầm, không được làm crash hay chặn phần mềm khởi động.
            self._confetti = None

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._confetti is not None:
            try:
                self._confetti.setGeometry(self.rect())
            except RuntimeError:
                # Overlay đã tự hủy (deleteLater) sau khi hiệu ứng kết thúc.
                self._confetti = None

    # ------------------------------------------------------- kích thước cửa sổ
    def _splash_msg(self, text):

        """Cập nhật dòng trạng thái trên màn hình chờ (nếu có) và ép vẽ lại

        ngay lập tức - để người dùng thấy tiến trình khởi động thực sự thay

        vì một dòng chữ đứng yên trong nhiều giây."""

        if self._splash is not None:

            self._splash.showMessage(text, Qt.AlignBottom | Qt.AlignHCenter, QColor("#ffffff"))

            QApplication.processEvents()



    def _fit_to_screen(self):

        """Đảm bảo cửa sổ luôn vừa với màn hình hiện tại, không bị tràn ra ngoài

        (đặc biệt trên laptop màn hình nhỏ hoặc độ phân giải thấp)."""

        screen = QApplication.primaryScreen()

        avail = screen.availableGeometry() if screen else None

        if avail is not None:

            w = max(self.minimumWidth(), min(1440, int(avail.width() * 0.94)))

            h = max(self.minimumHeight(), min(880, int(avail.height() * 0.90)))

            self.resize(w, h)

            self.move(avail.x() + (avail.width() - w) // 2, avail.y() + (avail.height() - h) // 2)

            if avail.width() < 1150 or avail.height() < 760:

                self.showMaximized()

        else:

            self.resize(1280, 800)



    # ---------------------------------------------------------------- toolbar

    def _build_toolbar(self):

        tb = QToolBar("Công cụ chính")
        tb.setMovable(False)
        tb.setIconSize(QSize(22, 22))

        # Luôn hiện CẢ icon lẫn chữ trên mỗi nút toolbar - mặc định Qt chỉ
        # hiện icon khi QAction có icon thật (như "Ánh xạ cột", "Thoát" sau
        # khi gắn ảnh icon riêng), khiến các nút này bị "mất chữ" so với các
        # nút còn lại chỉ có emoji làm text. Đặt rõ TextBesideIcon để đồng
        # nhất toàn bộ toolbar, mọi nút đều hiện icon + chữ như nhau.
        tb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)

        self.addToolBar(tb)

        # Logo + tên phần mềm đặt NGAY ĐẦU toolbar (trước nút "Mở file
        # Excel") để nhận diện thương hiệu ngay từ cái nhìn đầu tiên.
        logo_lbl = QLabel()
        logo_lbl.setPixmap(make_tree_pixmap(26))
        tb.addWidget(logo_lbl)

        title = QLabel("  TDForest - Trắc đồ ngang / dọc rừng 2D & 3D")
        title.setStyleSheet("font-weight:700; color:#1b5e20; font-size:11.5pt;")
        tb.addWidget(title)

        tb.addSeparator()

        # Ghi chú: nút "Mở file Excel" KHÔNG đặt trên toolbar nữa - đã có nút
        # tương ứng ngay trong khối "① Dữ liệu đầu vào" ở panel trái (xem
        # btn_open trong _build_data_group), để tránh trùng lặp chức năng.

        act_report = QAction("📑  Xuất báo cáo", self)
        act_report.setToolTip("Xuất báo cáo PDF (A4, có trang bìa, bảng chỉ tiêu, tổ thành loài, "
                              "trắc đồ, biểu đồ) và/hoặc Excel nhiều sheet")
        act_report.triggered.connect(self.export_report)
        tb.addAction(act_report)
        _btn_rpt = tb.widgetForAction(act_report)
        if _btn_rpt is not None:
            _btn_rpt.setObjectName("primaryAction")
        act_export = QAction("🖼️  Xuất ảnh", self)
        act_export.triggered.connect(self.export_current_view)
        tb.addAction(act_export)

        act_export_paper = QAction("🖨️  Xuất khổ giấy", self)
        act_export_paper.setToolTip(
            "Xuất trắc đồ ngang/dọc ra file in đúng khổ giấy A2/A3/A4/A5, "
            "đúng hướng (đứng/ngang) và đúng tỷ lệ 1:N")
        act_export_paper.triggered.connect(self.export_paper_size)
        tb.addAction(act_export_paper)

        tb.addSeparator()

        act_about = QAction("ℹ️  Thông tin", self)
        act_about.setToolTip("Xem thông tin phần mềm và liên hệ Admin")
        act_about.triggered.connect(self._show_about_dialog)
        tb.addAction(act_about)

        act_update = QAction("🔄  Cập nhật", self)
        act_update.setToolTip("Mở Google Drive chứa bản cập nhật phần mềm mới nhất")
        act_update.triggered.connect(self._open_update_link)
        tb.addAction(act_update)



        tb.addSeparator()

        act_exit = QAction(make_exit_qicon(28), "  Thoát", self)
        act_exit.setToolTip("Đóng phần mềm (có xác nhận trước khi thoát)")
        act_exit.triggered.connect(self._exit_app)
        tb.addAction(act_exit)

        tb.addSeparator()

        # Ghi chú: "Dữ liệu Demo", "Giải thích cột", "Xuất Demo → Excel" và
        # "Xóa toàn bộ dữ liệu" KHÔNG đặt trên toolbar nữa - đã có nút tương
        # ứng ngay trong khối "① Dữ liệu đầu vào" ở panel trái (xem
        # _build_data_group), để tránh trùng lặp chức năng ở hai nơi.

        # Spacer đẩy nhãn phiên bản sang sát mép phải của toolbar.
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        tb.addWidget(spacer)

        # Nhãn phiên bản: kiểu "huy hiệu" (badge) nhỏ nằm sát mép phải toolbar,
        # dùng tông xanh navy - xanh ngọc đồng bộ với hộp thoại "Thông tin"
        # (thay cho màu cam trước đây vốn lạc tông với bảng màu chủ đạo).
        self.lbl_app_version = QLabel(f"Phiên bản v{APP_VERSION}")
        self.lbl_app_version.setStyleSheet(
            "color:#ffffff;"
            "background: qlineargradient(x1:0,y1:0,x2:1,y2:0,"
            "stop:0 #16213e, stop:0.5 #1f4287, stop:1 #278ea5);"
            "font-size:10px; font-weight:800;"
            "padding:3px 11px; border-radius:9px; margin-right:6px;"
        )
        tb.addWidget(self.lbl_app_version)

    # ------------------------------------------------------- thông tin & thoát

    def _show_about_dialog(self):

        AboutDialog(self).exec()

    def _open_update_link(self):
        """Mở thư mục Google Drive chứa bản cập nhật phần mềm mới nhất bằng
        trình duyệt mặc định của hệ điều hành."""
        url = QUrl("https://drive.google.com/drive/folders/1qbJQGJhYygF_37ucR9nTZexJ8pgAthSQ")
        ok = QDesktopServices.openUrl(url)
        if not ok:
            show_notice(
                self, "error", "Cập nhật",
                "Không thể mở trình duyệt để truy cập link cập nhật.\n"
                "Vui lòng mở thủ công địa chỉ sau:\n"
                "https://drive.google.com/drive/folders/1qbJQGJhYygF_37ucR9nTZexJ8pgAthSQ"
            )



    def _exit_app(self):

        """Đóng hẳn phần mềm sau khi người dùng xác nhận (tránh thoát nhầm khi lỡ

        tay bấm)."""

        reply = QMessageBox.question(

            self, "Xác nhận thoát", "Bạn có chắc chắn muốn thoát phần mềm không?\n"

            "Mọi dữ liệu chưa lưu (nếu có) sẽ bị mất.",

            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,

        )

        if reply == QMessageBox.Yes:

            self.close()



    def _lock_all_features(self, reason: str = ""):

        """Khóa các tab chức năng (Trắc đồ ngang/dọc, 3D, Phân tích cấu trúc) khi
        bản phát hành cần cập nhật. Giữ lại tab "📋 Dữ liệu" (index 0)."""

        if hasattr(self, "tabs"):

            for i in range(self.tabs.count()):

                if i != 0:

                    self.tabs.setTabEnabled(i, False)

        if reason:

            style = (

                "color:#7A4F00; background:#FFF8E1; border:1px solid #FFD54F;"

                "border-radius:6px; padding:6px 12px; font-size:14px; font-weight:600;"

            )

            self._license_banner_text = reason

            self._license_banner_style = style

            lbl = getattr(self, "lbl_license_banner", None)

            if lbl is not None:

                try:

                    lbl.setText(reason)

                    lbl.setStyleSheet(style)

                except RuntimeError:

                    pass



    def _unlock_all_features(self):

        if hasattr(self, "tabs"):

            for i in range(self.tabs.count()):

                self.tabs.setTabEnabled(i, True)
    def _show_startup_notices(self):
        """Gọi sau khi hộp thoại "Thông tin" tự đóng: nếu tới hạn bản phát hành
        thì báo có phiên bản mới, ngược lại hiện thông báo phạm vi sử dụng (1 lần)."""
        if _release_needs_update():
            self._apply_update_required()
            return
        settings = get_settings()
        if not settings.value("app/scope_notice_shown", False, type=bool):
            show_notice(self, "info", "Phạm vi sử dụng phần mềm", _SCOPE_NOTICE)
            settings.setValue("app/scope_notice_shown", True)
            settings.sync()

    def _periodic_release_check(self):
        """Kiểm tra định kỳ (app mở liên tục qua ngày tới hạn vẫn được nhắc)."""
        if not self._update_notice_shown and _release_needs_update():
            self._apply_update_required()

    def _apply_update_required(self):
        self._lock_all_features()
        self.statusBar().showMessage(
            "🔄  Đã có phiên bản TDForest mới - vui lòng liên hệ Admin để nhận bản cập nhật.")
        if not self._update_notice_shown:
            self._update_notice_shown = True
            show_notice(self, "info", "Có phiên bản mới", _UPDATE_REQUIRED_MSG)




    def _on_notification_loaded(self, res: "NotificationResult"):

        """Hiện thông báo hệ thống (nếu có) và cảnh báo bản cập nhật (nếu có phiên

        bản mới hơn APP_VERSION hiện tại). Không phụ thuộc kết quả xác thực bản

        quyền."""

        if not res.ok:

            return  # lỗi mạng/chưa có dữ liệu -> im lặng, không làm phiền người dùng



        if res.thongbao:

            show_notice(self, "info", "Thông báo", res.thongbao)



        if (res.phienbanmoinhat and res.phienbanmoinhat != APP_VERSION
                and not self._update_notice_shown):
            show_notice(
                self, "info", "Có phiên bản mới",
                f"🎉  Đã có phiên bản mới: {res.phienbanmoinhat}\n"
                f"Bạn đang dùng phiên bản: {APP_VERSION}\n\n"
                "Vui lòng liên hệ Admin để nhận bản cập nhật:\n" + _ADMIN_INFO
                + "\n\n──────────────────────────────\n" + _SCOPE_NOTICE
            )



    def closeEvent(self, event):

        """Dừng luồng nền (notification) trước khi đóng cửa sổ để

        tránh cảnh báo QThread bị hủy khi vẫn đang chạy."""

        for w in (self._notification_loader,):

            if w is not None and w.isRunning():

                w.wait(200)

        super().closeEvent(event)



    # --------------------------------------------------------------- central

    def _build_central(self):

        splitter = QSplitter(Qt.Horizontal)

        splitter.setChildrenCollapsible(False)

        self.setCentralWidget(splitter)



        # Panel trái: mở lên đủ rộng để thấy hết chức năng ngay từ đầu, và vẫn

        # kéo giãn/thu hẹp được bằng tay sau đó (kéo thanh chia splitter).

        # Có thanh cuộn dọc khi nội dung dài hơn màn hình, và cuộn ngang dự

        # phòng nếu người dùng kéo panel hẹp hơn nội dung cần.

        left_scroll = QScrollArea()

        left_scroll.setWidgetResizable(True)

        left_scroll.setMinimumWidth(300)

        left_scroll.setMaximumWidth(680)

        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        left_scroll.setFrameShape(QFrame.NoFrame)

        left_panel = QWidget()

        left_panel.setMinimumWidth(400)

        left_layout = QVBoxLayout(left_panel)

        left_layout.setSpacing(12)

        left_layout.setContentsMargins(4, 4, 10, 12)

        left_layout.addWidget(self._build_data_group())

        left_layout.addWidget(self._build_plot_group())

        left_layout.addWidget(self._build_style_group())

        left_layout.addWidget(self._build_labels_group())

        left_layout.addWidget(self._build_legend_group())

        left_layout.addWidget(self._build_north_scale_group())

        left_layout.addStretch(1)

        left_scroll.setWidget(left_panel)

        splitter.addWidget(left_scroll)



        self.tabs = QTabWidget()

        self._splash_msg("  Đang tạo các tab trắc đồ 2D/3D...")

        self.tabs.addTab(self._build_table_tab(), "📋  Dữ liệu")

        self.tabs.addTab(self._build_plan_tab(), "🌳  Trắc đồ ngang (2D)")

        self.tabs.addTab(self._build_profile_tab(), "📐  Trắc đồ dọc (2D)")

        self.tabs.addTab(self._build_3d_tab(), "🧊  Mô phỏng 3D")

        self.tabs.addTab(self._build_structure_tab(), "🌲  Phân tích cấu trúc rừng")




        splitter.addWidget(self.tabs)

        splitter.setStretchFactor(0, 0)

        splitter.setStretchFactor(1, 1)

        splitter.setSizes([430, 1000])



    # ---- nhóm: dữ liệu đầu vào ----

    def _build_data_group(self):

        box = QGroupBox("① Dữ liệu đầu vào")

        lay = QVBoxLayout(box)



        self.file_combo = QComboBox()

        self.file_combo.setPlaceholderText("Chưa có file nào")

        self.file_combo.currentIndexChanged.connect(lambda _: self.load_current_file(ask_mapping=False))

        lay.addWidget(QLabel("File Excel đang xem:"))

        lay.addWidget(self.file_combo)



        row = QHBoxLayout()

        btn_open = QPushButton("📂 Mở file...")

        btn_open.clicked.connect(self.open_files)

        btn_map = QPushButton(make_mapping_icon(18), " Ánh xạ cột...")

        btn_map.setObjectName("secondary")

        btn_map.clicked.connect(self.open_mapping_dialog)

        row.addWidget(btn_open); row.addWidget(btn_map)

        lay.addLayout(row)



        btn_demo = QPushButton(make_demo_icon(20), " Dùng dữ liệu Demo (không cần file)")

        btn_demo.setObjectName("secondary")

        btn_demo.setToolTip("Nạp ngay dữ liệu mẫu tích hợp sẵn trong phần mềm (40 cây, 8 loài) để dùng thử.")

        btn_demo.clicked.connect(self.load_demo_dataset)

        lay.addWidget(btn_demo)



        btn_refresh = QPushButton("🔄 Làm mới (Refresh)")

        btn_refresh.clicked.connect(self.refresh)

        lay.addWidget(btn_refresh)



        row2 = QHBoxLayout()

        btn_explain = QPushButton(make_explain_icon(20), " Giải thích cột")

        btn_explain.setObjectName("secondary")

        btn_explain.setToolTip("Xem ý nghĩa, đơn vị của từng cột dữ liệu (TT, D1.3, Hvn, Hdc, X, Y...).")

        btn_explain.clicked.connect(self.open_column_explain_dialog)

        btn_export_demo = QPushButton(make_exportdemo_icon(20), " Xuất Demo → Excel")

        btn_export_demo.setObjectName("secondary")

        btn_export_demo.setToolTip("Xuất dữ liệu Demo tích hợp sẵn ra file Excel để dùng làm mẫu định dạng nhập liệu.")

        btn_export_demo.clicked.connect(self.export_demo_to_excel)

        row2.addWidget(btn_explain); row2.addWidget(btn_export_demo)

        lay.addLayout(row2)



        lay.addWidget(self._hline())

        row3 = QHBoxLayout()

        btn_save_cfg = QPushButton(make_savecfg_icon(20), " Lưu cấu hình...")

        btn_save_cfg.setObjectName("secondary")

        btn_save_cfg.setToolTip(
            "Lưu toàn bộ dữ liệu cây đã nạp + mọi chỉnh sửa/cài đặt trắc đồ hiện tại "
            "(màu sắc, kích thước ô, nhãn, la bàn, thước tỷ lệ, tuỳ biến biểu đồ...) ra "
            "1 file .geojson. Lần sau chỉ cần \"Tải cấu hình\" để khôi phục lại y hệt, "
            "kể cả khi file Excel gốc đã bị xóa hoặc chỉnh sửa.")

        btn_save_cfg.clicked.connect(self.save_config_geojson)

        btn_load_cfg = QPushButton(make_loadcfg_icon(20), " Tải cấu hình...")

        btn_load_cfg.setObjectName("secondary")

        btn_load_cfg.setToolTip("Tải lại dữ liệu cây + toàn bộ cài đặt trắc đồ đã lưu trước đó từ file .geojson.")

        btn_load_cfg.clicked.connect(self.load_config_geojson)

        row3.addWidget(btn_save_cfg); row3.addWidget(btn_load_cfg)

        lay.addLayout(row3)



        btn_clear = QPushButton("🗑️ Xóa toàn bộ dữ liệu")

        btn_clear.setObjectName("danger")

        btn_clear.setToolTip("Xóa toàn bộ dữ liệu hiện tại, đưa phần mềm về trạng thái như khi mới mở.")

        btn_clear.clicked.connect(self.clear_all_data)

        lay.addWidget(btn_clear)



        self.info_label = QLabel("Chưa có dữ liệu. Hãy mở file Excel hoặc bấm \"Dùng dữ liệu Demo\" để dùng thử ngay.")

        self.info_label.setWordWrap(True)

        self.info_label.setStyleSheet("color:#4b5c4d; font-style:italic; margin-top:4px;")

        lay.addWidget(self.info_label)

        return box



    # ---- nhóm: ô tiêu chuẩn ----

    def _build_plot_group(self):

        box = QGroupBox("② Ô tiêu chuẩn")

        form = QFormLayout(box)



        self.showplot_chk = QCheckBox("Hiển thị khung ô tiêu chuẩn")

        self.showplot_chk.setChecked(True)

        self.showplot_chk.stateChanged.connect(self.on_style_changed)

        form.addRow(self.showplot_chk)



        self.plotw_spin = QDoubleSpinBox(); self.plotw_spin.setRange(1, 2000)

        self.plotw_spin.setValue(20.0); self.plotw_spin.setSuffix(" m")

        self.plotw_spin.valueChanged.connect(self.on_style_changed)

        form.addRow("Chiều rộng (X):", self.plotw_spin)



        self.ploth_spin = QDoubleSpinBox(); self.ploth_spin.setRange(1, 2000)

        self.ploth_spin.setValue(25.0); self.ploth_spin.setSuffix(" m")

        self.ploth_spin.valueChanged.connect(self.on_style_changed)

        form.addRow("Chiều dài (Y):", self.ploth_spin)



        self.area_label = QLabel("Diện tích: 500 m²")

        self.area_label.setStyleSheet("color:#1b5e20; font-weight:600;")

        form.addRow(self.area_label)

        self.plotw_spin.valueChanged.connect(self._update_area_label)

        self.ploth_spin.valueChanged.connect(self._update_area_label)



        form.addRow(self._hline())

        self.scale_combo = QComboBox(); self.scale_combo.addItems(SCALE_OPTIONS)

        self.scale_combo.currentIndexChanged.connect(self.on_scale_changed)

        form.addRow("Tỷ lệ bản vẽ:", self.scale_combo)



        self.scale_custom_spin = QDoubleSpinBox()

        self.scale_custom_spin.setRange(1, 100000); self.scale_custom_spin.setDecimals(0)

        self.scale_custom_spin.setValue(250); self.scale_custom_spin.setPrefix("1 : ")

        self.scale_custom_spin.setVisible(False)

        self.scale_custom_spin.valueChanged.connect(self.on_scale_changed)

        form.addRow("Tỷ lệ tuỳ chỉnh:", self.scale_custom_spin)



        self.scale_hint = QLabel("Tỷ lệ chỉ áp dụng khi in \"Actual size / 100%\" (không dùng \"Fit to page\").")

        self.scale_hint.setWordWrap(True)

        self.scale_hint.setStyleSheet("color:#6b7d6d; font-size:8.4pt; font-style:italic;")

        form.addRow(self.scale_hint)

        return box



    def _hline(self):

        line = QFrame(); line.setFrameShape(QFrame.HLine); line.setStyleSheet("color:#e3ece0;")

        return line



    def _update_area_label(self):

        self.area_label.setText(f"Diện tích: {self.plotw_spin.value() * self.ploth_spin.value():.0f} m²")



    def on_scale_changed(self):

        text = self.scale_combo.currentText()

        is_custom = text.startswith("Tuỳ chỉnh")

        self.scale_custom_spin.setVisible(is_custom)

        if text.startswith("Tự động"):

            self.style["scale_denom"] = None

        elif is_custom:

            self.style["scale_denom"] = self.scale_custom_spin.value()

        else:

            self.style["scale_denom"] = float(text.split(":")[1])

        self.refresh_views()



    # ---- nhóm: style cây ----

    def _build_style_group(self):

        box = QGroupBox("③ Kiểu hiển thị cây")

        form = QFormLayout(box)



        self.plan_shape_combo = QComboBox(); self.plan_shape_combo.addItems(PLAN_SHAPE_OPTIONS)

        self.plan_shape_combo.currentIndexChanged.connect(self.on_style_changed)

        form.addRow("Hình chiếu tán (mặt bằng):", self.plan_shape_combo)



        self.crownform_combo = QComboBox(); self.crownform_combo.addItems(CROWN_FORM_OPTIONS)

        self.crownform_combo.currentIndexChanged.connect(self.on_style_changed)

        form.addRow("Dạng khối tán (dọc/3D):", self.crownform_combo)



        self.crownform_perspecies_btn = QPushButton("🌲 Dạng tán riêng theo loài...")

        self.crownform_perspecies_btn.setObjectName("secondary")

        self.crownform_perspecies_btn.clicked.connect(self.edit_crown_form_per_species)

        form.addRow(self.crownform_perspecies_btn)



        self.colormode_combo = QComboBox(); self.colormode_combo.addItems(COLOR_MODES)

        self.colormode_combo.currentIndexChanged.connect(self.on_style_changed)

        form.addRow("Chế độ màu:", self.colormode_combo)



        self.uniform_btn = QPushButton("")

        self.uniform_btn.setFixedHeight(26)

        self._set_swatch(self.uniform_btn, self.style["uniform_color"])

        self.uniform_btn.clicked.connect(self.choose_uniform_color)

        self.uniform_btn.setEnabled(False)

        form.addRow("Màu đồng nhất:", self.uniform_btn)



        self.trunk_btn = QPushButton("")

        self.trunk_btn.setFixedHeight(26)

        self._set_swatch(self.trunk_btn, self.style["trunk_color"])

        self.trunk_btn.clicked.connect(self.choose_trunk_color)

        form.addRow("Màu thân cây:", self.trunk_btn)



        self.alpha_slider = QSlider(Qt.Horizontal)

        self.alpha_slider.setRange(20, 100); self.alpha_slider.setValue(85)

        self.alpha_slider.valueChanged.connect(self.on_style_changed)

        form.addRow("Độ trong tán lá:", self.alpha_slider)



        form.addRow(self._hline())

        self.crownedge_chk = QCheckBox("Dùng màu viền tán riêng")

        self.crownedge_chk.stateChanged.connect(self.on_crownedge_toggled)

        form.addRow(self.crownedge_chk)



        self.crownedge_btn = QPushButton("")

        self.crownedge_btn.setFixedHeight(26)

        self._set_swatch(self.crownedge_btn, "#2e7d32")

        self.crownedge_btn.clicked.connect(self.choose_crownedge_color)

        self.crownedge_btn.setEnabled(False)

        form.addRow("Màu viền tán:", self.crownedge_btn)



        self.crownedge_width_spin = QDoubleSpinBox()

        self.crownedge_width_spin.setRange(0.3, 4.0); self.crownedge_width_spin.setSingleStep(0.1)

        self.crownedge_width_spin.setValue(self.style["crown_edge_width"])

        self.crownedge_width_spin.valueChanged.connect(self.on_style_changed)

        form.addRow("Độ dày viền tán (nét):", self.crownedge_width_spin)



        form.addRow(self._hline())

        self.trunk_spin = QDoubleSpinBox()

        self.trunk_spin.setRange(0.5, 15.0); self.trunk_spin.setSingleStep(0.5)

        self.trunk_spin.setValue(1.0)

        self.trunk_spin.valueChanged.connect(self.on_style_changed)

        form.addRow("Phóng đại thân cây:", self.trunk_spin)



        self.radius_chk = QCheckBox("Dtán là BÁN KÍNH (bỏ chọn nếu là đường kính)")

        self.radius_chk.setChecked(True)

        self.radius_chk.stateChanged.connect(self.on_style_changed)

        form.addRow(self.radius_chk)



        self.labels_chk = QCheckBox("Hiển thị nhãn cây")

        self.labels_chk.stateChanged.connect(self.on_style_changed)

        form.addRow(self.labels_chk)



        self.labelmode_combo = QComboBox()

        self.labelmode_combo.addItems(["Số TT", "Tên loài"])

        self.labelmode_combo.currentIndexChanged.connect(self.on_style_changed)

        form.addRow("Nội dung nhãn:", self.labelmode_combo)



        self.labelcolor_btn = QPushButton("")

        self.labelcolor_btn.setFixedHeight(26)

        self._set_swatch(self.labelcolor_btn, self.style["label_color"])

        self.labelcolor_btn.clicked.connect(self.choose_label_color)

        form.addRow("Màu nhãn cây:", self.labelcolor_btn)



        self.grid_chk = QCheckBox("Hiển thị lưới toạ độ")

        self.grid_chk.setChecked(True)

        self.grid_chk.stateChanged.connect(self.on_style_changed)

        form.addRow(self.grid_chk)



        self.specieslegend_chk = QCheckBox("Hiển thị chú dẫn màu loài trên biểu đồ")

        self.specieslegend_chk.setChecked(True)

        self.specieslegend_chk.stateChanged.connect(self.on_style_changed)

        form.addRow(self.specieslegend_chk)



        hint = QLabel("Có thể kéo chú dẫn tới vị trí bất kỳ ngay trên biểu đồ bằng chuột.")

        hint.setWordWrap(True)

        hint.setStyleSheet("color:#6b7d6d; font-size:8.6pt;")

        form.addRow(hint)



        self.legendfontsize_spin = QDoubleSpinBox()

        self.legendfontsize_spin.setRange(5.0, 16.0)

        self.legendfontsize_spin.setSingleStep(0.5)

        self.legendfontsize_spin.setValue(self.style["legend_font_size"])

        self.legendfontsize_spin.valueChanged.connect(self.on_style_changed)

        form.addRow("Cỡ chữ chú dẫn:", self.legendfontsize_spin)



        self.legendfontcolor_btn = QPushButton("")

        self.legendfontcolor_btn.setFixedHeight(26)

        self._set_swatch(self.legendfontcolor_btn, self.style.get("legend_font_color") or self.style["font_color"])

        self.legendfontcolor_btn.clicked.connect(self.choose_legendfont_color)

        form.addRow("Màu chữ chú dẫn:", self.legendfontcolor_btn)



        reset_pos_btn = QPushButton("↺  Đặt lại vị trí chú dẫn (góc trên-phải)")

        reset_pos_btn.setObjectName("secondary")

        reset_pos_btn.clicked.connect(self.reset_legend_position)

        form.addRow(reset_pos_btn)



        # ---- Khoảng đệm (margin) quanh dữ liệu - đây là vùng "đất trống"
        # THẬT SỰ mà chú dẫn có thể được kéo vào để không đè lên tán cây, và
        # là khoảng đệm ĐƯỢC LƯU VÀO self.style nên áp dụng giống hệt nhau ở
        # màn hình xem trước lẫn khi xuất PDF/khổ giấy. KHÔNG dùng nút "Edit
        # axis" (biểu tượng cờ-lê) trên thanh công cụ matplotlib để nới trục
        # tạo chỗ cho chú dẫn - thao tác đó chỉ đổi tạm thời trên màn hình,
        # không được lưu lại nên khi xuất PDF trục sẽ tính lại theo giá trị
        # dưới đây và chú dẫn sẽ "nhảy" về vị trí khác. Muốn có thêm chỗ cho
        # chú dẫn, hãy TĂNG giá trị bên dưới rồi kéo chú dẫn vào vùng trống
        # mới - vị trí đó sẽ giữ nguyên khi in.
        margin_hint = QLabel("Khoảng đệm quanh dữ liệu (chỗ trống thật cho chú dẫn, "
                              "GIỮ NGUYÊN khi in PDF - không dùng nút \"Edit axis\" "
                              "của thanh công cụ để nới trục):")
        margin_hint.setWordWrap(True)
        margin_hint.setStyleSheet("color:#6b7d6d; font-size:8.6pt;")
        form.addRow(margin_hint)

        self.legendmargin_spin = QDoubleSpinBox()
        self.legendmargin_spin.setRange(0.0, 1.0)
        self.legendmargin_spin.setSingleStep(0.02)
        self.legendmargin_spin.setDecimals(2)
        self.legendmargin_spin.setValue(self.style.get("legend_margin_frac", 0.14))
        self.legendmargin_spin.valueChanged.connect(self._on_legend_margin_changed)
        form.addRow("Đệm quanh Trắc đồ ngang:", self.legendmargin_spin)

        self.profilemarginx_spin = QDoubleSpinBox()
        self.profilemarginx_spin.setRange(0.0, 1.0)
        self.profilemarginx_spin.setSingleStep(0.02)
        self.profilemarginx_spin.setDecimals(2)
        self.profilemarginx_spin.setValue(self.style.get("profile_margin_x", 0.08))
        self.profilemarginx_spin.valueChanged.connect(self._on_profile_margin_changed)
        form.addRow("Đệm ngang Trắc đồ dọc:", self.profilemarginx_spin)

        self.profilemargintop_spin = QDoubleSpinBox()
        self.profilemargintop_spin.setRange(1.0, 2.0)
        self.profilemargintop_spin.setSingleStep(0.02)
        self.profilemargintop_spin.setDecimals(2)
        self.profilemargintop_spin.setValue(self.style.get("profile_margin_y_top", 1.12))
        self.profilemargintop_spin.valueChanged.connect(self._on_profile_margin_changed)
        form.addRow("Đệm trên Trắc đồ dọc:", self.profilemargintop_spin)



        return box



    def choose_legendfont_color(self):

        current = QColor(self.style.get("legend_font_color") or self.style["font_color"])

        c = QColorDialog.getColor(current, self, "Chọn màu chữ chú dẫn")

        if c.isValid():

            self.style["legend_font_color"] = c.name()

            self._set_swatch(self.legendfontcolor_btn, c.name())

            self.refresh_views()



    def reset_legend_position(self):
        # Đặt lại vị trí chú dẫn về mặc định (góc trên-phải) CHO CẢ 2 trắc
        # đồ (ngang lẫn dọc), vì nút này nằm ở mục ⑤ chung, không tách riêng
        # theo tab đang xem.
        for key in ("legend_bbox", "legend_bbox_datalim_frac",
                    "plan_legend_bbox", "plan_legend_bbox_datalim_frac",
                    "profile_legend_bbox", "profile_legend_bbox_datalim_frac"):
            self.style[key] = None

        self.refresh_views()



    def _on_legend_margin_changed(self):
        # Đổi khoảng đệm quanh Trắc đồ ngang -> vùng trống cho chú dẫn của
        # RIÊNG trắc đồ ngang thay đổi, nên chỉ đặt lại vị trí chú dẫn của
        # trắc đồ ngang (không đụng tới vị trí đã kéo ở trắc đồ dọc).
        self.style["legend_margin_frac"] = self.legendmargin_spin.value()
        self.style["plan_legend_bbox"] = None
        self.style["plan_legend_bbox_datalim_frac"] = None
        self.refresh_views()

    def _on_profile_margin_changed(self):
        # Tương tự trên, chỉ đặt lại vị trí chú dẫn của trắc đồ dọc.
        self.style["profile_margin_x"] = self.profilemarginx_spin.value()
        self.style["profile_margin_y_top"] = self.profilemargintop_spin.value()
        self.style["profile_legend_bbox"] = None
        self.style["profile_legend_bbox_datalim_frac"] = None
        self.refresh_views()



    # ---- nhóm: la bàn, thước tỷ lệ & chú thích tỷ lệ ----

    def _build_north_scale_group(self):

        box = QGroupBox("⑥ La bàn, Thước tỷ lệ & Chú thích tỷ lệ")

        form = QFormLayout(box)



        self.northarrow_chk = QCheckBox("Hiển thị la bàn (hướng Bắc)")

        self.northarrow_chk.setChecked(True)

        self.northarrow_chk.stateChanged.connect(self.on_style_changed)

        form.addRow(self.northarrow_chk)



        self.northarrowsize_spin = QDoubleSpinBox()

        self.northarrowsize_spin.setRange(0.3, 4.0)

        self.northarrowsize_spin.setSingleStep(0.1)

        self.northarrowsize_spin.setValue(self.style["north_arrow_size"])

        self.northarrowsize_spin.valueChanged.connect(self.on_style_changed)

        form.addRow("  Kích thước la bàn:", self.northarrowsize_spin)



        self.northarrowcolor_btn = QPushButton("")

        self.northarrowcolor_btn.setFixedHeight(26)

        self._set_swatch(self.northarrowcolor_btn, self.style["north_arrow_color"])

        self.northarrowcolor_btn.clicked.connect(self.choose_northarrow_color)

        form.addRow("  Màu la bàn:", self.northarrowcolor_btn)



        form.addRow(self._hline())



        self.scalebar_chk = QCheckBox("Hiển thị thước tỷ lệ đồ hoạ")

        self.scalebar_chk.setChecked(True)

        self.scalebar_chk.stateChanged.connect(self.on_style_changed)

        form.addRow(self.scalebar_chk)



        self.scalebarsize_spin = QDoubleSpinBox()

        self.scalebarsize_spin.setRange(0.3, 4.0)

        self.scalebarsize_spin.setSingleStep(0.1)

        self.scalebarsize_spin.setValue(self.style["scale_bar_size"])

        self.scalebarsize_spin.valueChanged.connect(self.on_style_changed)

        form.addRow("  Kích thước thước tỷ lệ:", self.scalebarsize_spin)



        self.scalebarcolor_btn = QPushButton("")

        self.scalebarcolor_btn.setFixedHeight(26)

        self._set_swatch(self.scalebarcolor_btn, self.style.get("scale_bar_color") or self.style["font_color"])

        self.scalebarcolor_btn.clicked.connect(self.choose_scalebar_color)

        form.addRow("  Màu thước tỷ lệ:", self.scalebarcolor_btn)



        form.addRow(self._hline())



        self.scalenote_chk = QCheckBox("Hiển thị chú thích \"Tỷ lệ 1:N\" (kể cả khi để \"Tự động (vừa khung)\" - lúc đó tỷ lệ có dấu ≈)")

        self.scalenote_chk.setChecked(True)

        self.scalenote_chk.stateChanged.connect(self.on_style_changed)

        form.addRow(self.scalenote_chk)



        self.scalenotesize_spin = QDoubleSpinBox()

        self.scalenotesize_spin.setRange(0.3, 4.0)

        self.scalenotesize_spin.setSingleStep(0.1)

        self.scalenotesize_spin.setValue(self.style["scale_note_size"])

        self.scalenotesize_spin.valueChanged.connect(self.on_style_changed)

        form.addRow("  Kích thước chữ:", self.scalenotesize_spin)



        self.scalenotecolor_btn = QPushButton("")

        self.scalenotecolor_btn.setFixedHeight(26)

        self._set_swatch(self.scalenotecolor_btn, self.style.get("scale_note_color") or self.style["font_color"])

        self.scalenotecolor_btn.clicked.connect(self.choose_scalenote_color)

        form.addRow("  Màu chữ:", self.scalenotecolor_btn)



        self.scalenotefont_combo = QComboBox()

        self.scalenotefont_combo.addItem("(theo font chung mục ⑤)")

        self.scalenotefont_combo.addItems(FONT_FAMILY_OPTIONS)

        current_font = self.style.get("scale_note_font_family")

        if current_font and current_font in FONT_FAMILY_OPTIONS:

            self.scalenotefont_combo.setCurrentIndex(FONT_FAMILY_OPTIONS.index(current_font) + 1)

        self.scalenotefont_combo.currentIndexChanged.connect(self.on_style_changed)

        form.addRow("  Font chữ:", self.scalenotefont_combo)



        hint = QLabel("Có thể kéo la bàn, thước tỷ lệ và chú thích tỷ lệ tới bất kỳ vị trí nào "

                       "ngay trên Trắc đồ ngang / Trắc đồ dọc bằng chuột, giống chú dẫn loài cây. "

                       "Click vào 1 cây trên trắc đồ để xem thông tin chi tiết của cây đó.")

        hint.setWordWrap(True)

        hint.setStyleSheet("color:#6b7d6d; font-size:8.6pt;")

        form.addRow(hint)



        reset_btn = QPushButton("↺  Đặt lại vị trí la bàn, thước tỷ lệ & chú thích tỷ lệ")

        reset_btn.setObjectName("secondary")

        reset_btn.clicked.connect(self.reset_north_scale_position)

        form.addRow(reset_btn)



        return box



    def choose_northarrow_color(self):

        c = QColorDialog.getColor(QColor(self.style["north_arrow_color"]), self, "Chọn màu la bàn")

        if c.isValid():

            self.style["north_arrow_color"] = c.name()

            self._set_swatch(self.northarrowcolor_btn, c.name())

            self.refresh_views()



    def choose_scalebar_color(self):

        current = QColor(self.style.get("scale_bar_color") or self.style["font_color"])

        c = QColorDialog.getColor(current, self, "Chọn màu thước tỷ lệ")

        if c.isValid():

            self.style["scale_bar_color"] = c.name()

            self._set_swatch(self.scalebarcolor_btn, c.name())

            self.refresh_views()



    def choose_scalenote_color(self):

        current = QColor(self.style.get("scale_note_color") or self.style["font_color"])

        c = QColorDialog.getColor(current, self, "Chọn màu chữ chú thích tỷ lệ")

        if c.isValid():

            self.style["scale_note_color"] = c.name()

            self._set_swatch(self.scalenotecolor_btn, c.name())

            self.refresh_views()



    def reset_north_scale_position(self):

        for key in ("plan_north_arrow_pos", "profile_north_arrow_pos",

                    "plan_scalebar_pos", "profile_scalebar_pos",

                    "plan_scale_note_pos", "profile_scale_note_pos"):

            self.style[key] = None

        self.refresh_views()



    def _set_swatch(self, btn, hex_color):

        btn.setStyleSheet(

            f"background-color:{hex_color}; border:1.4px solid #8fa38a; border-radius:6px;"

        )



    def choose_uniform_color(self):

        c = QColorDialog.getColor(QColor(self.style["uniform_color"]), self, "Chọn màu tán cây")

        if c.isValid():

            self.style["uniform_color"] = c.name()

            self._set_swatch(self.uniform_btn, c.name())

            self.refresh_views()



    def choose_trunk_color(self):

        c = QColorDialog.getColor(QColor(self.style["trunk_color"]), self, "Chọn màu thân cây")

        if c.isValid():

            self.style["trunk_color"] = c.name()

            self._set_swatch(self.trunk_btn, c.name())

            self.refresh_views()



    def choose_label_color(self):

        c = QColorDialog.getColor(QColor(self.style["label_color"]), self, "Chọn màu nhãn cây")

        if c.isValid():

            self.style["label_color"] = c.name()

            self._set_swatch(self.labelcolor_btn, c.name())

            self.refresh_views()



    def on_crownedge_toggled(self):

        checked = self.crownedge_chk.isChecked()

        self.crownedge_btn.setEnabled(checked)

        self.style["crown_edge_color"] = self._crown_edge_last_color if checked else None

        self.refresh_views()



    def choose_crownedge_color(self):

        current = QColor(self.style.get("crown_edge_color") or self._crown_edge_last_color)

        c = QColorDialog.getColor(current, self, "Chọn màu viền tán")

        if c.isValid():

            self._set_swatch(self.crownedge_btn, c.name())

            self._crown_edge_last_color = c.name()

            self.style["crown_edge_color"] = c.name()

            self.refresh_views()



    # ---- nhóm: chú giải loài ----

    def _build_legend_group(self):

        box = QGroupBox("④ Chú giải loài cây")

        lay = QVBoxLayout(box)

        hint = QLabel("Nhấp đúp vào 1 loài để tự chọn màu riêng cho loài đó.")

        hint.setWordWrap(True)

        hint.setStyleSheet("color:#6b7d6d; font-size:8.6pt;")

        lay.addWidget(hint)

        self.legend_list = QListWidget()

        self.legend_list.setFixedHeight(150)

        self.legend_list.itemDoubleClicked.connect(self.edit_species_color)

        lay.addWidget(self.legend_list)

        reset_btn = QPushButton("↺  Đặt lại màu loài mặc định")

        reset_btn.clicked.connect(self.reset_species_colors)

        lay.addWidget(reset_btn)

        return box



    def edit_species_color(self, item):

        """Cho phép người dùng tự chọn màu riêng cho 1 loài cây bằng cách

        nhấp đúp vào mục tương ứng trong danh sách chú giải."""

        if not self.color_map:

            return

        idx = self.legend_list.row(item)

        species_names = list(self.color_map.keys())

        if idx < 0 or idx >= len(species_names):

            return

        species = species_names[idx]

        current = QColor(self.color_map.get(species, "#558b2f"))

        c = QColorDialog.getColor(current, self, f"Chọn màu cho loài: {species}")

        if c.isValid():

            self.color_overrides[species] = c.name()

            self.color_map[species] = c.name()

            self._update_legend()

            self.refresh_views()



    def reset_species_colors(self):

        """Xoá các màu tự chọn riêng, quay lại bảng màu mặc định của phần mềm."""

        if not self.color_overrides:

            return

        self.color_overrides = {}

        if self.df is not None and not self.df.empty:

            self.color_map = build_species_colors(self.df["Loai"].tolist(), self.color_overrides)

            self._update_legend()

            self.refresh_views()



    def edit_crown_form_per_species(self):

        """Mở hộp thoại cho phép gán dạng khối tán riêng cho từng loài cây

        (mặt cắt dọc / 3D) - không có dữ liệu thì báo và thoát."""

        if self.df is None or self.df.empty:

            show_notice(self, "info", "Dạng tán theo loài",

                        "Vui lòng mở dữ liệu cây rừng trước khi gán dạng tán riêng theo loài.")

            return

        species_list = sorted(set(str(s) for s in self.df["Loai"].tolist()))

        dlg = SpeciesCrownFormDialog(species_list, self.crown_form_map, self)

        if dlg.exec() == QDialog.Accepted:

            self.crown_form_map = dlg.get_map()

            self.style["crown_form_map"] = dict(self.crown_form_map)

            self.refresh_views()



    # ---- tab: bảng dữ liệu ----

    def _build_table_tab(self):

        w = QWidget(); lay = QVBoxLayout(w)

        self.table_model = PandasModel()

        self.table_view = QTableView()

        self.table_view.setModel(self.table_model)

        self.table_view.setAlternatingRowColors(True)

        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.table_view.setEditTriggers(QTableView.NoEditTriggers)

        lay.addWidget(self.table_view)

        return w



    # ---- tab: trắc đồ ngang ----

    def _build_plan_tab(self):

        w = QWidget(); lay = QVBoxLayout(w)

        self.canvas_plan = MplCanvas(three_d=False)

        self._wire_overlay_events(self.canvas_plan, "plan")

        toolbar = NavigationToolbar(self.canvas_plan, w)

        lay.addWidget(toolbar)

        lay.addWidget(self._make_tree_click_hint())

        self.plan_scroll = QScrollArea()

        self.plan_scroll.setWidgetResizable(True)

        self.plan_scroll.setFrameShape(QFrame.NoFrame)

        self.plan_scroll.setWidget(self.canvas_plan)

        lay.addWidget(self.plan_scroll)

        return w



    # ---- tab: trắc đồ dọc ----

    def _build_profile_tab(self):

        w = QWidget(); lay = QVBoxLayout(w)

        controls = QHBoxLayout()

        controls.addWidget(QLabel("Trục chiếu mặt cắt:"))

        self.axis_combo = QComboBox()

        self.axis_combo.addItems(["X (Đông - Tây)", "Y (Nam - Bắc)"])

        self.axis_combo.currentIndexChanged.connect(self.on_axis_changed)

        controls.addWidget(self.axis_combo)

        controls.addStretch(1)

        lay.addLayout(controls)

        self.canvas_profile = MplCanvas(three_d=False)

        self._wire_overlay_events(self.canvas_profile, "profile")

        toolbar = NavigationToolbar(self.canvas_profile, w)

        lay.addWidget(toolbar)

        lay.addWidget(self._make_tree_click_hint())

        self.profile_scroll = QScrollArea()

        self.profile_scroll.setWidgetResizable(True)

        self.profile_scroll.setFrameShape(QFrame.NoFrame)

        self.profile_scroll.setWidget(self.canvas_profile)

        lay.addWidget(self.profile_scroll)

        return w



    def on_axis_changed(self, idx):

        self.style["profile_axis"] = "X" if idx == 0 else "Y"

        self.refresh_views()



    def _make_tree_click_hint(self):

        hint = QLabel("💡 Click vào 1 cây trên trắc đồ để xem thông tin chi tiết "

                       "(TT, loài, D1.3, Hvn, Hdc, đường kính tán, toạ độ...). "

                       "Kéo la bàn 🧭 hoặc thước tỷ lệ tới vị trí khác bằng chuột nếu muốn.")

        hint.setWordWrap(True)

        hint.setStyleSheet("color:#6b7d6d; font-size:8.6pt; padding:1px 2px;")

        return hint



    def _make_tree_click_hint_simple(self):

        """Giống _make_tree_click_hint() nhưng KHÔNG nhắc tới la bàn/thước tỷ

        lệ (2 biểu đồ Độ che phủ tán và Sơ đồ phân tầng đứng không có 2 vật

        thể đó) - dùng cho các tab Phân tích cấu trúc rừng."""

        hint = QLabel("💡 Click vào 1 cây trên biểu đồ để xem thông tin chi tiết "

                       "(TT, loài, D1.3, Hvn, Hdc, đường kính tán, toạ độ...).")

        hint.setWordWrap(True)

        hint.setStyleSheet("color:#6b7d6d; font-size:8.6pt; padding:1px 2px;")

        return hint



    # ==========================================================================

    # LA BÀN / THƯỚC TỶ LỆ - KÉO THẢ & CLICK XEM THÔNG TIN CÂY (mục ⑥)

    # ==========================================================================



    def _wire_overlay_events(self, canvas, key):

        """Kết nối 1 LẦN DUY NHẤT (khi tạo canvas) các sự kiện chuột cho: kéo

        thả la bàn/thước tỷ lệ, và click vào tán cây để xem thông tin. Canvas

        và đối tượng Axes của nó được TÁI SỬ DỤNG xuyên suốt vòng đời phần

        mềm (draw_plan_view/draw_profile_view chỉ gọi ax.clear() chứ không

        tạo Axes mới), nên không cần ngắt/kết nối lại mỗi lần vẽ."""

        ax = canvas.ax

        canvas.mpl_connect("button_press_event",

                            lambda e, ax=ax, k=key: self._on_overlay_press(e, ax, k))

        canvas.mpl_connect("motion_notify_event",

                            lambda e, ax=ax, k=key: self._on_overlay_motion(e, ax, k))

        canvas.mpl_connect("button_release_event",

                            lambda e, ax=ax, k=key: self._on_overlay_release(e, ax, k))

        canvas.mpl_connect("pick_event", self.on_tree_picked)



    @staticmethod

    def _event_axes_frac(ax, event):

        """Chuyển toạ độ pixel màn hình của sự kiện chuột sang toạ độ phân

        số theo hệ trục (0..1) - dùng chung cho la bàn & thước tỷ lệ vì cả

        2 đều được vẽ bằng ax.transAxes."""

        if event.x is None or event.y is None:

            return None

        try:

            return ax.transAxes.inverted().transform((event.x, event.y))

        except Exception:

            return None



    def _on_overlay_press(self, event, ax, key):

        if event.inaxes != ax or event.button != 1:

            return

        frac = self._event_axes_frac(ax, event)

        if frac is None:

            return

        fx, fy = frac

        state = self._overlay_drag[key]

        na_bbox = getattr(ax, "_north_arrow_bbox", None)

        sb_bbox = getattr(ax, "_scale_bar_bbox", None)

        if na_bbox and _point_in_bbox(fx, fy, na_bbox):

            pos = self.style.get(f"{key}_north_arrow_pos") or _DEFAULT_NORTH_POS

            state["target"] = "north"

            state["offset"] = (fx - pos[0], fy - pos[1])

        elif sb_bbox and _point_in_bbox(fx, fy, sb_bbox):

            pos = self.style.get(f"{key}_scalebar_pos") or _DEFAULT_SCALEBAR_POS

            state["target"] = "scale"

            state["offset"] = (fx - pos[0], fy - pos[1])

        else:

            state["target"] = None



    def _on_overlay_motion(self, event, ax, key):

        state = self._overlay_drag[key]

        target = state.get("target")

        if not target or event.inaxes != ax:

            return

        frac = self._event_axes_frac(ax, event)

        if frac is None:

            return

        fx, fy = frac

        ox, oy = state["offset"]

        new_pos = (min(max(fx - ox, 0.0), 1.0), min(max(fy - oy, 0.0), 1.0))

        if target == "north":

            self.style[f"{key}_north_arrow_pos"] = new_pos

            for art in getattr(ax, "_north_arrow_artists", []):

                try:

                    art.remove()

                except Exception:

                    pass

            _draw_north_arrow(ax, self.style, key)

        else:

            self.style[f"{key}_scalebar_pos"] = new_pos

            for art in getattr(ax, "_scale_bar_artists", []):

                try:

                    art.remove()

                except Exception:

                    pass

            _draw_scale_bar(ax, self.style, key)

        ax.figure.canvas.draw_idle()



    def _on_overlay_release(self, event, ax, key):

        self._overlay_drag[key]["target"] = None



    def on_tree_picked(self, event):

        """Bật hộp thoại thông tin cây khi người dùng click vào 1 cây - dùng

        chung cho CẢ 4 biểu đồ có gắn picker: Trắc đồ ngang (2D), Trắc đồ dọc

        (2D), Độ che phủ tán và Sơ đồ phân tầng đứng.



        2 trường hợp artist mang thông tin cây (xem nơi vẽ, hàm draw_*):

          - `_tree_row` : artist vẽ RIÊNG cho 1 cây (đa giác tán trên trắc đồ

            ngang/dọc) -> lấy thẳng ra dòng dữ liệu (pandas Series) của cây đó.

          - `_tree_rows_df` : artist gộp NHIỀU cây trong 1 lần vẽ (chấm điểm

            trên Độ che phủ tán, hoặc scatter theo tầng trên Sơ đồ phân tầng

            đứng) -> dùng `event.ind` (chỉ số điểm được click, do matplotlib

            trả về) để tra đúng dòng dữ liệu tương ứng trong DataFrame đã lưu

            kèm theo artist đó."""

        mouse_event = getattr(event, "mouseevent", None)

        if mouse_event is not None and mouse_event.button != 1:

            return

        # Bỏ qua nếu đang trong thao tác kéo thả la bàn/thước tỷ lệ, tránh

        # vừa kéo vừa bật hộp thoại thông tin cây không mong muốn.

        if any(state.get("target") for state in self._overlay_drag.values()):

            return

        row = getattr(event.artist, "_tree_row", None)

        if row is None:

            rows_df = getattr(event.artist, "_tree_rows_df", None)

            ind = getattr(event, "ind", None)

            if rows_df is not None and ind is not None and len(ind):

                pos = int(ind[0])

                if 0 <= pos < len(rows_df):

                    row = rows_df.iloc[pos]

        if row is None:

            return

        dlg = TreeInfoDialog(row, self)

        dlg.exec()



    # ---- tab: 3D ----

    def _build_3d_tab(self):

        w = QWidget(); lay = QVBoxLayout(w)



        # Hàng điều khiển được đặt trong QScrollArea cuộn ngang - khi cửa sổ

        # thu nhỏ, các nút/slider không bị cắt hay tràn ra ngoài mà có thể

        # kéo ngang để xem hết.

        controls_widget = QWidget()

        controls = QHBoxLayout(controls_widget)

        controls.setContentsMargins(2, 2, 2, 2)



        controls.addWidget(QLabel("Góc nhìn (Elevation):"))

        self.elev_slider = QSlider(Qt.Horizontal)

        self.elev_slider.setRange(-90, 90); self.elev_slider.setValue(25)

        self.elev_slider.setFixedWidth(110)

        self.elev_slider.valueChanged.connect(self.on_3d_angle_changed)

        controls.addWidget(self.elev_slider)



        controls.addWidget(QLabel("Xoay (Azimuth):"))

        self.azim_slider = QSlider(Qt.Horizontal)

        self.azim_slider.setRange(0, 360); self.azim_slider.setValue(300)

        self.azim_slider.setFixedWidth(140)

        self.azim_slider.valueChanged.connect(self.on_3d_angle_changed)

        controls.addWidget(self.azim_slider)



        self.autorotate_chk = QCheckBox("🔁 Tự động xoay")

        self.autorotate_chk.stateChanged.connect(self.on_autorotate_toggled)

        controls.addWidget(self.autorotate_chk)



        btn_reset = QPushButton("Đặt lại góc nhìn")

        btn_reset.setObjectName("secondary")

        btn_reset.clicked.connect(self.reset_3d_view)

        controls.addWidget(btn_reset)

        controls.addStretch(1)



        controls_scroll = QScrollArea()

        controls_scroll.setWidgetResizable(True)

        controls_scroll.setWidget(controls_widget)

        controls_scroll.setFixedHeight(52)

        controls_scroll.setFrameShape(QFrame.NoFrame)

        controls_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        controls_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        lay.addWidget(controls_scroll)



        self.canvas_3d = MplCanvas(three_d=True)

        toolbar = NavigationToolbar(self.canvas_3d, w)

        lay.addWidget(toolbar)

        lay.addWidget(self.canvas_3d)

        return w



    def on_3d_angle_changed(self):

        self.canvas_3d.ax.view_init(elev=self.elev_slider.value(), azim=self.azim_slider.value())

        self.canvas_3d.draw_idle()



    def on_autorotate_toggled(self, state):

        if self.autorotate_chk.isChecked():

            self.rotate_timer.start()

        else:

            self.rotate_timer.stop()



    def _auto_rotate_tick(self):

        az = (self.azim_slider.value() + 2) % 360

        self.azim_slider.blockSignals(True)

        self.azim_slider.setValue(az)

        self.azim_slider.blockSignals(False)

        self.canvas_3d.ax.view_init(elev=self.elev_slider.value(), azim=az)

        self.canvas_3d.draw_idle()



    def reset_3d_view(self):

        self.elev_slider.setValue(25)

        self.azim_slider.setValue(300)

        self.on_3d_angle_changed()



    # ---- tab: phân tích cấu trúc rừng ----

    def _build_structure_tab(self):

        w = QWidget()

        outer = QHBoxLayout(w)

        outer.setContentsMargins(0, 0, 0, 0)



        # ----- panel trái: tham số + bảng chỉ tiêu tổng hợp -----

        # Trước đây left_scroll bị setFixedWidth(300) + tắt cuộn ngang nên

        # nội dung rộng hơn 300px (ô màu, spin cỡ chữ...) bị CHE MẤT không

        # cách nào xem được. Nay đặt trong QSplitter (giống sidebar chính)

        # để người dùng tự kéo ngang nới rộng/thu hẹp panel trái theo ý,

        # đồng thời vẫn giữ QScrollArea để cuộn dọc khi nội dung dài, và bật

        # cuộn ngang dự phòng (AsNeeded) phòng khi panel bị kéo quá hẹp.

        struct_splitter = QSplitter(Qt.Horizontal)

        self.struct_splitter = struct_splitter



        left_scroll = QScrollArea()

        left_scroll.setWidgetResizable(True)

        left_scroll.setMinimumWidth(260)

        left_scroll.setMaximumWidth(900)

        left_scroll.setFrameShape(QFrame.NoFrame)

        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)



        left = QWidget()

        left.setMinimumWidth(300)

        left_lay = QVBoxLayout(left)

        left_lay.setContentsMargins(0, 0, 6, 0)



        param_box = QGroupBox("Tham số cấu trúc ngang")

        param_form = QFormLayout(param_box)

        self.struct_binwidth_spin = QDoubleSpinBox()

        self.struct_binwidth_spin.setRange(1.0, 20.0)

        self.struct_binwidth_spin.setSingleStep(0.5)

        self.struct_binwidth_spin.setValue(self.style["struct_bin_width"])

        self.struct_binwidth_spin.valueChanged.connect(self.on_struct_params_changed)

        param_form.addRow("Cấp kính D1.3 (cm):", self.struct_binwidth_spin)



        self.struct_gridres_spin = QDoubleSpinBox()

        self.struct_gridres_spin.setDecimals(0)

        self.struct_gridres_spin.setRange(40, 400)

        self.struct_gridres_spin.setSingleStep(20)

        self.struct_gridres_spin.setValue(self.style["struct_grid_res"])

        param_form.addRow("Độ phân giải lưới:", self.struct_gridres_spin)



        left_lay.addWidget(param_box)



        vparam_box = QGroupBox("Tham số cấu trúc đứng (Hvn)")

        vparam_form = QFormLayout(vparam_box)

        self.vstruct_binwidth_spin = QDoubleSpinBox()

        self.vstruct_binwidth_spin.setRange(0.5, 10.0)

        self.vstruct_binwidth_spin.setSingleStep(0.5)

        self.vstruct_binwidth_spin.setValue(self.style["vstruct_bin_width"])

        self.vstruct_binwidth_spin.valueChanged.connect(self.on_struct_params_changed)

        vparam_form.addRow("Cấp chiều cao Hvn (m):", self.vstruct_binwidth_spin)



        self.vstruct_kstd_spin = QDoubleSpinBox()

        self.vstruct_kstd_spin.setRange(0.2, 3.0)

        self.vstruct_kstd_spin.setSingleStep(0.1)

        self.vstruct_kstd_spin.setDecimals(2)

        self.vstruct_kstd_spin.setValue(self.style["vstruct_k_std"])

        self.vstruct_kstd_spin.setToolTip(

            "Hệ số k trong ngưỡng phân tầng Htb ± k·σ:\n"

            "  • Hvn > Htb + k·σ  → Tầng trội (A1)\n"

            "  • Htb − k·σ ≤ Hvn ≤ Htb + k·σ  → Tầng ưu thế (A2)\n"

            "  • Hvn < Htb − k·σ  → Tầng dưới tán (A3)\n"

            "k càng nhỏ, tầng trội/dưới tán càng rộng (nhiều cây rơi vào 2 tầng này).")

        self.vstruct_kstd_spin.valueChanged.connect(self.on_struct_params_changed)

        vparam_form.addRow("Hệ số phân tầng k (×σ):", self.vstruct_kstd_spin)



        left_lay.addWidget(vparam_box)



        summary_box = QGroupBox("Chỉ tiêu cấu trúc ngang")

        summary_outer = QVBoxLayout(summary_box)

        summary_form = QFormLayout()

        summary_form.setLabelAlignment(Qt.AlignLeft)

        self.struct_labels = {}

        rows = [

            ("n_area", "Số cây / Diện tích OTC:"),

            ("density", "Mật độ cây/ha:"),

            ("d13_mean", "D1.3 trung bình:"),

            ("dg", "Đường kính bình quân Dg:"),

            ("g_mean", "G trung bình (1 cây):"),

            ("g_per_ha", "G/ha:"),

            ("crown_mean", "Diện tích tán trung bình:"),

            ("crown_total", "Tổng diện tích tán (ΣAi):"),

            ("cover_pct", "Độ che phủ tán:"),

            ("overlap_idx", "Chỉ số chồng lấn tán (CI):"),

        ]

        self._struct_summary_rows = rows

        for key, label in rows:

            val_lbl = QLabel("-")

            val_lbl.setStyleSheet("font-weight:600; color:#1b5e20;")

            val_lbl.setWordWrap(True)

            self.struct_labels[key] = val_lbl

            lab = QLabel(label)

            lab.setWordWrap(True)

            summary_form.addRow(lab, val_lbl)

        summary_outer.addLayout(summary_form)

        btn_copy_h = QPushButton("📋  Copy chỉ tiêu cấu trúc ngang")

        btn_copy_h.clicked.connect(self.copy_struct_summary_horizontal)

        summary_outer.addWidget(btn_copy_h)

        left_lay.addWidget(summary_box)



        vsummary_box = QGroupBox("Chỉ tiêu cấu trúc đứng (theo Hvn)")

        vsummary_outer = QVBoxLayout(vsummary_box)

        vsummary_form = QFormLayout()

        vsummary_form.setLabelAlignment(Qt.AlignLeft)

        self.vstruct_labels = {}

        vrows = [

            ("hvn_mean", "Hvn trung bình:"),

            ("hvn_max", "Hvn max:"),

            ("hvn_min", "Hvn min:"),

            ("canopy_height", "Chiều cao tầng rừng (tán chính):"),

            ("layer_a1", "Tầng trội (A1):"),

            ("layer_a2", "Tầng ưu thế (A2):"),

            ("layer_a3", "Tầng dưới tán (A3):"),

        ]

        self._vstruct_summary_rows = vrows

        for key, label in vrows:

            val_lbl = QLabel("-")

            val_lbl.setStyleSheet("font-weight:600; color:#00695c;")

            val_lbl.setWordWrap(True)

            self.vstruct_labels[key] = val_lbl

            lab = QLabel(label)

            lab.setWordWrap(True)

            vsummary_form.addRow(lab, val_lbl)

        vsummary_outer.addLayout(vsummary_form)

        btn_copy_v = QPushButton("📋  Copy chỉ tiêu cấu trúc đứng")

        btn_copy_v.clicked.connect(self.copy_struct_summary_vertical)

        vsummary_outer.addWidget(btn_copy_v)

        left_lay.addWidget(vsummary_box)



        btn_canopy = QPushButton("🧮  Tính độ che phủ - chồng lấn tán")

        btn_canopy.clicked.connect(self.compute_canopy_and_draw)

        left_lay.addWidget(btn_canopy)

        note = QLabel("Độ che phủ & chỉ số chồng lấn tán cần tính riêng "

                       "(rasterize toàn bộ OTC) nên không tự chạy mỗi lần "

                       "đổi tham số khác - bấm nút trên để tính lại.")

        note.setWordWrap(True)

        note.setStyleSheet("color:#6b7d6d; font-size:8.6pt;")

        left_lay.addWidget(note)



        left_lay.addWidget(self._build_struct_chart_style_group())

        left_lay.addStretch(1)



        left_scroll.setWidget(left)

        struct_splitter.addWidget(left_scroll)



        # ----- panel phải: các biểu đồ phân bố -----

        self.struct_tabs = QTabWidget()
        self.struct_tabs.setObjectName("subTabs")



        nd_w = QWidget(); nd_lay = QVBoxLayout(nd_w)

        self.canvas_struct_nd = MplCanvas(three_d=False)

        nd_lay.addWidget(NavigationToolbar(self.canvas_struct_nd, nd_w))

        nd_lay.addWidget(self.canvas_struct_nd)

        self.struct_tabs.addTab(nd_w, "Phân bố N/D")



        ndg_w = QWidget(); ndg_lay = QVBoxLayout(ndg_w)

        self.canvas_struct_ndg = MplCanvas(three_d=False)

        ndg_lay.addWidget(NavigationToolbar(self.canvas_struct_ndg, ndg_w))

        ndg_lay.addWidget(self.canvas_struct_ndg)

        self.struct_tabs.addTab(ndg_w, "Phân bố N/Dg")



        g_w = QWidget(); g_lay = QVBoxLayout(g_w)

        self.canvas_struct_g = MplCanvas(three_d=False)

        g_lay.addWidget(NavigationToolbar(self.canvas_struct_g, g_w))

        g_lay.addWidget(self.canvas_struct_g)

        self.struct_tabs.addTab(g_w, "Phân bố tiết diện ngang (G)")



        cover_w = QWidget(); cover_lay = QVBoxLayout(cover_w)

        self.canvas_struct_cover = MplCanvas(three_d=False)

        # Click vào chấm điểm 1 cây trên biểu đồ Độ che phủ tán để xem thông
        # tin chi tiết - cùng cơ chế picker/pick_event với trắc đồ ngang/dọc
        # (xem draw_canopy_coverage và MainWindow.on_tree_picked).
        self.canvas_struct_cover.mpl_connect("pick_event", self.on_tree_picked)

        cover_lay.addWidget(NavigationToolbar(self.canvas_struct_cover, cover_w))

        cover_lay.addWidget(self._make_tree_click_hint_simple())

        cover_lay.addWidget(self.canvas_struct_cover)

        self.struct_tabs.addTab(cover_w, "Độ che phủ tán")



        # ----- các tab con thuộc CẤU TRÚC ĐỨNG (theo Hvn) -----

        nh_w = QWidget(); nh_lay = QVBoxLayout(nh_w)

        self.canvas_struct_nh = MplCanvas(three_d=False)

        nh_lay.addWidget(NavigationToolbar(self.canvas_struct_nh, nh_w))

        nh_lay.addWidget(self.canvas_struct_nh)

        self.struct_tabs.addTab(nh_w, "Phân bố N/H")



        layers_w = QWidget(); layers_lay = QVBoxLayout(layers_w)

        self.canvas_struct_layers = MplCanvas(three_d=False)

        # Click vào chấm điểm 1 cây trên Sơ đồ phân tầng đứng để xem thông
        # tin chi tiết - cùng cơ chế picker/pick_event với trắc đồ ngang/dọc
        # (xem draw_forest_layers và MainWindow.on_tree_picked).
        self.canvas_struct_layers.mpl_connect("pick_event", self.on_tree_picked)

        layers_lay.addWidget(NavigationToolbar(self.canvas_struct_layers, layers_w))

        layers_lay.addWidget(self._make_tree_click_hint_simple())

        layers_lay.addWidget(self.canvas_struct_layers)

        self.struct_tabs.addTab(layers_w, "Sơ đồ phân tầng đứng")
        self.struct_tabs.addTab(self._build_composition_subtab(), "🌿 Tổ thành & đa dạng loài")



        # Danh sách canvas theo ĐÚNG thứ tự các sub-tab ở trên, dùng để tra

        # cứu canvas đang hiển thị khi xuất ảnh (xem export_current_view).

        self.struct_canvas_list = [

            self.canvas_struct_nd, self.canvas_struct_ndg,

            self.canvas_struct_g, self.canvas_struct_cover,

            self.canvas_struct_nh, self.canvas_struct_layers,
            self.canvas_struct_comp,
        ]



        struct_splitter.addWidget(self.struct_tabs)

        struct_splitter.setStretchFactor(0, 0)

        struct_splitter.setStretchFactor(1, 1)

        struct_splitter.setCollapsible(0, False)

        struct_splitter.setCollapsible(1, False)

        struct_splitter.setSizes([320, 900])

        outer.addWidget(struct_splitter)

        return w



    def _build_struct_chart_style_group(self):

        """Khối "🎨 Tiêu đề / nhãn trục / màu sắc / kiểu chữ" - cho phép

        chỉnh riêng từng biểu đồ con trong tab "Phân tích cấu trúc rừng":

        chọn biểu đồ ở combo trên cùng, các ô bên dưới tự nạp giá trị đang

        có của biểu đồ đó và áp dụng ngay khi gõ/chọn (cùng cơ chế với mục

        ⑤ Nhãn & phông chữ biểu đồ của trắc đồ ngang/dọc/3D)."""

        box = QGroupBox("🎨 Tiêu đề / nhãn trục / màu sắc / kiểu chữ")

        lay = QVBoxLayout(box)



        self.struct_chart_combo = QComboBox()

        for key, label, _has_x, _has_c, _has_legend in STRUCT_CHART_OPTIONS:

            self.struct_chart_combo.addItem(label, userData=key)

        self.struct_chart_combo.currentIndexChanged.connect(self.on_struct_chart_selected)

        lay.addWidget(self.struct_chart_combo)



        form = QFormLayout()

        self.struct_title_edit = QLineEdit()

        self.struct_title_edit.textChanged.connect(self.on_struct_chart_field_changed)

        form.addRow("Tiêu đề:", self.struct_title_edit)



        self.struct_xlabel_edit = QLineEdit()

        self.struct_xlabel_edit.textChanged.connect(self.on_struct_chart_field_changed)

        self.struct_xlabel_row = self._row_widgets(form, "Tên trục X:", self.struct_xlabel_edit)



        self.struct_ylabel_edit = QLineEdit()

        self.struct_ylabel_edit.textChanged.connect(self.on_struct_chart_field_changed)

        form.addRow("Tên trục Y:", self.struct_ylabel_edit)



        self.struct_color_btn = QPushButton("")

        self.struct_color_btn.setFixedHeight(24)

        self.struct_color_btn.clicked.connect(self.choose_struct_chart_color)

        self.struct_color_row = self._row_widgets(form, "Màu sắc biểu đồ:", self.struct_color_btn)



        self.struct_fontcolor_btn = QPushButton("")

        self.struct_fontcolor_btn.setFixedHeight(24)

        self.struct_fontcolor_btn.clicked.connect(self.choose_struct_chart_fontcolor)

        form.addRow("Màu chữ:", self.struct_fontcolor_btn)



        self.struct_font_combo = QComboBox()

        self.struct_font_combo.addItem("(theo mục ⑤ chung)", userData="")

        for f in FONT_FAMILY_OPTIONS:

            self.struct_font_combo.addItem(f, userData=f)

        self.struct_font_combo.currentIndexChanged.connect(self.on_struct_chart_field_changed)

        form.addRow("Kiểu chữ (font):", self.struct_font_combo)



        # Cỡ chữ 0 = "theo mục ⑤ chung" (title_size/label_size toàn cục) -

        # spin cho phép về 0 để người dùng "bỏ trống" cỡ chữ riêng dễ dàng.

        self.struct_titlesize_spin = QDoubleSpinBox()

        self.struct_titlesize_spin.setRange(0, 24)

        self.struct_titlesize_spin.setSpecialValueText("(theo mục ⑤ chung)")

        self.struct_titlesize_spin.valueChanged.connect(self.on_struct_chart_field_changed)

        form.addRow("Cỡ chữ tiêu đề:", self.struct_titlesize_spin)



        self.struct_labelsize_spin = QDoubleSpinBox()

        self.struct_labelsize_spin.setRange(0, 20)

        self.struct_labelsize_spin.setSpecialValueText("(theo mục ⑤ chung)")

        self.struct_labelsize_spin.valueChanged.connect(self.on_struct_chart_field_changed)

        form.addRow("Cỡ chữ trục:", self.struct_labelsize_spin)



        # Chỉ "Độ che phủ tán" (hộp thông tin) và "Sơ đồ phân tầng đứng"

        # (chú dẫn tầng rừng) mới có chữ chú dẫn riêng để chỉnh cỡ - ẩn/hiện

        # hàng này theo has_legend của biểu đồ đang chọn (giống Tên trục X).

        self.struct_legendsize_spin = QDoubleSpinBox()

        self.struct_legendsize_spin.setRange(0, 24)

        self.struct_legendsize_spin.setSpecialValueText("(theo Cỡ chữ trục)")

        self.struct_legendsize_spin.valueChanged.connect(self.on_struct_chart_field_changed)

        self.struct_legendsize_row = self._row_widgets(form, "Cỡ chữ chú dẫn:", self.struct_legendsize_spin)

        # Bật/tắt nhãn từng cây (số TT / tên loài - dùng chung "Nội dung
        # nhãn" ở khung ⑤ của trắc đồ ngang/dọc) ngay trên biểu đồ - giống
        # cơ chế "Hiển thị nhãn cây" của trắc đồ ngang/dọc. Chỉ có ý nghĩa
        # với "Độ che phủ tán" và "Sơ đồ phân tầng đứng" (2 biểu đồ có vẽ
        # từng chấm/cây riêng lẻ) nên ẩn/hiện theo has_legend, trùng đúng 2
        # biểu đồ này.
        self.struct_showlabels_chk = QCheckBox("Hiển thị nhãn cây trên biểu đồ")
        self.struct_showlabels_chk.stateChanged.connect(self.on_struct_chart_field_changed)
        self.struct_showlabels_row = self._row_widgets(form, "", self.struct_showlabels_chk)

        lay.addLayout(form)



        btn_reset = QPushButton("↺  Khôi phục mặc định biểu đồ này")

        btn_reset.clicked.connect(self.reset_struct_chart_style)

        lay.addWidget(btn_reset)



        note = QLabel("Để trống Tiêu đề / Tên trục sẽ dùng nội dung mặc định. "

                       "Màu chữ / kiểu chữ / cỡ chữ để trống (hoặc 0) sẽ dùng theo mục ⑤ chung. "

                       "Riêng \"Độ che phủ tán\" và \"Sơ đồ phân tầng đứng\" có thêm ô "

                       "\"Cỡ chữ chú dẫn\" - chỉnh riêng cỡ chữ trong hộp chú dẫn/thông tin, "

                       "để 0 sẽ dùng theo \"Cỡ chữ trục\".")

        note.setWordWrap(True)

        note.setStyleSheet("color:#6b7d6d; font-size:8.4pt;")

        lay.addWidget(note)



        self._loading_struct_chart_fields = False

        self.on_struct_chart_selected()

        return box



    @staticmethod

    def _row_widgets(form, label_text, widget):

        """Thêm 1 hàng vào QFormLayout, trả về (QLabel, widget) để có thể

        ẩn/hiện cả hàng (setVisible) khi biểu đồ đang chọn không hỗ trợ

        trường đó (VD: sơ đồ phân tầng không có trục X / không có 1 màu

        chủ đạo duy nhất)."""

        lab = QLabel(label_text)

        form.addRow(lab, widget)

        return (lab, widget)



    def _current_struct_chart_key(self):

        idx = self.struct_chart_combo.currentIndex()

        return self.struct_chart_combo.itemData(idx) if idx >= 0 else "nd"



    def on_struct_chart_selected(self):

        """Nạp lại nội dung các ô nhập theo biểu đồ vừa chọn trong combo,

        đồng thời ẩn/hiện Tên trục X và Màu sắc nếu biểu đồ đó không hỗ trợ."""

        key = self._current_struct_chart_key()

        opt = next((o for o in STRUCT_CHART_OPTIONS if o[0] == key), STRUCT_CHART_OPTIONS[0])

        has_xlabel, has_color, has_legend = opt[2], opt[3], opt[4]

        cs = self.style.setdefault("struct_chart_style", {}).setdefault(key, {})



        self._loading_struct_chart_fields = True

        self.struct_title_edit.setText(cs.get("title", ""))

        self.struct_xlabel_edit.setText(cs.get("xlabel", ""))

        self.struct_ylabel_edit.setText(cs.get("ylabel", ""))

        color = cs.get("color") or "#43a047"

        self._set_swatch(self.struct_color_btn, color)

        fontcolor = cs.get("font_color") or self.style.get("font_color", "#1b5e20")

        self._set_swatch(self.struct_fontcolor_btn, fontcolor)

        fam_idx = self.struct_font_combo.findData(cs.get("font_family", ""))

        self.struct_font_combo.setCurrentIndex(fam_idx if fam_idx >= 0 else 0)

        self.struct_titlesize_spin.setValue(cs.get("title_size") or 0)

        self.struct_labelsize_spin.setValue(cs.get("label_size") or 0)

        self.struct_legendsize_spin.setValue(cs.get("legend_size") or 0)

        self.struct_showlabels_chk.setChecked(bool(cs.get("show_labels", False)))

        self._loading_struct_chart_fields = False



        self.struct_xlabel_row[0].setVisible(has_xlabel)

        self.struct_xlabel_row[1].setVisible(has_xlabel)

        self.struct_color_row[0].setVisible(has_color)

        self.struct_color_row[1].setVisible(has_color)

        self.struct_legendsize_row[0].setVisible(has_legend)

        self.struct_legendsize_row[1].setVisible(has_legend)

        self.struct_showlabels_row[0].setVisible(has_legend)

        self.struct_showlabels_row[1].setVisible(has_legend)



    def on_struct_chart_field_changed(self):

        if getattr(self, "_loading_struct_chart_fields", False):

            return

        key = self._current_struct_chart_key()

        cs = self.style.setdefault("struct_chart_style", {}).setdefault(key, {})

        cs["title"] = self.struct_title_edit.text().strip()

        cs["xlabel"] = self.struct_xlabel_edit.text().strip()

        cs["ylabel"] = self.struct_ylabel_edit.text().strip()

        cs["font_family"] = self.struct_font_combo.currentData() or ""

        cs["title_size"] = self.struct_titlesize_spin.value() or None

        cs["label_size"] = self.struct_labelsize_spin.value() or None

        cs["legend_size"] = self.struct_legendsize_spin.value() or None

        cs["show_labels"] = self.struct_showlabels_chk.isChecked()

        self.update_structure_analysis()



    def choose_struct_chart_color(self):

        key = self._current_struct_chart_key()

        cs = self.style.setdefault("struct_chart_style", {}).setdefault(key, {})

        current = QColor(cs.get("color") or "#43a047")

        c = QColorDialog.getColor(current, self, "Chọn màu sắc biểu đồ")

        if c.isValid():

            cs["color"] = c.name()

            self._set_swatch(self.struct_color_btn, c.name())

            self.update_structure_analysis()



    def choose_struct_chart_fontcolor(self):

        key = self._current_struct_chart_key()

        cs = self.style.setdefault("struct_chart_style", {}).setdefault(key, {})

        current = QColor(cs.get("font_color") or self.style.get("font_color", "#1b5e20"))

        c = QColorDialog.getColor(current, self, "Chọn màu chữ")

        if c.isValid():

            cs["font_color"] = c.name()

            self._set_swatch(self.struct_fontcolor_btn, c.name())

            self.update_structure_analysis()



    def reset_struct_chart_style(self):

        key = self._current_struct_chart_key()

        default_colors = {o[0]: None for o in STRUCT_CHART_OPTIONS}

        default_colors.update({"nd": "#43a047", "ndg": "#7cb342", "g": "#ef6c00",

                                "cover": "#66bb6a", "nh": "#00838f", "layers": "#2e7d32"})

        self.style.setdefault("struct_chart_style", {})[key] = {

            "title": "", "xlabel": "", "ylabel": "",

            "color": default_colors.get(key, "#43a047"), "font_color": "", "font_family": "",

            "title_size": None, "label_size": None, "legend_size": None, "show_labels": False,

        }

        self.on_struct_chart_selected()

        self.update_structure_analysis()



    def _copy_struct_summary(self, rows, labels_dict, title, status_label=None):

        """Ghép các dòng "Nhãn: Giá trị" (rows lấy từ _struct_summary_rows /

        _vstruct_summary_rows, labels_dict là struct_labels / vstruct_labels)

        thành văn bản, copy vào clipboard hệ thống - dán được thẳng vào

        Word/Excel/Zalo... Mỗi dòng cũng ngăn cách bằng TAB giữa nhãn và giá

        trị nên dán vào Excel sẽ tự tách thành 2 cột."""

        lines = [title]

        for key, label in rows:

            val = labels_dict.get(key)

            val_txt = val.text() if val is not None else "-"

            lab_txt = label[:-1] if label.endswith(":") else label

            lines.append(f"{lab_txt}\t{val_txt}")

        text = "\n".join(lines)

        QApplication.clipboard().setText(text)

        show_notice(self, "info", "Đã copy",

                    f"Đã copy \"{title}\" vào clipboard.\nBạn dán được thẳng vào Word/Excel.")



    def copy_struct_summary_horizontal(self):

        if not hasattr(self, "struct_labels") or self.df is None or self.df.empty:

            show_notice(self, "warning", "Chưa có dữ liệu",

                        "Chưa có kết quả chỉ tiêu cấu trúc ngang để copy.")

            return

        self._copy_struct_summary(self._struct_summary_rows, self.struct_labels,

                                   "CHỈ TIÊU CẤU TRÚC NGANG")



    def copy_struct_summary_vertical(self):

        if not hasattr(self, "vstruct_labels") or self.df is None or self.df.empty:

            show_notice(self, "warning", "Chưa có dữ liệu",

                        "Chưa có kết quả chỉ tiêu cấu trúc đứng để copy.")

            return

        self._copy_struct_summary(self._vstruct_summary_rows, self.vstruct_labels,

                                   "CHỈ TIÊU CẤU TRÚC ĐỨNG (theo Hvn)")



    def on_struct_params_changed(self):

        self.style["struct_bin_width"] = self.struct_binwidth_spin.value()

        self.style["vstruct_bin_width"] = self.vstruct_binwidth_spin.value()

        self.style["vstruct_k_std"] = self.vstruct_kstd_spin.value()

        self.update_structure_analysis()



    def _update_structure_analysis_core(self):

        """Cập nhật bảng chỉ tiêu tổng hợp + 3 biểu đồ phân bố (N/D, N/Dg,

        G) - nhẹ, không rasterize nên có thể gọi lại mỗi khi refresh_views().

        Độ che phủ tán KHÔNG tính lại ở đây (xem compute_canopy_and_draw)."""

        if not hasattr(self, "struct_labels"):

            return

        if self.df is None or self.df.empty:

            for lbl in self.struct_labels.values():

                lbl.setText("-")

            if hasattr(self, "vstruct_labels"):

                for lbl in self.vstruct_labels.values():

                    lbl.setText("-")

            self._vstruct_result = None

            empty_df = self.df if self.df is not None else pd.DataFrame()

            draw_nd_distribution(self.canvas_struct_nd.ax, empty_df, self.style, self.style["struct_bin_width"])

            self.canvas_struct_nd.draw_idle()

            draw_ndg_distribution(self.canvas_struct_ndg.ax, empty_df, self.style)

            self.canvas_struct_ndg.draw_idle()

            draw_g_distribution(self.canvas_struct_g.ax, empty_df, self.style, self.style["struct_bin_width"])

            self.canvas_struct_g.draw_idle()

            draw_canopy_coverage(self.canvas_struct_cover.ax, empty_df, self.style, None)

            self.canvas_struct_cover.draw_idle()

            if hasattr(self, "canvas_struct_nh"):

                draw_nh_distribution(self.canvas_struct_nh.ax, empty_df, self.style,

                                      self.style["vstruct_bin_width"], None)

                self.canvas_struct_nh.draw_idle()

                draw_forest_layers(self.canvas_struct_layers.ax, empty_df, self.style, None)

                self.canvas_struct_layers.draw_idle()

            return



        summary = compute_structure_summary(self.df, self.style)

        L = self.struct_labels

        area_txt = f"{summary['area_ha']:.3f} ha" if summary["area_ha"] else "chưa nhập kích thước ô"

        L["n_area"].setText(f"{summary['n']} cây / {area_txt}")

        L["density"].setText(f"{summary['density_ha']:.0f} cây/ha" if summary["density_ha"] else "-")

        L["d13_mean"].setText(f"{summary['d13_mean']:.1f} cm")

        L["dg"].setText(f"{summary['dg']:.1f} cm" if summary["dg"] else "-")

        L["g_mean"].setText(f"{summary['g_mean_cm2']:.0f} cm²")

        L["g_per_ha"].setText(f"{summary['g_per_ha']:.2f} m²/ha" if summary["g_per_ha"] else "-")

        L["crown_mean"].setText(f"{summary['crown_area_mean']:.1f} m²")

        L["crown_total"].setText(

            f"{summary['crown_area_total']:.1f} m²"

            + (f"  ({summary['crown_area_per_ha']:.0f} m²/ha)" if summary["crown_area_per_ha"] else ""))

        if self._canopy_result is not None:

            L["cover_pct"].setText(f"{self._canopy_result['cover_pct']:.1f}%")

            L["overlap_idx"].setText(f"{self._canopy_result['overlap_index']:.2f}")

        else:

            L["cover_pct"].setText("chưa tính")

            L["overlap_idx"].setText("chưa tính")



        bin_width = self.style["struct_bin_width"]

        draw_nd_distribution(self.canvas_struct_nd.ax, self.df, self.style, bin_width)

        self.canvas_struct_nd.draw_idle()

        draw_ndg_distribution(self.canvas_struct_ndg.ax, self.df, self.style)

        self.canvas_struct_ndg.draw_idle()

        draw_g_distribution(self.canvas_struct_g.ax, self.df, self.style, bin_width)

        self.canvas_struct_g.draw_idle()

        draw_canopy_coverage(self.canvas_struct_cover.ax, self.df, self.style, self._canopy_result)

        self.canvas_struct_cover.draw_idle()



        # ----- cấu trúc đứng (theo Hvn): Hvn tb/max/min, phân tầng rừng,

        # phân bố N/H - tính lại mỗi khi refresh (nhẹ, không rasterize) -----

        if hasattr(self, "vstruct_labels"):

            k_std = self.style.get("vstruct_k_std", 1.0)

            vstruct = compute_vertical_structure(self.df, k_std=k_std)

            self._vstruct_result = vstruct

            VL = self.vstruct_labels

            if vstruct is not None:

                VL["hvn_mean"].setText(f"{vstruct['h_mean']:.1f} m")

                VL["hvn_max"].setText(f"{vstruct['h_max']:.1f} m")

                VL["hvn_min"].setText(f"{vstruct['h_min']:.1f} m")

                VL["canopy_height"].setText(f"{vstruct['canopy_height']:.1f} m")

                for code, key in (("A1", "layer_a1"), ("A2", "layer_a2"), ("A3", "layer_a3")):

                    stt = vstruct[code]

                    VL[key].setText(f"{stt['count']} cây ({stt['pct']:.1f}%)" if stt["count"] else "0 cây (0%)")

            else:

                for lbl in VL.values():

                    lbl.setText("-")



            vbin_width = self.style["vstruct_bin_width"]

            if hasattr(self, "canvas_struct_nh"):

                draw_nh_distribution(self.canvas_struct_nh.ax, self.df, self.style, vbin_width, vstruct)

                self.canvas_struct_nh.draw_idle()

                draw_forest_layers(self.canvas_struct_layers.ax, self.df, self.style, vstruct)

                self.canvas_struct_layers.draw_idle()



    def compute_canopy_and_draw(self):

        if self.df is None or self.df.empty:

            show_notice(self, "info", "Độ che phủ tán", "Vui lòng nạp dữ liệu cây rừng trước.")

            return

        if not self.style.get("field_w") or not self.style.get("field_h"):

            show_notice(self, "warning", "Độ che phủ tán",

                        "Vui lòng nhập kích thước ô tiêu chuẩn (mục ② Kích thước ô) trước khi tính.")

            return

        self.style["struct_grid_res"] = int(self.struct_gridres_spin.value())

        self.statusBar().showMessage("Đang tính độ che phủ & chồng lấn tán...")

        QApplication.processEvents()

        self._canopy_result = compute_canopy_coverage(self.df, self.style, self.style["struct_grid_res"])

        self.update_structure_analysis()

        self.statusBar().showMessage("Đã tính xong độ che phủ & chồng lấn tán.", 5000)



    # ==========================================================================

    # LOGIC DỮ LIỆU

    # ==========================================================================



    def open_files(self):

        paths, _ = QFileDialog.getOpenFileNames(

            self, "Chọn file Excel dữ liệu cây rừng", "", "Excel Files (*.xlsx *.xls)")

        if not paths:

            return

        # Chặn tín hiệu currentIndexChanged khi thêm/đổi mục trong combo, vì

        # addItem() trên combo rỗng và setCurrentIndex() đều tự phát tín hiệu

        # này, gọi load_current_file(ask_mapping=False) "âm thầm" trước khi ta

        # kịp gọi load_current_file(ask_mapping=True) bên dưới - kết quả là

        # hộp thoại "Ánh xạ cột" bị hiện lên 2 lần liên tiếp cho cùng 1 file.

        self.file_combo.blockSignals(True)

        for p in paths:

            if p not in self.file_paths:

                self.file_paths.append(p)

                self.file_combo.addItem(os.path.basename(p), p)

        self.file_combo.setCurrentIndex(self.file_combo.count() - 1)

        self.file_combo.blockSignals(False)

        self.load_current_file(ask_mapping=True)



    def load_demo_dataset(self):

        """Nạp dữ liệu Demo tích hợp sẵn trong phần mềm (không cần file Excel

        ngoài) - thêm 1 mục riêng vào combo "File Excel đang xem" để người

        dùng vẫn thấy rõ đang xem dữ liệu gì, đồng thời có thể quay lại các

        file Excel đã mở trước đó bất kỳ lúc nào."""

        self.file_combo.blockSignals(True)

        idx = self.file_combo.findData(DEMO_MARKER)

        if idx < 0:

            self.file_combo.insertItem(0, "📊 Dữ liệu Demo (tích hợp sẵn - 40 cây)", DEMO_MARKER)

            idx = 0

        self.file_combo.setCurrentIndex(idx)

        self.file_combo.blockSignals(False)

        self.load_current_file(ask_mapping=False)

        self.statusBar().showMessage("Đã tải dữ liệu Demo tích hợp sẵn (40 cây, 8 loài) để dùng thử.", 6000)



    def load_current_file(self, ask_mapping=False):

        path = self.file_combo.currentData()

        if not path:

            return

        if path == DEMO_MARKER:

            # Dữ liệu Demo dựng sẵn trong bộ nhớ, tên cột đã khớp 100% với

            # định dạng chuẩn nên ánh xạ trực tiếp, không cần hộp thoại.

            self.raw_df = get_demo_dataframe()

            self.raw_columns = [str(c) for c in self.raw_df.columns]

            self.column_mapping = dict(DEMO_COLUMNS)

            self.apply_mapping()

            return

        if not os.path.exists(path):

            return

        try:

            self.raw_df = pd.read_excel(path)

        except Exception as e:

            show_notice(self, "error", "Lỗi đọc file", f"Không thể đọc file:\n{path}\n\n{e}")

            return

        self.raw_columns = [str(c) for c in self.raw_df.columns]



        need_mapping = ask_mapping or not self.column_mapping or not any(

            v in self.raw_columns for v in self.column_mapping.values() if v

        )

        if need_mapping:

            dlg = ColumnMappingDialog(self.raw_columns, self.column_mapping, self)

            if dlg.exec() == QDialog.Accepted:

                self.column_mapping = dlg.get_mapping()

            elif not self.column_mapping:

                show_notice(self, "warning", "Ánh xạ cột",

                            "Cần ánh xạ ít nhất các cột X, Y, Hvn, Hdc để vẽ trắc đồ.")

                return

        self.apply_mapping()



    def open_mapping_dialog(self):

        if self.raw_df is None:

            show_notice(self, "info", "Ánh xạ cột", "Hãy mở file Excel trước.")

            return

        dlg = ColumnMappingDialog(self.raw_columns, self.column_mapping, self)

        if dlg.exec() == QDialog.Accepted:

            self.column_mapping = dlg.get_mapping()

            self.apply_mapping()



    def refresh(self):

        if self.file_combo.currentData() is None:

            show_notice(self, "info", "Làm mới", "Chưa có dữ liệu để làm mới. Hãy mở file Excel trước.")

            return

        self.load_current_file(ask_mapping=False)

        self.statusBar().showMessage("Đã làm mới dữ liệu từ file Excel.", 4000)



    def open_column_explain_dialog(self):

        """Mở hộp thoại giải thích ý nghĩa từng cột (biến) dữ liệu - dùng

        được bất kỳ lúc nào, kể cả khi chưa nạp dữ liệu gì."""

        dlg = ColumnExplainDialog(self)

        dlg.exec()



    def export_demo_to_excel(self):

        """Xuất dữ liệu Demo tích hợp sẵn (40 cây, 8 loài) ra file Excel -

        vừa để người dùng xem thử định dạng bảng dữ liệu chuẩn, vừa có thể

        dùng ngay làm khung mẫu (template) để điền dữ liệu thực địa của

        chính mình rồi mở lại bằng chức năng "Mở file Excel"."""

        path, _ = QFileDialog.getSaveFileName(

            self, "Xuất dữ liệu Demo ra Excel",

            "TDForest_DuLieu_Demo.xlsx", "Excel Files (*.xlsx)")

        if not path:

            return

        if not path.lower().endswith(".xlsx"):

            path += ".xlsx"

        try:

            get_demo_dataframe().to_excel(path, index=False, sheet_name="OTC_Demo")

        except Exception as e:

            show_notice(self, "error", "Lỗi xuất Excel", f"Không thể xuất file:\n{path}\n\n{e}")

            return

        show_notice(

            self, "success", "Xuất dữ liệu Demo",

            f"Đã xuất dữ liệu Demo (40 cây, 8 loài) ra file:\n{path}\n\n"

            "Bạn có thể dùng file này làm mẫu định dạng cột để nhập dữ liệu "

            "đo đếm thực tế của mình, sau đó mở lại bằng \"Mở file Excel\".")

        self.statusBar().showMessage(f"Đã xuất dữ liệu Demo ra: {path}", 6000)



    def clear_all_data(self):

        """Xóa toàn bộ dữ liệu hiện tại (file đã mở, ánh xạ cột, màu loài,

        dạng tán riêng theo loài, các trắc đồ đã vẽ...) - đưa phần mềm về

        đúng trạng thái như khi vừa khởi động, sẵn sàng nạp dữ liệu mới."""

        if self.df is None and self.raw_df is None and not self.file_paths:

            show_notice(self, "info", "Xóa dữ liệu", "Hiện chưa có dữ liệu nào để xóa.")

            return

        reply = QMessageBox.question(

            self, "Xóa toàn bộ dữ liệu",

            "Bạn có chắc chắn muốn xóa toàn bộ dữ liệu hiện tại?\n"

            "Phần mềm sẽ trở về trạng thái như khi mới mở (chưa nạp file nào).\n"

            "Thao tác này không thể hoàn tác.",

            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply != QMessageBox.Yes:

            return



        # ---- trạng thái dữ liệu ----

        self.file_paths = []

        self.raw_df = None

        self.raw_columns = []

        self.column_mapping = {}

        self.df = None

        self.color_map = {}

        self.color_overrides = {}

        self.crown_form_map = {}

        self.style["crown_form_map"] = {}

        self._canopy_result = None

        self._vstruct_result = None



        # ---- danh sách file & bảng dữ liệu ----

        self.file_combo.blockSignals(True)

        self.file_combo.clear()

        self.file_combo.blockSignals(False)

        self.table_model.setDataFrame(pd.DataFrame(columns=[k for k, _ in DISPLAY_COLUMNS]))

        self.legend_list.clear()



        # ---- xóa sạch cả 3 khung vẽ (ngang / dọc / 3D) ----

        for canvas in (self.canvas_plan, self.canvas_profile, self.canvas_3d):

            canvas.ax.clear()

            canvas.draw_idle()



        # ---- đưa tab "Phân tích cấu trúc rừng" về trạng thái rỗng ----

        self.update_structure_analysis()



        self.info_label.setText(

            "Chưa có dữ liệu. Hãy mở file Excel hoặc bấm \"Dùng dữ liệu Demo\" để dùng thử ngay.")

        self.statusBar().showMessage("Đã xóa toàn bộ dữ liệu. Sẵn sàng nạp dữ liệu mới.", 5000)



    def apply_mapping(self):

        if self.raw_df is None or not self.column_mapping:

            return

        mapping = self.column_mapping

        df = pd.DataFrame(index=self.raw_df.index)

        for key, _, _ in FIELD_DEFS:

            col = mapping.get(key)

            if col is not None and col in self.raw_df.columns:

                df[key] = self.raw_df[col]

            else:

                df[key] = "" if key in ("TT", "Loai") else np.nan



        for k in FIELD_KEYS_NUMERIC:

            df[k] = pd.to_numeric(df[k], errors="coerce")



        # điền giá trị mặc định hợp lý cho các trường không bắt buộc

        for k, default in [("D13", 15.0), ("DtanD", 1.5), ("DtanT", 1.5), ("DtanN", 1.5), ("DtanB", 1.5)]:

            if df[k].notna().any():

                df[k] = df[k].fillna(df[k].mean())

            else:

                df[k] = df[k].fillna(default)



        df["Loai"] = df["Loai"].astype(str).replace({"": "Không rõ", "nan": "Không rõ"})



        before = len(df)

        df = df.dropna(subset=FIELD_KEYS_CRITICAL)

        dropped = before - len(df)

        df = df.reset_index(drop=True)



        if df["TT"].astype(str).str.strip().eq("").all() or df["TT"].astype(str).eq("nan").all():

            df["TT"] = (df.index + 1).astype(str)

        else:

            df["TT"] = df["TT"].astype(str)



        self.df = df

        self.color_map = build_species_colors(df["Loai"].tolist(), self.color_overrides)

        self.table_model.setDataFrame(df)

        self._update_legend()



        n_species = df["Loai"].nunique() if not df.empty else 0

        msg = f"Đã tải {len(df)} cây thuộc {n_species} loài."

        if dropped:

            msg += f" (Bỏ qua {dropped} dòng thiếu X/Y/Hvn/Hdc.)"

        self.info_label.setText(msg)

        self.statusBar().showMessage(msg, 6000)



        if not df.empty:

            self.plotw_spin.blockSignals(True); self.ploth_spin.blockSignals(True)

            span_x = max(float(df["X"].max()) + 2, self.plotw_spin.value())

            span_y = max(float(df["Y"].max()) + 2, self.ploth_spin.value())

            if self.plotw_spin.value() < float(df["X"].max()):

                self.plotw_spin.setValue(round(span_x, 1))

            if self.ploth_spin.value() < float(df["Y"].max()):

                self.ploth_spin.setValue(round(span_y, 1))

            self.plotw_spin.blockSignals(False); self.ploth_spin.blockSignals(False)

            self._update_area_label()



        self.refresh_views()



    def _update_legend(self):

        self.legend_list.clear()

        if self.df is None or self.df.empty:

            return

        counts = self.df["Loai"].value_counts()

        for species, color in self.color_map.items():

            pix = QPixmap(14, 14)

            pix.fill(QColor(color))

            item = QListWidgetItem(QIcon(pix), f"  {species}   ({int(counts.get(species, 0))} cây)")

            self.legend_list.addItem(item)



    # ==========================================================================
    # LƯU / TẢI CẤU HÌNH (.geojson) - dữ liệu cây + toàn bộ cài đặt trắc đồ
    # ==========================================================================
    # Mục tiêu: sau khi đã nhập dữ liệu từ Excel và chỉnh sửa/tuỳ biến trắc đồ
    # (ánh xạ cột, màu sắc, kích thước ô, nhãn, la bàn, thước tỷ lệ, tuỳ biến
    # từng biểu đồ cấu trúc rừng...), người dùng có thể LƯU toàn bộ trạng thái
    # này ra 1 file .geojson duy nhất. Lần sau chỉ cần TẢI lại đúng file đó là
    # khôi phục lại y hệt - kể cả khi file Excel gốc đã bị xóa/di chuyển/sửa,
    # vì toàn bộ dữ liệu từng cây đã được nhúng sẵn trong file .geojson (mỗi
    # cây là 1 Feature dạng Point tại toạ độ (X, Y), các trường còn lại nằm
    # trong "properties"). Cài đặt/style không thuộc chuẩn GeoJSON được lưu
    # kèm dưới dạng "foreign member" ở khoá "tdforest" cấp cao nhất của
    # FeatureCollection - vẫn là file .geojson hợp lệ, các phần mềm GIS khác
    # mở lên vẫn đọc được vị trí + thuộc tính từng cây bình thường.
    # LƯU Ý: toạ độ (X, Y) ở đây là toạ độ cục bộ trong ô tiêu chuẩn (đơn vị
    # mét), KHÔNG phải kinh độ/vĩ độ địa lý - phù hợp với cách đo đạc thực
    # địa hiện tại của phần mềm (xem ĐỊNH DẠNG DỮ LIỆU ĐẦU VÀO ở đầu file).

    def _build_config_geojson(self):
        """Đóng gói self.df (từng cây -> 1 Feature GeoJSON) cùng self.style,
        color_overrides, column_mapping vào 1 FeatureCollection - trả về
        dict sẵn sàng json.dump ra file."""
        features = []
        if self.df is not None and not self.df.empty:
            for _, row in self.df.iterrows():
                props = {}
                for key, _, _ in FIELD_DEFS:
                    v = row.get(key)
                    if key in FIELD_KEYS_NUMERIC:
                        props[key] = None if pd.isna(v) else float(v)
                    else:
                        props[key] = "" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v)
                x = props.get("X")
                y = props.get("Y")
                geometry = {
                    "type": "Point",
                    "coordinates": [x if x is not None else 0.0, y if y is not None else 0.0],
                }
                features.append({"type": "Feature", "geometry": geometry, "properties": props})

        payload = {
            "type": "FeatureCollection",
            "features": features,
            # --- "foreign member" (RFC 7946 cho phép thêm khoá tuỳ ý ở cấp
            # FeatureCollection) - chứa toàn bộ cài đặt/tuỳ biến riêng của
            # TDForest, không ảnh hưởng tới việc phần mềm GIS khác đọc file. ---
            "tdforest": {
                "app": "TDForest",
                "config_version": TDFOREST_CONFIG_VERSION,
                "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "style": self.style,
                "color_overrides": self.color_overrides,
                "column_mapping": self.column_mapping,
                "source_files": [os.path.basename(p) for p in self.file_paths if p and p != DEMO_MARKER],
            },
        }
        return payload

    def save_config_geojson(self):
        """Lưu dữ liệu cây hiện tại (self.df) + toàn bộ cài đặt trắc đồ
        (self.style, màu loài tự chọn, ánh xạ cột) ra 1 file .geojson."""
        if self.df is None or self.df.empty:
            show_notice(
                self, "warning", "Lưu cấu hình",
                "Chưa có dữ liệu cây nào để lưu.\n"
                "Hãy mở file Excel hoặc bấm \"Dùng dữ liệu Demo\" trước khi lưu cấu hình.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Lưu cấu hình trắc đồ (.geojson)", "tdforest_cauhinh.geojson",
            "GeoJSON (*.geojson);;JSON (*.json)")
        if not path:
            return
        if not path.lower().endswith((".geojson", ".json")):
            path += ".geojson"

        try:
            payload = self._build_config_geojson()
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2, default=_json_safe_default)
        except Exception as e:
            show_notice(self, "error", "Lỗi lưu cấu hình", f"Không thể lưu file:\n{path}\n\n{e}")
            return

        show_notice(
            self, "success", "Lưu cấu hình",
            f"Đã lưu {len(self.df)} cây cùng toàn bộ cài đặt trắc đồ hiện tại "
            "(màu sắc, kích thước ô, nhãn, la bàn, thước tỷ lệ, tuỳ biến từng biểu đồ...) "
            f"vào file:\n{path}\n\n"
            "Lần sau chỉ cần bấm \"Tải cấu hình\" và chọn lại file này để khôi phục "
            "y hệt như hiện tại - kể cả khi file Excel gốc đã bị xóa hoặc thay đổi.")
        self.statusBar().showMessage(f"Đã lưu cấu hình: {path}", 6000)

    def load_config_geojson(self):
        """Tải lại dữ liệu cây + toàn bộ cài đặt trắc đồ đã lưu trước đó
        bằng save_config_geojson() - KHÔNG phụ thuộc file Excel gốc."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Tải cấu hình trắc đồ (.geojson)", "", "GeoJSON (*.geojson *.json)")
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            show_notice(self, "error", "Lỗi tải cấu hình", f"Không thể đọc file:\n{path}\n\n{e}")
            return

        if not isinstance(data, dict) or data.get("type") != "FeatureCollection" or "features" not in data:
            show_notice(
                self, "error", "Lỗi tải cấu hình",
                "File không đúng định dạng cấu hình TDForest (.geojson).\n"
                "Hãy chọn đúng file đã lưu bằng chức năng \"Lưu cấu hình\".")
            return

        # ---- dựng lại DataFrame từ các Feature ----
        rows = []
        for feat in (data.get("features") or []):
            props = dict(feat.get("properties") or {})
            geom = feat.get("geometry") or {}
            coords = geom.get("coordinates") or []
            row = {key: props.get(key) for key, _, _ in FIELD_DEFS}
            if row.get("X") in (None, "") and len(coords) >= 1:
                row["X"] = coords[0]
            if row.get("Y") in (None, "") and len(coords) >= 2:
                row["Y"] = coords[1]
            rows.append(row)

        if not rows:
            show_notice(self, "warning", "Tải cấu hình", "File cấu hình không có dữ liệu cây nào.")
            return

        df = pd.DataFrame(rows, columns=[k for k, _, _ in FIELD_DEFS])
        for k in FIELD_KEYS_NUMERIC:
            df[k] = pd.to_numeric(df[k], errors="coerce")

        before = len(df)
        df = df.dropna(subset=FIELD_KEYS_CRITICAL).reset_index(drop=True)
        dropped = before - len(df)

        df["Loai"] = df["Loai"].astype(str).replace({"": "Không rõ", "None": "Không rõ", "nan": "Không rõ"})
        if df["TT"].astype(str).str.strip().eq("").all() or df["TT"].astype(str).isin(["None", "nan"]).all():
            df["TT"] = (df.index + 1).astype(str)
        else:
            df["TT"] = df["TT"].astype(str)

        # ---- khôi phục self.style (gộp với mặc định hiện tại để không mất
        # các key mới hơn phiên bản file cũ), color_overrides, column_mapping ----
        meta = data.get("tdforest") or {}
        loaded_style = meta.get("style")
        if isinstance(loaded_style, dict):
            merged_style = copy.deepcopy(self.style)
            merged_style.update(loaded_style)
            if isinstance(loaded_style.get("struct_chart_style"), dict):
                scs = copy.deepcopy(self.style.get("struct_chart_style", {}))
                for k, v in loaded_style["struct_chart_style"].items():
                    scs.setdefault(k, {}).update(v or {})
                merged_style["struct_chart_style"] = scs
            # h_range lưu ra JSON thành list - đưa lại về tuple như style gốc
            if isinstance(merged_style.get("h_range"), list):
                merged_style["h_range"] = tuple(merged_style["h_range"])
            self.style = merged_style

        self.crown_form_map = dict(self.style.get("crown_form_map") or {})
        self.color_overrides = dict(meta.get("color_overrides") or {})
        if isinstance(meta.get("column_mapping"), dict):
            self.column_mapping = dict(meta["column_mapping"])

        # ---- dữ liệu từ file cấu hình là 1 "nguồn" riêng trong combo, KHÔNG
        # phụ thuộc file Excel gốc (có thể đã bị xóa/di chuyển) ----
        self.raw_df = None
        self.raw_columns = []
        self.file_paths = []
        self.file_combo.blockSignals(True)
        self.file_combo.clear()
        self.file_combo.addItem(f"⚙ Cấu hình đã tải: {os.path.basename(path)}", CONFIG_MARKER)
        self.file_combo.setCurrentIndex(0)
        self.file_combo.blockSignals(False)

        self.df = df
        self.color_map = build_species_colors(df["Loai"].tolist(), self.color_overrides)
        self.table_model.setDataFrame(df)
        self._update_legend()

        # đồng bộ lại toàn bộ control bên trái theo self.style vừa khôi phục,
        # tránh trường hợp người dùng chạm 1 control bất kỳ sau đó vô tình
        # ghi đè giá trị cũ (trên widget) lên cấu hình vừa tải.
        self._apply_style_to_ui()

        self._canopy_result = None
        self._vstruct_result = None

        n_species = df["Loai"].nunique() if not df.empty else 0
        msg = f"Đã tải cấu hình: {len(df)} cây thuộc {n_species} loài."
        if dropped:
            msg += f" (Bỏ qua {dropped} dòng thiếu X/Y/Hvn/Hdc.)"
        self.info_label.setText(msg)
        self.statusBar().showMessage(msg, 6000)

        self.refresh_views()

        show_notice(
            self, "success", "Tải cấu hình",
            f"Đã khôi phục {len(df)} cây thuộc {n_species} loài cùng toàn bộ cài đặt trắc đồ "
            f"từ file:\n{path}")

    def _apply_style_to_ui(self):
        """Đồng bộ lại TOÀN BỘ control ở panel bên trái (mục ②③⑤⑥ và tab
        "Phân tích cấu trúc rừng") theo đúng self.style hiện tại - gọi sau
        khi TẢI CẤU HÌNH (.geojson) để giao diện khớp với dữ liệu/cài đặt vừa
        khôi phục. Chặn tín hiệu (blockSignals) trong lúc set giá trị để
        không kích hoạt hàng loạt on_style_changed()/refresh_views() thừa và
        để không ghi đè lại self.style vừa tải bằng giá trị widget cũ."""
        s = self.style
        widgets = [
            self.plotw_spin, self.ploth_spin, self.scale_combo, self.scale_custom_spin,
            self.showplot_chk, self.plan_shape_combo, self.crownform_combo, self.colormode_combo,
            self.alpha_slider, self.crownedge_chk, self.crownedge_width_spin,
            self.trunk_spin, self.radius_chk, self.labels_chk, self.labelmode_combo,
            self.grid_chk, self.specieslegend_chk, self.legendfontsize_spin,
            self.northarrow_chk, self.northarrowsize_spin, self.scalebar_chk,
            self.scalebarsize_spin, self.scalenote_chk, self.scalenotesize_spin,
            self.scalenotefont_combo, self.title_plan_edit, self.title_profile_edit,
            self.title_3d_edit, self.xlabel_plan_edit, self.ylabel_plan_edit,
            self.ylabel_profile_edit, self.font_combo, self.titlesize_spin,
            self.labelsize_spin, self.axis_combo,
            self.struct_binwidth_spin, self.struct_gridres_spin,
            self.vstruct_binwidth_spin, self.vstruct_kstd_spin,
        ]
        for w in widgets:
            w.blockSignals(True)
        try:
            # ② ô tiêu chuẩn
            field_w = s.get("field_w") or 20.0
            field_h = s.get("field_h") or 25.0
            self.plotw_spin.setValue(field_w)
            self.ploth_spin.setValue(field_h)
            self.showplot_chk.setChecked(s.get("plot_w") is not None)

            denom = s.get("scale_denom")
            if denom is None:
                self.scale_combo.setCurrentIndex(0)
                self.scale_custom_spin.setVisible(False)
            else:
                preset = f"1:{int(denom)}"
                idx = SCALE_OPTIONS.index(preset) if preset in SCALE_OPTIONS else -1
                if idx >= 0:
                    self.scale_combo.setCurrentIndex(idx)
                    self.scale_custom_spin.setVisible(False)
                else:
                    self.scale_combo.setCurrentIndex(len(SCALE_OPTIONS) - 1)  # "Tuỳ chỉnh..."
                    self.scale_custom_spin.setValue(denom)
                    self.scale_custom_spin.setVisible(True)

            # ③ kiểu hiển thị cây
            if s.get("plan_shape") in PLAN_SHAPE_OPTIONS:
                self.plan_shape_combo.setCurrentText(s["plan_shape"])
            if s.get("crown_form") in CROWN_FORM_OPTIONS:
                self.crownform_combo.setCurrentText(s["crown_form"])
            if s.get("color_mode") in COLOR_MODES:
                self.colormode_combo.setCurrentText(s["color_mode"])
            self._set_swatch(self.uniform_btn, s.get("uniform_color") or "#43a047")
            self.uniform_btn.setEnabled(str(s.get("color_mode") or "").startswith("Một"))
            self._set_swatch(self.trunk_btn, s.get("trunk_color") or "#795548")
            self.alpha_slider.setValue(int(round((s.get("alpha") or 0.85) * 100)))

            has_crown_edge = bool(s.get("crown_edge_color"))
            self.crownedge_chk.setChecked(has_crown_edge)
            self.crownedge_btn.setEnabled(has_crown_edge)
            if has_crown_edge:
                self._crown_edge_last_color = s["crown_edge_color"]
            self._set_swatch(self.crownedge_btn, s.get("crown_edge_color") or self._crown_edge_last_color)
            self.crownedge_width_spin.setValue(s.get("crown_edge_width") or 1.2)

            self.trunk_spin.setValue(s.get("trunk_exagg") or 1.0)
            self.radius_chk.setChecked(bool(s.get("radius_mode", True)))
            self.labels_chk.setChecked(bool(s.get("show_labels", False)))
            self.labelmode_combo.setCurrentIndex(0 if s.get("label_mode", "TT") == "TT" else 1)
            self._set_swatch(self.labelcolor_btn, s.get("label_color") or "#2b3a2f")
            self.grid_chk.setChecked(bool(s.get("show_grid", True)))
            self.specieslegend_chk.setChecked(bool(s.get("show_species_legend", True)))
            self.legendfontsize_spin.setValue(s.get("legend_font_size") or 7.6)
            self._set_swatch(
                self.legendfontcolor_btn,
                s.get("legend_font_color") or s.get("font_color") or "#1b5e20")

            # ⑤ nhãn & phông chữ
            self.title_plan_edit.setText(s.get("title_plan") or "")
            self.title_profile_edit.setText(s.get("title_profile") or "")
            self.title_3d_edit.setText(s.get("title_3d") or "")
            self.xlabel_plan_edit.setText(s.get("xlabel_plan") or "")
            self.ylabel_plan_edit.setText(s.get("ylabel_plan") or "")
            self.ylabel_profile_edit.setText(s.get("ylabel_profile") or "")
            self._set_swatch(self.fontcolor_btn, s.get("font_color") or "#1b5e20")
            if s.get("font_family") in FONT_FAMILY_OPTIONS:
                self.font_combo.setCurrentText(s["font_family"])
            self.titlesize_spin.setValue(s.get("title_size") or 11.5)
            self.labelsize_spin.setValue(s.get("label_size") or 10.0)

            # ⑥ la bàn / thước tỷ lệ / chú thích tỷ lệ
            self.northarrow_chk.setChecked(bool(s.get("show_north_arrow", True)))
            self.northarrowsize_spin.setValue(s.get("north_arrow_size") or 1.0)
            self._set_swatch(self.northarrowcolor_btn, s.get("north_arrow_color") or "#111111")
            self.scalebar_chk.setChecked(bool(s.get("show_scale_bar", True)))
            self.scalebarsize_spin.setValue(s.get("scale_bar_size") or 1.0)
            self._set_swatch(
                self.scalebarcolor_btn, s.get("scale_bar_color") or s.get("font_color") or "#1b5e20")
            self.scalenote_chk.setChecked(bool(s.get("show_scale_note", True)))
            self.scalenotesize_spin.setValue(s.get("scale_note_size") or 1.0)
            self._set_swatch(
                self.scalenotecolor_btn, s.get("scale_note_color") or s.get("font_color") or "#1b5e20")
            note_font = s.get("scale_note_font_family")
            if note_font and note_font in FONT_FAMILY_OPTIONS:
                self.scalenotefont_combo.setCurrentIndex(FONT_FAMILY_OPTIONS.index(note_font) + 1)
            else:
                self.scalenotefont_combo.setCurrentIndex(0)

            # trục trắc đồ dọc (X hoặc Y)
            self.axis_combo.setCurrentIndex(0 if s.get("profile_axis", "X") == "X" else 1)

            # tham số tab "Phân tích cấu trúc rừng"
            self.struct_binwidth_spin.setValue(s.get("struct_bin_width") or 4.0)
            self.struct_gridres_spin.setValue(s.get("struct_grid_res") or 140)
            self.vstruct_binwidth_spin.setValue(s.get("vstruct_bin_width") or 2.0)
            self.vstruct_kstd_spin.setValue(s.get("vstruct_k_std") or 1.0)
        finally:
            for w in widgets:
                w.blockSignals(False)

        self._update_area_label()
        if hasattr(self, "on_struct_chart_selected"):
            self.on_struct_chart_selected()


    # ==========================================================================

    # STYLE & VẼ LẠI

    # ==========================================================================



    def on_style_changed(self):

        # Nếu người dùng đổi kích thước ô tiêu chuẩn, kết quả "Độ che phủ &

        # chồng lấn tán" đã tính trước đó (nếu có) không còn đúng với kích

        # thước mới -> hủy để tránh hiển thị số liệu sai, người dùng bấm lại

        # nút "Tính độ che phủ & chồng lấn tán" khi cần xem lại.

        if self._canopy_result is not None and (

                self.style.get("field_w") != self.plotw_spin.value()

                or self.style.get("field_h") != self.ploth_spin.value()):

            self._canopy_result = None

        self.style["plan_shape"] = self.plan_shape_combo.currentText()

        self.style["crown_form"] = self.crownform_combo.currentText()

        self.style["color_mode"] = self.colormode_combo.currentText()

        self.style["alpha"] = self.alpha_slider.value() / 100.0

        self.style["radius_mode"] = self.radius_chk.isChecked()

        self.style["show_labels"] = self.labels_chk.isChecked()

        self.style["label_mode"] = "TT" if self.labelmode_combo.currentIndex() == 0 else "Loai"

        self.style["show_grid"] = self.grid_chk.isChecked()

        self.style["show_species_legend"] = self.specieslegend_chk.isChecked()

        self.style["legend_font_size"] = self.legendfontsize_spin.value()

        self.style["show_north_arrow"] = self.northarrow_chk.isChecked()

        self.style["show_scale_bar"] = self.scalebar_chk.isChecked()

        self.style["north_arrow_size"] = self.northarrowsize_spin.value()

        self.style["scale_bar_size"] = self.scalebarsize_spin.value()

        self.style["show_scale_note"] = self.scalenote_chk.isChecked()

        self.style["scale_note_size"] = self.scalenotesize_spin.value()

        note_font_idx = self.scalenotefont_combo.currentIndex()

        self.style["scale_note_font_family"] = (

            FONT_FAMILY_OPTIONS[note_font_idx - 1] if note_font_idx > 0 else None

        )

        self.style["crown_form_map"] = dict(self.crown_form_map)

        self.style["trunk_exagg"] = self.trunk_spin.value()

        self.style["crown_edge_width"] = self.crownedge_width_spin.value()

        # Kích thước thực (dùng để TÍNH TỶ LỆ 1:N) luôn lấy theo ô nhập ②,

        # dù có tick hiển thị khung viền hay không.

        self.style["field_w"] = self.plotw_spin.value()

        self.style["field_h"] = self.ploth_spin.value()

        if self.showplot_chk.isChecked():

            self.style["plot_w"] = self.plotw_spin.value()

            self.style["plot_h"] = self.ploth_spin.value()

        else:

            self.style["plot_w"] = None

            self.style["plot_h"] = None

        self.uniform_btn.setEnabled(self.style["color_mode"].startswith("Một"))

        self.refresh_views()



    # ---- nhóm: nhãn & phông chữ biểu đồ ----

    def _build_labels_group(self):

        box = QGroupBox("⑤ Nhãn & phông chữ biểu đồ")

        form = QFormLayout(box)



        self.title_plan_edit = QLineEdit(); self.title_plan_edit.setPlaceholderText("(mặc định) TRẮC ĐỒ NGANG…")

        self.title_plan_edit.textChanged.connect(self.on_labels_changed)

        form.addRow("Tiêu đề trắc đồ ngang:", self.title_plan_edit)



        self.title_profile_edit = QLineEdit(); self.title_profile_edit.setPlaceholderText("(mặc định) TRẮC ĐỒ DỌC…")

        self.title_profile_edit.textChanged.connect(self.on_labels_changed)

        form.addRow("Tiêu đề trắc đồ dọc:", self.title_profile_edit)



        self.title_3d_edit = QLineEdit(); self.title_3d_edit.setPlaceholderText("(mặc định) MÔ PHỎNG 3D…")

        self.title_3d_edit.textChanged.connect(self.on_labels_changed)

        form.addRow("Tiêu đề mô phỏng 3D:", self.title_3d_edit)



        form.addRow(self._hline())

        self.xlabel_plan_edit = QLineEdit(); self.xlabel_plan_edit.setPlaceholderText("X (m)")

        self.xlabel_plan_edit.textChanged.connect(self.on_labels_changed)

        form.addRow("Nhãn trục X (ngang):", self.xlabel_plan_edit)



        self.ylabel_plan_edit = QLineEdit(); self.ylabel_plan_edit.setPlaceholderText("Y (m)")

        self.ylabel_plan_edit.textChanged.connect(self.on_labels_changed)

        form.addRow("Nhãn trục Y (ngang):", self.ylabel_plan_edit)



        self.ylabel_profile_edit = QLineEdit(); self.ylabel_profile_edit.setPlaceholderText("Chiều cao (m)")

        self.ylabel_profile_edit.textChanged.connect(self.on_labels_changed)

        form.addRow("Nhãn trục Y (dọc/độ cao):", self.ylabel_profile_edit)



        form.addRow(self._hline())

        self.fontcolor_btn = QPushButton("")

        self.fontcolor_btn.setFixedHeight(26)

        self._set_swatch(self.fontcolor_btn, self.style["font_color"])

        self.fontcolor_btn.clicked.connect(self.choose_font_color)

        form.addRow("Màu chữ:", self.fontcolor_btn)



        self.font_combo = QComboBox(); self.font_combo.addItems(FONT_FAMILY_OPTIONS)

        self.font_combo.currentIndexChanged.connect(self.on_labels_changed)

        form.addRow("Kiểu chữ (font):", self.font_combo)



        self.titlesize_spin = QDoubleSpinBox()

        self.titlesize_spin.setRange(7, 24); self.titlesize_spin.setValue(self.style["title_size"])

        self.titlesize_spin.valueChanged.connect(self.on_labels_changed)

        form.addRow("Cỡ chữ tiêu đề:", self.titlesize_spin)



        self.labelsize_spin = QDoubleSpinBox()

        self.labelsize_spin.setRange(6, 20); self.labelsize_spin.setValue(self.style["label_size"])

        self.labelsize_spin.valueChanged.connect(self.on_labels_changed)

        form.addRow("Cỡ chữ trục:", self.labelsize_spin)



        return box



    def on_labels_changed(self):

        self.style["title_plan"] = self.title_plan_edit.text().strip()

        self.style["title_profile"] = self.title_profile_edit.text().strip()

        self.style["title_3d"] = self.title_3d_edit.text().strip()

        self.style["xlabel_plan"] = self.xlabel_plan_edit.text().strip()

        self.style["ylabel_plan"] = self.ylabel_plan_edit.text().strip()

        self.style["ylabel_profile"] = self.ylabel_profile_edit.text().strip()

        self.style["font_family"] = self.font_combo.currentText()

        self.style["title_size"] = self.titlesize_spin.value()

        self.style["label_size"] = self.labelsize_spin.value()

        self.refresh_views()



    def choose_font_color(self):

        c = QColorDialog.getColor(QColor(self.style["font_color"]), self, "Chọn màu chữ")

        if c.isValid():

            self.style["font_color"] = c.name()

            self._set_swatch(self.fontcolor_btn, c.name())

            self.refresh_views()



    def _real_size_for_scale(self, kind):

        """Trả về (chiều rộng thực, chiều cao thực) tính bằng mét, dùng để

        quy đổi tỷ lệ 1:N - LUÔN dựa vào ô nhập ② (không phụ thuộc việc có

        tick hiển thị khung viền hay không). Logic dùng chung với

        _draw_scale_note (xem hàm module-level _real_size_for_scale)."""

        return _real_size_for_scale(self.style, kind)



    def _sync_canvas_scale(self, canvas, scroll_area, kind):

        """Khi đã chọn 1 tỷ lệ cụ thể (không phải 'Tự động'), đặt canvas về

        đúng kích thước vật lý theo tỷ lệ 1:N (bọc trong QScrollArea để cuộn

        xem nếu lớn hơn khung nhìn) - nhờ vậy đổi tỷ lệ sẽ thấy trắc đồ to/nhỏ

        lại rõ ràng ngay trên màn hình, không chỉ ảnh hưởng lúc xuất ảnh/in."""

        denom = self.style.get("scale_denom")

        real_w, real_h = self._real_size_for_scale(kind)

        if denom and real_w and real_h:

            dpi = canvas.fig.get_dpi()

            w_in, h_in = figsize_for_scale(real_w, real_h, denom)

            w_px = max(120, int(round(w_in * dpi)))

            h_px = max(120, int(round(h_in * dpi)))

            scroll_area.setWidgetResizable(False)

            canvas.setMinimumSize(w_px, h_px)

            canvas.setMaximumSize(w_px, h_px)

            canvas.resize(w_px, h_px)

        else:

            scroll_area.setWidgetResizable(True)

            canvas.setMinimumSize(0, 0)

            canvas.setMaximumSize(16777215, 16777215)



    def refresh_views(self):

        if self.df is None or self.df.empty:

            self.update_structure_analysis()

            return

        if self.df["Hvn"].notna().any():

            self.style["h_range"] = (float(self.df["Hvn"].min()), float(self.df["Hvn"].max()))



        draw_plan_view(self.canvas_plan.ax, self.df, self.color_map, self.style)

        self._sync_canvas_scale(self.canvas_plan, self.plan_scroll, "plan")

        self.canvas_plan.draw_idle()



        draw_profile_view(self.canvas_profile.ax, self.df, self.color_map, self.style)

        self._sync_canvas_scale(self.canvas_profile, self.profile_scroll, "profile")

        self.canvas_profile.draw_idle()



        draw_3d_view(self.canvas_3d.ax, self.df, self.color_map, self.style)

        self.canvas_3d.ax.view_init(elev=self.elev_slider.value(), azim=self.azim_slider.value())

        self.canvas_3d.draw_idle()



        # Tab "Phân tích cấu trúc rừng" - cập nhật đồng bộ mỗi khi dữ liệu

        # hoặc các tham số vẽ (kích thước ô, dạng tán...) thay đổi, vì các

        # chỉ tiêu (mật độ, G/ha, diện tích tán...) và phân bố N/D, N/Dg, G

        # đều phụ thuộc vào dữ liệu/tham số này. Độ che phủ tán KHÔNG tính

        # lại tự động ở đây (xem compute_canopy_and_draw) vì tốn nhiều thời

        # gian hơn (rasterize) - nhưng biểu đồ của nó vẫn được vẽ lại với

        # kết quả gần nhất đã tính (self._canopy_result), nếu có.

        self.update_structure_analysis()



    # ==========================================================================

    # XUẤT ẢNH

    # ==========================================================================



    def export_current_view(self):

        idx = self.tabs.currentIndex()

        canvas_map = {1: self.canvas_plan, 2: self.canvas_profile, 3: self.canvas_3d}

        canvas = canvas_map.get(idx)

        if canvas is None and idx == 4 and hasattr(self, "struct_canvas_list"):

            # Tab "Phân tích cấu trúc rừng" có 4 biểu đồ con (N/D, N/Dg, G,

            # Độ che phủ tán) - xuất đúng biểu đồ đang mở trong struct_tabs.

            sub_idx = self.struct_tabs.currentIndex()

            if 0 <= sub_idx < len(self.struct_canvas_list):

                canvas = self.struct_canvas_list[sub_idx]

        if canvas is None:

            show_notice(self, "info", "Xuất ảnh",

                        "Vui lòng mở tab Trắc đồ ngang / dọc / 3D / Phân tích cấu trúc rừng để xuất ảnh.")

            return

        path, _ = QFileDialog.getSaveFileName(

            self, "Lưu ảnh trắc đồ", "tracdo.png", "PNG (*.png);;PDF (*.pdf);;SVG (*.svg)")

        if not path:

            return



        # Nếu đã chọn tỷ lệ cụ thể (không phải "Tự động") và đang xuất trắc đồ

        # ngang/dọc 2D: đặt đúng kích thước hình theo tỷ lệ trước khi lưu, để

        # in ở chế độ "Actual size / 100%" cho ra đúng tỷ lệ 1:N mong muốn.

        # KHÔNG dùng bbox_inches="tight" trong trường hợp này vì nó sẽ cắt

        # theo nội dung và làm sai kích thước vật lý đã tính.

        denom = self.style.get("scale_denom")

        applied_scale = False

        old_size = canvas.fig.get_size_inches()

        kind = {1: "plan", 2: "profile"}.get(idx)

        if denom and kind:

            real_w, real_h = self._real_size_for_scale(kind)

            if real_w and real_h:

                w_in, h_in = figsize_for_scale(real_w, real_h, denom)

                canvas.fig.set_size_inches(w_in, h_in)

                applied_scale = True

                # QUAN TRỌNG: đổi kích thước Figure làm thay đổi tỷ lệ khung
                # hình (aspect) thật của Axes so với lúc xem trên màn hình Qt
                # (vốn có tỷ lệ khung cửa sổ tuỳ ý). Với Trắc đồ ngang, Axes
                # dùng set_aspect("equal", adjustable="datalim") nên tỷ lệ
                # khung mới này khiến matplotlib tính lại giới hạn hiển thị
                # (view limits) - PHẢI vẽ lại từ đầu (gọi lại draw_plan_view/
                # draw_profile_view) ở đúng kích thước Figure cuối cùng để
                # chú dẫn (neo theo ax.dataLim, xem _draw_species_legend_box)
                # và khoảng đệm ax.margins() được tính đúng theo khung hình
                # SẼ LƯU RA FILE, thay vì giữ nguyên bố cục đã tính cho khung
                # hình xem trước rồi chỉ phóng/co lại - nếu không, ảnh xuất
                # ra có thể lệch bố cục/chú dẫn so với những gì thấy trên màn
                # hình ngay trước khi bấm "Xuất ảnh".
                if kind == "plan":
                    draw_plan_view(canvas.ax, self.df, self.color_map, self.style)
                else:
                    draw_profile_view(canvas.ax, self.df, self.color_map, self.style)



        try:

            if applied_scale:

                canvas.fig.savefig(path, dpi=220, facecolor=canvas.fig.get_facecolor())

            else:

                canvas.fig.savefig(path, dpi=220, facecolor=canvas.fig.get_facecolor(), bbox_inches="tight")

        finally:

            if applied_scale:

                canvas.fig.set_size_inches(*old_size)

                if kind == "plan":
                    draw_plan_view(canvas.ax, self.df, self.color_map, self.style)
                else:
                    draw_profile_view(canvas.ax, self.df, self.color_map, self.style)

                canvas.draw_idle()



        msg = f"Đã lưu ảnh: {path}"

        if applied_scale:

            msg += f" - đúng tỷ lệ 1:{int(denom)} khi in ở chế độ 100% (Actual size)"

        self.statusBar().showMessage(msg, 6000)

    # ==========================================================================
    # XUẤT THEO KHỔ GIẤY (A2 / A3 / A4 / A5) - ĐÚNG TỶ LỆ, ĐÚNG HƯỚNG GIẤY
    # ==========================================================================

    def export_paper_size(self):
        """Mở hộp thoại chọn nội dung / khổ giấy / hướng giấy / tỷ lệ rồi xuất
        trắc đồ ngang hoặc dọc ra file PDF/PNG đúng khổ giấy tiêu chuẩn
        (A2/A3/A4/A5), đúng hướng (đứng/ngang) và đúng tỷ lệ 1:N - phục vụ in
        ấn hồ sơ điều tra rừng theo quy cách. Khác với "Xuất ảnh" (chỉ lưu
        nguyên khung vẽ hiện có), chức năng này VẼ LẠI trắc đồ trên khổ giấy
        chuẩn, có khung viền + khung tiêu đề bản vẽ, canh giữa trang."""
        if self.df is None or self.df.empty:
            show_notice(self, "info", "Xuất khổ giấy",
                        "Vui lòng mở dữ liệu (file Excel) trước khi xuất trắc đồ ra khổ giấy.")
            return

        kind_options = [
            ("plan", "Trắc đồ ngang (mặt bằng phân bố cây)"),
            ("profile", "Trắc đồ dọc (mặt cắt phân bố cây)"),
        ]
        idx = self.tabs.currentIndex()
        default_kind = {1: "plan", 2: "profile"}.get(idx, "plan")
        default_denom = self.style.get("scale_denom")

        dlg = ExportPaperDialog(kind_options, default_kind, default_denom, parent=self)
        if dlg.exec() != QDialog.Accepted:
            return
        vals = dlg.get_values()
        kind = vals["kind"]

        real_w, real_h = self._real_size_for_scale(kind)
        if not real_w or not real_h:
            show_notice(self, "error", "Xuất khổ giấy",
                        "Không xác định được kích thước thực của ô tiêu chuẩn. "
                        "Vui lòng kiểm tra lại thông tin ở khung ② Kích thước ô.")
            return

        paper_w_mm, paper_h_mm = PAPER_SIZES_MM[vals["paper"]]
        if vals["orientation"] == "landscape":
            paper_w_mm, paper_h_mm = paper_h_mm, paper_w_mm
        # Lề dành cho khung tiêu đề bản vẽ + lề ngoài (xem _render_paper_export).
        margin_cm = 2.2
        paper_w_cm = paper_w_mm / 10.0
        paper_h_cm = paper_h_mm / 10.0
        # Trừ thêm lề dành RIÊNG cho số + tên trục X/Y (xem giải thích ở khai
        # báo AXIS_LABEL_MARGIN_*_CM) - nếu không trừ, ở chế độ "vừa khổ giấy"
        # khung nội dung sẽ tính khít đúng bằng vùng in khả dụng và đẩy nhãn
        # trục ra ngoài mép giấy khi in.
        avail_w_cm = paper_w_cm - margin_cm - AXIS_LABEL_MARGIN_LEFT_CM - AXIS_LABEL_MARGIN_RIGHT_CM
        avail_h_cm = paper_h_cm - margin_cm - AXIS_LABEL_MARGIN_BOTTOM_CM - AXIS_LABEL_MARGIN_TOP_CM

        denom_input = vals["denom"]
        is_fit = (denom_input == "fit")

        if is_fit:
            # Phóng/thu vừa khổ giấy: KHÔNG theo tỷ lệ chuẩn - tính mẫu số
            # (denom) chính xác (không làm tròn) sao cho bản vẽ lấp đầy vùng
            # in khả dụng, giữ nguyên tỷ lệ khung hình (không méo hình).
            denom = max(real_w * 100.0 / avail_w_cm, real_h * 100.0 / avail_h_cm)
        elif denom_input:
            denom = denom_input
            content_w_cm = real_w * 100.0 / denom
            content_h_cm = real_h * 100.0 / denom
            if content_w_cm > avail_w_cm or content_h_cm > avail_h_cm:
                reply = QMessageBox.question(
                    self, "Nội dung vượt khổ giấy",
                    f"Với tỷ lệ 1:{int(denom)}, bản vẽ có kích thước khoảng "
                    f"{content_w_cm:.1f} × {content_h_cm:.1f} cm - LỚN HƠN vùng in "
                    f"khả dụng của khổ {vals['paper']} ({avail_w_cm:.1f} × {avail_h_cm:.1f} cm).\n\n"
                    "Có thể chọn khổ giấy lớn hơn, đổi hướng giấy, hoặc chọn tỷ lệ "
                    "nhỏ hơn (mẫu số lớn hơn), hoặc dùng chế độ \"Phóng/thu vừa khổ "
                    "giấy\". Vẫn tiếp tục xuất (nội dung có thể bị cắt ở mép giấy)?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply != QMessageBox.Yes:
                    return
        else:
            # Tự động: lấy tỷ lệ chuẩn (1:50, 1:100, ...) NHỎ NHẤT (bản vẽ to
            # nhất, gần tỷ lệ 1:1 nhất) mà vẫn vừa khít vùng in khả dụng.
            denom = None
            for d in sorted(STANDARD_SCALE_DENOMS):
                if (real_w * 100.0 / d) <= avail_w_cm and (real_h * 100.0 / d) <= avail_h_cm:
                    denom = d
                    break
            if denom is None:
                # Không có tỷ lệ chuẩn nào vừa - tính tỷ lệ tối thiểu cần thiết,
                # làm tròn lên hàng chục gần nhất cho "đẹp" (dễ đọc trên giấy).
                need_w = real_w * 100.0 / avail_w_cm
                need_h = real_h * 100.0 / avail_h_cm
                denom = math.ceil(max(need_w, need_h) / 10.0) * 10.0

        kind_label = "ngang" if kind == "plan" else "doc"
        if is_fit:
            default_name = f"tracdo_{kind_label}_{vals['paper']}_vuakho.{vals['fmt']}"
        else:
            default_name = f"tracdo_{kind_label}_{vals['paper']}_1-{int(denom)}.{vals['fmt']}"
        filt = "PDF (*.pdf)" if vals["fmt"] == "pdf" else "PNG (*.png)"
        path, _ = QFileDialog.getSaveFileName(self, "Lưu file trắc đồ theo khổ giấy", default_name, filt)
        if not path:
            return

        try:
            self._render_paper_export(path, kind, vals["paper"], vals["orientation"], denom, vals["fmt"],
                                       is_fit=is_fit)
        except Exception as exc:
            show_notice(self, "error", "Xuất khổ giấy", f"Có lỗi khi xuất file:\n{exc}")
            return

        orient_vn = "ngang" if vals["orientation"] == "landscape" else "đứng"
        content_vn = "ngang" if kind == "plan" else "dọc"
        if is_fit:
            show_notice(self, "success", "Xuất khổ giấy",
                        f"Đã xuất trắc đồ {content_vn} - khổ {vals['paper']} ({orient_vn}) - "
                        f"phóng vừa khổ giấy (không theo tỷ lệ chuẩn, tỷ lệ xấp xỉ 1:{round(denom)}) "
                        f"thành công:\n{path}")
        else:
            show_notice(self, "success", "Xuất khổ giấy",
                        f"Đã xuất trắc đồ {content_vn} - khổ {vals['paper']} ({orient_vn}) - "
                        f"tỷ lệ 1:{int(denom)} thành công:\n{path}")

    def _render_paper_export(self, path, kind, paper_name, orientation, denom, fmt, is_fit=False):
        """Vẽ trắc đồ vào 1 Figure có kích thước ĐÚNG khổ giấy đã chọn (đổi mm
        sang inch), đặt khung vẽ (axes) đúng tỷ lệ vật lý 1:denom và canh giữa
        trang, có khung viền + khung tiêu đề bản vẽ, rồi lưu ra PDF/PNG. Khi in
        file PDF ở chế độ "Actual size / 100%" (KHÔNG "Fit to page"), bản in sẽ
        đúng khổ giấy và đúng tỷ lệ 1:N đã chọn."""
        paper_w_mm, paper_h_mm = PAPER_SIZES_MM[paper_name]
        if orientation == "landscape":
            paper_w_mm, paper_h_mm = paper_h_mm, paper_w_mm
        paper_w_in = paper_w_mm / 25.4
        paper_h_in = paper_h_mm / 25.4

        real_w, real_h = self._real_size_for_scale(kind)
        content_w_in = (real_w * 100.0 / denom) / 2.54
        content_h_in = (real_h * 100.0 / denom) / 2.54

        dpi = 300
        fig = Figure(figsize=(paper_w_in, paper_h_in), dpi=dpi)
        fig.patch.set_facecolor("#ffffff")

        # Khung tiêu đề bản vẽ (title block) dành ~1.6cm phía dưới trang; lề
        # ngoài mỏng (~0.5cm) các cạnh còn lại.
        title_block_in = 1.6 / 2.54
        margin_in = 0.5 / 2.54

        # Lề dành RIÊNG cho số + tên trục X/Y (đổi từ AXIS_LABEL_MARGIN_*_CM
        # sang inch) - LUÔN được chừa ra bên trái/dưới khung nội dung, không
        # để khung nội dung dán sát mép giấy, tránh bị cắt/che nhãn trục khi in
        # (đặc biệt ở chế độ "Phóng/thu vừa khổ giấy" - is_fit=True - vốn
        # trước đây tính khung nội dung khít đúng bằng vùng in khả dụng).
        axis_margin_left_in = AXIS_LABEL_MARGIN_LEFT_CM / 2.54
        axis_margin_bottom_in = AXIS_LABEL_MARGIN_BOTTOM_CM / 2.54
        axis_margin_right_in = AXIS_LABEL_MARGIN_RIGHT_CM / 2.54
        axis_margin_top_in = AXIS_LABEL_MARGIN_TOP_CM / 2.54

        avail_w_in = paper_w_in - 2 * margin_in - axis_margin_left_in - axis_margin_right_in
        avail_h_in = paper_h_in - title_block_in - margin_in - axis_margin_bottom_in - axis_margin_top_in

        x0 = max(0.01, (margin_in + axis_margin_left_in
                         + max(0.0, avail_w_in - content_w_in) / 2.0) / paper_w_in)
        y0 = max(0.01, (title_block_in + axis_margin_bottom_in
                         + max(0.0, avail_h_in - content_h_in) / 2.0) / paper_h_in)
        w_frac = min(0.98, content_w_in / paper_w_in)
        h_frac = min(0.98, content_h_in / paper_h_in)

        ax = fig.add_axes([x0, y0, w_frac, h_frac])

        style = copy.deepcopy(self.style)
        style["scale_denom"] = denom
        if kind == "plan":
            draw_plan_view(ax, self.df, self.color_map, style)
        else:
            draw_profile_view(ax, self.df, self.color_map, style)

        # Khung nét mảnh quanh toàn bộ trang (đường biên khổ giấy) - hỗ trợ
        # căn chỉnh khi in/cắt giấy thủ công.
        border = 0.15 / 2.54
        bx0 = border / paper_w_in
        by0 = border / paper_h_in
        fig.add_artist(Rectangle((bx0, by0), 1 - 2 * bx0, 1 - 2 * by0, fill=False,
                                  edgecolor="#9aa79a", linewidth=0.8,
                                  transform=fig.transFigure, zorder=0))

        # Khung tiêu đề bản vẽ ở đáy trang: tên bản vẽ, tỷ lệ, khổ giấy,
        # hướng giấy, ngày xuất - theo quy cách hồ sơ kỹ thuật lâm nghiệp.
        tb_h_frac = title_block_in / paper_h_in
        fig.add_artist(Rectangle((bx0, by0), 1 - 2 * bx0, tb_h_frac, fill=False,
                                  edgecolor="#5c7a5c", linewidth=1.0,
                                  transform=fig.transFigure, zorder=0))
        ten_ban_ve = ("TRẮC ĐỒ NGANG - MẶT BẰNG PHÂN BỐ CÂY" if kind == "plan"
                      else "TRẮC ĐỒ DỌC - MẶT CẮT PHÂN BỐ CÂY")
        fig.text(bx0 + 0.01, by0 + tb_h_frac * 0.62, ten_ban_ve,
                  fontsize=9, fontweight="bold", color="#1b5e20",
                  ha="left", va="center", transform=fig.transFigure)
        orient_label = "Ngang" if orientation == "landscape" else "Đứng"
        scale_label = (f"Phóng vừa khổ giấy (≈1:{round(denom)}, không theo tỷ lệ chuẩn)"
                       if is_fit else f"Tỷ lệ 1:{int(denom)}")
        sub = (f"{scale_label}   •   Khổ {paper_name} ({orient_label})   •   "
               f"Xuất ngày {datetime.now().strftime('%d/%m/%Y')}   •   TDForest v{APP_VERSION}")
        fig.text(bx0 + 0.01, by0 + tb_h_frac * 0.22, sub,
                  fontsize=6.5, color="#33422f", ha="left", va="center",
                  transform=fig.transFigure)

        fig.savefig(path, dpi=dpi, facecolor=fig.get_facecolor())

    # ==========================================================================
    # TỔ THÀNH LOÀI & ĐA DẠNG SINH HỌC (tab con trong "Phân tích cấu trúc rừng")
    # ==========================================================================
    def update_structure_analysis(self):
        """Cập nhật toàn bộ phân tích cấu trúc + tab Tổ thành & đa dạng loài."""
        self._update_structure_analysis_core()
        self._update_composition_tab()

    def _build_composition_subtab(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(10)
        # ---- hàng thẻ chỉ số ----
        cards = QHBoxLayout()
        cards.setSpacing(10)
        self.comp_cards = {}
        for key, label, tip in (
                ("S", "Số loài (S)", "Tổng số loài cây trong ô tiêu chuẩn"),
                ("H", "Shannon H'", "H' = -Σ pi·ln(pi) - càng lớn càng đa dạng"),
                ("J", "Pielou J'", "Độ đồng đều J' = H' / ln(S), từ 0 đến 1"),
                ("D", "Simpson (1 - D)", "Xác suất 2 cây chọn ngẫu nhiên khác loài"),
                ("d", "Margalef d", "Độ phong phú loài d = (S - 1) / ln(N)"),
                ("M", "Trữ lượng (m³/ha)", "M/ha = Σ(G · Hvn · f) / diện tích OTC")):
            card = QFrame()
            card.setObjectName("kpiCard")
            card.setToolTip(tip)
            cl = QVBoxLayout(card)
            cl.setContentsMargins(14, 10, 14, 10)
            cl.setSpacing(2)
            val = QLabel("-")
            val.setObjectName("kpiValue")
            lab = QLabel(label)
            lab.setObjectName("kpiLabel")
            cl.addWidget(val)
            cl.addWidget(lab)
            self.comp_cards[key] = val
            cards.addWidget(card, 1)
        lay.addLayout(cards)
        # ---- công thức tổ thành ----
        self.lbl_comp_formula = QLabel("Công thức tổ thành: -")
        self.lbl_comp_formula.setObjectName("formulaBox")
        self.lbl_comp_formula.setWordWrap(True)
        self.lbl_comp_formula.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lay.addWidget(self.lbl_comp_formula)
        # ---- điều khiển ----
        ctl = QHBoxLayout()
        ctl.addWidget(QLabel("Hình số thân cây f:"))
        self.ff_spin = QDoubleSpinBox()
        self.ff_spin.setRange(0.30, 0.80)
        self.ff_spin.setSingleStep(0.01)
        self.ff_spin.setDecimals(2)
        self.ff_spin.setValue(float(self.style.get("form_factor", DEFAULT_FORM_FACTOR)))
        self.ff_spin.setToolTip("Dùng tính trữ lượng V = G · Hvn · f (rừng tự nhiên thường lấy 0,45)")
        self.ff_spin.valueChanged.connect(self._on_form_factor_changed)
        ctl.addWidget(self.ff_spin)
        ctl.addStretch(1)
        btn_copy = QPushButton("📋  Copy bảng tổ thành")
        btn_copy.setObjectName("secondary")
        btn_copy.clicked.connect(self._copy_composition)
        ctl.addWidget(btn_copy)
        btn_rpt = QPushButton("📑  Xuất báo cáo")
        btn_rpt.clicked.connect(self.export_report)
        ctl.addWidget(btn_rpt)
        lay.addLayout(ctl)
        # ---- bảng + biểu đồ ----
        split = QSplitter(Qt.Horizontal)
        self.comp_table = QTableWidget(0, 8)
        self.comp_table.setHorizontalHeaderLabels(
            ["STT", "Loài cây", "N", "N%", "G (m²)", "G%", "IV%", "Hvn tb"])
        self.comp_table.verticalHeader().setVisible(False)
        self.comp_table.setAlternatingRowColors(True)
        self.comp_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.comp_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        hh = self.comp_table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        split.addWidget(self.comp_table)
        chart_w = QWidget()
        cl = QVBoxLayout(chart_w)
        cl.setContentsMargins(0, 0, 0, 0)
        self.canvas_struct_comp = MplCanvas(three_d=False)
        self.canvas_struct_comp.fig.subplots_adjust(left=0.22, right=0.97, top=0.9, bottom=0.12)
        cl.addWidget(NavigationToolbar(self.canvas_struct_comp, chart_w))
        cl.addWidget(self.canvas_struct_comp)
        split.addWidget(chart_w)
        split.setSizes([430, 620])
        lay.addWidget(split, 1)
        return w

    def _on_form_factor_changed(self, v):
        self.style["form_factor"] = float(v)
        self._update_composition_tab()

    def _update_composition_tab(self):
        if not hasattr(self, "comp_table"):
            return
        df = self.df if (self.df is not None and not self.df.empty) else None
        comp = compute_species_composition(df)
        div = compute_diversity_indices(comp)
        fw, fh = self.style.get("field_w"), self.style.get("field_h")
        area = (fw * fh / 10000.0) if (fw and fh) else None
        vol = compute_stand_volume(df, area, self.style.get("form_factor", DEFAULT_FORM_FACTOR))
        C = self.comp_cards
        C["S"].setText(str(div["S"]) if div["S"] else "-")
        C["H"].setText(_vn(div["H"], 3))
        C["J"].setText(_vn(div["J"], 3))
        C["D"].setText(_vn(div["simpson"], 3))
        C["d"].setText(_vn(div["margalef"], 3))
        C["M"].setText(_vn(vol["M_ha"], 1))
        self.lbl_comp_formula.setText(
            f"<b>Công thức tổ thành (IV%):</b>&nbsp; {build_composition_formula(comp)}")
        self.comp_table.setRowCount(len(comp))
        for i, (_, r) in enumerate(comp.iterrows()):
            vals = [str(int(r["STT"])), str(r["Loai"]), str(int(r["N"])), _vn(r["N%"], 2), _vn(r["G"], 4),
                    _vn(r["G%"], 2), _vn(r["IV%"], 2), _vn(r["Hvn_tb"], 1)]
            dominant = r["IV%"] >= IV_DOMINANT_THRESHOLD
            for j, v in enumerate(vals):
                it = QTableWidgetItem(v)
                it.setTextAlignment((Qt.AlignLeft if j == 1 else Qt.AlignRight) | Qt.AlignVCenter)
                if dominant and j in (1, 6):
                    f = it.font()
                    f.setBold(True)
                    it.setFont(f)
                    it.setForeground(QColor("#1b5e20"))
                self.comp_table.setItem(i, j, it)
        draw_species_composition(self.canvas_struct_comp.ax, comp, self.color_map)
        self.canvas_struct_comp.draw_idle()
        self._comp_cache = (comp, div, vol)

    def _copy_composition(self):
        comp, div, vol = getattr(self, "_comp_cache", (None, None, None))
        if comp is None or comp.empty:
            show_notice(self, "info", "Tổ thành loài", "Chưa có dữ liệu để copy.")
            return
        lines = ["STT\tLoài cây\tN\tN%\tG (m²)\tG%\tIV%\tD1.3 tb\tHvn tb"]
        for _, r in comp.iterrows():
            lines.append("\t".join([str(int(r["STT"])), str(r["Loai"]), str(int(r["N"])), _vn(r["N%"], 2),
                                    _vn(r["G"], 4), _vn(r["G%"], 2), _vn(r["IV%"], 2), _vn(r["D13_tb"], 1),
                                    _vn(r["Hvn_tb"], 1)]))
        lines += ["", "Công thức tổ thành:\t" + build_composition_formula(comp),
                  f"Số loài S\t{div['S']}", f"Shannon H'\t{_vn(div['H'], 3)}", f"Pielou J'\t{_vn(div['J'], 3)}",
                  f"Simpson (1-D)\t{_vn(div['simpson'], 3)}", f"Margalef d\t{_vn(div['margalef'], 3)}",
                  f"Trữ lượng M/ha (f={_vn(vol['f'], 2)})\t{_vn(vol['M_ha'], 1)}"]
        QApplication.clipboard().setText("\n".join(lines))
        self.statusBar().showMessage("Đã copy bảng tổ thành loài - dán thẳng vào Excel/Word.", 5000)

    # ==========================================================================
    # XUẤT BÁO CÁO CHUYÊN NGHIỆP (PDF / EXCEL)
    # ==========================================================================
    def export_report(self):
        if self.df is None or self.df.empty:
            show_notice(self, "info", "Xuất báo cáo", "Vui lòng nạp dữ liệu cây rừng trước khi xuất báo cáo.")
            return
        dlg = ReportDialog(self, form_factor=self.style.get("form_factor", DEFAULT_FORM_FACTOR))
        if dlg.exec() != QDialog.Accepted:
            return
        meta, sections, fmt = dlg.get_values()
        try:
            meta["view3d"] = (float(self.canvas_3d.ax.elev), float(self.canvas_3d.ax.azim))
        except Exception:
            meta["view3d"] = (22, -60)
        self.style["form_factor"] = meta["form_factor"]
        if hasattr(self, "ff_spin"):
            self.ff_spin.blockSignals(True)
            self.ff_spin.setValue(meta["form_factor"])
            self.ff_spin.blockSignals(False)

        want_pdf = fmt in ("pdf", "both", "all")
        want_docx = fmt in ("word", "all")
        want_xlsx = fmt in ("excel", "both", "all")
        settings = get_settings()
        last_dir = str(settings.value("report/last_dir", os.path.expanduser("~")))
        tag = re.sub(r"[^A-Za-z0-9_-]+", "_", strip_accents(meta.get("otc") or "OTC")).strip("_") or "OTC"
        base = os.path.join(last_dir, f"BaoCao_{tag}_{datetime.now().strftime('%Y%m%d')}")
        if want_pdf:
            title, filt, ext = "Lưu báo cáo PDF", "PDF (*.pdf)", ".pdf"
        elif want_docx:
            title, filt, ext = "Lưu báo cáo Word", "Word (*.docx)", ".docx"
        else:
            title, filt, ext = "Lưu báo cáo Excel", "Excel (*.xlsx)", ".xlsx"
        chosen, _ = QFileDialog.getSaveFileName(self, title, base + ext, filt)
        if not chosen:
            return
        stem = os.path.splitext(chosen)[0]
        pdf_path = stem + ".pdf" if want_pdf else None
        docx_path = stem + ".docx" if want_docx else None
        xlsx_path = stem + ".xlsx" if want_xlsx else None
        settings.setValue("report/last_dir", os.path.dirname(chosen))
        settings.sync()

        # Tự tính độ che phủ tán nếu chưa tính (để báo cáo đầy đủ chỉ tiêu)
        if (self._canopy_result is None and self.style.get("field_w") and self.style.get("field_h")):
            self.statusBar().showMessage("Đang tính độ che phủ tán cho báo cáo...")
            QApplication.processEvents()
            try:
                self._canopy_result = compute_canopy_coverage(
                    self.df, self.style, int(self.style.get("struct_grid_res", 140)))
            except Exception:
                self._canopy_result = None
        vstruct = compute_vertical_structure(self.df, k_std=self.style.get("vstruct_k_std", 1.0))
        style_copy = copy.deepcopy(self.style)
        QApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        self.statusBar().showMessage("Đang dựng báo cáo, vui lòng đợi trong giây lát...")
        QApplication.processEvents()
        done = []
        try:
            if pdf_path:
                n_pages = build_pdf_report(pdf_path, self.df, style_copy, dict(self.color_map), meta, sections,
                                           canopy=self._canopy_result, vstruct=vstruct, version=APP_VERSION)
                done.append(f"📄  PDF ({n_pages} trang): {pdf_path}")
            if docx_path:
                build_word_report(docx_path, self.df, style_copy, dict(self.color_map), meta, sections,
                                  canopy=self._canopy_result, vstruct=vstruct, version=APP_VERSION)
                done.append(f"📝  Word: {docx_path}")
            if xlsx_path:
                build_excel_report(xlsx_path, self.df, style_copy, meta, canopy=self._canopy_result,
                                   vstruct=vstruct, version=APP_VERSION)
                done.append(f"📊  Excel: {xlsx_path}")
        except PermissionError:
            QApplication.restoreOverrideCursor()
            show_notice(self, "error", "Xuất báo cáo",
                        "Không ghi được file - có thể file đang được mở bằng chương trình khác "
                        "(Word, Excel, Adobe Reader...). Hãy đóng file đó rồi thử lại.")
            return
        except Exception as e:
            QApplication.restoreOverrideCursor()
            show_notice(self, "error", "Xuất báo cáo", f"Có lỗi khi xuất báo cáo:\n{e}")
            return
        QApplication.restoreOverrideCursor()
        self.update_structure_analysis()
        self.statusBar().showMessage("Đã xuất báo cáo thành công.", 6000)

        box = QMessageBox(self)
        box.setWindowTitle("Xuất báo cáo")
        box.setIconPixmap(make_notice_pixmap("success", 48))
        box.setText("✅  Đã xuất báo cáo thành công!\n\n" + "\n".join(done))
        btn_open = box.addButton("📂  Mở báo cáo", QMessageBox.AcceptRole)
        box.addButton("Đóng", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is btn_open:
            QDesktopServices.openUrl(QUrl.fromLocalFile(pdf_path or docx_path or xlsx_path))



# ==============================================================================

# 8. ĐIỂM VÀO CHƯƠNG TRÌNH

# ==============================================================================



def main():

    app = QApplication(sys.argv)

    app.setStyle("Fusion")

    app.setStyleSheet(LIGHT_QSS)

    app.setApplicationName("TDForest")

    app.setWindowIcon(make_tree_icon(64))  # icon taskbar / alt-tab



    # ---- Màn hình chờ: hiện NGAY logo phần mềm + trạng thái "đang tải" -
    # phần mềm mất khoảng 5-10s để dựng xong giao diện + dữ liệu mẫu lần
    # đầu, nếu không có màn hình chờ người dùng dễ tưởng phần mềm bị treo.
    splash = build_startup_splash()

    splash.show()

    app.processEvents()



    win = MainWindow(splash=splash)

    win.show()

    splash.finish(win)  # đóng màn hình chờ ngay khi cửa sổ chính sẵn sàng

    # ---- Hiệu ứng tung hoa chào mừng ngay khi vào phần mềm - phát sau một
    # nhịp rất ngắn để đảm bảo cửa sổ chính đã có kích thước thật (đã show).
    _st = get_settings()
    if _st.value("app/confetti_version", "") != APP_VERSION:
        QTimer.singleShot(120, win.show_confetti_effect)
        _st.setValue("app/confetti_version", APP_VERSION)
        _st.sync()

    # ---- Sau khi khởi động xong: hiện hộp thoại "Thông tin" trong 3 giây
    # rồi tự đóng, giúp người dùng biết chắc chắn phần mềm đã tải xong và
    # nắm được thông tin liên hệ/phiên bản ngay từ đầu.
    QTimer.singleShot(200, lambda: AboutDialog(win, auto_close_ms=3000).exec())



    sys.exit(app.exec())





if __name__ == "__main__":

    main()