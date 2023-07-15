'''
docstr
'''
# pylint: disable=W0611
# pylint: disable=W0402

from distutils.core import setup
import py2exe

py2exe.patch_distutils()

packages=[
    'reportlab',
    'reportlab.graphics.charts',
    'reportlab.graphics.samples',
    'reportlab.graphics.widgets',
    'reportlab.graphics.barcode',
    'reportlab.graphics',
    'reportlab.lib',
    'reportlab.pdfbase',
    'reportlab.pdfgen',
    'reportlab.platypus',
]

setup(windows=[
    {'script': 'main.py',
     'icon_resources': [(1, 'gui_icon.ico')],
     'dest_base': 'Ecopax Transwide Analysis Tool'}],
    options={"py2exe": {"typelibs": [('{00020813-0000-0000-C000-000000000046}',
             '{000208D5-0000-0000-C000-000000000046}', 0, 1, 4)],
             "packages": packages},
})
