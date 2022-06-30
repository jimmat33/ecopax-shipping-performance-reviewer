from distutils.core import setup
import py2exe


setup(windows=[{'script' : 'ecopax_shipping_performance_reviewer.py', "icon_resources": [(1, "gui_icon.ico")], "dest_base" : "Ecopax Transwide Analysis Tool"}])
