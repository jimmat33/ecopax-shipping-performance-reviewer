cd ecopax-shipping-performance-reviewer
python setup.py install
python setup.py py2exe
robocopy dist-cpy dist /E
ren dist Ecopax-Performance-Reviewer-Program-Files