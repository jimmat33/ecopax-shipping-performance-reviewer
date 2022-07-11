cd "C:\USERS\jmattison\Desktop\Performance Reviewer\ecopax-shipping-performance-reviewer
python setup.py install
python setup.py py2exe
xcopy /s "C:\Users\jmattison\Desktop\Performance Reviewer\ecopax-shipping-performance-reviewer\exe dependency files" "C:\Users\jmattison\Desktop\Performance Reviewer\ecopax-shipping-performance-reviewer\dist" /mir
ren dist Ecopax-Performance-Reviwer-Program-Files