rm *.oxt
zip -r PracticaIndex.oxt * -x *.pyc *.swp
unopkg remove PracticaIndex.oxt
unopkg add PracticaIndex.oxt
