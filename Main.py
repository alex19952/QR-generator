from sys import argv
import library


try:
    library.startGUI(*argv)
except Exception as exc:
    print(exc)
    input()
    

# компилировать в exe командой pyinstaller Main.py --hidden-import 'library' -w
# модуль library должен включать все библиотеки, используемые остальными модулями
