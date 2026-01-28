from functions import *
import sys


print("Selecione o primeiro campo no CX-Programmer")
var_name = read_varname()
window = get_window()
layout = read_layout()
set_layout(layout, var_name, window)
