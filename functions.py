import win32gui
import win32con
import win32com.client

def window_enumeration_handler(hwnd, results_list):
    if win32gui.IsWindowVisible(hwnd):
        window_title = win32gui.GetWindowText(hwnd)
        #List all Cx-Programmer Screens
        if "CX-Programmer" in window_title:
            results_list.append((hwnd, window_title))
    return True


def get_top_level_windows():
    windows = []
    win32gui.EnumWindows(window_enumeration_handler, windows)
    return windows


def read_layout():
    print("Cole o layout da etiqueta e aperte CTRL-D:")

    layout = ""
    i = 0

    while True:
        try:
            line = input()
        except EOFError:
            break
        if i > 0:
            layout += "\n" + line
        else:
            layout += line
            i = 1

    return layout.encode().hex()


def read_varname():
    return input("Digite o nome da variavel: ")


def get_window():
    windows_list = get_top_level_windows()
    for i in range(windows_list.__len__()):
        print(f"{i} - {windows_list[i][1]}")

    window_number = input("Digite o número da janela do programa: ")
    main_window = win32gui.FindWindow(None, windows_list[int(window_number)][1])
    return win32gui.GetWindow(main_window, win32con.GW_CHILD)


def set_layout(layout, var_name, child_window):
    from time import sleep
    from pynput.keyboard import Key, Controller

    shell = win32com.client.Dispatch("WScript.Shell")
    keyboard = Controller()
    shell.SendKeys('%')

    try:
        win32gui.ShowWindow(int(child_window), win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(child_window)
        sleep(1)
    except Exception as e:
        print(f"Erro ao focar: {e}")
        return

    cnt_var = 0

    for i in range(0, len(layout), 4):
        try:

            hex1 = str(layout[i])
            hex2 = str(layout[i + 1])
            hex3 = str(layout[i + 2])
            hex4 = str(layout[i + 3])

            script = str(cnt_var)
            keyboard.press("m")
            keyboard.press("O")
            keyboard.press("V")
            keyboard.press(" ")
            keyboard.press("#")
            sleep(0.1)
            keyboard.press(hex1)
            sleep(0.1)
            keyboard.press(hex2)
            sleep(0.1)
            keyboard.press(hex3)
            sleep(0.1)
            keyboard.press(hex4)
            sleep(0.1)
            keyboard.press(" ")
            for letter in var_name:
                sleep(0.1)
                keyboard.press(letter)
            keyboard.press("[")
            for n in range(len(script)):
                sleep(0.1)
                keyboard.press(script[n])
            keyboard.press("]")
            keyboard.press(Key.enter)
            sleep(0.3)
            keyboard.press(Key.enter)
            if cnt_var % 79 == 0 and cnt_var != 0:
                keyboard.press(Key.down)
            sleep(0.5)
            cnt_var += 1

        except IndexError:
            hex1 = str(layout[i])
            hex2 = str(layout[i + 1])
            script = str(cnt_var)

            keyboard.press("m")
            keyboard.press("O")
            keyboard.press("V")
            keyboard.press(" ")
            keyboard.press("#")
            sleep(0.1)
            keyboard.press(hex1)
            sleep(0.1)
            keyboard.press(hex2)
            keyboard.press(" ")
            for letter in var_name:
                sleep(0.1)
                keyboard.press(letter)
            keyboard.press("[")
            for n in range(len(script)):
                sleep(0.1)
                keyboard.press(script[n])
            keyboard.press("]")
            keyboard.press(Key.enter)
            sleep(0.5)
            keyboard.press(Key.enter)
            keyboard.press(Key.end)
            sleep(0.5)
            cnt_var += 1

        except Exception as e:
            print(f"Erro: {e}")


