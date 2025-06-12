import pyautogui as pa
from pyautogui import ImageNotFoundException
from time import sleep, time
from pathlib import Path

BASE_DIR=Path(__file__).resolve().parent

#Image Paths:
iqs_icon1=BASE_DIR/"icons"/"iqs-icon.png"
iqs_icon2=BASE_DIR/"icons"/"iqs-icon-2.png"
iqs_icon3=BASE_DIR/"icons"/"iqs-icon-3.png"
modules_icons=BASE_DIR/"icons"/"modules-icons.png"
closes_reminder=BASE_DIR/"icons"/"close-reminder-icon.png"
mm_module_icon2=BASE_DIR/"icons"/"mm-module-icon2.png"
mm_module_icon1=BASE_DIR/"icons"/"MM-module-icon.png"
employee_icon1=BASE_DIR/"icons"/"employee-icon.png"
employee_icon2=BASE_DIR/"icons"/"mm-module-icon2.png"
export_option=BASE_DIR/"icons"/"export-options.png"



pa.PAUSE=0.5
pa.FAILSAFE=True

def wait_image(path,timeout=30,confidence=0.95):
    start = time()
    while time() - start < timeout:
        try:
            loc = pa.locateOnScreen(path, confidence=confidence, grayscale=True)
        except ImageNotFoundException:
            loc=None
        if loc:
            return pa.center(loc)
    raise ImageNotFoundException(f"Imagem '{path}' não apareceu em {timeout}s.")

def open_iqs():
    pa.hotkey('win','d')
    sleep(3)
    try:
        iqs_icon=pa.locateOnScreen(str(iqs_icon1),confidence=0.9)
    except ImageNotFoundException:
        try:
            iqs_icon=pa.locateOnScreen(str(iqs_icon2),confidence=0.9)
        except ImageNotFoundException:
            iqs_icon=pa.locateOnScreen(str(iqs_icon3),confidence=0.9)
    pa.moveTo(iqs_icon)
    pa.doubleClick()
    pa.hotkey('win','d')
    opened_iqs= False
    attempts=0
    while  not opened_iqs:
        try:
            module_icon=pa.locateOnScreen(str(modules_icons),confidence=0.9)
        except ImageNotFoundException:
            module_icon=None
            sleep(5)
            attempts+=1
            if attempts>20:
                print('mais de 20 tentativas realizadas!')
                break
        if module_icon:
            opened_iqs = True
    return module_icon
def open_mm_module(module_icon):
    if module_icon:
        pa.moveTo(module_icon)
        pa.click()
        sleep(2)
        close_reminder=wait_image(str(closes_reminder))
        pa.moveTo(close_reminder)
        pa.click()
        sleep(2)
        try:
            mm_module=wait_image(str(mm_module_icon2))
        except ImageNotFoundException:
            mm_module=wait_image(str(mm_module_icon1))
        pa.click(mm_module)
        try:
            wait_image(str(employee_icon1))
        except ImageNotFoundException:
            pass
        pa.click(module_icon)
        mm_module_ok=wait_image(str(employee_icon2))
        if mm_module_ok:
            organizations=(202,417)
            folder=(36,248)
            enercon=(117,270)
            close_organizations=(276,174)
            pa.click(organizations)
            pa.click(folder)
            pa.doubleClick(enercon)
            pa.click(close_organizations)
def export_table():
    sleep(4)
    table_option=(1198,243)
    pa.moveTo(table_option)
    pa.click()
    sleep(2)
    export_bt=(1260,382)
    pa.moveTo(export_bt)
    pa.click()
    #########
    next_1=(1013,773)
    pa.moveTo(next_1)
    pa.click()
    file_options=pa.locateOnScreen(str(export_option))
    if file_options:
        pa.moveTo(file_options)
        pa.click()
    for _ in range(6):
        pa.press('tab')
    pa.press('enter')
    pa.write(r"C:/Users/00071228/OneDrive - ENERCON/QA Team - Follow up - Qualidade - databases")
    pa.press("enter")
    for _ in range(7):
        pa.press('tab')
    db_name='actions_db'
    pa.write(db_name)
    pa.press('Enter')
    pa.press('left')
    pa.press('Enter')
    pa.moveTo(next_1)
    pa.click()
    finish_opt=(772,409)
    pa.moveTo(finish_opt)
    pa.click()
    finish_bt=(1109,772)
    pa.moveTo(finish_bt)
    pa.click()
    pa.click(1665,12,2,1)
def main():
    step1=open_iqs()
    open_mm_module(step1)
    export_table()

if __name__=="__main__":
    main()