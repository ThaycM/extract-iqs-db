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
qi_icon_ns=BASE_DIR/"icons"/"qi_mm_ns.png"
qi_icon_s=BASE_DIR/"icons"/"qi_mm_s.png"
reports_icon=BASE_DIR/"icons"/"reports.png"
analysis_icon=BASE_DIR/"icons"/"show_a.png"
search_icon=BASE_DIR/"icons"/"search_icon.png"
complaint_icon=BASE_DIR/"icons"/"complaints_year.png"
confirm_complaint= BASE_DIR/"icons"/"complaints_opened.png"
working_list_icon= BASE_DIR/"icons"/"working_list.png"
update_icon= BASE_DIR/"icons"/"update.png"
options_icon= BASE_DIR/"icons"/"options_icons.png"
excel_icon=BASE_DIR/"icons"/"excel_icon.png"
all_modules_icon=BASE_DIR/"icons"/"all-modules.png"
mm_module_icon3v1=BASE_DIR/"icons"/"mm-module-icon3v1.png"
mm_module_icon3v2=BASE_DIR/"icons"/"mm-module-icon3v2.png"
customize_icon=BASE_DIR/"icons"/"customize_icon.png"
max_ribbon=BASE_DIR/"icons"/"maximize_ribbon_icon.png"
general_icon=BASE_DIR/"icons"/"general_icon.png"
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

def _maximize_ribbon():
    customize=pa.locateOnScreen(str(customize_icon))
    pa.click(customize)
    try:
        maximize=pa.locateOnScreen(str(max_ribbon))
        pa.click(maximize)
    except ImageNotFoundException:
        pa.click()
        pass
    else:
        maximize=wait_image(str(max_ribbon))

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

def open_mm_module():
        
        module_icon=pa.locateOnScreen(str(modules_icons),confidence=0.9)
        pa.moveTo(module_icon)
        pa.click()
        sleep(2)
        try:
            close_reminder=wait_image(str(closes_reminder))
            pa.moveTo(close_reminder)
            pa.click()
            sleep(2)
        except:
            pass
        
        _maximize_ribbon()    
        all_modules=pa.locateOnScreen(str(all_modules_icon),confidence=0.9)
        pa.click(all_modules)
        try:
            mm_module=wait_image(str(mm_module_icon3v1))
        except ImageNotFoundException:
            mm_module=wait_image(str(mm_module_icon3v2))
        pa.click(mm_module)
        try:
            wait_image(str(employee_icon1))
        except ImageNotFoundException:
            pass
        module_icon=pa.locateOnScreen(str(modules_icons),confidence=0.9)
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
    table_option=(1264,244)
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
    
def qi_module():
    _maximize_ribbon()
    module_icon=pa.locateOnScreen(str(modules_icons),confidence=0.9)
    pa.click(module_icon)
    # pa.click(module_icon)
    all_modules=wait_image(str(all_modules_icon))
    pa.click(all_modules)
    try:
        qi_mm= wait_image(str(qi_icon_ns))
    except ImageNotFoundException:
        qi_mm=wait_image(str(qi_icon_s))
    pa.doubleClick(qi_mm)

    try:
        general=wait_image(str(general_icon))
        pa.click(general)
    except ImageNotFoundException:
        pass
    try:
        wait_image(str(reports_icon))
    except ImageNotFoundException:
        qi_module()
    try:
        analysis=wait_image(str(analysis_icon))
    except ImageNotFoundException:
        pass
    pa.click(analysis)
    try:
        search=wait_image(str(search_icon))
    except ImageNotFoundException:
        search= (36,202)
        print("Posição localizada pelas coordenadas.")
    pa.click(search)
    pa.write("Complaints per year and method")
    pa.press("Enter")
    try:
        complaint=wait_image(str(complaint_icon))
    except ImageNotFoundException:
        pass
    pa.doubleClick(complaint)
    try:
        wait_image(str(confirm_complaint))
    except:
        pass
    try:
        working_list=wait_image(str(working_list_icon))
    except:
        pass
    pa.doubleClick(working_list)
    pa.press("Enter")
    sleep(2)
    update= wait_image(str(update_icon),confidence=0.9)
    pa.click(update)
    pa.click(1463,444)
    try:
        options=wait_image(str(options_icon))
    except ImageNotFoundException:
        pa.click(1463,444)
        try:
            options=wait_image(str(options_icon))
        except ImageNotFoundException:
            print("icone de options não encontrado.")
    sleep(1)
    pa.click(options)
    excel=wait_image(str(excel_icon))
    pa.click(excel)
####################################################################################
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
    db_name='problems_db'
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
def close_iqs():
    pa.click(1665,12,2,1)    
def main():
    open_iqs()
    open_mm_module()
    export_table()
    sleep(2)
    qi_module()
    close_iqs()

if __name__=="__main__":
    main()