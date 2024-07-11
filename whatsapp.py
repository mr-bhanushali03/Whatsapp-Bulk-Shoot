import keyboard as k
import pyautogui
import time
import pandas as pd
import webbrowser as wb
from urllib.parse import quote

def send_whatsapp(data_file_excel,message_file_text,x_cord=969,y_cord=1024):
    df=pd.read_excel(data_file_excel,dtype={"Contact":str})
    name=df["Name"].values
    contact=df["Contact"].values
    files=message_file_text

    with open(files) as f:
        file_data=f.read()
        zipped=zip(name,contact)
        counter=0
        for(a,b) in zipped:
            msg = file_data.format(a)
            wb.open(f"https://web.whatsapp.com/send?phone=+91{b}&text={quote(msg)}")
            time.sleep(15)
            pyautogui.click(x_cord,y_cord)
            time.sleep(1)
            pyautogui.press("enter")
            time.sleep(10)
            k.press_and_release("ctrl+w")
            time.sleep(1)
            counter+=1
            print(f"Message sent to {a} at {b}")
        print(f"Total {counter} messages sent")
        print("Task Completed")

excel_path=r"C:\xampp\htdocs\Whatsapp Bulk\whatsapp list.xlsx"
text_path=r"C:\xampp\htdocs\Whatsapp Bulk\whatsdraft.txt"
send_whatsapp(excel_path,text_path)