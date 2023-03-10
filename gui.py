from fengzhuang import Li_spider
#selenium浏览器驱动下载地址 ：http://chromedriver.storage.googleapis.com/index.html
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import sys
import time
import xlwt
# This file was generated by the Tkinter Designer by Parth Jadhav
# https://github.com/ParthJadhav/Tkinter-Designer
from pathlib import Path
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage
import os


OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("./assets")


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)


window = Tk()

window.geometry("1000x500")
window.configure(bg = "#FFFFFF")


canvas = Canvas(
    window,
    bg = "#FFFFFF",
    height = 500,
    width = 1000,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)

canvas.place(x = 0, y = 0)
image_image_1 = PhotoImage(
    file=relative_to_assets("image_1.png"))
image_1 = canvas.create_image(
    500.0,
    350.0,
    image=image_image_1
)
LONE=Li_spider('https://www.fis-ski.com/DB/general/statistics.html?sectorcode=JP','JP.xls')
button_image_1 = PhotoImage(
    file=relative_to_assets("button_1.png"))
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=lambda:LONE.run_spider() ,
    relief="flat",
)
# print("button_1 clicked")
button_1.place(
    x=772.0,
    y=96.0,
    width=180.0,
    height=46.0
)
Ltwo=Li_spider('https://www.fis-ski.com/DB/general/statistics.html?sectorcode=CC','CC.xls')
button_image_2 = PhotoImage(
    file=relative_to_assets("button_2.png"))
button_2 = Button(
    image=button_image_2,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: Ltwo.run_spider(),
    relief="flat"
)
button_2.place(
    x=772.0,
    y=176.0,
    width=180.0,
    height=46.0
)
Lthree=Li_spider('https://www.fis-ski.com/DB/general/statistics.html?sectorcode=NK','NK.xls')
button_image_3 = PhotoImage(
    file=relative_to_assets("button_3.png"))
button_3 = Button(
    image=button_image_3,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: Lthree.run_spider(),
    relief="flat"
)
button_3.place(
    x=772.0,
    y=256.0,
    width=180.0,
    height=46.0
)


button_image_4 = PhotoImage(
    file=relative_to_assets("button_3.png"))
button_4 = Button(
    image=button_image_4,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: os.system('python xlwing.py'),
    relief="flat"
)
button_4.place(
    x=772.0,
    y=336.0,
    width=180.0,
    height=46.0
)

canvas.create_text(
    200.0,
    85.0,
    anchor="nw",
    text="Download SkiJump",
    fill="#FFFFFF",
    font=("Roboto Medium", 48 * -1)
)
canvas.create_text(
    200.0,
    170.0,
    anchor="nw",
    text="Download cross_country",
    fill="#FFFFFF",
    font=("Roboto Medium", 48 * -1)
)
canvas.create_text(
    200.0,
    260.0,
    anchor="nw",
    text="Download nodic_combined",
    fill="#FFFFFF",
    font=("Roboto Medium", 48 * -1)
)
canvas.create_text(
    200.0,
    335.0,
    anchor="nw",
    text="Analysis",
    fill="#FFFFFF",
    font=("Roboto Medium", 48 * -1)
)
button_1.bind('<Button-1>', print('1'))
window.resizable(False, False)
window.mainloop()

