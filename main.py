import time
import sys
from bs4 import BeautifulSoup
from requests import get
import tkinter as tk
from openpyxl import load_workbook
from pycoingecko import CoinGeckoAPI
from tkinter.filedialog import askopenfilename

cg = CoinGeckoAPI()

window = tk.Tk()
window.title("Pobieracz Kursów 0.8 Beta")
bg_colour = 'lightblue'
window.configure(background=bg_colour)
window.wm_iconbitmap('icon.ico')
logo = tk.PhotoImage(file="logo.gif")


# RYSOWANIE MENU:


def draw_window():
    clear_frame()
    menu = tk.Menu(window)
    window.config(menu=menu)
    menu.add_command(label="    ZMIEŃ LOKALIZACJĘ ARKUSZA   ", command=lambda: change_xlsx_local())
    window_height = 310
    window_width = 620
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    window.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
    window.resizable(False, False)

    right_side = tk.Frame(window, background=bg_colour)
    right_side.pack(side=tk.RIGHT)
    title = tk.Label(right_side, image=logo)
    title.pack()
    space = tk.Label(right_side, text="\n", background=bg_colour)
    space.pack()
    button_go = tk.Button(right_side, activebackground='gold', text='AKTUALIZUJ ARKUSZ',
                          command=lambda: update_everything(), bg='green', fg="yellow",
                          font='Helvetica 11 bold', width=22)
    button_go.pack()

    left_side = tk.Frame(window)
    left_side.pack(side=tk.LEFT)

    button_usd = tk.Button(left_side, activebackground='gold', text='USD / PLN', command=lambda: change_token_menu(1),
                           bg='silver', fg="brown",
                           font='Helvetica 10 bold', width=25)
    button_usd.pack()

    button_eur = tk.Button(left_side, activebackground='gold', text='EUR / PLN', command=lambda: change_token_menu(2),
                           bg='silver', fg="brown",
                           font='Helvetica 10 bold', width=25)
    button_eur.pack()

    button_gbp = tk.Button(left_side, activebackground='gold', text='GBP / PLN', command=lambda: change_token_menu(3),
                           bg='silver', fg="brown",
                           font='Helvetica 10 bold', width=25)
    button_gbp.pack()

    button_gold = tk.Button(left_side, activebackground='gold', text='GOLD / USD', command=lambda: change_token_menu(4),
                            bg='silver', fg="brown",
                            font='Helvetica 10 bold', width=25)
    button_gold.pack()

    button_swda = tk.Button(left_side, activebackground='gold', text='SWDA ETF / GBP',
                            command=lambda: change_token_menu(5), bg='silver', fg="brown",
                            font='Helvetica 10 bold', width=25)
    button_swda.pack()

    button_emim = tk.Button(left_side, activebackground='gold', text='EMIM ETF / GBP',
                            command=lambda: change_token_menu(6), bg='silver', fg="brown",
                            font='Helvetica 10 bold', width=25)
    button_emim.pack()

    button_btc = tk.Button(left_side, activebackground='gold', text=ticker(7), command=lambda: change_token_menu(7),
                           bg='silver', fg="brown",
                           font='Helvetica 10 bold', width=25)
    button_btc.pack()

    button_eth = tk.Button(left_side, activebackground='gold', text=ticker(8), command=lambda: change_token_menu(8),
                           bg='silver', fg="brown",
                           font='Helvetica 10 bold', width=25)
    button_eth.pack()

    button_0 = tk.Button(left_side, activebackground='gold', text=ticker(9), command=lambda: change_token_menu(9),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_0.pack()

    button_1 = tk.Button(left_side, activebackground='gold', text=ticker(10), command=lambda: change_token_menu(10),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_1.pack()

    button_2 = tk.Button(left_side, activebackground='gold', text=ticker(11), command=lambda: change_token_menu(11),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_2.pack()

    left2_side = tk.Frame(window)
    left2_side.pack(side=tk.LEFT)

    button_3 = tk.Button(left2_side, activebackground='gold', text=ticker(12), command=lambda: change_token_menu(12),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_3.pack()

    button_4 = tk.Button(left2_side, activebackground='gold', text=ticker(13), command=lambda: change_token_menu(13),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_4.pack()

    button_5 = tk.Button(left2_side, activebackground='gold', text=ticker(14), command=lambda: change_token_menu(14),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_5.pack()

    button_6 = tk.Button(left2_side, activebackground='gold', text=ticker(15), command=lambda: change_token_menu(15),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_6.pack()

    button_7 = tk.Button(left2_side, activebackground='gold', text=ticker(16), command=lambda: change_token_menu(16),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_7.pack()

    button_8 = tk.Button(left2_side, activebackground='gold', text=ticker(17), command=lambda: change_token_menu(17),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_8.pack()

    button_9 = tk.Button(left2_side, activebackground='gold', text=ticker(18), command=lambda: change_token_menu(18),
                         bg='silver', fg="brown",
                         font='Helvetica 10 bold', width=25)
    button_9.pack()

    button_10 = tk.Button(left2_side, activebackground='gold', text=ticker(19), command=lambda: change_token_menu(19),
                          bg='silver', fg="brown",
                          font='Helvetica 10 bold', width=25)
    button_10.pack()

    button_11 = tk.Button(left2_side, activebackground='gold', text=ticker(20), command=lambda: change_token_menu(20),
                          bg='silver', fg="brown",
                          font='Helvetica 10 bold', width=25)
    button_11.pack()

    button_12 = tk.Button(left2_side, activebackground='gold', text=ticker(21), command=lambda: change_token_menu(21),
                          bg='silver', fg="brown",
                          font='Helvetica 10 bold', width=25)
    button_12.pack()

    button_13 = tk.Button(left2_side, activebackground='gold', text=ticker(22), command=lambda: change_token_menu(22),
                          bg='silver', fg="brown",
                          font='Helvetica 10 bold', width=25)
    button_13.pack()


def change_token_menu(id):
    load_track()
    global ident_t
    global sheet_t
    global cell_t
    clear_frame()
    menu = tk.Menu(window)
    window.config(menu=menu)
    menu.add_command(label="    ZMIEŃ LOKALIZACJĘ ARKUSZA   ", command=lambda: change_xlsx_local())
    window_height = 310
    window_width = 620
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    window.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
    window.resizable(False, False)

    right_side = tk.Frame(window, background=bg_colour)
    right_side.pack(side=tk.RIGHT)
    title = tk.Label(right_side, image=logo)
    title.pack()
    space = tk.Label(right_side, text="\n", background=bg_colour)
    space.pack()
    button_go = tk.Button(right_side, activebackground='gold', text='MENU GŁÓWNE', command=lambda: draw_window(),
                          bg='green',
                          fg="yellow",
                          font='Helvetica 11 bold', width=22)
    button_go.pack()

    left_side = tk.Frame(window, background=bg_colour)
    left_side.pack(side=tk.LEFT)
    if not str(id) in ['1', '2', '3', '4', '5', '6']:
        name = tk.Label(left_side, text="Podaj API id tokena z Coingecko:", fg="brown", font='Helvetica 11 bold',
                        background=bg_colour)
        name.pack()
        ident_t = tk.Entry(left_side, width=40)
        ident_t.pack()
        ident_t.focus_set()
        ident_t.insert(0, ticker(id))
    name1 = tk.Label(left_side, text="Wybierz arkusz w zeszycie:", fg="brown", font='Helvetica 11 bold',
                     background=bg_colour)
    name1.pack()

    sheet_base = wb.sheetnames
    sheet_base.remove('data')
    sheet_t = tk.StringVar(left_side)
    if sheet(id) == 'None':
        sheet_t.set(sheet_base[0])
    else:
        sheet_t.set(sheet(id))
    sheet1 = tk.OptionMenu(left_side, sheet_t, *sheet_base)
    sheet1.pack()

    name2 = tk.Label(left_side, text="Podaj komórkę w arkuszu (np. A1):", fg="brown", font='Helvetica 11 bold',
                     background=bg_colour)
    name2.pack()
    cell_t = tk.Entry(left_side, width=40)
    cell_t.pack()
    cell_t.insert(0, cell(id))
    space = tk.Label(left_side, text="\n", background=bg_colour)
    space.pack()
    button_get = tk.Button(left_side, activebackground='gold', text='ZMIEŃ / DODAJ', command=lambda: get_token(id),
                           bg='green',
                           fg="yellow",
                           font='Helvetica 11 bold', width=20)
    button_get.pack()


def clear_frame():
    for obiekty in window.winfo_children():
        obiekty.destroy()


def ticker(id):
    data = wb['data']
    ticker = data.cell(row=1, column=id)
    if str(ticker.value) == 'None':
        return 'TWÓJ TOKEN'
    return str(ticker.value)


def sheet(id):
    data = wb['data']
    ticker = data.cell(row=2, column=id)
    if str(ticker.value) == 'None':
        return ''
    return str(ticker.value)


def cell(id):
    data = wb['data']
    ticker = data.cell(row=3, column=id)
    if str(ticker.value) == 'None':
        return ''
    return str(ticker.value)


# EDYCJA TOKENÓW:


def get_token(id):
    if not str(id) in ['1', '2', '3', '4', '5', '6']:
        ticker = ident_t.get()
    sheet = sheet_t.get()
    cell = cell_t.get()

    data = wb['data']
    if not str(id) in ['1', '2', '3', '4', '5', '6']:
        data.cell(row=1, column=id).value = ticker
    data.cell(row=2, column=id).value = sheet
    data.cell(row=3, column=id).value = cell

    wb.save(trak)
    draw_window()


# ZBIERANIE NOTOWAŃ:


def get_token_price_from_coingecko(id):
    try:
        data = wb['data']
        ticker = data.cell(row=1, column=id)
        ticker = str(ticker.value)

        price = cg.get_price(ids=ticker, vs_currencies='usd')
        price_exact = price[ticker]['usd']

        data = wb['data']
        sheet = data.cell(row=2, column=id)
        cell = data.cell(row=3, column=id)
        sheet_exact = wb[str(sheet.value)]
        sheet_exact[str(cell.value)] = price_exact
        wb.save(trak)
    except:
        None


def get_fiat_price(link, id):
    try:
        url = link
        page = get(url)
        bs = BeautifulSoup(page.content, 'html.parser')
        c = 0

        for nastronie in bs.find_all('div', class_='left'):
            if c == 0:
                price = nastronie.find('span', itemprop='price')
                price = str(price)
                price = price.replace('<span content="', '')
                price = price.replace('" itemprop="price">', '')
                price = price.replace('</span>', '')
                price = price[0:6]
                price = price.replace('.', ',')
                c += 1

        data = wb['data']
        sheet = data.cell(row=2, column=id)
        cell = data.cell(row=3, column=id)
        sheet_exact = wb[str(sheet.value)]
        sheet_exact[str(cell.value)] = price
        wb.save(trak)
    except:
        None


def get_gold_price(link, id):
    try:
        global notowania
        url = link
        page = get(url)
        bs = BeautifulSoup(page.content, 'html.parser')

        for nastronie in bs.find_all('div', class_='data-blk bid'):
            price = nastronie.find('span').get_text()
            price = price.replace(",", "")
            price = price.replace(".", ",")

        data = wb['data']
        sheet = data.cell(row=2, column=id)
        cell = data.cell(row=3, column=id)
        sheet_exact = wb[str(sheet.value)]
        sheet_exact[str(cell.value)] = price
        wb.save(trak)
    except:
        None


def get_swda_price(link, id):
    try:
        global notowania
        url = link
        page = get(url)
        bs = BeautifulSoup(page.content, 'html.parser')

        for nastronie in bs.find_all('span', id="aq_swda.uk_c2"):
            nastronie = str(nastronie)
            price = nastronie.replace('<span id="aq_swda.uk_c2">', "")
            price = price.replace('</span>', "")
            price = price.replace('.', ",")

        data = wb['data']
        sheet = data.cell(row=2, column=id)
        cell = data.cell(row=3, column=id)
        sheet_exact = wb[str(sheet.value)]
        sheet_exact[str(cell.value)] = price
        wb.save(trak)
    except:
        None


def get_emim_price(link, id):
    try:
        global notowania
        url = link
        page = get(url)
        bs = BeautifulSoup(page.content, 'html.parser')

        for nastronie in bs.find_all('span', id="aq_emim.uk_c1"):
            nastronie = str(nastronie)
            price = nastronie.replace('<span id="aq_emim.uk_c1">', "")
            price = price.replace('</span>', "")
            price = price.replace('.', ",")

        data = wb['data']
        sheet = data.cell(row=2, column=id)
        cell = data.cell(row=3, column=id)
        sheet_exact = wb[str(sheet.value)]
        sheet_exact[str(cell.value)] = price
        wb.save(trak)
    except:
        None


def update_everything():
    load_track()
    get_fiat_price('https://e-kursy-walut.pl/kurs-dolara/', 1)
    get_fiat_price('https://e-kursy-walut.pl/kurs-euro/', 2)
    get_fiat_price('https://e-kursy-walut.pl/kurs-funta/', 3)
    get_gold_price('https://www.kitco.com/charts/livegold.html', 4)
    get_swda_price('https://stooq.pl/q/?s=swda.uk', 5)
    get_emim_price('https://stooq.pl/q/?s=emim.uk', 6)
    get_token_price_from_coingecko(7)
    get_token_price_from_coingecko(8)
    get_token_price_from_coingecko(9)
    get_token_price_from_coingecko(10)
    get_token_price_from_coingecko(11)
    get_token_price_from_coingecko(12)
    get_token_price_from_coingecko(13)
    get_token_price_from_coingecko(14)
    get_token_price_from_coingecko(15)
    get_token_price_from_coingecko(16)
    get_token_price_from_coingecko(17)
    get_token_price_from_coingecko(18)
    get_token_price_from_coingecko(19)
    get_token_price_from_coingecko(20)
    get_token_price_from_coingecko(21)
    get_token_price_from_coingecko(22)
    sys.exit()


# OPERACJE ZE ŚCIEŻKĄ DO ARKUSZA:


def load_track():
    global trak
    try:
        temp = open('data.txt', 'r')
        trak = temp.readline()
        temp.close()
    except:
        check_file_xlsx


def check_file_xlsx():
    try:
        global wb
        wb = load_workbook(trak)
        draw_window()
    except:
        global track
        global ident
        clear_frame()
        window_height = 310
        window_width = 620
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))
        window.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
        window.resizable(False, False)

        right_side = tk.Frame(window, background=bg_colour)
        right_side.pack(side=tk.RIGHT)
        title = tk.Label(right_side, image=logo)
        title.pack()
        space = tk.Label(right_side, text="\n", background=bg_colour)
        space.pack()
        left_side = tk.Frame(window, background=bg_colour)
        left_side.pack(side=tk.LEFT)
        name = tk.Label(left_side, text="    Podaj pełną ścieżkę do arkusza w formacie xlsx:", fg="brown",
                        font='Helvetica 11 bold', background=bg_colour)
        name.pack()
        track1 = askopenfilename(title="Wybierz zeszyt")
        track = tk.Entry(left_side, width=50)
        track.pack()
        track.focus_set()
        track.insert(0, track1)
        but = tk.Button(left_side, activebackground='gold', text='...', command=lambda: check_file_xlsx(),
                               bg='green', fg="yellow", font='Helvetica 11 bold', width=4)
        but.pack()
        space = tk.Label(left_side, text="\n", background=bg_colour)
        space.pack()
        button_get = tk.Button(left_side, activebackground='gold', text='DODAJ ŚCIEŻKĘ', command=lambda: add_track(),
                               bg='green',
                               fg="yellow",
                               font='Helvetica 11 bold', width=20)
        button_get.pack()


def change_xlsx_local():
    global track
    global ident
    clear_frame()
    menu = tk.Menu(window)
    window.config(menu=menu)
    menu.add_command(label="    ZMIEŃ LOKALIZACJĘ ARKUSZA   ", command=lambda: change_xlsx_local())
    window_height = 310
    window_width = 620
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    window.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
    window.resizable(False, False)

    right_side = tk.Frame(window, background=bg_colour)
    right_side.pack(side=tk.RIGHT)
    title = tk.Label(right_side, image=logo)
    title.pack()
    space = tk.Label(right_side, text="\n", background=bg_colour)
    space.pack()
    button_go = tk.Button(right_side, activebackground='gold', text='MENU GŁÓWNE', command=lambda: draw_window(),
                          bg='green',
                          fg="yellow",
                          font='Helvetica 11 bold', width=22)
    button_go.pack()
    left_side = tk.Frame(window, background=bg_colour)
    left_side.pack(side=tk.LEFT)
    name = tk.Label(left_side, text="    Podaj pełną ścieżkę do arkusza w formacie xlsx:", fg="brown",
                    font='Helvetica 11 bold', background=bg_colour)
    name.pack()
    time.sleep(0.5)
    track1 = askopenfilename(title="Wybierz zeszyt")
    track = tk.Entry(left_side, width=50)
    track.pack()
    track.focus_set()
    track.insert(0, track1)
    space = tk.Label(left_side, text="\n", background=bg_colour)
    space.pack()
    button_get = tk.Button(left_side, activebackground='gold', text='DODAJ ŚCIEŻKĘ', command=lambda: add_track(),
                           bg='green',
                           fg="yellow",
                           font='Helvetica 11 bold', width=20)
    button_get.pack()


def add_track():
    global wb
    try:
        tracker = track.get()
        wb = load_workbook(tracker)
        temp = open('data.txt', 'w')
        temp.write(tracker)
        if 'data' in wb.sheetnames:
            None
        else:
            wb.create_sheet('data')
            hidden = wb['data']
            hidden.sheet_state = 'hidden'
            wb.save(tracker)
        wb = load_workbook(tracker)
        draw_window()
    except:
        try:
            load_track()
        except:
            check_file_xlsx()


# WYWOŁANIE PROGRAMU:


load_track()
check_file_xlsx()
window.mainloop()
