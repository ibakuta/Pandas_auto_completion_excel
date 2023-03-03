import tkinter as tk
from tkinter import ttk
import excel_pndas_book_income
from tkinter import filedialog
from tkinter.ttk import Progressbar
import time

def open_file_bank():
    filename = filedialog.askopenfilename(filetypes = (("Text files","*.xlsx"),("all files","*.*")))
    global df
    df = excel_pndas_book_income.open_excel_bank(filename)
      
def open_file_book():
    filename = filedialog.askopenfilename(filetypes = (("Text files","*.xlsx"),("all files","*.*")))
    excel_pndas_book_income.save__excel_for_calc(df, filename)

def calculation():
    df_summary = excel_pndas_book_income.calculation_excel_book('книга доходів.xlsx')
    excel_pndas_book_income.save__excel_cumulative_total(df_summary,'книга доходів.xlsx')
    excel_pndas_book_income.delete_Sheet11('книга доходів.xlsx')
    
root = tk.Tk()
root.geometry('500x300+100+100')
root.title('Hey, automation!')


progress_bar1 = ttk.Progressbar(root, orient='horizontal', mode='determinate', maximum=90, value=0)
progress_bar1.grid(column=3, row=0, columnspan=2, padx=45, pady=20)
root.update()
progress_bar1['value'] = 0
root.update()

def step():
    while progress_bar1['value'] < 90:
        progress_bar1['value'] += 30
        root.update()
        time.sleep(0.1)

bt1 = ttk.Button(text="Открыть файл (банковская выписка)", command = lambda:[open_file_bank(), step()])
bt1.grid(row = 0, column = 2, padx = 45, pady = 20)


progress_bar2 = ttk.Progressbar(root, orient='horizontal', mode='determinate', maximum=90, value=0)
progress_bar2.grid(column=3, row=1, columnspan=2, padx=45, pady=20)
root.update()
progress_bar2['value'] = 0
root.update()

def step2():
    while progress_bar2['value'] < 90:
        progress_bar2['value'] += 30
        root.update()
        time.sleep(0.1)

bt2 = ttk.Button(text="Открыть файл (книга учета доходов)", command = lambda:[open_file_book(), step2()])
bt2.grid(row = 1, column = 2, padx = 45, pady = 10)


progress_bar3 = ttk.Progressbar(root, orient='horizontal', mode='determinate', maximum=90, value=0)
progress_bar3.grid(column=3, row=2, columnspan=2, padx=45, pady=20)
root.update()
progress_bar3['value'] = 0
root.update()

def step3():
    while progress_bar3['value'] < 90:
        progress_bar3['value'] += 30
        root.update()
        time.sleep(0.1)

bt3 = tk.Button(text="Рассчитать", background="ivory4", height=2, command = lambda:[calculation(), step3()])
bt3.grid(row = 2, column = 2, padx = 45, pady = 25, sticky= tk.W+tk.E)


root.mainloop()

