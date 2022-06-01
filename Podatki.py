#python3
import xlrd
import pandas as pd
from tkinter import *
import tkinter as tk
import textwrap as tw
import tkinter.font as font
import sys, os
import xlsxwriter
from PIL import Image,ImageTk
#pyinstaller -F --hidden-import='xlsxwriter' scratch.py
df = pd.read_excel(os.getcwd() + '\Podatki.xlsx')
lista_testy = ['Przepisy O Doradztwie Podatkowym I Etyka Zawodowa','Ewidencja Podatkowa I Zasady Prowadzenia Ksiąg Rachunkowych','Rachunkowość','Organizacja I Funkcjonowanie Krajowej Administracji Skarbowej','Prawo Karne Skarbowe','Prawo Dewizowe','Międzynarodowe, Wspólnotowe I Krajowe ','Postępowanie Egzekucyjne W Administracji','Ustawa O Zasadach Ewidencji I Identyfikacji Podatników I Płatników','Wydawanie Zaświadczeń Przez Organy Podatkowe','Czynności Sprawdzające I Kontrola Podatkowa','Postępowanie Podatkowe','Kodeks Postępowania Administracyjnego','Opłata Skarbowa I Inne Opłaty Samorządowe','Podatek Od Spadków I Darowizn','Podatek Od Czynności Cywilnoprawnych','Podatek Rolny I Podatek Leśny','Podatek Od Środków Transportowych','Podatki I Opłaty Samorządowe','Podatek Od Niektórych Instytucji Finansowych','Podatek Od Gier','Podatek Od Wydobycia Niektórych Kopalin','Podatek Tonażowy','Podatek Dochodowy Od Osób Prawnych','Podatek Dochodowy Od Osób Fizycznych','Podatek Od Towarów I Usług','Tajemnica Skarbowa','V Zobowiązania Podatkowe','IV Materialne Prawo Podatkowe Zagadnienia Wspólne','III Podstawy Międzynarodowego Oraz Wspólnotowego Prawa Podatkowego','II Analiza Podatkowa','I Źródła Prawa I Wykładnia Prawa']

root = Tk()
root.title('Testy podatkowe')

mainframe = Frame(root)
mainframe.grid(column=3, row=6, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=5)
mainframe.rowconfigure(0, weight=5)
mainframe.grid(pady=100, padx=100)

# dorb box
clicked = StringVar()
clicked.set("Wybierz temat")
drop = OptionMenu(root, clicked, *lista_testy)
droped_list_value = clicked.get()
questions = df.loc[df['Grupa'] == droped_list_value]

lp = 0
correct_answers = 0
incorrect_answers = 0
question_number = 0


def opk():
    droped_list_value = clicked.get()
    global questions
    global questions_pac_number

    questions = df.loc[df['Grupa'] == droped_list_value]
    global nt
    nt = questions
    questions_pac_number = len(questions.index)

global p
p = 0

def nowe():
    try:
        myLabelodpa.destroy()

    except:
        pass


def buttonaa():
    nowe()
    global myButtood
    global df
    global correct_answers
    global incorrect_answers
    if questions.iloc[t,6] == "a":

        myButtood["text"] = str('Prawidłowa odpowiedź') +  '\n '+ tw.fill(str(questions.iloc[t,7]),width=100)
        myButtood["foreground"] = "green"
        myButtona["background"] = "green"
        myButtonb["background"] = "red"
        myButtonc["background"] = "red"
        df.iloc[lp, 9] += 1
        correct_answers += 1
        # Score number for currect seesion
        odpa = "Prawidłowe odp.: " + str(correct_answers) + '/' + '\n' + "Nieprawidłowe odp.: " + str(incorrect_answers) + '/' + '\n' + "Pytanie: " + str(question_number) + '/' + str(questions_pac_number)
        myLabelodpa = Label(root, text=odpa)
        myLabelodpa.grid(row=3, column=1)
    else:
        myButtood["text"] = str('Nieprawidłowa odpowiedź') + '\n ' + tw.fill(str(questions.iloc[t, 7]), width=100)
        myButtood["foreground"] = "red"
        df.iloc[lp, 10] += 1
        incorrect_answers += 1
        # Score number for currect seesion
        odpa = "Prawidłowe odp.: " + str(correct_answers) + '/' + '\n' + "Nieprawidłowe odp.: " + str(incorrect_answers) + '/' + '\n' + "Pytanie: " + str(question_number) + '/' + str(questions_pac_number)
        myLabelodpa = Label(root, text=odpa)
        myLabelodpa.grid(row=3, column=1)
        if questions.iloc[t,6] == "b":
            myButtona["background"] = "red"
            myButtonb["background"] = "green"
            myButtonc["background"] = "red"
        elif questions.iloc[t,6] == "c":
            myButtona["background"] = "red"
            myButtonb["background"] = "red"
            myButtonc["background"] = "green"

def buttonab():
    nowe()
    global myLabelodpa
    global df
    global correct_answers
    global incorrect_answers
    if questions.iloc[t, 6] == "b":
        myButtood["text"] = str('Prawidłowa odpowiedź') + '\n ' + tw.fill(str(questions.iloc[t, 7]), width=100)
        myButtood["foreground"] = "green"
        myButtona["background"] = "red"
        myButtonb["background"] = "green"
        myButtonc["background"] = "red"
        df.iloc[lp, 9] += 1
        correct_answers += 1
        # Score number for currect seesion
        odpa = "Prawidłowe odp.: " + str(correct_answers) + '/' + '\n' + "Nieprawidłowe odp.: " + str(incorrect_answers) + '/' + '\n' + "Pytanie: " + str(question_number)+ '/' + str(questions_pac_number)
        myLabelodpa = Label(root, text=odpa)
        myLabelodpa.grid(row=3, column=1)
    else:
        myButtood["text"] = str('Nieprawidłowa odpowiedź') + '\n ' + tw.fill(str(questions.iloc[t, 7]), width=100)
        myButtood["foreground"] = "red"
        df.iloc[lp, 10] += 1
        incorrect_answers += 1
        # Score number for currect seesion
        odpa = "Prawidłowe odp.: " + str(correct_answers) + '/' + '\n' + "Nieprawidłowe odp.: " + str(incorrect_answers) + '/' + '\n' + "Pytanie: " + str(question_number) + '/' + str(questions_pac_number)
        myLabelodpa = Label(root, text=odpa)
        myLabelodpa.grid(row=3, column=1)
        if questions.iloc[t, 6] == "a":
            myButtona["background"] = "green"
            myButtonb["background"] = "red"
            myButtonc["background"] = "red"
        elif questions.iloc[t, 6] == "c":
            myButtona["background"] = "red"
            myButtonb["background"] = "red"
            myButtonc["background"] = "green"
            
def buttonac():
    nowe()
    global myLabelodpa
    global df
    global correct_answers
    global incorrect_answers
    if questions.iloc[t, 6] == "c":
        myButtood["text"] = str('Prawidłowa odpowiedź') + '\n ' + tw.fill(str(questions.iloc[t, 7]), width=100)
        myButtood["foreground"] = "green"
        myButtona["background"] = "red"
        myButtonb["background"] = "red"
        myButtonc["background"] = "green"
        df.iloc[lp,9] += 1
        correct_answers += 1
        # Score number for currect seesion
        odpa = "Prawidłowe odp.: " + str(correct_answers) + '/' + '\n' + "Nieprawidłowe odp.: " + str(incorrect_answers) + '/' + '\n' + "Pytanie: " + str(question_number) + '/' + str(questions_pac_number)
        myLabelodpa = Label(root, text=odpa)
        myLabelodpa.grid(row=3, column=1)
    else:
        myButtood["text"] = str('Nieprawidłowa odpowiedź') + '\n ' + tw.fill(str(questions.iloc[t, 7]), width=100)
        myButtood["foreground"] = "red"
        df.iloc[lp,10] += 1
        incorrect_answers += 1
        # Score number for currect seesion
        odpa = "Prawidłowe odp.: " + str(correct_answers) + '/' + '\n' + "Nieprawidłowe odp.: " + str(incorrect_answers) + '/' + '\n' + "Pytanie: " + str(question_number) + '/' + str(questions_pac_number)
        myLabelodpa = Label(root, text=odpa)
        myLabelodpa.grid(row=3, column=1)
        if questions.iloc[t, 6] == "a":
            myButtona["background"] = "green"
            myButtonb["background"] = "red"
            myButtonc["background"] = "red"

        elif questions.iloc[t, 6] == "b":
            myButtona["background"] = "red"
            myButtonb["background"] = "green"
            myButtonc["background"] = "red"
            
def pytanie():
    global lp
    global p
    global t
    global myLabel
    global kto
    global myLabelodpa
    global myButtonbb
    global myButtoncc
    global buttonod
    global question_number
    global odpa
    t = p
  
    # rest coulor and score

    myButtona["background"] = "SystemButtonFace"
    myButtonb["background"] = "SystemButtonFace"
    myButtonc["background"] = "SystemButtonFace"
    myButtood["foreground"] = "black"

    # buttons questions display
    myButton1["text"] = tw.fill(str(questions.iloc[t,2]),width=100)
    myButtona["text"] = tw.fill(str(questions.iloc[t,3]),width=100)
    myButtonb["text"] = tw.fill(str(questions.iloc[t,4]),width=100)
    myButtonc["text"] = tw.fill(str(questions.iloc[t,5]),width=100)

    #Score saved in excel

    lp = int(questions.iloc[t,0]) - 1
    df.iloc[lp, 8] += 1
    writer = pd.ExcelWriter('Podatki.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False)
    writer.save()
    question_number += 1
    p += 1

t = p

myButtonaa = "A"
myButtonbb = "B"
myButtoncc = "C"
myButtonodd = "==>"
odpa = 0

myFont = font.Font(size=15)

myButton = Button(root, text="Zatwierdź listę", padx=40, pady=20, command=opk)
myButton1 = Button(root, text="Pytanie", width=100, height=3)
myButtona = Button(root, text=myButtonaa, width=100, height=3, command=buttonaa)
myButtonb = Button(root, text=myButtonbb, width=100, height=3, command=buttonab)
myButtonc = Button(root, text=myButtoncc, width=100, height=3, command=buttonac)
myButtood = Button(root, text=myButtonodd, width=100, height=4, command=pytanie)

myButton1['font'] = myFont
myButtona['font'] = myFont
myButtonb['font'] = myFont
myButtonc['font'] = myFont
myButtood['font'] = myFont

drop.grid(row=0, column=2)
myButton.grid(row=1, column=1)

myButton1.grid(row=1, column=2)
myButtona.grid(row=2, column=2)
myButtonb.grid(row=3, column=2)
myButtonc.grid(row=4, column=2)
myButtood.grid(row=5, column=2)

root.mainloop()
