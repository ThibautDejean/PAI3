import imaplib, email,datetime
from tkinter import * 
import tkinter as tk
from tkinter import ttk
import openpyxl as xl
import sqlite3
import time
from pytz import timezone
from PIL import ImageTk,Image
import re
import winsound

   
def tri_geographique(corps,id_aeronef,decoupage) : 

    res = True
    ORG_EMAIL = "@outlook.fr" 
    usernm = "test.pai3" + ORG_EMAIL 
    passwd = "Tomblanchard3."
    conn = imaplib.IMAP4_SSL('outlook.office365.com')
    conn.login(usernm,passwd)
    conn.select('Inbox')

    # Recuperation arrivee, depart et noms de ville

    ligne=corps[4].split('-')
    depart=[ligne[1][1:5]]
    
    
    ligne2=corps[8].split('-')
    arrivee=[ligne2[1][1:5]]
    

    ligne3 = corps[6].split(' ')
    chemin = ligne3[2:len(ligne3)-1]
    

    liste_geo = depart + arrivee 
    
    # Est ce que les villes sont dans la zone de surveillance ? 

    base = sqlite3.connect('/Users/thibautdejean/Downloads/PAI-git/PAI-3/Aerodromes.sqlite')
    cur = base.cursor()
    
    for lieu in liste_geo : 
        cur.execute(f'''SELECT {decoupage} FROM "Liste_des_Aerodromes_en_France"  WHERE "CodeOACI" = ? ''', (lieu,))
        a = cur.fetchall()[0][0]
        print(a)
        if a != '1' :
            res = False
        else :
            res = True
    
    base.close()
    #Recherche de l'identifiant du mail
    
    if res == False : 

        _, messages = conn.search(None, 'ALL')

        # Recuperer l'identifiant du dernier message
        latest_message_id = messages[0].split()[-1]

        # Recuperer l'en-tete du dernier message
        _, msg_headers = conn.fetch(latest_message_id, '(BODY.PEEK[HEADER])')

        # Analyser l'en-tete pour savoir si le message est lu ou non
        msg = email.message_from_bytes(msg_headers[0][1])
        
        # Recuperer le contenu du dernier message
        _, msg_content = conn.fetch(latest_message_id, '(BODY[TEXT])')
        content = msg_content[0][1].decode()

        conn.copy(latest_message_id, 'Hors_zone')
        conn.store(latest_message_id, '+FLAGS', '\\Deleted')


    conn.expunge()
    conn.close()
    conn.logout()
    return(res)


def decoup1():
    self.__valid1.config(state=tk.NORMAL)
    self.__img = ImageTk.PhotoImage(Image.open('norm.png')) 
    self.__Canva.create_image(150, 145, image=self.__img)
    self.__decoupage = "NORM"

def decoup2(): 
    self.__valid1.config(state=tk.NORMAL)
    self.__img = ImageTk.PhotoImage(Image.open('LY00.png')) 
    self.__Canva.create_image(150, 145, image=self.__img)
    self.__decoupage = "LY00"

def decoup3(): 
    self.__valid1.config(state=tk.NORMAL)
    self.__img = ImageTk.PhotoImage(Image.open('LY1T.png')) 
    self.__Canva.create_image(150, 145, image=self.__img)
    self.__decoupage = "LY1T"

def decoup4(): 
    self.__valid1.config(state=tk.NORMAL)
    self.__img = ImageTk.PhotoImage(Image.open('LY10.png')) 
    self.__Canva.create_image(150, 145, image=self.__img)
    self.__decoupage = "LY10"

def decoup5(): 
    self.__valid1.config(state=tk.NORMAL)
    self.__img = ImageTk.PhotoImage(Image.open('LY11.png')) 
    self.__Canva.create_image(150, 145, image=self.__img)
    self.__decoupage = "LY11"

def decoup6():
    self.__valid1.config(state=tk.NORMAL)
    self.__img = ImageTk.PhotoImage(Image.open('MM1L.png')) 
    self.__Canva.create_image(150, 145, image=self.__img)
    self.__decoupage = "MM1L"

def decoup7(): 
    self.__valid1.config(state=tk.NORMAL)
    self.__img = ImageTk.PhotoImage(Image.open('TR00.png')) 
    self.__Canva.create_image(150, 145, image=self.__img)
    self.__decoupage = "TR00"

def decoup8(): 
    self.__valid1.config(state=tk.NORMAL)
    self.__img = ImageTk.PhotoImage(Image.open('TR10.png')) 
    self.__Canva.create_image(150, 145, image=self.__img)
    self.__decoupage = "TR10"

def decoup9(): 
    self.__valid1.config(state=tk.NORMAL)
    self.__img = ImageTk.PhotoImage(Image.open('TR11.png')) 
    self.__Canva.create_image(150, 145, image=self.__img)
    self.__decoupage = "TR11"


