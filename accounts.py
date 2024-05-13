import imaplib, email,datetime
from tkinter import * 
import tkinter as tk
from tkinter import ttk
import openpyxl as xl
import sqlite3
import time
from pytz import timezone
import re
import winsound


def est_adresse_email(email):
         # Expression reguliere pour verifier si l'adresse e-mail est valide
         regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
         
         # Verifier si l'adresse e-mail correspond a  l'expression reguliere
         if re.match(regex, email):
             return True
         else:
             return False  
