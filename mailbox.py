import imaplib
import email
import datetime


def connexion(servername = None, username = None, password = None): 

    servername='outlook.office365.com'
    ORG_EMAIL = "@outlook.fr" 
    username = "test.pai3" + ORG_EMAIL 
    password = "Tomblanchard3."

    conn = imaplib.IMAP4_SSL(servername)
    conn.login(username,password)
    conn.select('Inbox')

    result, data = conn.uid('search', None, "UNSEEN") # (ALL/UNSEEN)
    i = len(data[0].split())
    
    return(i,data,conn)

