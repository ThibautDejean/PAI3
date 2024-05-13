import imaplib, email,datetime

def connexion(servername, ): 
    #gestion des mot de passe et user (introduire une table de hashage (voir double table pour plus de securite))
    ORG_EMAIL = "@outlook.fr" 
    usernm = "test.pai3" + ORG_EMAIL 
    passwd = "Tomblanchard3."
    conn = imaplib.IMAP4_SSL(servername)
    conn.login(usernm,passwd)
    conn.select('Inbox')
    result, data = conn.uid('search', None, "UNSEEN") # (ALL/UNSEEN)
    i = len(data[0].split())
    return(i,data,conn)