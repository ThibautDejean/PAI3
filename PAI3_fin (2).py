### PAI N°3 : Programme python ###
### Importation des modules ###
import imaplib, email,datetime
from tkinter import * 
import tkinter as tk
from tkinter import ttk
import openpyxl as xl
import sqlite3
import time
from pytz import timezone
from PIL import ImageTk, Image
import re
import winsound

from mailbox import connexion
from alerts import iter_retard_avion
from accounts import est_adresse_email, new_id
from zones_sort import tri_geographique, decoup1, decoup2, decoup3, decoup4, decoup5, decoup6, decoup7, decoup8, decoup9
from database_updtae import plan_de_vol, message_arrive
from update_excel import ecriture_excel, message_changement, message_delai, message_annulation, message_depart, message_arrive


### Recuperation du corps des mails ###


### Interface Graphique (choix des parametres) ### 
       
class FenPrincipale(Tk):    

    def reconnaissance(self, corps):
        corps=corps.split("(")
        corps=corps[1].split('\n')

        ligne=corps[0].split('-')
        type_message=ligne[0].strip(' ')
        id_aeronef=ligne[1]
        decoupage = self.__decoupage

        if type_message=='FPL':
            plan_de_vol(corps,id_aeronef)
            ecriture_excel(corps,id_aeronef)
            tri_geographique(corps,id_aeronef,decoupage)
        elif type_message=='DLA':
            message_delai(corps,id_aeronef)
        elif type_message=='CHG':
            message_changement(corps,id_aeronef)
        elif type_message=='CNL':
            message_annulation(corps,id_aeronef)
        elif type_message=='DEP':
            message_depart(corps,id_aeronef)
        elif type_message=='ARR':
            message_arrive(corps,id_aeronef)
        elif type_message=='REFUS':
            self.message_refus(corps,id_aeronef)
        elif type_message=='ACP':
            self.message_acceptation(corps,id_aeronef)
        elif type_message=='SPL':
            self.plan_de_vol_complementaire(corps,id_aeronef)   
    
    ## interface graphique pour le traitement des mails re
    def quitter(self): 
        self.__fen_decoup.destroy()
        self.__fen_vols.destroy()

    def quitter_bis(self):
        self.__fen_new_id.destroy()   
        
    def create(self):
        # premiere fenetre :
          self.__fen_decoup = Toplevel(self, padx=2, pady=2)
          self.__fen_decoup.title("Plan de Vol")
           # La barre d'outils compose de 2 boutons :
          self.__BO1 = tk.Frame(self.__fen_decoup)
          self.__BO1.pack(side=tk.BOTTOM,padx=5, pady=5)
          self.__valid1=tk.Button(self.__BO1,text='Valider')
          self.__valid1.pack(side=tk.LEFT, padx=5, pady=5)
          self.__valid1.config(state=tk.DISABLED)
          self.__valid1.config(command=self.nouvelle_fen)
          self.__BO2=tk.Frame(self.__fen_decoup,borderwidth=2,bg='white')
          self.__BO2.pack(side=tk.TOP,padx=5,pady=2)
         # Configuration du Label de l'en-tete qui sert a  donner des indications
          self.__label_enTete1=tk.Label(self.__BO2,text="Definissez le decoupage du territoire :", bg='white')
          self.__label_enTete1.config(font=("Arial", 20, "underline"))
          self.__label_enTete1.pack(side=tk.LEFT, padx=20,pady=8) 
         # Configuration du choix du decoupage du territoire
          self.__menu_zone = tk.Menubutton ( self.__BO2 , text = "Choix du decoupage")
          self.__menu_zone.pack(side=tk.LEFT, padx=5, pady=5)
          self.__option = tk.Menu ( self.__menu_zone )
          self.__option.add_command ( label = "Plan NORM" , command = self.decoup1)
          self.__option.add_command ( label = "Plan LY00" , command = self.decoup2)
          self.__option.add_command ( label = "Plan LY1T" , command = self.decoup3)
          self.__option.add_command ( label = "Plan LY10" , command = self.decoup4)
          self.__option.add_command ( label = "Plan LY11" , command = self.decoup5)
          self.__option.add_command ( label = "Plan MM1L" , command = self.decoup6)
          self.__option.add_command ( label = "Plan TR00" , command = self.decoup7)
          self.__option.add_command ( label = "Plan TR10" , command = self.decoup8)
          self.__option.add_command ( label = "Plan TR11" , command = self.decoup9)
          self.__menu_zone [ "menu" ] = self.__option
          # Le canvas pour afficher le plan du decoupage
          self.__Canva =tk.Canvas(self.__fen_decoup, width = 300,height = 290,bg='white')  
          self.__Canva.pack()
          
         # deuxieme fenetre :
          self.__fen_vols=tk.Toplevel(self)
          self.__fen_vols.title("Vol en Cours")
          self.__fen_vols.geometry("1355x700")
          self.__fen_vols.geometry("+0+0")
          self.__BO3 = tk.Frame(self.__fen_vols)
          self.__BO3.pack(side=tk.BOTTOM, padx=5, pady=5)
          self.__Quit1 = tk.Button(self.__BO3, text ='Quitter',command= self.quitter,activeforeground = "blue",activebackground = "yellow", width=13)
          self.__Quit1.pack(side=tk.RIGHT, padx=5, pady=5)
          # Configuration du label de l'en-tete qui sert a  donner des indications
          self.__BO4=tk.Frame(self.__fen_vols,borderwidth=2,bg='white')
          self.__BO4.pack(side=tk.TOP,padx=5,pady=2)
          self.__label_enTete2=tk.Label(self.__BO4,text="Trafic aerien actuel : ", bg='white',fg="black",font=("Arial",15))
          self.__label_enTete2.pack(side=tk.LEFT, padx=20,pady=8) 
 
    def create_bis(self) :
        self.__fen_new_id = tk.Toplevel(self,padx=130,pady=100)
        self.__fen_new_id.title("Nouvel Identifiant")
        self.__info = tk.Frame(self.__fen_new_id)
        
        self.__prenom_label = tk.Label(self.__info, text="Prenom_militaire")
        self.__prenom_label.pack()
        self.__prenom_entry = tk.Entry(self.__info)
        self.__prenom_entry.pack()
       
        self.__nom_label = tk.Label(self.__info, text="Nom_militaire")
        self.__nom_label.pack()
        self.__nom_entry = tk.Entry(self.__info)
        self.__nom_entry.pack()
        
        self.__mail_label = tk.Label(self.__info, text="Mail_militaire")
        self.__mail_label.pack()
        self.__mail_entry = tk.Entry(self.__info)
        self.__mail_entry.pack()
        
        self.__id_bis_label = tk.Label(self.__info, text="id_militaire")
        self.__id_bis_label.pack()
        self.__id_bis_entry = tk.Entry(self.__info)
        self.__id_bis_entry.pack()
        
        self.__passwordnew_label = tk.Label(self.__info, text="Mot de passe")
        self.__passwordnew_label.pack()
        self.__passwordnew_entry = tk.Entry(self.__info)
        self.__passwordnew_entry.pack()
        self.__info.pack()
        
        
        self.__passwordnew_confirmation_label = tk.Label(self.__info, text="Confirmation mot de passe")
        self.__passwordnew_confirmation_label.pack()
        self.__passwordnew_confirmation_entry = tk.Entry(self.__info)
        self.__passwordnew_confirmation_entry.pack()
        
        self.__textAffiche = tk.StringVar()
        self.__textAffiche.set("veuillez vous identifier")
        self.__message_erreur= tk.Label(self.__info, textvariable=self.__textAffiche, font=('Times', '16', 'bold'),fg="blue")
        self.__message_erreur.pack()
        
       
        self.__boutonValider2=tk.Button(self.__info,text='Valider',command=self.new_id)
        self.__boutonValider2.pack()
        
        self.__Quit2 = tk.Button(self.__info, text ='Quitter',command=self.quitter_bis,activeforeground = "blue",activebackground = "yellow", width=13)
        self.__Quit2.pack(side=tk.BOTTOM, padx=5, pady=5)
        self.__Quit2.config(state=tk.DISABLED)

                
    def new_id(self):
        current_time = str(datetime.datetime.now())
        prenom = self.__prenom_entry.get()
        nom = self.__nom_entry.get()
        id_bisg = self.__id_bis_entry.get()
        mail = self.__mail_entry.get()
        password = self.__passwordnew_entry.get()
        passwordnew_confirmation = self.__passwordnew_confirmation_entry.get()
        prenom=str(prenom)
        nom=str(nom)
        id_ =str(id_bisg)
        mail=str(mail)
        mdp = str(password)
        mdp_c= str(passwordnew_confirmation)
        self.__conn = sqlite3.connect('id_pai3.db')
        curseur = self.__conn.cursor()
        curseur.execute("SELECT id FROM mdp WHERE id = '{}'".format(id_bisg.strip()))
        liste = curseur.fetchall()
        self.__conn.close()
        if len(liste)!=0: 
            self.__textAffiche.set("Vous avez deja  un identifiant : veuillez demander a verifier dans la base de donnee ou creez un nouvel identifiant")
            return None
        if prenom==None or nom==None or id_bisg==None or mail==None or password==None :
            self.__textAffiche.set("Vous avez oubliez de remplir une case")
            return None
            
        if len(prenom)==0 or  len(nom)==0 or len(id_bisg)==0 or len(mail)==0 or len(password)==0 :
            self.__textAffiche.set("Vous avez oublie de remplir une case")
            return None
        if mdp_c != mdp:
            self.__textAffiche.set("Erreur mot de passe et confirmation mot de passe differents")
            return None
        try: 
            if self.est_adresse_email(mail)==False:
                raise TypeError
                return(None) 
        except TypeError:
            self.__textAffiche.set("Erreur dans la saisie de l'adresse mail")
            return None
        try: 
            if nom.isdigit() or prenom.isdigit():
                raise TypeError
                return(None) 
        except TypeError:
            self.__textAffiche.set("Il ne faut pas de chiffre dans le nom ou le prenom")
            return None

        else : 
            self.__conn = sqlite3.connect('id_pai3.db')
            curseur = self.__conn.cursor()
            curseur.execute("INSERT INTO creation_id(prenom, nom, mail,id,mot_de_passe,date_creation) VALUES ('{}','{}','{}','{}','{}','{}')".format(prenom.strip(), nom.strip(), mail.strip(), id_.strip(), mdp.strip(),current_time.strip()))
            curseur.execute("INSERT INTO mdp(id, mot_de_passe) VALUES ('{}', '{}')".format(id_.strip(), mdp.strip()))
            self.__conn.commit()
            self.__textAffiche.set("Vous êtes enregistrés comme utilisateur")
            self.__conn.close()
            self.__boutonValider2.config(state=tk.DISABLED)
            self.__Quit2.config(state=tk.NORMAL)
            

            

        ## fenetre d'identification ## 
    def get_id_db(self): 
            self.__conn = sqlite3.connect('id_pai3.db')
            curseur = self.__conn.cursor()
            username = str(self.__username_entry.get())
            password = str(self.__password_entry.get())
            curseur.execute("SELECT id,mot_de_passe From  mdp where id='{}' AND mot_de_passe='{}' ".format(username.strip(), password.strip()))
            liste = curseur.fetchall()
            self.__conn.close()
            if len(liste) != 0 :
                return True
            else: 
                return False
            
    def login(self):
        current_time = str(datetime.datetime.now())
        username = str(self.__username_entry.get())
        # Verification des informations d'identification ici

        if self.get_id_db():
            self.__message_label.config(text="Connexion reussie !")
            self.__login_button.config(state=tk.DISABLED)
            self.__logout_button.config(state=tk.NORMAL)
            self.create()
        else:
            self.__message_label.config(text="Nom d'utilisateur ou mot de passe incorrect")
        self.__conn = sqlite3.connect('id_pai3.db')
        curseur = self.__conn.cursor()
        curseur.execute("INSERT INTO tab_debut_connexion(id,debut_connexion) VALUES ('{}','{}')".format(username.strip(),current_time.strip()))
        self.__conn.commit()
        self.__conn.close()
            
    def logout(self):
            current_time = str(datetime.datetime.now())
            self.__conn = sqlite3.connect('id_pai3.db')
            curseur = self.__conn.cursor()
            username = str(self.__username_entry.get())
            curseur.execute("INSERT INTO tab_fin_connexion(id,fin_connexion) VALUES ('{}','{}')".format(username.strip(),current_time.strip()))
            # se connecter a  la base de donnee

            # executer la requete avec un numero de connexion donne (remplacer 1 par le numero de connexion souhaite)
            curseur.execute("""INSERT INTO Historique_connexion (id, debut_connexion, fin_connexion)
                            SELECT id, debut_connexion, fin_connexion
                            FROM (
                                SELECT id, debut_connexion, NULL AS fin_connexion
                                FROM tab_debut_connexion
                                ORDER BY debut_connexion DESC
                                LIMIT 1
                                ) AS t1
                            UNION ALL
                            SELECT id, NULL AS debut_connexion, fin_connexion
                            FROM (
                                SELECT id, NULL AS debut_connexion, fin_connexion
                                FROM tab_fin_connexion
                                ORDER BY fin_connexion DESC
                                LIMIT 1
                                ) AS t2;
                            """)
            # valider les changements et fermer la connexion
            self.__conn.commit()
            self.__conn.close()
            self.__message_label.config(text="deconnexion reussie !")
            self.__login_button.config(state=tk.NORMAL)
            self.___logout_button.config(state=tk.DISABLED)
            self.quitter()
            
    def __init__(self,base):
            global etat
            etat=True 
            
            # Configuration de la base de donnees
            self.__conn = sqlite3.connect(base)
            # on est connecte a  la base de donnees
            ##Pour executer une requete :
            # curseur.execute("")
            # curseur.fetchall()[0][0]
            
        ## Mise en place de l'interface graphique de la fenetre d'identification  
            Tk.__init__(self)
            self.title("fenetre d'identification")
            self.geometry('400x200')
            
            # Creation des champs de saisie
            self.__username_label = tk.Label(self, text="Nom d'utilisateur")
            self.__username_label.pack()
            self.__username_entry = tk.Entry(self)
            self.__username_entry.pack()
            
            self.__password_label = tk.Label(self, text="Mot de passe")
            self.__password_label.pack()
            self.__password_entry = tk.Entry(self, show="*")
            self.__password_entry.pack()
            
            # Creation du bouton de validation
            self.__login_button = tk.Button(self, text="Se connecter", command=self.login)
            self.__login_button.pack()
            # Creation du bouton de deconnexion 
            self.__logout_button = tk.Button(self, text="Se deconnecter", command=self.logout)
            self.__logout_button.pack()
            self.__logout_button.config(state=tk.DISABLED)
            #bouton en cas d'oubli de mot de passe
            self.__oubli_button = tk.Button(self, text="mot de passe oublie", command=self.create_bis)
            self.__oubli_button.pack()
            # Creation d'une etiquette pour afficher le message de connexion
            self.__message_label = tk.Label(self, text="")
            self.__message_label.pack()

    
            
    ### Interface Graphique (affichage des vols en cours) ###

    def nouvelle_fen(self):
            global etat
            #Fermeture de la premiere fenetre
            self.withdraw()
            # self._fen.deiconify()
            self.lecture_mail()
            
            
        ### definnir un cycle qui relance la fonction a  un intervalle de temps precis afin de lire les nouveaux mails.
    def lecture_mail(self):
            global etat
            servername='outlook.office365.com'
            self.__fen_vols.deiconify()
            ### Lecture des mails ###
            start_time = time.time()
            interval = 10 #on recupere des nouveaux mails toutes les 10 secondes
            j=0
            while etat==True:
                j+=interval
                time.sleep(start_time + j - time.time())
                (i,data,conn)=connexion(servername)
                self.mail(i,data,conn)
                iter_retard_avion()
        
    def mail(self,i,data,conn):
            #On parcours les mails 1 par 1.
            for x in range(i):
                latest_email_uid = data[0].split()[x]
                result, email_data = conn.uid('fetch', latest_email_uid, '(RFC822)')
                # result, email_data = conn.store(num,'-FLAGS','\\Seen') 
                # this might work to set flag to seen, if it doesn't already
                raw_email = email_data[0][1]
                raw_email_string = raw_email.decode('utf-8')
                email_message = email.message_from_string(raw_email_string)
            
                ### informations sur le sujet du mail :
                # Header Details
                # date_tuple = email.utils.parsedate_tz(email_message['Date'])
                # if date_tuple:
                #     local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                #     local_message_date = "%s" %(str(local_date.strftime("%a, %d %b %Y %H:%M:%S")))
                # email_from = str(email.header.make_header(email.header.decode_header(email_message['From'])))
                # email_to = str(email.header.make_header(email.header.decode_header(email_message['To'])))
                # subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))
            
                # Body details
                for part in email_message.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True)
                        corps=body.decode('UTF-8')
                        print(corps)
                        self.reconnaissance(corps)
                        # on obtient de chaine de caractere qu'il faut maintenan traiter.
                        
                        ###Possibilite de stocker chaque mail dans un fichier txt : 
                        # file_name = "email_" + str(x) + ".txt"
                        # output_file = open(file_name, 'w')
                        # output_file.write("From: %s\nTo: %s\nDate: %s\nSubject: %s\n\nBody: \n\n%s" %(email_from, email_to,local_message_date, subject, body.decode('utf-8')))
                        # output_file.close()
                        
                    else:
                        continue
            print("mails traites, pas de mails non lus")
                    
            
 
                

if __name__ == '__main__':

    fen = FenPrincipale("nom_base.db")
    fen.mainloop()  