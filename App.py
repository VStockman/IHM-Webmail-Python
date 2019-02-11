from tkinter import *
import requests
from html.parser import HTMLParser
import re
import tkinter.ttk as ttk
from email.mime.text import MIMEText
import smtplib
import sys
mail = []
host = 'smtp.gmail.com'
server = smtplib.SMTP(host, 587)
server.ehlo()
server.starttls()
server.ehlo()
server.login("pythonwebmail@gmail.com", "itescia2019")

#fenêtre principale, contenant les boutons ci-dessous
def principalwindow(oldwindow, mail):
    newwindow = Tk()
    oldwindow.destroy()
    liste = Listbox(newwindow)
    nb = len(mail)
    for i in range(0, nb):
        liste.insert(i, mail[i])
    liste.pack()
    frame3 = Frame(newwindow, borderwidth=2, relief=GROOVE)
    frame3.pack(side=LEFT, padx=30, pady=30)
    frame4 = Frame(newwindow, borderwidth=2, relief=GROOVE)
    frame4.pack(side=LEFT, padx=10, pady=10)
    frame5 = Frame(newwindow, borderwidth=0, relief=GROOVE)
    frame5.pack(side=BOTTOM, padx=5, pady=5)
    bouton13 = Button(newwindow, text='Supprimer',  command=lambda: deletewindow(mail, str(liste.curselection())[1], newwindow))
    bouton13.pack(side=BOTTOM)
    bouton2 = Button(frame3, text='Dédoublon', command=lambda: deduplication(mail, newwindow))
    bouton2.pack()
    bouton3 = Button(frame4, text='Vérification', command=lambda: verifymail(mail, newwindow))
    bouton3.pack()
    bouton4 = Button(frame3, text='Import CSV', command=lambda: importcsv(mail,newwindow))
    bouton4.pack()
    bouton5 = Button(frame4, text='Import URL', command=lambda: importurl(mail,newwindow))
    bouton5.pack()
    bouton6 = Button(frame5, text='OK', command=lambda: mailsettings(mail, newwindow))
    bouton6.pack()


#fonction contenant la fenêtre d'envoi de mail avec l'expediteur, l'objet, et le message
def mailsettings(mails,oldwindow11):
    if len(mail) != 0:
        oldwindow11.destroy()
        liste = []
        email = list(set(mails))
        nb = len(mails)
        for i in range(0, nb):
            if "@" in mails[i] and (email[i].endswith('.fr') or email[i].endswith('.com')):
                liste.insert(i, email[i])
        class Mailer(Tk):
            def __init__(self, mail):
                Tk.__init__(self)
                self.geometry("700x520")
                self.mail = mail
                self.liste_couleurs = ["Noir", "Rouge", "Vert", "Bleu"]
                self.liste_couleurs_resolve = ["black", "red", "green", "blue"]
                self.couleur = self.liste_couleurs_resolve[0]
                ## TOP = entêtes ##
                fTop = Frame(self)
                top = PanedWindow(fTop)
                Label(top, text="Expéditeur").grid(row=0, column=0)
                self.expediteur = Entry(top)
                self.expediteur.insert(INSERT, "")
                self.expediteur.get()
                self.expediteur.grid(row=0, column=1)
                Label(top, text="Objet").grid(row=1, column=0)
                self.objet = Entry(top)
                self.objet.insert(INSERT, "")
                self.objet.grid(row=1, column=1)
                self.objet.get()
                top.pack(side=TOP)
                fTop.pack(side=TOP)

                ## BOTTOM = corps ##
                fBot = PanedWindow(self)

                Label(fBot, text="Message :").pack()
                self.message = Text(fBot, wrap=WORD, undo=TRUE)
                self.message.insert(INSERT, "")
                self.message.pack(side=BOTTOM)

                bot = PanedWindow(fBot)
                Button(bot, text='B', command=lambda: self.actionListener_widget('gras')).pack(side=LEFT)
                Button(bot, text="I", command=lambda: self.actionListener_widget('italique')).pack(side=LEFT)
                Button(bot, text="S", command=lambda: self.actionListener_widget('souligner')).pack(side=LEFT)
                Button(bot, text="Couleur", command=lambda: self.actionListener_widget('couleur')).pack(side=LEFT)
                cb = ttk.Combobox(bot, values=(
                self.liste_couleurs[0], self.liste_couleurs[1], self.liste_couleurs[2], self.liste_couleurs[3]),
                                  state='readonly', width=8)
                cb.set(self.liste_couleurs[0])
                cb.pack(side=LEFT)
                cb.bind('<<ComboboxSelected>>', self.actionListener_couleur)
                Button(bot, text="Redo", command=self.message.edit_redo).pack(side=RIGHT)
                Button(bot, text="Undo", command=self.message.edit_undo).pack(side=RIGHT)
                bot.pack(side=TOP, fill=BOTH)

                fBot.pack(side=TOP)

                ## BOUTON continuer ##
                fFooter = Frame(self)
                Button(fFooter, text="Continuer", command=lambda: sendmail(liste, self.expediteur.get(), self.objet.get(), self.message.get('1.0', END), mailwindow)).pack()
                fFooter.pack(side=BOTTOM)

            def actionListener_couleur(self, event=None):
                if event:  # <-- this works only with bind because `command=` doesn't send event
                    self.couleur = self.liste_couleurs_resolve[
                        self.liste_couleurs.index(event.widget.get())]  # self.actionListener_widget('couleur')

            def actionListener_widget(self, context):
                index = self.getIndex()
                if (context == 'gras'):
                    self.message.insert(index[1], '</b>')
                    self.message.insert(index[0], '<b>')
                elif (context == 'italique'):
                    self.message.insert(index[1], '</i>')
                    self.message.insert(index[0], '<i>')
                elif (context == 'souligner'):
                    self.message.insert(index[1], '</u>')
                    self.message.insert(index[0], '<u>')
                elif (context == 'couleur'):
                    self.message.insert(index[1], '</font>')
                    self.message.insert(index[0], '<font color="' + self.couleur + '">')

            # retourne True si du texte a été sélectionné, False sinon
            def isSelection(self):
                try:
                    print(self.message.selection_own())
                    return self.message.selection_get() != ''
                except:
                    return False

            # retourne les coordonnées de la sélection
            def getSelectionIndexs(self):
                return [self.message.index(SEL_FIRST), self.message.index(SEL_LAST)]

            # retourne les coordonnées de la sélection, ou bien du curseur s'il n'y a pas de sélection
            def getIndex(self):
                if self.isSelection():
                    return self.getSelectionIndexs()
                else:
                    return [self.message.index(INSERT), self.message.index(INSERT)]
        mailwindow = Mailer(Tk)
    else:
        errormailwindow(oldwindow11, mail)


#fonction appelée lors du clic sur continuer de la fenêtre des paramètres du mail, et permet d'envoyer les mails aux différentes adresses, et envoie vers la fenêtre de fin d'envoi de mails
def sendmail(mail, expediteur, object, message,lastwindow):
    nb = len(mail)
    msg = MIMEText(message)
    msg['Subject'] = object
    msg['From'] = expediteur
    for i in range(0, nb):
        msg['To'] = mail[i]
        server.sendmail(expediteur, mail[i], msg.as_string())
        print('Email sent at ', mail[i])
    server.quit()
    successwindow(lastwindow)


#fenêtre appelée juste après l'envoi de mails
def successwindow(oldwindow15):
    oldwindow15.destroy()
    window14 = Tk()
    frame24 = Frame(window14, borderwidth=5)
    frame24.pack(side=LEFT, padx=30, pady=30)
    Label(frame24, text="Mails envoyés avec succès !",fg="green").pack(padx=10, pady=10)
    frame25 = Frame(frame24, bg="white", borderwidth=0, relief=GROOVE)
    frame25.pack(side=RIGHT, padx=0, pady=0)
    bouton15 = Button(frame25, text='OK', command=lambda: sys.exit())
    bouton15.pack()


#fenêtre d'erreur apparaissant si on clique sur le bouton OK de la fenêtre principale alors qu'il n'ya aucune adresse mail dans la liste
def errormailwindow(oldwindow12,mail):
    window13 = Tk()
    oldwindow12.destroy()
    frame22 = Frame(window13, borderwidth=5)
    frame22.pack(side=LEFT, padx=30, pady=30)
    Label(frame22,text="Veuillez avoir au moins une adresse mail dans la liste afin d'envoyer un mail !",fg="red").pack(padx=10, pady=10)
    frame23 = Frame(frame22, bg="white", borderwidth=0, relief=GROOVE)
    frame23.pack(side=RIGHT, padx=0, pady=0)
    bouton9 = Button(frame23, text='OK', command=lambda: principalwindow(window13, mail))
    bouton9.pack()


#fonction d'import du CSV
def importcsv(mail,oldwindow2):
    window2 = Tk()
    oldwindow2.destroy()
    frame6 = Frame(window2, borderwidth=5)
    frame6.pack(side=LEFT, padx=30, pady=30)
    Label(frame6, text="Tapez un nom de fichier csv à importer :").pack(padx=10, pady=10)
    frame7 = Frame(frame6, bg="white", borderwidth=0, relief=GROOVE)
    frame7.pack(side=RIGHT, padx=0, pady=0)
    value2 = StringVar()
    entree2 = Entry(frame6, textvariable=value2, width=30)
    entree2.pack()
    entree2.get()
    bouton7 = Button(frame7, text='OK', command=lambda: checkcsv(mail,value2.get(), window2, 'false'))
    bouton7.pack()


#classe permettant de parser du HTML, utilisée lors de l'import URL
class MyHTMLParser(HTMLParser):
    def handle_starttag(self, tag, attrs):
        print(tag)

    def handle_endtag(self, tag):
        print(tag)

    def handle_data(self, data):
        print(data)


#fonction permettant de ne garder que les adresses mails correctes, appelée lors du clic sur le bouton vérifier sur la fenêtre principake
def verifymail(email, oldwindow9):
    liste = []
    email = list(set(email))
    nb = len(email)
    for i in range(0, nb):
        if "@" in email[i] and (email[i].endswith('.fr') or email[i].endswith('.com')):
            liste.insert(i, email[i])
    window9 = Tk()
    oldwindow9.destroy()
    frame16 = Frame(window9, borderwidth=5)
    frame16.pack(side=LEFT, padx=30, pady=30)
    Label(frame16, text="Vérification terminée !").pack(padx=10, pady=10)
    frame17 = Frame(frame16, bg="white", borderwidth=0, relief=GROOVE)
    frame17.pack(side=RIGHT, padx=0, pady=0)
    bouton12 = Button(frame17, text='OK', command=lambda: principalwindow(window9, liste))
    bouton12.pack()


#fonction de vérification de l'URL saisie, qui est appelée juste après avoir cliqué sur le bouton OK après avoir saisi une URL
def checkurl(mail, window4, url):
    if (url.startswith("http://") or url.startswith("https://")) and (url.endswith(".com") or url.endswith(".fr") or url.endswith(".html")):
        html = requests.get(url)
        result = re.findall('mailto:[A-Za-z]+.[A-Za-z]+.[A-Za-z]+.[A-Za-z]+.[a-z]+', html.text)
        for i in result:
            result[result.index(i)] = i.replace('mailto:', '')
        nb = len(result)
        for i in range(0, nb):
            mail.append(result[i])
        principalwindow(window4, mail)
    else:
        window5 = Tk()
        window4.destroy()
        frame10 = Frame(window5, borderwidth=5)
        frame10.pack(side=LEFT, padx=30, pady=30)
        Label(frame10,
              text="Tapez une URL valide s'il vous plaît ! Commençant par http/https et terminant par .fr/.com",
              fg="red").pack(padx=10, pady=10)
        frame11 = Frame(frame10, bg="white", borderwidth=0, relief=GROOVE)
        frame11.pack(side=RIGHT, padx=0, pady=0)
        bouton10 = Button(frame11, text='OK', command=lambda: importurl(mail, window5))
        bouton10.pack()


#fonction d'import de l'URL, qui est appelée lors du clic sur Import URL de ela fenêtre principale
def importurl(mail,oldwindow3):
    window3 = Tk()
    oldwindow3.destroy()
    frame8 = Frame(window3, borderwidth=5)
    frame8.pack(side=LEFT, padx=30, pady=30)
    Label(frame8, text="Tapez une URL valide à importer :").pack(padx=10, pady=10)
    frame9 = Frame(frame8, bg="white", borderwidth=0, relief=GROOVE)
    frame9.pack(side=RIGHT, padx=0, pady=0)
    value3 = StringVar()
    entree3 = Entry(frame8, textvariable=value3, width=30)
    entree3.pack()
    entree3.get()
    bouton8 = Button(frame9, text='OK', command=lambda: checkurl(mail,window3, value3.get()))
    bouton8.pack()


#fonction qui permet de déduplication des adresses mails, appelée lors du clic sur dédoublon sur la fenêtre principale
def deduplication(email,oldwindow8):
    email = list(set(email))
    window8 = Tk()
    oldwindow8.destroy()
    frame14 = Frame(window8, borderwidth=5)
    frame14.pack(side=LEFT, padx=30, pady=30)
    Label(frame14, text="Dédoublonnage effectué !").pack(padx=10, pady=10)
    frame15 = Frame(frame14, bg="white", borderwidth=0, relief=GROOVE)
    frame15.pack(side=RIGHT, padx=0, pady=0)
    bouton11 = Button(frame15, text='OK', command=lambda: principalwindow(window8, email))
    bouton11.pack()


#fonction de vérification du fichier sur l'espace de stockage, après la saisie du potentiel fichier, avec un cas spécial si le premier fichier saisi n'existe pas, avec la variable createfile passée en paramètre
def checkcsv(mail, filename, oldwindow7, createfile):
    try:
        with open(filename, "r"):
            readfile(mail,filename, oldwindow7)
    except FileNotFoundError:
        if createfile == 'true':
            file = open(filename, "w")
            file.write("")
            file.close()
            principalwindow(window, mail)
        else:
            window6 = Tk()
            oldwindow7.destroy()
            frame12 = Frame(window6, borderwidth=5)
            frame12.pack(side=LEFT, padx=30, pady=30)
            Label(frame12, text="Veuillez sélectionner un fichier présent sur votre espace de stockage s'il vous plaît !",
                  fg="red").pack(padx=10, pady=10)
            frame13 = Frame(frame12, bg="white", borderwidth=0, relief=GROOVE)
            frame13.pack(side=RIGHT, padx=0, pady=0)
            bouton9 = Button(frame13, text='OK', command=lambda: importcsv(mail, window6))
            bouton9.pack()


#fonction de lecture du fichier si le fichier existe bien, après la fonction checkCSV
def readfile(mail, filename, oldwindow7):
    file = open(filename, "r")
    openfile = file.read()
    file.close()
    openfile = list(openfile.split())
    nb = len(openfile)
    for i in range(0, nb):
        mail.append(openfile[i])
    principalwindow(oldwindow7, mail)


#fonction appelée lors du clic sur le bouton supprimer de la fenêtre principale, qui supprime de la liste l'adresse mail sélectionée dans la liste
def deletewindow(mail, selectedmail, oldwindow10):
    if selectedmail != ')':
        selectedmail = int(selectedmail)
        window10 = Tk()
        oldwindow10.destroy()
        frame20 = Frame(window10, borderwidth=5)
        frame20.pack(side=LEFT, padx=30, pady=30)
        Label(frame20, text="Suppression de l'adresse mail " + mail[selectedmail] + ' effectuée !').pack(padx=10, pady=10)
        mail.pop(selectedmail)
        frame21 = Frame(frame20, bg="white", borderwidth=0, relief=GROOVE)
        frame21.pack(side=RIGHT, padx=0, pady=0)
        bouton14 = Button(frame21, text='OK', command=lambda: principalwindow(window10, mail))
        bouton14.pack()


window = Tk()
label = Label(window, text="NomCampagne")
label.pack()
frame1 = Frame(window, borderwidth=5)
frame1.pack(side=LEFT, padx=30, pady=30)
Label(frame1, text="Tapez un nom de fichier csv à analyser :").pack(padx=10, pady=10)
frame2 = Frame(frame1, bg="white", borderwidth=0, relief=GROOVE)
frame2.pack(side=RIGHT, padx=0, pady=0)
value = StringVar()
entree = Entry(frame1, textvariable=value, width=30)
entree.pack()
entree.get()
bouton = Button(frame2, text='OK', command=lambda: checkcsv(mail, value.get(), window, 'true'))
bouton.pack()
window.mainloop()
