from tkinter import *
# from tkinter.ttk import *
from tkinter import messagebox

class Gui_relance_client():
    def __init__(self):
        super().__init__()
        # paramètre : 
        couleur = "#E4AB5B"

        # fenêtre
        self.window=Tk()
        menubar = Menu(self.window)
        menubar.add_cascade(label="Notice d'utilisation", command=self.help)
        self.window.title("INSTADOC")
        # self.window.geometry("808x700")
        self.window.resizable(0,0)
        self.window.config(menu=menubar)

        # self.window.iconbitmap("img/iconRC.ico")
        # self.window.config(background=couleur)

        # # background
        # self.Background_window = PhotoImage(file = 'img/bg.png')
        # self.ib = Label(self.window, background=couleur, image = self.Background_window)
        # self.ib.grid(x=-15,y=280)

        # # gestion du style des widget
        # style = Style()
        # style.configure("BW.TLabel", foreground="white", background=couleur, font=("Arial",12, ) )
        # style.configure("BC.TLabel", foreground="white", background=couleur, font=("Arial",9, ))
        # style.configure("BF.TLabel", foreground="white", background=couleur, font=("Arial",10, 'bold'))
        # style.configure("TButton", foreground="green", background="#ccc", padding=6)
        # style.map("TButton" ,foreground=[('pressed',"red"),('active','blue')] )
        # style.configure("TSeparator",  background="orange", borderwidth=2)

        # # Images
        # self.Logo_envoi_mail = PhotoImage(file= 'img/send_mail.png')
        # self.logo_excel = PhotoImage(file= 'img/excelbtn.png')

        # # Button
        # self.xl_all_dest =Button(self.window, text="Valider", compound = LEFT) #, image = self.logo_excel
        # self.xl_all_dest.grid(row =6,  column= 0 , columnspan=7)

        # Label : 
        lb_recherche = Label(self.window, text="Tri :", justify= LEFT)
        lb_recherche.grid(row=1, column=0)
        lb_millesime = Label(self.window, text="Millésime :", justify= LEFT)
        lb_millesime.grid(row=3, column= 0, columnspan=2)
        lb_Pc = Label(self.window, text="Compte comptable de rattachement", justify= LEFT)
        lb_Pc.grid(row=3, column=2 ,  columnspan=2)
        lb_Pc = Label(self.window, text="\ud83d\udd0e", justify= LEFT, font=('bold',15))
        lb_Pc.grid(row=1, column=7 )
        lb_Pc = Label(self.window, text="\u261c", justify= LEFT, font=('bold',20))
        lb_Pc.grid(row=2, column=7 )
        lb_Pc = Label(self.window, text="\ud83d\udd0e", justify= LEFT, font=('bold',15))
        lb_Pc.grid(row=3, column=7 )
        lb_Pc = Label(self.window, text="\u261C", justify= LEFT, font=('bold',20))
        lb_Pc.grid(row=4, column=7 )

        # Liste : 
        self.list_clt = Listbox(self.window,
                                            selectbackground=couleur,
                                            selectforeground='white',
                                            highlightbackground=couleur,
                                            highlightcolor=couleur,width=103 , height = 5
                                            )
        self.list_clt.grid(row=2, column=1, columnspan=5)
        self.list_millesime = Listbox(self.window,
                                            selectbackground=couleur,
                                            selectforeground='white',
                                            highlightbackground=couleur,
                                            highlightcolor=couleur,width=20 , height = 5
                                            )
        self.list_millesime.grid(row=4, column=0 , columnspan=2)
        self.list_plancomptable = Listbox(self.window,
                                            selectbackground=couleur,
                                            selectforeground='white',
                                            highlightbackground=couleur,
                                            highlightcolor=couleur,width=90 , height = 5
                                            )
        self.list_plancomptable.grid(row=4, column=2 , columnspan=4)


        # Saisie utilisateur
        self.saisie_dossier = Entry(self.window, width=103)
        self.saisie_dossier.grid(row=1,  column=1 , columnspan=6 )
        self.saisie_PCG = Entry(self.window, width=40)
        self.saisie_PCG.grid(row=3,  column=5 , columnspan=2 )



        # Messagebox : 
    def message_box(self, nb_relance):
        res = messagebox.askokcancel("Compte-rendu de l'envoie : ", f"""Nous avons envoyé : {nb_relance} mails de relances.


Voulez-vous quitter l'application ?""")
        return res


    def show(self):
        self.window.mainloop()

    def help(self):
        pass

if __name__ == "__main__":
    app = Gui_relance_client()
    app.show()
  