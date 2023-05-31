import json
import tkinter as tk
from ProjectClosing_Procedure import SQLMailSender
from tkinter import END
from tkinter import PhotoImage

class SQLMailSenderGUI(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.sender = SQLMailSender() 
        self.master = master
        self.master.title("Project Closing (debug)")
        self.gear_img = PhotoImage(file="image/gear_img.png")

        self.create_widgets()
        
        with open('config.json', 'r', encoding="utf-8") as f:
            self.config = json.load(f)
    

    def create_widgets(self):
        
        # dodanie dużego labela z nazwą programu
        title_label = tk.Label(self.master, text="Project Closing (debug)", font=("Helvetica", 20, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=10)

        # stworzenie przycisków do pobrania danych z SQL
        self.button0 = tk.Button(self.master, text="Get Data from SQL", width=20, font=("Helvetica", 12, "bold"), command=lambda: self.sender.get_data_from_SQL())
        self.button0.grid(row=0, column=4, columnspan=2, padx=10, pady=10)

        # stworzenie etykiet kolumn
        self.label1 = tk.Label(self.master, text="Production Orders", font=("Helvetica", 12, "bold"))
        self.label1.grid(row=1, column=0)

        self.label2 = tk.Label(self.master, text="Project Inventory", font=("Helvetica", 12, "bold"))
        self.label2.grid(row=1, column=2)

        self.label3 = tk.Label(self.master, text="PRP & other details", font=("Helvetica", 12, "bold"))
        self.label3.grid(row=1, column=4)

        # stworzenie przycisków do wywołania funkcji obiektu sender
        self.button1 = tk.Button(self.master, text="Send Prod Orders BU1", width=30, command=lambda: self.sender.send_prod_ord('BU1'))
        self.button2 = tk.Button(self.master, text="Send Prod Orders BU2", width=30, command=lambda: self.sender.send_prod_ord('BU2'))
        self.button3 = tk.Button(self.master, text="Send Prod Orders BU3", width=30, command=lambda: self.sender.send_prod_ord('BU3'))
        self.button4 = tk.Button(self.master, text="Send Planned Prod Orders BU1", width=30, command=lambda: self.sender.send_planned_prod_ord('BU1'))
        self.button5 = tk.Button(self.master, text="Send Planned Prod Orders BU2", width=30, command=lambda: self.sender.send_planned_prod_ord('BU2'))
        self.button6 = tk.Button(self.master, text="Send Planned Prod Orders BU3", width=30, command=lambda: self.sender.send_planned_prod_ord('BU3'))
        self.button7 = tk.Button(self.master, text="Send Machine MM Inv W20-W50 BU1", width=30, command=lambda: self.sender.send_machine_MM_inv_WHR20_WHR50('BU1'))
        self.button8 = tk.Button(self.master, text="Send Machine MM Inv W20-W50 BU2", width=30, command=lambda: self.sender.send_machine_MM_inv_WHR20_WHR50('BU2'))
        self.button9 = tk.Button(self.master, text="Send Machine MM Inv W20-W50 BU3", width=30, command=lambda: self.sender.send_machine_MM_inv_WHR20_WHR50('BU3'))
        self.button10 = tk.Button(self.master, text="Send Machine EM Inv W20-W50 BU1", width=30, command=lambda: self.sender.send_machine_EM_inv_WHR20_WHR50('BU1'))
        self.button11 = tk.Button(self.master, text="Send Machine EM Inv W20-W50 BU2", width=30, command=lambda: self.sender.send_machine_EM_inv_WHR20_WHR50('BU2'))
        self.button12 = tk.Button(self.master, text="Send Machine EM Inv W20-W50 BU3", width=30, command=lambda: self.sender.send_machine_EM_inv_WHR20_WHR50('BU3'))
        self.button13 = tk.Button(self.master, text="Send Machine PRP BU1", width=30, command=lambda: self.sender.send_machine_prp('BU1'))
        self.button14 = tk.Button(self.master, text="Send Machine PRP BU2", width=30, command=lambda: self.sender.send_machine_prp('BU2'))
        self.button15 = tk.Button(self.master, text="Send Machine PRP BU3", width=30, command=lambda: self.sender.send_machine_prp('BU3'))
        self.button16 = tk.Button(self.master, text="Send Service Inv PRP", width=30, command=lambda: self.sender.send_service_inv_prp())
        self.button17 = tk.Button(self.master, text="Send Inv W60-W70", width=30, command=lambda: self.sender.send_inv_WHR60_WHR70())
        self.button18 = tk.Button(self.master, text="Wyślij wszystko!", width=50, font=("Helvetica", 12, "bold"), command=lambda: self.sender.KAMIKAZE())
        
        # stworzenie przycisków konfiguracji funkcji
        self.gear_button1 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_prod_ord('BU1')"))
        self.gear_button2 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_prod_ord('BU2')"))
        self.gear_button3 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_prod_ord('BU3')"))
        self.gear_button4 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_planned_prod_ord('BU1')"))
        self.gear_button5 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_planned_prod_ord('BU2')"))
        self.gear_button6 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_planned_prod_ord('BU3')"))
        self.gear_button7 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_machine_MM_inv_WHR20_WHR50('BU1')"))
        self.gear_button8 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_machine_MM_inv_WHR20_WHR50('BU2')"))
        self.gear_button9 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_machine_MM_inv_WHR20_WHR50('BU3')"))
        self.gear_button10 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_machine_EM_inv_WHR20_WHR50('BU1')"))
        self.gear_button11 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_machine_EM_inv_WHR20_WHR50('BU2')"))
        self.gear_button12 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_machine_EM_inv_WHR20_WHR50('BU3')"))
        self.gear_button13 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_machine_prp('BU1')"))
        self.gear_button14 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_machine_prp('BU2')"))
        self.gear_button15 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_machine_prp('BU3')"))
        self.gear_button16 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_service_inv_prp()"))
        self.gear_button17 = tk.Button(self.master, image=self.gear_img, width=20, height=20, command=lambda: self.open_config_window("send_inv_WHR60_WHR70()"))

        # ustawienie layoutu przycisków
        self.button1.grid(row=2, column=0, padx=(10, 0), pady=10)
        self.button2.grid(row=3, column=0, padx=(10, 0), pady=10)
        self.button3.grid(row=4, column=0, padx=(10, 0), pady=10)
        self.button4.grid(row=5, column=0, padx=(10, 0), pady=10)
        self.button5.grid(row=6, column=0, padx=(10, 0), pady=10)
        self.button6.grid(row=7, column=0, padx=(10, 0), pady=10)
        self.button7.grid(row=2, column=2, padx=(10, 0), pady=10)
        self.button8.grid(row=3, column=2, padx=(10, 0), pady=10)
        self.button9.grid(row=4, column=2, padx=(10, 0), pady=10)
        self.button10.grid(row=5, column=2, padx=(10, 0), pady=10)
        self.button11.grid(row=6, column=2, padx=(10, 0), pady=10)
        self.button12.grid(row=7, column=2, padx=(10, 0), pady=10)
        self.button13.grid(row=2, column=4, padx=(10, 0), pady=10)
        self.button14.grid(row=3, column=4, padx=(10, 0), pady=10)
        self.button15.grid(row=4, column=4, padx=(10, 0), pady=10)
        self.button16.grid(row=5, column=4, padx=(10, 0), pady=10)
        self.button17.grid(row=6, column=4, padx=(10, 0), pady=10)
        self.button18.grid(row=8, column=0, padx=10, pady=15, columnspan=5)
        
        # umieszczenie przycisku z kółkiem zębatym obok przycisku do wysłania maila
        self.gear_button1.grid(row=2, column=1)
        self.gear_button2.grid(row=3, column=1)
        self.gear_button3.grid(row=4, column=1)
        self.gear_button4.grid(row=5, column=1)
        self.gear_button5.grid(row=6, column=1)
        self.gear_button6.grid(row=7, column=1)
        self.gear_button7.grid(row=2, column=3)
        self.gear_button8.grid(row=3, column=3)
        self.gear_button9.grid(row=4, column=3)
        self.gear_button10.grid(row=5, column=3)
        self.gear_button11.grid(row=6, column=3)
        self.gear_button12.grid(row=7, column=3)
        self.gear_button13.grid(row=2, column=5, padx=(0, 10))
        self.gear_button14.grid(row=3, column=5, padx=(0, 10))
        self.gear_button15.grid(row=4, column=5, padx=(0, 10))
        self.gear_button16.grid(row=5, column=5, padx=(0, 10))
        self.gear_button17.grid(row=6, column=5, padx=(0, 10))


    def open_config_window(self, function):
        # odczytanie danych z configu
        for setting in self.config['mail']:
            if setting['function_name'] == function:
                # utworzenie okna konfiguracji
                config_window = tk.Toplevel(root)
                config_window.title("Mail Configuration")

                # dodanie pól tekstowych do wprowadzenia danych przez użytkownika
                to_label = tk.Label(config_window, text="To:")
                to_label.grid(row=0, column=0, sticky='e', padx=(10, 0), pady=1)
                to_entry = tk.Entry(config_window, width=100)
                to_entry.insert(0, setting['to'].strip())
                to_entry.grid(row=0, column=1, sticky='w', padx=(0, 10), pady=1)

                cc_label = tk.Label(config_window, text="Cc:")
                cc_label.grid(row=1, column=0, sticky='e', padx=(10, 0), pady=1)
                cc_entry = tk.Entry(config_window, width=100)
                cc_entry.insert(0, setting['cc'].strip())
                cc_entry.grid(row=1, column=1, sticky='w', padx=(0, 10), pady=1)

                subject_label = tk.Label(config_window, text="Subject:")
                subject_label.grid(row=2, column=0, sticky='e', padx=(10, 0), pady=1)
                subject_entry = tk.Entry(config_window, width=100)
                subject_entry.insert(0, setting['subject'].strip())
                subject_entry.grid(row=2, column=1, sticky='w', padx=(0, 10), pady=1)

                body_label = tk.Label(config_window, text="Text:")
                body_label.grid(row=3, column=0, sticky='e', padx=(10, 0), pady=1)
                body_text = tk.Text(config_window, wrap='word', width=75, height=10)
                body_text.insert('1.0', setting['text'].strip())
                body_text.grid(row=3, column=1, padx=(0, 10), pady=5)

                # dodanie przycisku do zapisania konfiguracji
                save_btn = tk.Button(config_window, text="Save", width=20, command=lambda: self.save_config(function,
                                                                                                            to_entry.get(),
                                                                                                            cc_entry.get(),
                                                                                                            subject_entry.get(),
                                                                                                            body_text.get(1.0, END)
                                                                                                            ))
                save_btn.grid(row=4, column=1, pady=5)


    def save_config(self, function, to='', cc='', subject='', text=''):
        for item in self.config['mail']:
            if item['function_name'] == function:
                item['to'] = to
                item['cc'] = cc
                item['subject'] = subject
                item['text'] = text[:-1]

            with open('config.json', 'w', encoding='utf-8') as fw:
                json.dump(self.config, fw, ensure_ascii=False, indent=4)


if __name__ == "__main__":
    root = tk.Tk()
    app = SQLMailSenderGUI(master=root)
    app.mainloop()
