import pandas as pd
import pyodbc
import win32com.client as win32
import json
from ProjectClosing_Query import SQL

class SQLMailSender:

    def __init__(self):
        
        # roboczo do excela jeśli dane nie są pobierane bezpośrednio z query
        # self.get_data_to_excel()

        # deklaracja configu
        with open('config.json', 'r', encoding="utf-8") as f:
            self.config = json.load(f)

    def get_data_to_excel(self):
        sql = SQL()
        self.df = sql.get_data_from_SQL_to_DataFrame()

    def get_data_to_excel(self):
        sql = SQL()
        sql.get_data_from_SQL_to_excel()
        self.df = pd.read_excel("Sample_Query.xlsx", index_col=False)

    def send_prod_ord(self, bu):
        #odfiltrowanie wyników
        prod_ord_df = self.df.loc[
                                (self.df['Production_Orders'] != 0) & 
                                (self.df['Planned_Prod_Ord'] == 0) & 
                                (self.df['Project'].str.startswith(bu)) &
                                (self.df['Details'].str.contains('Active|Released|Completed', case=False)) &
                                (self.df['Checked_Area'] == 'Production Orders'), 
                                ['Project', 'ProjectDescription', 'PM', 'Details']
                                ]
        
        if prod_ord_df.empty:
            return
        else:
            for mail in self.config['mail']:
                if mail['function_name'] == f"send_prod_ord('{bu}')":
                    self._send_mail(prod_ord_df, subject=mail['subject'], text=mail['text'], send_to=mail['to'], send_cc=mail['cc'])
    

    def send_planned_prod_ord(self, bu):
        #odfiltrowanie wyników
        prod_ord_df = self.df.loc[
                            (self.df['Planned_Prod_Ord'] != 0) & 
                            (self.df['Project'].str.startswith(bu)) & 
                            (self.df['Details'].str.endswith('Planned')) & 
                            (self.df['Checked_Area'] == 'Production Orders'), 
                            ['Project', 'ProjectDescription', 'PM', 'Details']
                            ]
        
        if prod_ord_df.empty:
            return
        else:
            for mail in self.config['mail']:
                if mail['function_name'] == f"send_planned_prod_ord('{bu}')":
                    self._send_mail(prod_ord_df, subject=mail['subject'], text=mail['text'], send_to=mail['to'], send_cc=mail['cc'])


    def send_machine_MM_inv_WHR20_WHR50(self, bu):
        #odfiltrowanie wyników
        inv_df = self.df.loc[
                        (self.df['Production_Orders'] == 0) & 
                        (self.df['Project'].str.contains(bu)) & 
                        (self.df['Details'].str.contains('WHR20|WHR30|WHR40|WHR50', case=False)) &
                        (self.df['Checked_Area'] == 'Project Inventory'), 
                        ['Project', 'ProjectDescription', 'PM', 'Details']
                        ]

        if inv_df.empty:
            return
        else:
            split_df = inv_df['Details'].str.split('|', expand=True)
            split_df.columns = ['Whr', 'Item', 'Qty', 'Desc', 'Item_Group', 'SearchKey']
            split_df['Qty'] = split_df['Qty'].str.replace('Qty: ', '')

            inv_df = pd.concat([inv_df, split_df], axis=1)
            inv_df = inv_df[['Item', 'Desc', 'Qty', 'Whr', 'Item_Group', 'SearchKey', 'Project', 'ProjectDescription']]

            for mail in self.config['mail']:
                if mail['function_name'] == f"send_machine_MM_inv_WHR20_WHR50('{bu}')":
                    self._send_mail(inv_df, subject=mail['subject'], text=mail['text'], send_to=mail['to'], send_cc=mail['cc'])


    def send_machine_EM_inv_WHR20_WHR50(self, bu):
        #odfiltrowanie wyników
        inv_df = self.df.loc[
                        (self.df['Production_Orders'] == 0) & 
                        (self.df['Project'].str.contains(bu)) & 
                        (self.df['Details'].str.contains('WHR20|WHR30|WHR40|WHR50', case=False)) &
                        (self.df['Details'].str.contains(self.electrical, case=False)) &
                        (self.df['Checked_Area'] == 'Project Inventory'), 
                        ['Project', 'ProjectDescription', 'PM', 'Details']
                        ]

        if inv_df.empty:
            return
        else:
            split_df = inv_df['Details'].str.split('|', expand=True)
            split_df.columns = ['Whr', 'Item', 'Qty', 'Desc', 'Item_Group', 'SearchKey']
            split_df['Qty'] = split_df['Qty'].str.replace('Qty: ', '')

            inv_df = pd.concat([inv_df, split_df], axis=1)
            inv_df = inv_df[['Item', 'Desc', 'Qty', 'Whr', 'Item_Group', 'SearchKey', 'Project', 'ProjectDescription']]

            for mail in self.config['mail']:
                if mail['function_name'] == f"send_machine_EM_inv_WHR20_WHR50('{bu}')":
                    self._send_mail(inv_df, subject=mail['subject'], text=mail['text'], send_to=mail['to'], send_cc=mail['cc'])


    def send_machine_prp(self, bu):
        #odfiltrowanie wyników
        inv_df = self.df.loc[
                        (self.df['Production_Orders'] == 0) & 
                        (self.df['Project'].str.contains(bu)) & 
                        (self.df['Checked_Area'] == 'PRP Warehouse Orders'), 
                        ['Project', 'ProjectDescription', 'PM', 'Details']
                        ]

        if inv_df.empty:
            return
        else:
            for mail in self.config['mail']:
                if mail['function_name'] == f"send_machine_prp('{bu}')":
                    self._send_mail(inv_df, subject=mail['subject'], text=mail['text'], send_to=mail['to'], send_cc=mail['cc'])


    def send_service_inv_prp(self, bu):
        #odfiltrowanie wyników
        inv_df = self.df.loc[
                        (self.df['Production_Orders'] == 0) & 
                        (self.df['Project'].str.contains(bu)) & 
                        (self.df['Details'].str.contains('WHR20|WHR30|WHR40|WHR50|PRP0', case=False)) & 
                        (self.df['Checked_Area'].isin(['Project Inventory', 'PRP Warehouse Orders'])), 
                        ['Project', 'ProjectDescription', 'PM', 'Details']
                        ]

        if inv_df.empty:
            return
        else:
            for mail in self.config['mail']:
                if mail['function_name'] == f"send_service_inv_prp()":
                    self._send_mail(inv_df, subject=mail['subject'], text=mail['text'], send_to=mail['to'], send_cc=mail['cc'])


    def send_inv_WHR60_WHR70(self):
        #odfiltrowanie wyników
        inv_df = self.df.loc[
                        (self.df['Production_Orders'] == 0) & 
                        (self.df['Details'].str.contains('WHR60|WHR70', case=False)) & 
                        (self.df['Checked_Area'] == 'Project Inventory'), 
                        ['Project', 'ProjectDescription', 'PM', 'Details']
                        ]

        if inv_df.empty:
            return
        else:
            split_df = inv_df['Details'].str.split('|', expand=True)
            split_df.columns = ['Whr', 'Item', 'Qty', 'Desc', 'Item_Group', 'SearchKey']
            split_df['Qty'] = split_df['Qty'].str.replace('Qty: ', '')

            inv_df = pd.concat([inv_df, split_df], axis=1)
            inv_df = inv_df[['Item', 'Desc', 'Qty', 'Whr', 'Item_Group', 'SearchKey', 'Project', 'ProjectDescription', 'PM']]

            for mail in self.config['mail']:
                if mail['function_name'] == f"send_inv_WHR60_WHR70()":
                    self._send_mail(inv_df, subject=mail['subject'], text=mail['text'], send_to=mail['to'], send_cc=mail['cc'])

    def KAMIKAZE(self):
        # Wysłanie wszystkich maili na raz
        self.send_prod_ord('BU1')
        self.send_prod_ord('BU2')
        self.send_prod_ord('BU3')
        self.send_planned_prod_ord('BU1')
        self.send_planned_prod_ord('BU2')
        self.send_planned_prod_ord('BU3')
        self.send_machine_MM_inv_WHR20_WHR50('BU1')
        self.send_machine_MM_inv_WHR20_WHR50('BU2')
        self.send_machine_MM_inv_WHR20_WHR50('BU3')
        self.send_machine_EM_inv_WHR20_WHR50('BU1')
        self.send_machine_EM_inv_WHR20_WHR50('BU2')
        self.send_machine_EM_inv_WHR20_WHR50('BU3')
        self.send_machine_prp('BU1')
        self.send_machine_prp('BU2')
        self.send_machine_prp('BU3')
        self.send_service_inv_prp()
        self.send_inv_WHR60_WHR70()


    def _send_mail(self, table, subject, text, send_to, send_cc):
        # konwersja DataFrame na tabelę HTML
        html_table = table.to_html(index=False)
        
        # określenie stylu tabeli HTML
        styl = '''
        <style>
        table, th, td {
        border: 1px solid black;
        border-collapse: collapse;
        }
        th, td {
          white-space: nowrap;
          padding-left: 5px;
          padding-right: 5px;
          text-align: left;
        }
        </style>
        <table
        '''
        styled_table = html_table.replace('<table', styl)
        
        # nawiązanie połączenia z aplikacją Outlook
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        
        #send_to = 'us.er@ex.ample'
        #send_cc = 'us.er@ex.ample'
        # dodanie odbiorcy, tematu i treści e-maila
        mail.To = send_to
        mail.Cc = send_cc
        mail.Subject = subject
        mail.HTMLBody = f'<html><body><p>{text}</p>{styled_table}</body></html>'
        
        # wysłanie e-maila
        mail.Send()
    

#if __name__ == "__main__":
#
#    # Deklaracja obiektu klasy SQLMailSender
#    sender = SQLMailSender()
    