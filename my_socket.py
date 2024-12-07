import websocket
from threading import Thread
import time
import json
import pandas as pd
from datetime import datetime, timedelta

forts = {
    7: {"level": 60, "defense": False},
    8: {"level": 70, "defense": False},
    9: {"level": 80, "defense": False},
    10: {"level": 40, "defense": True},
    11: {"level": 50, "defense": True},
    12: {"level": 60, "defense": True},
    13: {"level": 70, "defense": True},
    14: {"level": 80, "defense": True}
}


class MySocket(websocket.WebSocketApp):
    def __init__(self, window, url, serveur_header, royaume, nom, mdp, filepath, intervalle):
        super().__init__(url, on_open=self.on_open, on_message=self.on_message, on_error=self.on_error, on_close=self.on_close)
        self.window = window
        self.serveur_header = serveur_header
        self.royaume = royaume
        self.nom = nom
        self.mdp = mdp
        self.filepath = filepath
        self.intervalle = intervalle
        self.forts = []
        self.iles = []
        self.next_scan = ""
        self.last_x = -1
        self.last_y = -1
        self.last_request = -1
        self.nb_fail = 0

    def on_open(self, ws):
        print("### socket connected ###")
        time.sleep(1)
        self.send(f"""<msg t='sys'><body action='login' r='0'><login z='{self.serveur_header}'><nick><![CDATA[]]></nick><pword><![CDATA[1065004%fr%0]]></pword></login></body></msg>""")
        self.send(f"""%xt%{self.serveur_header}%lli%1%{{"CONM":175,"RTM":24,"ID":0,"PL":1,"NOM":"{self.nom}","PW":"{self.mdp}","LT":null,"LANG":"fr","DID":"0","AID":"1674256959939529708","KID":"","REF":"https://empire.goodgamestudios.com","GCI":"","SID":9,"PLFID":1}}%""")

    def run(self):
        while True:
            Thread(target=self.start_scan_map, daemon=True).start()
            self.next_scan = (datetime.now() + timedelta(seconds=self.intervalle * 60)).strftime('%H:%M:%S')
            for i in range(self.intervalle):
                self.send(f"""%xt%{self.serveur_header}%pin%1%<RoundHouseKick>%""")
                for j in range(6):
                    if self.last_request != -1 and time.time() - self.last_request > 10:
                        self.nb_fail += 1
                        if self.nb_fail >= 3:
                            self.scan_map_cells(self.last_x // 13 + self.last_y // 1170, (self.last_y // 13 + 10) % 100)
                        else:
                            self.scan_map_cells(self.last_x, self.last_y)
                    time.sleep(10)

    def start_scan_map(self):
        self.window.scan_state.set(f"Scan de la carte en cours : 0%")
        self.scan_map_cells(0, 0)

    def scan_map_cells(self, x, y):
        self.last_x = x
        self.last_y = y
        self.last_request = int(time.time())
        self.nb_fail = 0
        try:
            for i in range(10 - y // 90):
                self.send(f"""%xt%{self.serveur_header}%gaa%1%{{"KID":{self.royaume},"AX1":{13 * x},"AY1":{13 * (y + i)},"AX2":{13 * x + 12},"AY2":{13 * (y + i) + 12}}}%""")
        except websocket.WebSocketConnectionClosedException as e:
            if self.window is not None:
                self.window.scan_state.set("Connexion perdue. Reconnexion en cours...")
            raise e

    def finish_scan_map(self):
        self.window.scan_state.set(f"Scan de la carte en cours : 100%")
        self.last_x = -1
        self.last_y = -1
        self.last_request = -1
        self.nb_fail = 0
        try:
            with pd.ExcelWriter(self.filepath, engine="openpyxl", datetime_format="dd/mm/yyyy hh:mm:ss", date_format="dd/mm/yyyy") as writer:
                df = pd.DataFrame(self.forts, columns=["Coord X", "Coord Y", "Niveau", "Défendu", "Attaques restantes", "Heure de vérification", "Temps Restant"])
                df.loc[df["Temps Restant"] > timedelta(0), "Attaques restantes"] = 0
                df["Heure de disponibilité"] = df["Heure de vérification"] + df["Temps Restant"]
                df.sort_values(by=["Temps Restant", "Niveau", "Défendu"], ascending=[True, False, True], inplace=True)
                df.to_excel(writer, sheet_name="Forts", index=False, freeze_panes=(1, 0))
                worksheet = writer.sheets["Forts"]
                for cell in worksheet["F"]:
                    cell.number_format = "dd/mm/yyyy hh:mm:ss"
                for cell in worksheet["G"]:
                    cell.number_format = "hh:mm:ss"
                for cell in worksheet["H"]:
                    cell.number_format = "dd/mm/yyyy hh:mm:ss"
                worksheet.column_dimensions["E"].width = 14
                worksheet.column_dimensions["F"].width = 20
                worksheet.column_dimensions["G"].width = 14
                worksheet.column_dimensions["H"].width = 20

                df = pd.DataFrame(self.iles, columns=["Coord X", "Coord Y", "Type", "Libre", "Heure de vérification", "Temps Restant"])
                df.loc[(df["Temps Restant"] > timedelta(0)) & (df["Libre"] == False) & (df["Type"] == "grande"), "Temps Restant"] += timedelta(hours=72)
                df.loc[(df["Temps Restant"] > timedelta(0)) & (df["Libre"] == False) & (df["Type"] == "petite"), "Temps Restant"] += timedelta(hours=36)
                df["Heure de disponibilité"] = df["Heure de vérification"] + df["Temps Restant"]
                df.sort_values(by=["Temps Restant", "Type"], ascending=[True, True], inplace=True)
                df.to_excel(writer, sheet_name="Iles", index=False, freeze_panes=(1, 0))
                worksheet = writer.sheets["Iles"]
                for cell in worksheet["E"]:
                    cell.number_format = "dd/mm/yyyy hh:mm:ss"
                for cell in worksheet["F"]:
                    cell.number_format = "[h]:mm:ss"
                for cell in worksheet["G"]:
                    cell.number_format = "dd/mm/yyyy hh:mm:ss"
                worksheet.column_dimensions["E"].width = 20
                worksheet.column_dimensions["F"].width = 14
                worksheet.column_dimensions["G"].width = 20

            self.window.scan_state.set(f"Fichier Excel mis à jour à {datetime.now().strftime('%H:%M:%S')}. Prochain scan à {self.next_scan}.")
        except PermissionError:
            self.window.show_error(f"Impossible d'ouvrir le fichier Excel. Vérifiez qu'il n'est pas déjà ouvert par un autre programme puis réessayez.\n\nEmplacement du fichier :\n{self.filepath}")

    def on_message(self, ws, message):
        message = message.decode('UTF-8')
        if message[:12] == "%xt%lli%1%0%":
            Thread(target=self.run, daemon=True).start()
        elif message[:10] == "%xt%lli%1%" and message[10] != "0":
            self.window.show_error("La connexion au serveur a échoué. Vérifiez que le nom d'utilisateur et le mot de passe sont corrects et que vous avez sélectionné le bon serveur.")
            self.close()
        elif message[:12] == "%xt%gaa%1%0%":
            data = json.loads(message[12:-1])
            for castle in data["AI"]:
                if castle[0] == 25:
                    self.forts.append([castle[1], castle[2], forts[castle[5]]["level"], forts[castle[5]]["defense"], 10 - castle[7], datetime.now(), timedelta(seconds=castle[8])])
                elif castle[0] == 24 and castle[8] in [3, 6]:
                    self.iles.append([castle[1], castle[2], "grande" if castle[8] == 3 else "petite", castle[4] == -403, datetime.now(), timedelta(seconds=castle[9])])
            if data["AI"][0][2] // 13 == 98:
                self.window.scan_state.set(f"""Scan de la carte en cours : {data["AI"][0][1] // 13 + 1}%""")
                if data["AI"][0][1] // 13 == 98:
                    self.finish_scan_map()
            if (data["AI"][0][2] // 13) % 10 == 0 and (data["AI"][0][1] // 13 != 98 or data["AI"][0][2] // 13 != 90):
                self.scan_map_cells(data["AI"][0][1] // 13 + data["AI"][0][2] // 1170, (data["AI"][0][2] // 13 + 10) % 100)
        elif message[:10] == "%xt%gaa%1%" and message[10] != "0":
            self.window.show_error("Ce compte n'a pas de chateau dans ce royaume.")
            self.close()

    def on_error(self, ws, error):
        print("### Error ###")
        print(error)

    def on_close(self, ws, close_status_code, close_msg):
        print("### Socket disconnected ###")

    def close(self):
        super().close()
        self.window = None
