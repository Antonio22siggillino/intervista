import tkinter as tk
import sqlite3
import os
import datetime
from fpdf import FPDF
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from tkinter import ttk

def salva_dati():
    dati_cliente = {
        "nome": entry_nome_var.get(),
        "telefono": entry_telefono_var.get(),
        "citta": entry_citta_var.get(),
        "tipologia_intervento": tipologia_intervento_var.get(),
        "tempi_consegna": tempi_consegna_var.get(),
        "bonus": bonus_var.get(),
        "canale_contatto": canale_contatto_var.get(),
        "richiesta": richiesta_var.get(),
        "accessori": accessori_var.get(),
        "smontaggio": smontaggio_var.get(),
        "colore": entry_colore_var.get(),
        "è_già_stato_da_un_altro_serramentista": è_già_stato_da_un_altro_serramentista_var.get(),
        "riconoscimento": riconoscimento_var.get(),
        "cellulare": entry_cellulare_var.get(),
        "indirizzo": entry_indirizzo_var.get(),
        "ha_la_cila_scia": ha_la_cila_scia_var.get(),
        "presenza_impresa": presenza_impresa_var.get(),
        "note_bonus": entry_note_bonus_var.get(),
        "consigliato_da": entry_consigliato_da_var.get(),
        "richiesta_alternativa": richiesta_alternativa_var.get(),
        "fornitore_accessori": fornitore_accessori_var.get(),
        "smaltimento": smaltimento_var.get(),
        "tipologia_posa": tipologia_posa_var.get(),
        "cifra_precisa": entry_cifra_precisa_var.get(),
        "chi": entry_chi_var.get(),
        "servito_da": servito_da_var.get(),
        "email": entry_email_var.get(),
        "accessibilita_consegna": accessibilita_consegna_var.get(),
        "data_cila_scia": data_var.get(),
        "tipo_abitazione": tipo_abitazione_var.get(),
        "quantita_pezzi": quantita_pezzi_var.get(),
        "mestiere": entry_mestiere_var.get(),
        "vetro": vetro_var.get(),
        "note_accessori": entry_note_accessori_var.get(),
        "taglio_termico": taglio_termico_var.get(),
        "note": entry_note_var.get(),
        "invio_foto_misure": entry_invio_var.get(),
        "interesse": interesse_var.get()
    }

    db_file_path = os.path.join(os.path.expanduser("~"), "Desktop", "3.0", "clienti.db")
    conn = sqlite3.connect(db_file_path)
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS clienti
                 (nome TEXT, telefono TEXT, citta TEXT)''')

    c.execute("INSERT INTO clienti VALUES (:nome, :telefono, :citta, :tipologia_intervento, :tempi_consegna, :bonus, :canale_contatto, :richiesta, :accessori, :smontaggio, :colore, :è_già_stato_da_un_altro_serramentista, :riconoscimento, :cellulare, :indirizzo, :ha_la_cila_scia, :presenza_impresa, :note_bonus, :consigliato_da, :richiesta_alternativa, :fornitore_accessori, :smaltimento, :tipologia_posa, :cifra_precisa, :chi, :servito_da, :email, :accessibilita_consegna, :data_cila_scia, :tipo_abitazione, :quantita_pezzi, :mestiere, :vetro, :note_accessori, :taglio_termico, :note, :invio_foto_misure, :interesse)", dati_cliente)
    conn.commit()
    
    cliente_id = c.lastrowid
    conn.close()

    export_pdf(cliente_id, dati_cliente["nome"])  
    messagebox.showinfo("Successo", "I dati sono stati salvati con successo.")

def export_pdf(cliente_id, nome_cliente):
    pdf_folder_path = os.path.join(os.path.expanduser("~"), "Desktop", "3.0")
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    pdf_file_path = os.path.join(pdf_folder_path, f"{nome_cliente}_{cliente_id}_{timestamp}.pdf")  

    db_file_path = os.path.join(os.path.expanduser("~"), "Desktop", "3.0", "clienti.db")
    conn = sqlite3.connect(db_file_path)
    c = conn.cursor()

    c.execute("SELECT * FROM clienti WHERE rowid=?", (cliente_id,))
    cliente_info = c.fetchone()

    conn.close()

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    if cliente_info:
        pdf.set_font("Arial", style="B", size=14)
        pdf.cell(0, 10, txt=f"Modulo di Raccolta Informazioni Cliente - {nome_cliente}", ln=True, align="C")  # Aggiorna il titolo del PDF con il nome del cliente
        pdf.ln(10)
        
        # Impostazione tabellare dei titoli delle colonne
        col_width = 60
        col_titles = [
            "Nome", "Telefono", "Città", "Tipologia Intervento", "Tempi di Consegna", "Bonus",
            "Canale Contatto", "Richiesta", "Accessori", "Smontaggio", "Colore", "Budget",
            "Stato dalla Concorrenza", "Riconoscimento", "Cellulare", "Indirizzo",
            "Ha la CILA/SCIA", "Presenza Impresa", "Note Bonus", "Consigliato Da", "Richiesta Alternativa",
            "Fornitore Accessori", "Smaltimento", "Tipologia Posa", "Cifra Precisa", "Chi",
            "Servito Da", "Email", "Accessibilità Consegna", "Data CILA/SCIA", "Tipo Abitazione",
            "Quantità Pezzi", "Mestiere", "Vetro", "Note Accessori", "Taglio Termico", "Note",
            "Invio Foto Misure", "Interesse"
        ]

        row_counter = 3
        for title, data in zip(col_titles, cliente_info):
            pdf.cell(col_width, 6, title + ":", border=0, ln=0, align="L")
            pdf.cell(col_width, 6, str(data), border=0, ln=0, align="L")
            row_counter += 3
            # Aggiungi una nuova riga dopo ogni titolo
            if row_counter % 1 == 0:
                pdf.ln()
                # Se siamo arrivati a 26 righe, aggiungi una nuova pagina
                if row_counter % 43 == 0:
                    pdf.add_page()

        # Salvataggio del PDF
        pdf.output(pdf_file_path)
    else:
        messagebox.showerror("Errore", "Nessuna informazione trovata per il cliente.")

def crea_cartella(servito_da, nome_cliente, data_ora_compilazione):
    # Percorso specifico basato sulla risposta alla domanda "servito da"
    percorsi_servito_da = {
        "luca": "\\192.168.1.81\public\Condivisa\A1 PREVENTIVI\1 LUCA Preventivi 2024",
        "donatella": "\\192.168.1.81\public\Condivisa\A1 PREVENTIVI\2 DONATELLA Preventivi 2024",
        "antonio": "\\192.168.1.81\public\Condivisa\A1 PREVENTIVI\4 ANTONIO Preventivi 2024",
        "pasquale": "\\192.168.1.81\public\Condivisa\A1 PREVENTIVI\3 PASQUALE Preventivi 2024",
        "gina": "\\192.168.1.81\public\Condivisa\A1 PREVENTIVI\5 GINA Preventivi 2024"
    }

    if servito_da.lower() in percorsi_servito_da:
        percorso_base = percorsi_servito_da[servito_da.lower()]
        percorso_cartella = os.path.join(percorso_base, f"{nome_cliente}_{data_ora_compilazione}")
        os.makedirs(percorso_cartella)
        return percorso_cartella
    else:
        messagebox.showerror("Errore", f"Percorso non trovato per '{servito_da}'.")

def create_new_pdf(client_name, data):
    # Nome del nuovo PDF
    pdf_filename = f"{client_name}_abaco_infissi.pdf"

    # Creazione del documento PDF
    c = canvas.Canvas(pdf_filename, pagesize=letter)

    # Definizione delle dimensioni della tabella
    col_widths = [100, 70, 50, 50, 50, 180]
    row_height = 20

    # Definizione dei titoli delle colonne
    column_titles = ["Riferimento infissi", "Tipo", "Lmm", "Hmm", "Risposte", "Note"]

    # Aggiunta dei titoli alla prima riga della tabella
    for i, title in enumerate(column_titles):
        c.drawString((i * col_widths[i]) + 10, 770, title)

    # Aggiunta delle informazioni dal DataFrame
    for row_index, row_data in enumerate(data):
        for col_index, cell_data in enumerate(row_data):
            c.drawString((col_index * col_widths[col_index]) + 10, 750 - ((row_index + 1) * row_height), str(cell_data))

    # Aggiunta del titolo del documento
    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, 800, "Abaco Infissi - " + client_name)

    # Salvataggio del PDF
    c.save()

# Esempio di utilizzo
client_name = "Nome Cliente"
data = [
    ["Riferimento 1", "Tipo 1", "100", "200", "Spuntato", "Note 1"],
    ["Riferimento 2", "Tipo 2", "150", "250", "Non spuntato", "Note 2"],
    # Aggiungere altre righe di dati se necessario
]

create_new_pdf(client_name, data)

root = tk.Tk()
root.title("Modulo di Raccolta Informazioni Cliente")
root.geometry("1000x400")

root = tk.Tk()
root.title("Modulo di Raccolta Informazioni Cliente")
root.geometry("1000x400")

entry_nome_var = tk.StringVar(root)
entry_telefono_var = tk.StringVar(root)
entry_citta_var = tk.StringVar(root)
entry_colore_var = tk.StringVar(root)
entry_cellulare_var = tk.StringVar(root)
entry_indirizzo_var = tk.StringVar(root)
entry_note_bonus_var = tk.StringVar(root)
entry_consigliato_da_var = tk.StringVar(root)
entry_cifra_precisa_var = tk.StringVar(root)
entry_chi_var = tk.StringVar(root)
entry_email_var = tk.StringVar(root)
entry_mestiere_var = tk.StringVar(root)
entry_note_accessori_var = tk.StringVar(root)
entry_note_var = tk.StringVar(root)
entry_invio_var = tk.StringVar(root)

tipologia_intervento_var = tk.StringVar(root)
tempi_consegna_var = tk.StringVar(root)
bonus_var = tk.StringVar(root)
canale_contatto_var = tk.StringVar(root)
richiesta_var = tk.StringVar(root)
accessori_var = tk.StringVar(root)
smontaggio_var = tk.StringVar(root)
è_già_stato_da_un_altro_serramentista_var = tk.StringVar(root)
riconoscimento_var = tk.StringVar(root)
ha_la_cila_scia_var = tk.StringVar(root)
presenza_impresa_var = tk.StringVar(root)
richiesta_alternativa_var = tk.StringVar(root)
fornitore_accessori_var = tk.StringVar(root)
smaltimento_var = tk.StringVar(root)
tipologia_posa_var = tk.StringVar(root)
servito_da_var = tk.StringVar(root)
accessibilita_consegna_var = tk.StringVar(root)
tipo_abitazione_var = tk.StringVar(root)
quantita_pezzi_var = tk.StringVar(root)
vetro_var = tk.StringVar(root)
taglio_termico_var = tk.StringVar(root)
interesse_var = tk.StringVar(root)

richiesta_alternativa_boolean_var = tk.BooleanVar(root)
ha_la_cila_scia_boolean_var = tk.BooleanVar(root)
presenza_impresa_boolean_var = tk.BooleanVar(root)
taglio_termico_boolean_var = tk.BooleanVar(root)
interesse_boolean_var = tk.BooleanVar(root)

data_var = tk.StringVar(root)

canvas = tk.Canvas(root)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(root, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

scrollbar = ttk.Scrollbar(root, orient=tk.HORIZONTAL, command=canvas.xview)
scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

mainframe = ttk.Frame(canvas)
canvas.create_window((0, 0), window=mainframe, anchor=tk.NW)

frame_colonna1 = ttk.Frame(mainframe)
frame_colonna1.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.S), padx=10)

frame_colonna2 = ttk.Frame(mainframe)
frame_colonna2.grid(column=1, row=0, sticky=(tk.N, tk.W, tk.S), padx=10)

frame_colonna3 = ttk.Frame(mainframe)
frame_colonna3.grid(column=2, row=0, sticky=(tk.N, tk.W, tk.S), padx=10)

frame_colonna4 = ttk.Frame(mainframe)
frame_colonna4.grid(column=3, row=0, sticky=(tk.N, tk.W, tk.S), padx=10)

frame_colonna5 = ttk.Frame(mainframe)
frame_colonna5.grid(column=4, row=0, sticky=(tk.N, tk.W, tk.S), padx=10)

frame_colonna6 = ttk.Frame(mainframe)
frame_colonna6.grid(column=5, row=0, sticky=(tk.N, tk.W, tk.S), padx=10)

frame_colonna7 = ttk.Frame(mainframe)
frame_colonna7.grid(column=6, row=0, sticky=(tk.N, tk.W, tk.S), padx=10)

frame_colonna8 = ttk.Frame(mainframe, padding="30")
frame_colonna8.grid(column=7, row=0, sticky=(tk.N, tk.W, tk.S), padx=20)

frame_colonna9 = ttk.Frame(mainframe)
frame_colonna9.grid(column=8, row=0, sticky=(tk.N, tk.W, tk.S), padx=10)

ttk.Label(frame_colonna1, text="CLIENTE:").grid(column=0, row=0, sticky=tk.W)
entry_nome = ttk.Entry(frame_colonna1)
entry_nome.grid(column=0, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="NUMERO DI TELEFONO:").grid(column=0, row=2, sticky=tk.W)
entry_telefono = ttk.Entry(frame_colonna1)
entry_telefono.grid(column=0, row=3, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="CITTÀ:").grid(column=0, row=4, sticky=tk.W)
entry_citta = ttk.Entry(frame_colonna1)
entry_citta.grid(column=0, row=5, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="TIPOLOGIA D'INTERVENTO:").grid(column=0, row=6, sticky=tk.W)
tipologia_intervento_var = tk.StringVar(root)
tipologia_intervento_var.set("")  
option_menu_tipologia_intervento = ttk.Combobox(frame_colonna1, textvariable=tipologia_intervento_var, values=["", "Ristrutturazione", "Manutenzione", "Nuova costruzione", "Altro"], width=21)
option_menu_tipologia_intervento.grid(column=0, row=7, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="TEMPI DI CONSEGNA:").grid(column=0, row=8, sticky=tk.W)
tempi_consegna_var = tk.StringVar(root)
tempi_consegna_var.set("")  
option_menu_tempi_consegna = ttk.Combobox(frame_colonna1, textvariable=tempi_consegna_var, values=["", "MENO DI 3 MESI", "MENO DI 6 MESI", "MENO DI 9 MESI", "UN ANNO"], width=21)
option_menu_tempi_consegna.grid(column=0, row=9, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="BONUS?").grid(column=0, row=10, sticky=tk.W)
bonus_var = tk.StringVar(root)
bonus_var.set("")
option_menu_bonus = ttk.Combobox(frame_colonna1, textvariable=bonus_var, values=["", "BONUS 75", "BONUS 50", "DETRAZIONE PERSONALE 50%", "DETRAZIONE PERSONALE 75%", "BONUS CASA", "ECO BONUS"], width=20)
option_menu_bonus.grid(column=0, row=11, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="CANALE DI CONTATTO").grid(column=0, row=12, sticky=tk.W)
canale_contatto_var = tk.StringVar(root)
canale_contatto_var.set("")
option_menu_canale_contatto = ttk.Combobox(frame_colonna1, textvariable=canale_contatto_var, values=["", "Sito web", "Expo", "Telefono", "Whatsapp", "E-mail", "Conoscenza", "Consigliato"], width=20)
option_menu_canale_contatto.grid(column=0, row=13, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="RICHIESTA").grid(column=0, row=14, sticky=tk.W)
richiesta_var = tk.StringVar(root)
richiesta_var.set("")
option_menu_richiesta = ttk.Combobox(frame_colonna1, textvariable=richiesta_var, values=["", "QFORT 4 STARS", "QFORT 4 STARS VIEW", "QFORT 7 STARS", "QFORT 7 STARS VIEW", "QFORT 5 STARS", "QFORT 5 STARS VIEW", "QFORT 6 STARS", "QFORT 6 STARS VIEW", "ARROGANCE", "ISIK A", "ISIK A PLUS", "ISIK Ae", "ISIK Ae PLUS" , "ISIK SE", "ISIK SE PLUS", "OKULTA", "THERMA", "ASTRA", "DE CARLO HIDDEN", "DE CARLO 68 ARTE OG", "DE CARLO 68 DESIGN", "DE CARLO LEGNO", "DE CARLO LEGNO/ALU", "ALTRO"], width=20)
option_menu_richiesta.grid(column=0, row=15, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="ACCESSORI:").grid(column=0, row=16, sticky=tk.W)
accessori_var = tk.StringVar(root)
accessori_var.set("")  
option_menu_accessori = ttk.Combobox(frame_colonna1, textvariable=accessori_var, values=["", "Cassonetti", "Avvolgibili", "Cassonetti e Avvolgibili", "Zanzariere", "Cassonetti e Zanzariere", "Avvolgibili e Zanzariere", "Cassonetti, Avvolgibili e Zanzariere"], width=21)
option_menu_accessori.grid(column=0, row=17, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="SMONTAGGIO:").grid(column=0, row=18, sticky=tk.W)
smontaggio_var = tk.StringVar(root)
smontaggio_var.set("")  
option_menu_smontaggio = ttk.Combobox(frame_colonna1, textvariable=smontaggio_var, values=["", "SI", "NO"], width=21)
option_menu_smontaggio.grid(column=0, row=19, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="COLORE:").grid(column=0, row=20, sticky=tk.W)
entry_colore = ttk.Entry(frame_colonna1)
entry_colore.grid(column=0, row=21, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="BUDGET:").grid(column=0, row=22, sticky=tk.W)
budget_var = tk.StringVar(root)
budget_var.set("")  
option_menu_budget = ttk.Combobox(frame_colonna1, textvariable=budget_var, values=["", "MENO DI 5K", "TRA 5K E 10K", "TRA 10K E 15K", "TRA 15K E 20K", "TRA 20K E 25K", "TRA 25K E 30K"], width=21)
option_menu_budget.grid(column=0, row=23, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna1, text="È GIÀ STATO DA UN ALTRO SERRAMENTISTA?:").grid(column=0, row=24, sticky=tk.W)
è_già_stato_da_un_altro_serramentista_var = tk.StringVar(root)
è_già_stato_da_un_altro_serramentista_var.set("")  
option_menu_è_già_stato_da_un_altro_serramentista = ttk.Combobox(frame_colonna1, textvariable=è_già_stato_da_un_altro_serramentista_var, values=["", "SI","NO"], width=21)
option_menu_è_già_stato_da_un_altro_serramentista.grid(column=0, row=25, sticky=(tk.W, tk.E), padx=5, pady=5)

# Etichette e campi di input per la seconda colonna
ttk.Label(frame_colonna2, text="RICONOSCIMENTO").grid(column=0, row=0, sticky=tk.W)
riconoscimento_var = ttk.Entry(frame_colonna2)
riconoscimento_var.grid(column=0, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="CELLULARE:").grid(column=0, row=2, sticky=tk.W)
entry_cellulare = ttk.Entry(frame_colonna2)
entry_cellulare.grid(column=0, row=3, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="INDIRIZZO:").grid(column=0, row=4, sticky=tk.W)
entry_indirizzo = ttk.Entry(frame_colonna2)
entry_indirizzo.grid(column=0, row=5, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="HA LA CILA/SCIA?").grid(column=0, row=6, sticky=tk.W)
ha_la_cila_scia_var = tk.StringVar(root)
ha_la_cila_scia_var.set("")
option_menu_ha_la_cila_scia = ttk.Combobox(frame_colonna2, textvariable=ha_la_cila_scia_var, values=["", "SCIA", "CILA", "NO"], width=5)
option_menu_ha_la_cila_scia.grid(column=0, row=7, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="È PRESENTE L'IMPRESA?").grid(column=0, row=8, sticky=tk.W)
presenza_impresa_var = tk.StringVar(root)
presenza_impresa_var.set("")  
option_menu_presenza_impresa = ttk.Combobox(frame_colonna2, textvariable=presenza_impresa_var, values=["", "SI", "NO"], width=21)
option_menu_presenza_impresa.grid(column=0, row=9, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="NOTE SUL BONUS:").grid(column=0, row=10, sticky=tk.W)
entry_note_bonus = ttk.Entry(frame_colonna2)
entry_note_bonus.grid(column=0, row=11, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="SE CONSIGLIATO, DA CHI?").grid(column=0, row=12, sticky=tk.W)
entry_consigliato_da = ttk.Entry(frame_colonna2)
entry_consigliato_da.grid(column=0, row=13, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="RICHIESTA ALTERNATIVA").grid(column=0, row=14, sticky=tk.W)
richiesta_alternativa_var = tk.StringVar(root)
richiesta_alternativa_var.set("")
option_menu_richiesta_alternativa = ttk.Combobox(frame_colonna2, textvariable=richiesta_alternativa_var, values=["", "QFORT 4 STARS", "QFORT 4 STARS VIEW", "QFORT 7 STARS", "QFORT 7 STARS VIEW", "QFORT 5 STARS", "QFORT 5 STARS VIEW", "QFORT 6 STARS", "QFORT 6 STARS VIEW", "ARROGANCE", "ISIK A", "ISIK A PLUS", "ISIK Ae", "ISIK Ae PLUS" , "ISIK SE", "ISIK SE PLUS", "OKULTA", "THERMA", "ASTRA", "DE CARLO HIDDEN", "DE CARLO 68 ARTE OG", "DE CARLO 68 DESIGN", "DE CARLO LEGNO", "DE CARLO LEGNO/ALU", "ALTRO"], width=20)
option_menu_richiesta_alternativa.grid(column=0, row=15, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="FORNITORE ACCESSORI:").grid(column=0, row=16, sticky=tk.W)
fornitore_accessori_var = tk.StringVar(root)
fornitore_accessori_var.set("")  
option_menu_fornitore_accessori = ttk.Combobox(frame_colonna2, textvariable=fornitore_accessori_var, values=["", "QFORT", "SCIUKER", "TEKNIKA", "ISOLCASS", "EDILPLASTIC", "EFFE", "PRATIC", "ALIAS", "EDILPLASTIC e EFFE", "EDILPLASTIC e ISOLCASS", "EFFE e ISOLCASS", "EDILPLASTIC, EFFE e ISOLCASS", "EDILPLASTIC e ALIAS", "EFFE e ALIAS", "ISOLCASS e ALIAS", "EDILPLASTIC EFFE e ALIAS", "EDILPLASTIC ISOLCASS e ALIAS", "EFFE ISOLCASS E ALIAS", "EDILPLASTIC EFFE ISOLCASS e ALIAS"], width=21)
option_menu_fornitore_accessori.grid(column=0, row=17, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="SMALTIMENTO:").grid(column=0, row=18, sticky=tk.W)
smaltimento_var = tk.StringVar(root)
smaltimento_var.set("")  
option_menu_smaltimento = ttk.Combobox(frame_colonna2, textvariable=smaltimento_var, values=["", "SI", "NO"], width=21)
option_menu_smaltimento.grid(column=0, row=19, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="TIPOLOGIA DI POSA:").grid(column=0, row=20, sticky=tk.W)
tipologia_posa_var = tk.StringVar(root)
tipologia_posa_var.set("")  
option_menu_tipologia_posa = ttk.Combobox(frame_colonna2, textvariable=tipologia_posa_var, values=["", "POSA CLIMA PREMIUM PLUS", "POSA CLIMA PREMIUM", "POSA SILICONE E SCHIUMA"], width=21)
option_menu_tipologia_posa.grid(column=0, row=21, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="CIFRA PRECISA:").grid(column=0, row=22, sticky=tk.W)
entry_cifra_precisa = ttk.Entry(frame_colonna2)
entry_cifra_precisa.grid(column=0, row=23, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna2, text="CHI?:").grid(column=0, row=24, sticky=tk.W)
entry_chi = ttk.Entry(frame_colonna2)
entry_chi.grid(column=0, row=25, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="SERVITO DA?:").grid(column=0, row=0, sticky=tk.W)
servito_da_var = tk.StringVar(root)
servito_da_var.set("")
option_menu_servito_da = ttk.Combobox(frame_colonna3, textvariable=servito_da_var, values=["", "LUCA", "DONATELLA", "GINA", "PASQUALE", "ANTONIO"], width=20)
option_menu_servito_da.grid(column=0, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="INDIRIZZO EMAIL:").grid(column=0, row=2, sticky=tk.W)
entry_email = ttk.Entry(frame_colonna3)
entry_email.grid(column=0, row=3, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="ACCESSIBILITÀ PER LA CONSEGNA:").grid(column=0, row=4, sticky=tk.W)
accessibilita_consegna_var = tk.StringVar(root)
accessibilita_consegna_var.set("")  
option_menu_accessibilita_consegna = ttk.Combobox(frame_colonna3, textvariable=accessibilita_consegna_var, values=["", "A mano", "Con scala mobile", "Gru con braccio"], width=21)
option_menu_accessibilita_consegna.grid(column=0, row=5, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="DATA CILA/SCIA?").grid(column=0, row=6, sticky=tk.W)
data_var = ttk.Entry(frame_colonna3)
data_var.grid(column=0, row=7, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="TIPO DI ABITAZIONE:").grid(column=0, row=8, sticky=tk.W)
tipo_abitazione_var = tk.StringVar(root)
tipo_abitazione_var.set("")  
option_menu_tipo_abitazione = ttk.Combobox(frame_colonna3, textvariable=tipo_abitazione_var, values=["", "Casa primaria", "Casa secondaria", "Casa unifamiliare", "Casa al mare", "Casa in montagna", "Villetta", "Villetta al mare", "Villa"], width=21)
option_menu_tipo_abitazione.grid(column=0, row=9, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="QUANTITÀ DI PEZZI:").grid(column=0, row=10, sticky=tk.W)
quantita_pezzi_var = tk.StringVar(root)
quantita_pezzi_var.set("")
option_menu_quantita_pezzi = ttk.Combobox(frame_colonna3, textvariable=quantita_pezzi_var, values=["", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13"], width=21)
option_menu_quantita_pezzi.grid(column=0, row=11, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="MESTIERE:").grid(column=0, row=12, sticky=tk.W)
additional_text = "MESTIERE: lei è un'ingegnere, architetto, geometra, o altro?"
ttk.Label(frame_colonna3, text=additional_text, font=("Arial", 8)).grid(column=0, row=12, sticky=tk.W)
entry_mestiere = ttk.Entry(frame_colonna3, width=25)
entry_mestiere.grid(column=0, row=13, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="VETRO TRIPLO O CAMERA?:").grid(column=0, row=14, sticky=tk.W)
vetro_var = tk.StringVar(root)
vetro_var.set("")
option_menu_vetro = ttk.Combobox(frame_colonna3, textvariable=vetro_var, values=["", "TRIPLO", "CAMERA"], width=21)
option_menu_vetro.grid(column=0, row=15, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="NOTE PER GLI ACCESSORI:").grid(column=0, row=16, sticky=tk.W)
entry_note_accessori = ttk.Entry(frame_colonna3)
entry_note_accessori.grid(column=0, row=17, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="TAGLIO DEL PONTE TERMICO?:").grid(column=0, row=18, sticky=tk.W)
taglio_termico_var = tk.StringVar(root)
taglio_termico_var.set("")
option_menu_taglio_termico = ttk.Combobox(frame_colonna3, textvariable=taglio_termico_var, values=["", "SI", "NO", "SOLO SULLE FINESTRE"], width=21)
option_menu_taglio_termico.grid(column=0, row=19, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="NOTE:").grid(column=0, row=20, sticky=tk.W)
entry_note = ttk.Entry(frame_colonna3)
entry_note.grid(column=0, row=21, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="INVIO FOTO, MISURE E PLANIMETRIA A 0835 38 5742:").grid(column=0, row=22, sticky=tk.W)
entry_invio = ttk.Entry(frame_colonna3)
entry_invio.grid(column=0, row=23, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna3, text="INTERESSE:").grid(column=0, row=24, sticky=tk.W)
interesse_var = tk.StringVar(root)
interesse_var.set("")
option_menu_interesse = ttk.Combobox(frame_colonna3, textvariable=interesse_var, values=["", "ALTO", "MEDIO", "BASSO"], width=21)
option_menu_interesse.grid(column=0, row=25, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="RIFERIMENTO INFISSI:").grid(column=0, row=0, sticky=tk.W)
entry_riferimento = ttk.Entry(frame_colonna4)
entry_riferimento.grid(column=0, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="2:").grid(column=0, row=2, sticky=tk.W)
entry_2 = ttk.Entry(frame_colonna4)
entry_2.grid(column=0, row=3, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="3:").grid(column=0, row=4, sticky=tk.W)
entry_3 = ttk.Entry(frame_colonna4)
entry_3.grid(column=0, row=5, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="4:").grid(column=0, row=6, sticky=tk.W)
entry_4 = ttk.Entry(frame_colonna4)
entry_4.grid(column=0, row=7, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="5:").grid(column=0, row=8, sticky=tk.W)
entry_5 = ttk.Entry(frame_colonna4)
entry_5.grid(column=0, row=9, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="6:").grid(column=0, row=10, sticky=tk.W)
entry_6 = ttk.Entry(frame_colonna4)
entry_6.grid(column=0, row=11, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="7:").grid(column=0, row=12, sticky=tk.W)
entry_7 = ttk.Entry(frame_colonna4)
entry_7.grid(column=0, row=13, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="8:").grid(column=0, row=14, sticky=tk.W)
entry_8 = ttk.Entry(frame_colonna4)
entry_8.grid(column=0, row=15, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="9:").grid(column=0, row=16, sticky=tk.W)
entry_9 = ttk.Entry(frame_colonna4)
entry_9.grid(column=0, row=17, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="10:").grid(column=0, row=18, sticky=tk.W)
entry_10 = ttk.Entry(frame_colonna4)
entry_10.grid(column=0, row=19, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="11:").grid(column=0, row=20, sticky=tk.W)
entry_11 = ttk.Entry(frame_colonna4)
entry_11.grid(column=0, row=21, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="12:").grid(column=0, row=22, sticky=tk.W)
entry_12 = ttk.Entry(frame_colonna4)
entry_12.grid(column=0, row=23, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna4, text="13:").grid(column=0, row=24, sticky=tk.W)
entry_13 = ttk.Entry(frame_colonna4)
entry_13.grid(column=0, row=25, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=0, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=2, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=3, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=4, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=5, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=6, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=7, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=8, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=9, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=10, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=11, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=12, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=13, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=14, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=15, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=16, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=17, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=18, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=19, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=20, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=21, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=22, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=23, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna5, text="tipo").grid(column=0, row=24, sticky=tk.W)
tipo_var = tk.StringVar(root)
tipo_var.set("")
ttk.Combobox(frame_colonna5, textvariable=tipo_var, values=["", "F1", "F2", "F3", "F4", "PF1", "PF2", "PF3", "PF4", "SCORREVOLE", "SCORREVOLE TRASLANTE", "SCORREVOLE ALZANTE", "BLINDATA", "BLINDATA 2 ANTE", "PORTA D'INGRESSO", "PORTA VERANDA", "VASISTAS", "VASISTAS MOTORIZZATO"], width=2).grid(column=0, row=25, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=0, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=2, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=3, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=4, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=5, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=6, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=7, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=8, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=9, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=10, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=11, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=12, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=13, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=14, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=15, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=16, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=17, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=18, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=19, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=20, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=21, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=22, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=23, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna6, text="L mm:").grid(column=0, row=24, sticky=tk.W)
entry_L = ttk.Entry(frame_colonna6, width=5)
entry_L.grid(column=0, row=25, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=0, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=2, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=3, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=4, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=5, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=6, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=7, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=8, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=9, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=10, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=11, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=12, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=13, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=14, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=15, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=16, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=17, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=18, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=19, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=20, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=21, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=22, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=23, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna7, text="H mm:").grid(column=0, row=24, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna7, width=5)
entry_H.grid(column=0, row=25, sticky=(tk.W, tk.E), padx=5, pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 0
ttk.Label(frame_colonna8).grid(column=0, row=0, sticky=tk.W)
accessori_var = []
opzioni = ["cass", "avv", "zanz"]
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    accessori_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var).grid(row=0, column=i, sticky="w")

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 4
ttk.Label(frame_colonna8, text="ACCESSORI 2").grid(column=0, row=2, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 8
ttk.Label(frame_colonna8, text="ACCESSORI 3").grid(column=0, row=6, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 12
ttk.Label(frame_colonna8, text="ACCESSORI 4").grid(column=0, row=10, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 16
ttk.Label(frame_colonna8, text="ACCESSORI 5").grid(column=0, row=14, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 20
ttk.Label(frame_colonna8, text="ACCESSORI 6").grid(column=0, row=18, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 24
ttk.Label(frame_colonna8, text="ACCESSORI 7").grid(column=0, row=22, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 28
ttk.Label(frame_colonna8, text="ACCESSORI 8").grid(column=0, row=26, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 32
ttk.Label(frame_colonna8, text="ACCESSORI 9").grid(column=0, row=30, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 36
ttk.Label(frame_colonna8, text="ACCESSORI 10").grid(column=0, row=34, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 40
ttk.Label(frame_colonna8, text="ACCESSORI 11").grid(column=0, row=38, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 44
ttk.Label(frame_colonna8, text="ACCESSORI 12").grid(column=0, row=42, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

style = ttk.Style()
style.configure('Custom.TCheckbutton', indicatorsize=20)
colonna = 0
riga = 48
ttk.Label(frame_colonna8, text="ACCESSORI 13").grid(column=0, row=46, sticky=tk.W)
opzioni = ["cass", "avv", "zanz"]
risposte_var = []
for i, opzione in enumerate(opzioni):
    var = tk.BooleanVar()
    risposte_var.append(var)
    ttk.Checkbutton(frame_colonna8, text=opzione, variable=var, style='Custom.TCheckbutton').grid(row=riga, column=i, sticky="w", pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=0, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=2, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=3, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=4, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=5, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=6, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=7, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=8, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=9, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=10, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=11, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=12, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=13, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=14, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=15, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=16, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=17, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=18, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=19, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=20, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=21, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=22, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=23, sticky=(tk.W, tk.E), padx=5, pady=5)

ttk.Label(frame_colonna9, text="note:").grid(column=0, row=24, sticky=tk.W)
entry_H = ttk.Entry(frame_colonna9)
entry_H.grid(column=0, row=25, sticky=(tk.W, tk.E), padx=5, pady=5)

button_salva = ttk.Button(mainframe, text="Salva", command=salva_dati)
button_salva.grid(column=0, row=12, columnspan=3, pady=10)

root.mainloop()
