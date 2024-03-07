import sys
from tkinter import Tk, filedialog, Button, Toplevel, Label, Frame
from tkinter.ttk import Progressbar
from threading import Thread
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from tqdm import tqdm
import os

def choose_file():
    root = Tk()
    root.withdraw()  
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path


def show_progress(root):
    progress_window = Toplevel(root)
    progress_window.title("Progresso")
    progress_window.geometry("300x100")
    progress_frame = Frame(progress_window)
    progress_frame.pack(pady=10)
    progress_label = Label(progress_frame, text="Geocodificação em progresso...")
    progress_label.pack()
    progress_bar = Progressbar(progress_frame, orient="horizontal", length=200, mode="determinate")
    progress_bar.pack(pady=10)
    return progress_window, progress_bar

def geocode_addresses(file_path, progress_bar):
    df = pd.read_excel(file_path)
    geolocator = Nominatim(user_agent="AdresstoPoint")
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

    def get_coordinates(address):
        location = geocode(address)
        if location:
            return location.latitude, location.longitude
        else:
            return None, None

    tqdm.pandas()
    total_rows = len(df)
    progress_bar["maximum"] = total_rows
    for index, row in tqdm(df.iterrows(), total=total_rows, desc="Geocodificação"):
        latitude, longitude = get_coordinates(f"{row['rua']}, {row['bairro']}, {row['cidade']}, {row['estado']}, {row['pais']}")
        df.at[index, 'latitude'] = latitude
        df.at[index, 'longitude'] = longitude
        progress_bar["value"] = index + 1
        progress_bar.update()

    df['WKT'] = df.apply(lambda row: f"POINT ({row['longitude']} {row['latitude']})", axis=1)
    df['descrição'] = df.apply(lambda row: f"{row['rua']}, {row['bairro']}", axis=1)  
    return df

def start_geocoding_thread(file_path, progress_window, progress_bar, start_button):
    progress_bar["value"] = 0
    start_button.config(state="disabled")  

    def run_geocoding():
        df = geocode_addresses(file_path, progress_bar)
        output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        df.to_excel(output_file_path, columns=['descrição', 'WKT'], index=False)  
        print("Conversão concluída. O arquivo foi salvo em:", output_file_path)
        progress_window.destroy()
        start_button.config(state="normal")  

    thread = Thread(target=run_geocoding)
    thread.start()

root = Tk()
root.title("Geocodificação de Endereços")
root.geometry("400x200")

def choose_file_and_start():
    file_path = choose_file()
    if file_path:
        progress_window, progress_bar = show_progress(root)
        start_geocoding_thread(file_path, progress_window, progress_bar, start_button)

start_button = Button(root, text="Carregar arquivo e iniciar geocodificação", command=choose_file_and_start)
start_button.pack(pady=70)

icon_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
icon_path = os.path.join(icon_path, 'icone.ico')  # Substitua 'icone.ico' pelo nome do seu ícone

if os.path.isfile(icon_path):
    root.iconbitmap(icon_path)

root.mainloop()