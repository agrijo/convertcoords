"""
Created on Tue Jul  4 20:03:13 2023

@author: agrijo
"""
# -*- coding: latin-1 -*-

import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pyproj import Proj, transform
import utm
import sys
from PIL import Image, ImageTk
import os 


basedir = r'C:\Users\agrijo\Desktop\transcord'

epsg_options = {
    "4326": "WGS84",
    "4225":"Corrego Alegre",
    "4618":"SAD69",
    "4674":"SIRGAS 2000",
    "4326":"WGS84",
    "31972":"SIRGAS 2000/UTM Zona 18N",
    "31978":"SIRGAS 2000/UTM Zona 18S",
    "31973":"SIRGAS 2000/UTM Zona 19N",
    "31979":"SIRGAS 2000/UTM Zona 19S",
    "31974":"SIRGAS 2000/UTM Zona 20N",
    "31980":"SIRGAS 2000/UTM Zona 20S",
    "31981":"SIRGAS 2000/UTM Zona 21S",
    "31982":"SIRGAS 2000/UTM Zona 22S",
    "31983":"SIRGAS 2000/UTM Zona 23S",
    "31984":"SIRGAS 2000/UTM Zona 24S",
    "31985":"SIRGAS 2000/UTM Zona 25S",
    "32618":"WGS84/UTM Zona 18N",
    "32718":"WGS84/UTM Zona 18S",
    "32619":"WGS84/UTM Zona 19N",
    "32719":"WGS84/UTM Zona 19S",
    "32620":"WGS84/UTM Zona 20N",
    "32720":"WGS84/UTM Zona 20S",
    "32721":"WGS84/UTM Zona 21S",
    "32722":"WGS84/UTM Zona 22S",
    "32723":"WGS84/UTM Zona 23S",
    "32724":"WGS84/UTM Zona 24S",
    "32725":"WGS84/UTM Zona 25S"

    # Adicione mais opções de EPSG e seus rótulos explicativos aqui
    }

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def transform_coordinates(lat, lon, from_epsg, to_epsg):
    in_proj = Proj(init=f"epsg:{from_epsg}")
    out_proj = Proj(init=f"epsg:{to_epsg}")
    x, y = transform(in_proj, out_proj, lon, lat)
    return y, x

def convert_to_utm(lat, lon):
    if lat < -80 or lat > 84:
        # Latitude is out of UTM range
        return None, None, None, None

    easting, northing, zone_number, zone_letter = utm.from_latlon(lat, lon)
    return easting, northing, zone_number, zone_letter

def convert_decimal_degrees_to_gms(latitude, longitude):
    lat_deg = int(latitude)
    lat_min = int((latitude - lat_deg) * 60)
    lat_sec = ((latitude - lat_deg) * 60 - lat_min) * 60
    if lat_deg > 0:
        letra = 'N'
    else:
        letra = 'S'
    lat_final = f'{abs(lat_deg)}° {abs(lat_min)}\' {abs(lat_sec):.3f}"{letra}'

    lon_deg = int(longitude)
    lon_min = int((longitude - lon_deg) * 60)
    lon_sec = ((longitude - lon_deg) * 60 - lon_min) * 60
    if lon_deg > 0:
        letra = 'E'
    else:
        letra = 'W'
    lon_final = f'{abs(lon_deg)}° {abs(lon_min)}\' {abs(lon_sec):.3f}"{letra}'

    return lat_deg, lat_min, lat_sec, lon_deg, lon_min, lon_sec, lat_final, lon_final

def convert_decimal_degrees_to_gmd(latitude, longitude):
    lat_deg = int(latitude)
    lat_min_decimal = (latitude - lat_deg) * 60
    if lat_deg > 0:
        letra = 'N'
    else:
        letra = 'S'
    lat_final_dec = f'{abs(lat_deg)}° {abs(lat_min_decimal):.3f}\'{letra}'

    lon_deg = int(longitude)
    lon_min_decimal = (longitude - lon_deg) * 60
    if lon_deg > 0:
        letra = 'E'
    else:
        letra = 'W'
    lon_final_dec = f'{abs(lon_deg)}° {abs(lon_min_decimal):.3f}\'{letra}'

    return lat_deg, lat_min_decimal, lon_deg, lon_min_decimal, lat_final_dec, lon_final_dec


def transform_file():
    latitude_col = lat_col_entry.get()
    longitude_col = lon_col_entry.get()
    from_epsg = from_epsg_var.get()
    to_epsg = to_epsg_var.get()
    input_type = input_type_var.get()
    zoneN = zoneN_entry.get()
    zoneL = zoneL_entry.get()

    if not file_path:
        result_label.config(text="Nenhum arquivo selecionado.", foreground="red", wraplength=400)
        return

    try:
        df = pd.read_excel(file_path, header = 0)
        df["Latitude UTM"] = ""
        df["Longitude UTM"] = ""
        df["Latitude (GMS)"] = ""
        df["Longitude (GMS)"] = ""
        df["Latitude (GMD)"] = ""
        df["Longitude (GMD)"] = ""
        df["Zona UTM"] = ""
        df["Letra UTM"] = ""
        
        for index, row in df.iterrows():
            latitude = row[latitude_col]
            longitude = row[longitude_col]
            
            if input_type == "UTM":
                
                lat, lon = utm.to_latlon(longitude, latitude, int(zoneN), str(zoneL))       
                easting, northing, zone_number, zone_letter = convert_to_utm(lat, lon)
                df.at[index, 'Latitude UTM'] = lat
                df.at[index, 'Longitude UTM'] = lon
                lat_deg, lat_min, lat_sec, lon_deg, lon_min, lon_sec, lat_final, lon_final = convert_decimal_degrees_to_gms(
                    lat, lon)
                df.at[index, 'Latitude (GMS)'] = lat_final
                df.at[index, 'Longitude (GMS)'] = lon_final
                lat_deg, lat_min_decimal, lon_deg, lon_min_decimal, lat_final_dec, lon_final_dec = convert_decimal_degrees_to_gmd(
                    lat, lon)
                df.at[index, 'Latitude (GMD)'] = lat_final_dec
                df.at[index, 'Longitude (GMD)'] = lon_final_dec
                df.at[index, 'Zona UTM'] = zone_number
                df.at[index, 'Letra UTM'] = zone_letter

                if easting is None or northing is None:
                    # Handle latitude out of range
                    continue
            else:
                
                easting, northing, zone_number, zone_letter = convert_to_utm(latitude, longitude)
                df.at[index, 'Latitude UTM'] = northing
                df.at[index, 'Longitude UTM'] = easting
                lat_deg, lat_min, lat_sec, lon_deg, lon_min, lon_sec, lat_final, lon_final = convert_decimal_degrees_to_gms(
                    latitude, longitude)
                df.at[index, 'Latitude (GMS)'] = lat_final
                df.at[index, 'Longitude (GMS)'] = lon_final
                lat_deg, lat_min_decimal, lon_deg, lon_min_decimal, lat_final_dec, lon_final_dec = convert_decimal_degrees_to_gmd(
                    latitude, longitude)
                df.at[index, 'Latitude (GMD)'] = lat_final_dec
                df.at[index, 'Longitude (GMD)'] = lon_final_dec
                df.at[index, 'Zona UTM'] = zone_number
                df.at[index, 'Letra UTM'] = zone_letter
            
            # Rest of the code for GMS and GMD conversions
        files = {'Arquivo Excel':'*.xlsx'}
        save_file_path = filedialog.asksaveasfilename(title="Salvar arquivo XLSX", filetypes = files)
        if save_file_path:
            df.to_excel(save_file_path+'.xlsx', index=False)
            # implementação necessária pro CSV funcionar corretamente  
            # encoding="cp1252"
            result_label.config(text="Conversão concluída. Arquivo salvo com sucesso.", foreground="blue")
        else:
            result_label.config(text="Conversão concluída.")
    except Exception as e:
        result_label.config(text=f"Erro ao processar o arquivo: {str(e)}", foreground="red")

def quit_app():
    root.destroy()
    
def preview_file():
    global file_path
    file_path = filedialog.askopenfilename(title="Selecione o arquivo XLSX")
    if file_path:
        try:
            df = pd.read_excel(file_path, header = 0)
            columns_preview = ", ".join(df.columns)
            columns_preview_label.config(text=f"Colunas do Arquivo Selecionado: {columns_preview}", foreground="blue")
        except Exception as e:
            columns_preview_label.config(text=f"Erro ao abrir o arquivo: {str(e)}",foreground="red")
    else:
        columns_preview_label.config(text="Nenhum arquivo selecionado.",foreground="red")

def show_image_in_new_window():
    
    new_window = tk.Toplevel(root)  # Crie uma nova janela
    new_window.title("Brasil UTM")  # Defina um título para a nova janela

    # Carregue a imagem usando a biblioteca PIL (Pillow)
    image_path = r"C:\Users\agrijo\Documents\GitHub\convertcoords\BrasilUTM.png"
    image = Image.open(image_path)  # Substitua pelo caminho da sua imagem
    photo = ImageTk.PhotoImage(image)

    # Crie um rótulo para exibir a imagem na nova janela
    image_label = tk.Label(new_window, image=photo)
    image_label.image = photo  # Salve uma referência para evitar que a imagem seja coletada pelo garbage collector
    image_label.pack()

    close_button = tk.Button(new_window, text="Fechar Figura", command=new_window.destroy)
    close_button.pack()
    
root = tk.Tk()
root.title("Calculadora de Coordenadas")



# root.geometry("400x400")

# Labels
tk.Label(root, text="Coluna Latitude:").grid(row=0, column=0, sticky="e")
tk.Label(root, text="Coluna Longitude:").grid(row=1, column=0, sticky="e")
tk.Label(root, text="EPSG de Entrada:").grid(row=2, column=0, sticky="e")
tk.Label(root, text="EPSG de Saída:").grid(row=3, column=0, sticky="e")
tk.Label(root, text="Tipo de Coordenada:").grid(row=4, column=0, sticky="e")
tk.Label(root, text="Número da Zona UTM:").grid(row=5, column=0, sticky="e")
tk.Label(root, text="Letra da Zona UTM:").grid(row=6, column=0, sticky="e")

result_label = tk.Label(root, text="")
result_label.grid(row=11, column=0, columnspan=2, pady=10)

columns_preview_label = tk.Label(root, text="Colunas do Arquivo: ")
columns_preview_label.grid(row=10, column=0, columnspan=2, pady=10)

# Entry Fields
lat_col_entry = tk.Entry(root)
lat_col_entry.grid(row=0, column=1)
lon_col_entry = tk.Entry(root)
lon_col_entry.grid(row=1, column=1)
zoneN_entry = tk.Entry(root)
zoneN_entry.grid(row=5, column=1)
zoneL_entry = tk.Entry(root)
zoneL_entry.grid(row=6, column=1)
# EPSG Options
from_epsg_var = tk.StringVar(root)
from_epsg_var.set(next(iter(epsg_options.values())))  # Set the first EPSG option as default
from_epsg_dropdown = tk.OptionMenu(root, from_epsg_var, *epsg_options.values())
from_epsg_dropdown.grid(row=2, column=1)

to_epsg_var = tk.StringVar(root)
to_epsg_var.set(next(iter(epsg_options.values())))  # Set the first EPSG option as default
to_epsg_dropdown = tk.OptionMenu(root, to_epsg_var, *epsg_options.values())
to_epsg_dropdown.grid(row=3, column=1)

input_type_var = tk.StringVar(root)
input_type_var.set("UTM")
input_type_dropdown = tk.OptionMenu(root, input_type_var, "UTM", "Graus Decimais")
input_type_dropdown.grid(row=4, column=1)

# Buttons
preview_button = tk.Button(root, text="Abrir Arquivo", command=preview_file)
preview_button.grid(row=8, column=0, pady=10)

show_image_button = tk.Button(root, text="Brasil UTM", command=show_image_in_new_window)
show_image_button.grid(row=4, column=2, pady=10)

transform_button = tk.Button(root, text="Transformar Coordenadas", command=transform_file)
transform_button.grid(row=8, column=1, pady=10)

quit_button = tk.Button(root, text="Sair", command=quit_app)
quit_button.grid(row=8, column=2, pady=10)

root.mainloop()
