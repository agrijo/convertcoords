"""
Created on Tue Jul  4 20:03:13 2023

@author: agrijo
"""

import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pyproj import Proj, transform
import utm
import sys

epsg_options = {
    "4326": "WGS84",
    "4225":"Corrego Alegre",
    "4618":"SAD69",
    "4674":"SIRGAS 2000",
    "4326":"WGS84",
    "22521":"Corrego Alegre/UTM Zona 21S",
    "22522":"Corrego Alegre/UTM Zona 22S",
    "22523":"Corrego Alegre/UTM Zona 23S",
    "22524":"Corrego Alegre/UTM Zona 24S",
    "22525":"Corrego Alegre/UTM Zona 25S",
    "29168":"SAD69/UTM Zona 18N",
    "29188":"SAD69/UTM Zona 18S",
    "29169":"SAD69/UTM Zona 19N",
    "29189":"SAD69/UTM Zona 19S",
    "29170":"SAD69/UTM Zona 20N",
    "29190":"SAD69/UTM Zona 20S",
    "29191":"SAD69/UTM Zona 21S",
    "29192":"SAD69/UTM Zona 22S",
    "29193":"SAD69/UTM Zona 23S",
    "29194":"SAD69/UTM Zona 24S",
    "29195":"SAD69/UTM Zona 25S",
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
        df = pd.read_csv(file_path, header=0, sep=";")
        df["Nova Latitude"] = ""
        df["Nova Longitude"] = ""
        df["Nova Latitude (GMS)"] = ""
        df["Nova Longitude (GMS)"] = ""
        df["Nova Latitude (GMD)"] = ""
        df["Nova Longitude (GMD)"] = ""
        df["Zona"] = ""
        df["Letra"] = ""
        
        for index, row in df.iterrows():
            latitude = row[latitude_col]
            longitude = row[longitude_col]
            
            if input_type == "UTM":
                
                lat, lon = utm.to_latlon(longitude, latitude, int(zoneN), str(zoneL))       
                easting, northing, zone_number, zone_letter = convert_to_utm(lat, lon)
                df.at[index, 'Nova Latitude'] = lat
                df.at[index, 'Nova Longitude'] = lon
                lat_deg, lat_min, lat_sec, lon_deg, lon_min, lon_sec, lat_final, lon_final = convert_decimal_degrees_to_gms(
                    lat, lon)
                df.at[index, 'Nova Latitude (GMS)'] = lat_final
                df.at[index, 'Nova Longitude (GMS)'] = lon_final
                lat_deg, lat_min_decimal, lon_deg, lon_min_decimal, lat_final_dec, lon_final_dec = convert_decimal_degrees_to_gmd(
                    lat, lon)
                df.at[index, 'Nova Latitude (GMD)'] = lat_final_dec
                df.at[index, 'Nova Longitude (GMD)'] = lon_final_dec
                df.at[index, 'Zona'] = zone_number
                df.at[index, 'Letra'] = zone_letter

                if easting is None or northing is None:
                    # Handle latitude out of range
                    continue
            else:
                
                easting, northing, zone_number, zone_letter = convert_to_utm(latitude, longitude)
                df.at[index, 'Nova Latitude'] = northing
                df.at[index, 'Nova Longitude'] = easting
                lat_deg, lat_min, lat_sec, lon_deg, lon_min, lon_sec, lat_final, lon_final = convert_decimal_degrees_to_gms(
                    latitude, longitude)
                df.at[index, 'Nova Latitude (GMS)'] = lat_final
                df.at[index, 'Nova Longitude (GMS)'] = lon_final
                lat_deg, lat_min_decimal, lon_deg, lon_min_decimal, lat_final_dec, lon_final_dec = convert_decimal_degrees_to_gmd(
                    latitude, longitude)
                df.at[index, 'Nova Latitude (GMD)'] = lat_final_dec
                df.at[index, 'Nova Longitude (GMD)'] = lon_final_dec
                df.at[index, 'Zona'] = zone_number
                df.at[index, 'Letra'] = zone_letter
            
            # Rest of the code for GMS and GMD conversions
            
        save_file_path = filedialog.asksaveasfilename(title="Salvar arquivo CSV")
        if save_file_path:
            df.to_csv(save_file_path, index=False, encoding="cp1252")
            result_label.config(text="Conversão concluída. Arquivo salvo com sucesso.", foreground="blue")
        else:
            result_label.config(text="Conversão concluída.")
    except Exception as e:
        result_label.config(text=f"Erro ao processar o arquivo: {str(e)}", foreground="red")

def quit_app():
    root.destroy()
    
def preview_file():
    global file_path
    file_path = filedialog.askopenfilename(title="Selecione o arquivo CSV")
    if file_path:
        try:
            df = pd.read_csv(file_path, nrows=5, sep=';')
            columns_preview = ", ".join(df.columns)
            columns_preview_label.config(text=f"Prévia das Colunas: {columns_preview}", foreground="blue")
        except Exception as e:
            columns_preview_label.config(text=f"Erro ao abrir o arquivo: {str(e)}",foreground="red")
    else:
        columns_preview_label.config(text="Nenhum arquivo selecionado.",foreground="red")

root = tk.Tk()
root.title("Calculadora de Coordenadas")
# root.geometry("400x400")

# Labels
tk.Label(root, text="Coluna Latitude:").grid(row=0, column=0, sticky="e")
tk.Label(root, text="Coluna Longitude:").grid(row=1, column=0, sticky="e")
tk.Label(root, text="EPSG de Origem:").grid(row=2, column=0, sticky="e")
tk.Label(root, text="EPSG de Destino:").grid(row=3, column=0, sticky="e")
tk.Label(root, text="Tipo de Coordenada:").grid(row=4, column=0, sticky="e")
tk.Label(root, text="Número da Zona:").grid(row=5, column=0, sticky="e")
tk.Label(root, text="Letra da Zona:").grid(row=6, column=0, sticky="e")
result_label = tk.Label(root, text="")
result_label.grid(row=11, column=0, columnspan=2, pady=10)

columns_preview_label = tk.Label(root, text="Prévia das Colunas: ")
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
preview_button = tk.Button(root, text="Pré-visualizar Arquivo", command=preview_file)
preview_button.grid(row=8, column=0, pady=10)

transform_button = tk.Button(root, text="Transformar", command=transform_file)
transform_button.grid(row=8, column=1, pady=10)

quit_button = tk.Button(root, text="Sair", command=quit_app)
quit_button.grid(row=8, column=2, pady=10)

root.mainloop()
