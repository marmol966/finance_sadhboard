# -*- coding: utf-8 -*-
"""
Created on Sun Mar 23 22:46:27 2025

@author: mmtar96
"""

import pandas as pd
from flask import Flask, jsonify
from flask_cors import CORS
import os
from datetime import datetime, timedelta
import random

app = Flask(__name__)
CORS(app, resources={r"/portfolio/*": {"origins": "*"}})

# Obtener el directorio del script actual
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Estado financiero_v6.1.xlsx')

# Cache variables
cached_data = {}
last_update = None
UPDATE_INTERVAL = timedelta(minutes=5)  # Actualizar datos cada 5 minutos

def should_update_cache():
    return last_update is None or datetime.now() - last_update > UPDATE_INTERVAL

def process_data():
    global cached_data, last_update
    
    dataset = pd.read_excel(EXCEL_PATH, sheet_name='datasetv1').iloc[:,0:14]
    dataset_inv = pd.read_excel(EXCEL_PATH, sheet_name='datasetv1').iloc[:1,14:18]
    
    # Procesar datos de cripto_old
    cripto_old = dataset[['Tipo', 'Precio', 'Activo', 'COld_Ud', 'COld_$']]
    cripto_old.set_index('Activo', inplace=True)
    cripto_old['Porcentaje'] = (cripto_old['COld_$'] / cripto_old['COld_$'].sum()) * 100
    
    # Procesar datos de cripto_new
    cripto_new = dataset[['Tipo', 'Precio', 'Activo', 'CNew_Ud', 'CNew_$']]
    cripto_new.set_index('Activo', inplace=True)
    cripto_new['Porcentaje'] = (cripto_new['CNew_$'] / cripto_new['CNew_$'].sum()) * 100
    
    # Procesar datos de defi
    defi = dataset[['Tipo', 'Precio', 'Activo', 'DEFI_Ud', 'DEFI_$']]
    defi.set_index('Activo', inplace=True)
    defi['Porcentaje'] = (defi['DEFI_$'] / defi['DEFI_$'].sum()) * 100
    
    # Procesar datos de hype
    hype = dataset[['Tipo', 'Precio', 'Activo', 'HYPE_Ud', 'HYPE_$']]
    hype.set_index('Activo', inplace=True)
    hype['Porcentaje'] = (hype['HYPE_$'] / hype['HYPE_$'].sum()) * 100
    
    # Calcular resumen de carteras
    resumen_carteras = pd.DataFrame({
        'Cripto Old': [
            cripto_old['COld_$'].sum(),
            dataset_inv['Cripto Old'][0],
            cripto_old[cripto_old['COld_$'] != 0]['COld_$'].count()
        ],
        'Cripto New': [
            cripto_new['CNew_$'].sum(),
            dataset_inv['Cripto New'][0],
            cripto_new[cripto_new['CNew_$'] != 0]['CNew_$'].count()
        ],
        'DEFI': [
            defi['DEFI_$'].sum(),
            dataset_inv['DEFI'][0],
            defi[defi['DEFI_$'] != 0]['DEFI_$'].count()
        ],
        'HYPE': [
            hype['HYPE_$'].sum(),
            dataset_inv['HYPE'][0],
            hype[hype['HYPE_$'] != 0]['HYPE_$'].count()
        ]
    })
    resumen_carteras.index = ['Capital Actual', 'Capital Invertido', 'N Pos']
    
    # Calcular resumen por tipo
    resumen_tipo = dataset.groupby('Tipo')['$ Total'].sum().reset_index()
    resumen_tipo.set_index('Tipo', inplace=True)
    resumen_tipo['Porcentaje'] = (resumen_tipo['$ Total'] / resumen_tipo['$ Total'].sum()) * 100
    resumen_tipo['N Pos'] = dataset[dataset['$ Total'] != 0].groupby('Tipo')['$ Total'].count()
    
    # Actualizar caché
    cached_data.update({
        'dataset': dataset,
        'dataset_inv': dataset_inv,
        'cripto_old': cripto_old,
        'cripto_new': cripto_new,
        'defi': defi,
        'hype': hype,
        'resumen_carteras': resumen_carteras,
        'resumen_tipos': resumen_tipo
    })
    last_update = datetime.now()

def get_cached_data():
    if should_update_cache():
        process_data()
    
    # Convertir DataFrames a diccionarios
    return {
        'dataset': cached_data['dataset'].to_dict(),
        'dataset_inv': cached_data['dataset_inv'].to_dict(),
        'cripto_old': cached_data['cripto_old'].to_dict(),
        'cripto_new': cached_data['cripto_new'].to_dict(),
        'defi': cached_data['defi'].to_dict(),
        'hype': cached_data['hype'].to_dict(),
        'resumen_carteras': cached_data['resumen_carteras'].to_dict(),
        'resumen_tipos': cached_data['resumen_tipos'].to_dict()
    }

@app.route('/portfolio/all', methods=['GET'])
def get_all_data():
    """Endpoint para obtener todos los datos (mantiene compatibilidad con versión anterior)"""
    print("Solicitud recibida en /portfolio/all")
    try:
        data = get_cached_data()
        print("Datos procesados correctamente")
        return jsonify(data)
    except Exception as e:
        print(f"Error al procesar datos: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/portfolio/dataset_complete', methods=['GET'])
def get_dataset_complete():
    """Endpoint para obtener solo el resumen de carteras"""
    return jsonify(get_cached_data()['dataset'])

@app.route('/portfolio/dataset_inv', methods=['GET'])
def get_dataset_inv():
    """Endpoint para obtener solo el resumen de carteras"""
    return jsonify(get_cached_data()['dataset_inv'])


@app.route('/portfolio/resumen-carteras', methods=['GET'])
def get_resumen_carteras():
    """Endpoint para obtener solo el resumen de carteras"""
    return jsonify(get_cached_data()['resumen_carteras'])

@app.route('/portfolio/resumen-tipos', methods=['GET'])
def get_resumen_tipos():
    """Endpoint para obtener solo el resumen por tipos"""
    return jsonify(get_cached_data()['resumen_tipos'])

@app.route('/portfolio/cripto-old', methods=['GET'])
def get_cripto_old():
    """Endpoint para obtener datos de cripto old"""
    return jsonify(get_cached_data()['cripto_old'])

@app.route('/portfolio/cripto-new', methods=['GET'])
def get_cripto_new():
    """Endpoint para obtener datos de cripto new"""
    return jsonify(get_cached_data()['cripto_new'])

@app.route('/portfolio/defi', methods=['GET'])
def get_defi():
    """Endpoint para obtener datos de DEFI"""
    return jsonify(get_cached_data()['defi'])

@app.route('/portfolio/hype', methods=['GET'])
def get_hype():
    """Endpoint para obtener datos de HYPE"""
    return jsonify(get_cached_data()['hype'])

@app.route('/portfolio/bubble-map', methods=['GET'])
def get_bubble_map_data():
    """Endpoint para obtener datos para el bubble map de todas las posiciones"""
    try:
        if should_update_cache():
            process_data()
        
        # Filtrar solo las posiciones con valor
        dataset = cached_data['dataset']
        positions = dataset[dataset['$ Total'] > 0].copy()
        
        # Preparar los datos para el bubble map
        bubble_data = positions[['Activo', '$ Total', 'Tipo']].copy()
        
        # Calcular el porcentaje de cada posición respecto al total
        total_value = bubble_data['$ Total'].sum()
        bubble_data['Porcentaje'] = (bubble_data['$ Total'] / total_value * 100).round(2)
        
        # Añadir parámetros para la animación flotante
        for i, row in bubble_data.iterrows():
            # Velocidad de movimiento (valores pequeños para movimiento lento)
            bubble_data.loc[i, 'vx'] = random.uniform(-0.5, 0.5)
            bubble_data.loc[i, 'vy'] = random.uniform(-0.5, 0.5)
            
            # Posición inicial aleatoria (el frontend puede ajustar estos valores según sus dimensiones)
            bubble_data.loc[i, 'x'] = random.uniform(10, 90)  # porcentaje del ancho
            bubble_data.loc[i, 'y'] = random.uniform(10, 90)  # porcentaje del alto
            
            # Parámetros para movimiento orgánico
            bubble_data.loc[i, 'amplitude'] = random.uniform(1, 5)  # amplitud de oscilación
            bubble_data.loc[i, 'frequency'] = random.uniform(0.001, 0.005)  # frecuencia de oscilación
            
            # Color basado en el tipo de activo (el frontend puede usar esta información)
            # Asignar un grupo de color basado en el tipo
            bubble_data.loc[i, 'colorGroup'] = bubble_data.loc[i, 'Tipo']
        
        # Ordenar por valor descendente
        bubble_data = bubble_data.sort_values('$ Total', ascending=False)
        
        return jsonify(bubble_data.to_dict(orient='records'))
    except Exception as e:
        print(f"Error al procesar datos para bubble map: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/test', methods=['GET'])
def test():
    """Endpoint de prueba"""
    return jsonify({"status": "ok", "message": "Backend funcionando correctamente"})

# Inicializar datos al arrancar el servidor
process_data()

if __name__ == '__main__':
    app.run(debug=True, port=5000)