# -*- coding: utf-8 -*-
"""
Módulo para recolectar y almacenar datos históricos de precios de criptomonedas
utilizando la API de CoinGecko.
"""

import os
import time
import json
import pandas as pd
import requests
from datetime import datetime
import pyarrow as pa
import pyarrow.parquet as pq
import logging
from apscheduler.schedulers.background import BackgroundScheduler

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("price_collector.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("price_collector")

# Configuración de directorios
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'historical_data')
PARQUET_FILE = os.path.join(DATA_DIR, 'price_history.parquet')
MAPPING_FILE = os.path.join(DATA_DIR, 'coingecko_mapping.json')

# Asegurar que el directorio de datos existe
os.makedirs(DATA_DIR, exist_ok=True)

# Configuración de la API de CoinGecko
COINGECKO_API_URL = "https://api.coingecko.com/api/v3"
API_KEY = "TU_API_KEY_AQUI"  # Reemplaza con tu API key de CoinGecko

# Límites de la API gratuita de CoinGecko
# 10-30 llamadas por minuto dependiendo del endpoint
API_RATE_LIMIT = 10  # llamadas por minuto
SLEEP_TIME = 60 / API_RATE_LIMIT  # segundos entre llamadas

class PriceCollector:
    def __init__(self):
        self.coin_mapping = self._load_coin_mapping()
        self.price_history = self._load_price_history()
        self.scheduler = BackgroundScheduler()
        
    def _load_coin_mapping(self):
        """Carga el mapeo de símbolos a IDs de CoinGecko"""
        if os.path.exists(MAPPING_FILE):
            with open(MAPPING_FILE, 'r') as f:
                return json.load(f)
        return {}
    
    def _save_coin_mapping(self):
        """Guarda el mapeo de símbolos a IDs de CoinGecko"""
        with open(MAPPING_FILE, 'w') as f:
            json.dump(self.coin_mapping, f, indent=2)
    
    def _load_price_history(self):
        """Carga el historial de precios desde el archivo parquet"""
        if os.path.exists(PARQUET_FILE):
            return pd.read_parquet(PARQUET_FILE)
        return pd.DataFrame(columns=['timestamp', 'symbol', 'price_usd', 'market_cap', 'volume_24h'])
    
    def _save_price_history(self):
        """Guarda el historial de precios en el archivo parquet"""
        self.price_history.to_parquet(PARQUET_FILE, index=False)
        logger.info(f"Datos guardados en {PARQUET_FILE}")
    
    def update_coin_mapping(self, symbols):
        """Actualiza el mapeo de símbolos a IDs de CoinGecko"""
        missing_symbols = [s for s in symbols if s not in self.coin_mapping]
        
        if not missing_symbols:
            return
        
        logger.info(f"Actualizando mapeo para {len(missing_symbols)} símbolos")
        
        try:
            # Obtener lista completa de monedas de CoinGecko
            response = requests.get(
                f"{COINGECKO_API_URL}/coins/list",
                headers={"X-CG-Pro-API-Key": API_KEY} if API_KEY else {}
            )
            response.raise_for_status()
            all_coins = response.json()
            
            # Crear un diccionario para búsqueda rápida
            coin_dict = {coin['symbol'].upper(): coin['id'] for coin in all_coins}
            
            # Actualizar el mapeo
            for symbol in missing_symbols:
                if symbol.upper() in coin_dict:
                    self.coin_mapping[symbol] = coin_dict[symbol.upper()]
                    logger.info(f"Mapeo encontrado: {symbol} -> {coin_dict[symbol.upper()]}")
                else:
                    logger.warning(f"No se encontró mapeo para {symbol}")
            
            # Guardar el mapeo actualizado
            self._save_coin_mapping()
            
        except Exception as e:
            logger.error(f"Error al actualizar el mapeo: {str(e)}")
    
    def collect_current_prices(self, symbols):
        """Recolecta los precios actuales para los símbolos dados"""
        timestamp = datetime.now()
        new_data = []
        
        # Asegurar que tenemos el mapeo para todos los símbolos
        self.update_coin_mapping(symbols)
        
        # Agrupar símbolos en lotes para reducir el número de llamadas a la API
        mapped_symbols = [s for s in symbols if s in self.coin_mapping]
        symbol_batches = [mapped_symbols[i:i+50] for i in range(0, len(mapped_symbols), 50)]
        
        for batch in symbol_batches:
            try:
                # Obtener IDs de CoinGecko para este lote
                coin_ids = [self.coin_mapping[s] for s in batch]
                ids_param = ",".join(coin_ids)
                
                # Llamar a la API de CoinGecko
                response = requests.get(
                    f"{COINGECKO_API_URL}/coins/markets",
                    params={
                        "vs_currency": "usd",
                        "ids": ids_param,
                        "per_page": 250,
                        "page": 1
                    },
                    headers={"X-CG-Pro-API-Key": API_KEY} if API_KEY else {}
                )
                response.raise_for_status()
                coins_data = response.json()
                
                # Procesar los datos recibidos
                for coin in coins_data:
                    # Encontrar el símbolo original (preservando mayúsculas/minúsculas)
                    original_symbol = next((s for s in batch if self.coin_mapping[s] == coin['id']), None)
                    if original_symbol:
                        new_data.append({
                            'timestamp': timestamp,
                            'symbol': original_symbol,
                            'price_usd': coin['current_price'],
                            'market_cap': coin['market_cap'],
                            'volume_24h': coin['total_volume']
                        })
                
                # Respetar el límite de tasa de la API
                time.sleep(SLEEP_TIME)
                
            except Exception as e:
                logger.error(f"Error al recolectar precios para lote: {str(e)}")
        
        # Añadir los nuevos datos al DataFrame
        if new_data:
            new_df = pd.DataFrame(new_data)
            self.price_history = pd.concat([self.price_history, new_df], ignore_index=True)
            logger.info(f"Recolectados {len(new_data)} precios a las {timestamp}")
            
            # Guardar los datos actualizados
            self._save_price_history()
    
    def start_scheduled_collection(self, symbols, interval_minutes=5):
        """Inicia la recolección programada de precios"""
        logger.info(f"Iniciando recolección programada cada {interval_minutes} minutos")
        
        # Configurar el trabajo programado
        self.scheduler.add_job(
            self.collect_current_prices,
            'interval',
            minutes=interval_minutes,
            args=[symbols]
        )
        
        # Iniciar el planificador
        self.scheduler.start()
        
        # Ejecutar una recolección inicial inmediatamente
        self.collect_current_prices(symbols)
    
    def stop_scheduled_collection(self):
        """Detiene la recolección programada de precios"""
        self.scheduler.shutdown()
        logger.info("Recolección programada detenida")
    
    def get_price_history(self, symbol=None, start_date=None, end_date=None):
        """Obtiene el historial de precios filtrado por símbolo y fechas"""
        df = self.price_history.copy()
        
        if symbol:
            df = df[df['symbol'] == symbol]
        
        if start_date:
            df = df[df['timestamp'] >= pd.to_datetime(start_date)]
        
        if end_date:
            df = df[df['timestamp'] <= pd.to_datetime(end_date)]
        
        return df.sort_values('timestamp')

# Instancia global del recolector de precios
price_collector = None

def initialize_price_collector(symbols, interval_minutes=5):
    """Inicializa el recolector de precios global"""
    global price_collector
    if price_collector is None:
        price_collector = PriceCollector()
        price_collector.start_scheduled_collection(symbols, interval_minutes)
    return price_collector

def get_price_collector():
    """Obtiene la instancia global del recolector de precios"""
    global price_collector
    if price_collector is None:
        raise RuntimeError("El recolector de precios no ha sido inicializado")
    return price_collector

# Función para pruebas
if __name__ == "__main__":
    # Lista de ejemplo de símbolos para probar
    test_symbols = ["BTC", "ETH", "SOL", "LINK", "DOT"]
    
    # Inicializar y probar el recolector
    collector = PriceCollector()
    collector.collect_current_prices(test_symbols)
    
    # Mostrar los datos recolectados
    print(collector.price_history) 