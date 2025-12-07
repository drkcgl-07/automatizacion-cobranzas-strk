"""
SISTEMA ARPC - AUTOMATIZACIÓN DEL REPORTE DE PROYECCIONES DE COBRANZAS
UNIVERSIDAD PERUANA DE CIENCIAS APLICADAS - FUNDAMENTOS DE PROGRAMACIÓN 2
"""

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from abc import ABC, abstractmethod
from typing import List, Dict
import sys

# ==================== EXCEPCIONES PERSONALIZADAS ====================
class ErrorProcesamiento(Exception):
    pass

class ErrorArchivoSAP(ErrorProcesamiento):
    pass

# ==================== INTERFACES ====================
class IProcesable(ABC):
    @abstractmethod
    def procesar(self) -> Dict:
        pass

class IExportable(ABC):
    @abstractmethod
    def a_diccionario(self) -> Dict:
        pass

# ==================== CLASE BASE ====================
class DocumentoBase:
    def __init__(self, datos: pd.Series):
        self._datos = datos
    
    def validar(self) -> bool:
        return True

# ==================== CLASE DOCUMENTO SAP ====================
# DESARROLLADOR: Josemir Poma - Implementación de algoritmos de cálculo
class DocumentoSAP(DocumentoBase, IProcesable, IExportable):
    
    def __init__(self, datos_fila: pd.Series, mes_reporte: datetime):
        super().__init__(datos_fila)
        self.mes_reporte = mes_reporte
        self._inicializar_atributos()
    
    def _inicializar_atributos(self):
        self.cd = self._obtener_valor('CD')
        self.dias_mora = self._obtener_valor('Mora', int, 0)
        self.sectorista = self._obtener_valor('Sectorista', default='SIN GESTOR')
        self.monto = self._obtener_valor('Imp. ML2 Pend.', float, 0.0)
        self.vencimiento = self._obtener_valor('Vencimiento neto')
        self.base_plazo = self._obtener_valor('Base p.plazo pago')
        self.ref_letra = self._obtener_valor('Ref. Letra')
        self.clv_ref = self._obtener_valor('Clv.ref.(cabecera) 2')
        
        self.tramo = ""
        self.estatus = ""
        self.proyeccion = ""
        self._calcular_atributos()
    
    def _obtener_valor(self, columna: str, tipo=str, default=None):
        try:
            valor = self._datos.get(columna, default)
            if pd.isna(valor):
                return default
            return tipo(valor) if valor != "" else default
        except:
            return default
    
    def procesar(self) -> Dict:
        return {
            'tramo': self._calcular_tramo(),
            'estatus': self._calcular_estatus(),
            'proyeccion': self._calcular_proyeccion(),
            'es_valido': self._es_valido_para_procesar()
        }
    
    def _calcular_atributos(self):
        resultados = self.procesar()
        self.tramo = resultados['tramo']
        self.estatus = resultados['estatus']
        self.proyeccion = resultados['proyeccion']
    
    def _calcular_tramo(self) -> str:
        if self.dias_mora <= 0:
            return "Por Vencer"
        elif self.dias_mora < 31:
            return "1 a 30"
        elif self.dias_mora < 61:
            return "31 a 60"
        elif self.dias_mora < 91:
            return "61 a 90"
        elif self.dias_mora < 121:
            return "91 a 120"
        elif self.dias_mora < 181:
            return "121 a 180"
        elif self.dias_mora < 361:
            return "181 a 360"
        else:
            return "360+"
    
    def _calcular_estatus(self) -> str:
        if self.dias_mora > 60:
            return "EN GESTIÓN"
        elif self.dias_mora <= 0:
            return "POR VENCER"
        else:
            return "PROYECTADO"
    
    def _calcular_proyeccion(self) -> str:
        if self.cd == 'DL':
            return self._calcular_proyeccion_dl()
        else:
            return self._calcular_proyeccion_dr()
    
    def _calcular_proyeccion_dr(self) -> str:
        if self.dias_mora <= 0:
            return self._determinar_semana_por_fecha(self.vencimiento)
        elif self.dias_mora <= 7:
            return "SEMANA_1"
        elif self.dias_mora <= 14:
            return "SEMANA_2"
        elif self.dias_mora <= 21:
            return "SEMANA_3"
        elif self.dias_mora <= 28:
            return "SEMANA_4"
        else:
            return "SEMANA_5"
    
    def _calcular_proyeccion_dl(self) -> str:
        # Solo DL válidas (con referencias)
        if pd.isna(self.ref_letra) or self.ref_letra == "":
            return "NO_PROCESAR"
        
        try:
            fecha_base = self.base_plazo
            if isinstance(fecha_base, str):
                fecha_base = datetime.strptime(fecha_base, "%d/%m/%Y")
            
            # Lógica DL: mes anterior + 8 días
            from datetime import timedelta
            fecha_ajustada = fecha_base + timedelta(days=8)
            dia = fecha_ajustada.day
            
            if dia <= 7:
                return "SEMANA_1"
            elif dia <= 14:
                return "SEMANA_2"
            elif dia <= 21:
                return "SEMANA_3"
            elif dia <= 28:
                return "SEMANA_4"
            else:
                return "SEMANA_5"
        except:
            return "SEMANA_1"
    
    def _determinar_semana_por_fecha(self, fecha) -> str:
        try:
            if pd.isna(fecha):
                return "POR VENCER"
            
            if isinstance(fecha, str):
                fecha_obj = datetime.strptime(fecha, "%d/%m/%Y")
            else:
                fecha_obj = fecha
            
            dia = fecha_obj.day
            
            if dia <= 7:
                return "SEMANA_2"
            elif dia <= 14:
                return "SEMANA_3"
            elif dia <= 21:
                return "SEMANA_4"
            elif dia <= 28:
                return "SEMANA_5"
            else:
                return "SEMANA_5"
        except:
            return "POR VENCER"
    
    def _es_valido_para_procesar(self) -> bool:
        if pd.isna(self.sectorista):
            return False
        if self.cd not in ['DR', 'DL']:
            return False
        # Solo procesar DL con referencias válidas
        if self.cd == 'DL' and (pd.isna(self.ref_letra) or self.ref_letra == ""):
            return False
        return True
    
    def a_diccionario(self) -> Dict:
        dict_base = self._datos.to_dict()
        dict_base['TRAMO'] = self.tramo
        dict_base['ESTATUS 1'] = self.estatus
        dict_base['PROYECCIÓN'] = self.proyeccion
        return dict_base
    
    def __str__(self) -> str:
        return f"{self.cd} | {self.sectorista} | ${self.monto:,.2f}"

# ==================== GESTOR DE DOCUMENTOS ====================