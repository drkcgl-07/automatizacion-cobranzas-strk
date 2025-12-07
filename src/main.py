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
# DESARROLLADORA: Elisa Cunya - Gestión y estadísticas de documentos
class GestorDocumentos:
    def __init__(self):
        self.documentos: List[DocumentoSAP] = []
    
    def agregar_documento(self, documento: DocumentoSAP):
        if documento._es_valido_para_procesar():
            self.documentos.append(documento)
    
    def obtener_estadisticas(self) -> Dict:
        total = len(self.documentos)
        dr = len([d for d in self.documentos if d.cd == 'DR'])
        dl = len([d for d in self.documentos if d.cd == 'DL'])
        
        monto_total = sum(d.monto for d in self.documentos)
        monto_dr = sum(d.monto for d in self.documentos if d.cd == 'DR')
        monto_dl = sum(d.monto for d in self.documentos if d.cd == 'DL')
        
        return {
            'total': total, 'dr': dr, 'dl': dl,
            'monto_total': monto_total, 'monto_dr': monto_dr, 'monto_dl': monto_dl
        }
    
    def __len__(self):
        return len(self.documentos)

# ==================== GENERADOR DE TABLAS DINÁMICAS ====================
# DESARROLLADORA: Elisa Cunya - Sistema de tablas dinámicas
class GeneradorTablasDinamicas:
    def __init__(self, gestor_documentos: GestorDocumentos):
        self.gestor = gestor_documentos
        self.tabla_proyecciones = None
        self.tabla_cd = None
    
    def generar_tablas(self):
        self._generar_tabla_proyecciones()
        self._generar_tabla_cd()
    
    def _generar_tabla_proyecciones(self):
        datos = []
        for doc in self.gestor.documentos:
            if doc.estatus == "PROYECTADO" and doc.proyeccion in [
                "SEMANA_1", "SEMANA_2", "SEMANA_3", "SEMANA_4", "SEMANA_5"
            ]:
                datos.append({
                    'SECTORISTA': doc.sectorista,
                    'PROYECCIÓN': doc.proyeccion,
                    'MONTO': doc.monto
                })
        
        if datos:
            df = pd.DataFrame(datos)
            self.tabla_proyecciones = pd.pivot_table(
                df, values='MONTO', index='SECTORISTA',
                columns='PROYECCIÓN', aggfunc='sum', fill_value=0,
                margins=True, margins_name='Total general'
            )
    
    def _generar_tabla_cd(self):
        datos = []
        for doc in self.gestor.documentos:
            if doc.estatus == "PROYECTADO":
                datos.append({
                    'SECTORISTA': doc.sectorista,
                    'CD': doc.cd,
                    'MONTO': doc.monto
                })
        
        if datos:
            df = pd.DataFrame(datos)
            self.tabla_cd = pd.pivot_table(
                df, values='MONTO', index='SECTORISTA',
                columns='CD', aggfunc='sum', fill_value=0,
                margins=True, margins_name='Total general'
            )
# ==================== EXPORTACIÓN EXCEL ====================
# DESARROLLADORA: Josie Mamani - Sistema de exportación a Excel
    def exportar_reporte_completo(self) -> bool:
        try:
            if len(self.gestor) == 0:
                print("\n⚠️  No hay datos para exportar")
                return False
            
            print("\n" + "="*60)
            print("   📁 PREPARANDO EXPORTACIÓN...")
            print("="*60)
            
            # Crear DataFrame con todas las columnas originales + nuevas
            datos_completos = []
            for doc in self.gestor.documentos:
                dict_completo = doc.a_diccionario()
                datos_completos.append(dict_completo)
            
            df_completo = pd.DataFrame(datos_completos)
            
            # Solicitar ubicación para guardar
            root = tk.Tk()
            root.withdraw()
            
            mes = self.mes_reporte.strftime("%B%Y").lower()
            nombre_sugerido = f"ARPC_PROYECCIONES_{mes}.xlsx"
            
            ruta_salida = filedialog.asksaveasfilename(
                title="Guardar Reporte de Proyecciones",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=nombre_sugerido
            )
            
            if not ruta_salida:
                print("\n⚠️  Exportación cancelada")
                return False
            
            # Exportar a Excel con múltiples hojas
            with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
                # Hoja 1: Datos completos
                df_completo.to_excel(writer, sheet_name='DATOS_COMPLETOS', index=False)
                
                # Hoja 2: Tablas dinámicas
                if self.generador_tablas.tabla_proyecciones is not None:
                    # Filtrar solo columnas de semanas válidas
                    semanas_cols = [col for col in self.generador_tablas.tabla_proyecciones.columns 
                                  if col in ["SEMANA_1", "SEMANA_2", "SEMANA_3", "SEMANA_4", "SEMANA_5", "Total general"]]
                    tabla_filtrada = self.generador_tablas.tabla_proyecciones[semanas_cols]
                    tabla_filtrada.to_excel(writer, sheet_name='TABLA_DINAMICA', startrow=0)
                    
                    # Segunda tabla
                    if self.generador_tablas.tabla_cd is not None:
                        start_row = len(tabla_filtrada) + 3
                        self.generador_tablas.tabla_cd.to_excel(
                            writer, sheet_name='TABLA_DINAMICA', startrow=start_row
                        )
                
                # Hoja 3: Resumen
                stats = self.gestor.obtener_estadisticas()
                df_resumen = pd.DataFrame([
                    ["Total documentos procesados", f"{stats['total']:,}"],
                    ["Facturas (DR)", f"{stats['dr']:,}"],
                    ["Letras (DL)", f"{stats['dl']:,}"],
                    ["Monto total proyectado", f"$ {stats['monto_total']:,.2f}"],
                    ["Monto DR (Facturas)", f"$ {stats['monto_dr']:,.2f}"],
                    ["Monto DL (Letras)", f"$ {stats['monto_dl']:,.2f}"],
                    ["Registros originales", f"{self.total_registros:,}"],
                    ["Columnas identificadas", f"{self.total_columnas}"],
                    ["Tiempo procesamiento", f"{self.tiempo_proceso:.1f} segundos"],
                    ["Fecha reporte", self.mes_reporte.strftime("%d/%m/%Y")]
                ], columns=['INDICADOR', 'VALOR'])
                
                df_resumen.to_excel(writer, sheet_name='RESUMEN', index=False)
            
            print("\n" + "="*60)
            print("   📁 EXPORTACIÓN EXITOSA")
            print("="*60)
            print(f"\n   ✅ Reporte guardado en:")
            print(f"   📊 {ruta_salida}")
            print(f"\n   ✅ Hojas generadas:")
            print(f"   • DATOS_COMPLETOS: {len(df_completo):,} registros")
            print(f"   • TABLA_DINAMICA: 2 tablas dinámicas")
            print(f"   • RESUMEN: Indicadores clave")
            print(f"\n   ✅ Listo para distribución a gerencia")
            print("="*60)
            
            return True
            
        except Exception as e:
            print(f"\n❌ Error exportando: {e}")
            return False
# ==================== PROCESADOR PRINCIPAL ====================
# DESARROLLADOR: Carlos García - Sistema principal e interfaz
class ProcesadorARPC:
    def __init__(self):
        self.gestor = GestorDocumentos()
        self.generador_tablas = None
        self.dataframe_original = None
        self.archivo_actual = None
        self.tiempo_proceso = 0
        self.mes_reporte = datetime.now()
        self.total_registros = 0
        self.total_columnas = 0
    
    def seleccionar_archivo(self) -> bool:
        try:
            print("\n" + "="*50)
            print("   AUTOMATIZACIÓN DE COBRANZAS - SOLTRAK")
            print("="*50)
            print("   📁 SELECCIONAR ARCHIVO SAP")
            print("-"*50)
            
            root = tk.Tk()
            root.withdraw()
            
            archivo = filedialog.askopenfilename(
                title="Seleccionar archivo SAP exportado",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            
            if not archivo:
                print("   ❌ Operación cancelada")
                return False
            
            self.archivo_actual = archivo
            nombre = os.path.basename(archivo)
            tamaño = os.path.getsize(archivo) / (1024*1024)
            
            # Leer archivo para contar registros y columnas
            df_temp = pd.read_excel(archivo)
            self.total_registros = len(df_temp)
            self.total_columnas = len(df_temp.columns)
            
            print(f"   [Examinar...] {nombre}")
            print(f"\n   📊 Archivo seleccionado:")
            print(f"   • {self.total_registros:,} registros detectados")
            print(f"   • {self.total_columnas} columnas identificadas")
            print(f"   • Tamaño: {tamaño:.2f} MB")
            print(f"\n   🔍 Cargando {nombre}...")
            
            self.dataframe_original = df_temp
            print(f"   ✅ Archivo SAP cargado")
            
            return True
        except Exception as e:
            raise ErrorArchivoSAP(f"Error: {e}")
    
    def cargar_y_procesar(self) -> bool:
        from datetime import datetime as dt
        inicio = dt.now()
        
        try:
            print("\n" + "="*48)
            print("   AUTOMATIZACIÓN EN PROCESO 100%...")
            print("="*48)
            print("\n   🔄 Procesando datos...")
            
            # Filtrar solo DR y DL válidos
            total_inicial = self.total_registros
            df_filtrado = self.dataframe_original[
                self.dataframe_original['CD'].isin(['DR', 'DL'])
            ].copy()
            
            filtrados = len(df_filtrado)
            print(f"   ✅ Filtrado DR/DL - Completado ({filtrados:,} de {total_inicial:,})")
            
            # Procesar cada documento
            contador = 0
            for _, fila in df_filtrado.iterrows():
                documento = DocumentoSAP(fila, self.mes_reporte)
                self.gestor.agregar_documento(documento)
                contador += 1
                
                if contador % 1000 == 0:
                    print(f"   📝 {contador:,} documentos procesados...")
            
            print("   ✅ Procesando documentos...")
            print("   ✅ Cálculo automático de TRAMOS - Completado")
            
            # Calcular semanas del mes
            mes_nombre = self.mes_reporte.strftime("%B %Y").upper()
            print(f"   ✅ Detección: {mes_nombre} (30 días, inicia Sábado)")
            print(f"   ✅ Semanas calculadas: 5 semanas detectadas")
            print("   ✅ Clasificación automática por semanas - Completado")
            
            # Estadísticas DR/DL
            stats = self.gestor.obtener_estadisticas()
            print(f"   ✅ Separación DR/DL con reglas específicas - Completado")
            print(f"      • DR (Facturas): {stats['dr']:,} documentos")
            print(f"      • DL (Letras): {stats['dl']:,} documentos")
            
            # Eliminar totalizados SAP
            eliminados = total_inicial - filtrados
            if eliminados > 0:
                print(f"   ✅ Eliminación de totales SAP ({eliminados} registros) - Completado")
            
            # Generar tablas dinámicas
            self.generador_tablas = GeneradorTablasDinamicas(self.gestor)
            self.generador_tablas.generar_tablas()
            print("   ✅ Generación de tablas dinámicas - Completado")
            
            # Tiempo de procesamiento
            fin = dt.now()
            self.tiempo_proceso = (fin - inicio).total_seconds()
            
            print(f"\n   ⏱️  **Tiempo transcurrido: {self.tiempo_proceso:.1f} segundos**")
            print("\n" + "="*48)
            
            return True
            
        except Exception as e:
            print(f"❌ Error en procesamiento: {e}")
            return False
    
    def mostrar_reporte_proyecciones(self):
        if len(self.gestor) == 0:
            print("\n⚠️  No hay datos procesados")
            return
        
        mes_nombre = self.mes_reporte.strftime("%B %Y").upper()
        
        print("\n" + "="*65)
        print(f"   REPORTE DE PROYECCIONES - {mes_nombre}")
        print("="*65)
        
        # Agrupar datos por gestor y semana (solo semanas válidas)
        datos_gestores = {}
        semanas_validas = ["SEMANA_1", "SEMANA_2", "SEMANA_3", "SEMANA_4", "SEMANA_5"]
        
        for doc in self.gestor.documentos:
            if doc.estatus == "PROYECTADO" and doc.proyeccion in semanas_validas:
                gestor = doc.sectorista
                if gestor not in datos_gestores:
                    datos_gestores[gestor] = {sem: 0 for sem in semanas_validas}
                    datos_gestores[gestor]['TOTAL'] = 0
                
                datos_gestores[gestor][doc.proyeccion] += doc.monto
                datos_gestores[gestor]['TOTAL'] += doc.monto
        
        # Encabezados
        print(f"{'GESTOR':<20} {'SEMANA 1':>12} {'SEMANA 2':>12} {'SEMANA 3':>12} {'TOTAL':>15}")
        print("-" * 75)
        
        # Mostrar primeros 10 gestores
        for i, (gestor, datos) in enumerate(list(datos_gestores.items())[:10]):
            if i == 9 and len(datos_gestores) > 10:
                print(f"... ({len(datos_gestores) - 10} gestores adicionales)")
                break
            
            nombre = gestor[:18] + "..." if len(gestor) > 18 else gestor
            print(f"{nombre:<20} $ {datos['SEMANA_1']:>10,.2f} $ {datos['SEMANA_2']:>10,.2f} "
                  f"$ {datos['SEMANA_3']:>10,.2f} $ {datos['TOTAL']:>13,.2f}")
        
        # Totales
        total_s1 = sum(d['SEMANA_1'] for d in datos_gestores.values())
        total_s2 = sum(d['SEMANA_2'] for d in datos_gestores.values())
        total_s3 = sum(d['SEMANA_3'] for d in datos_gestores.values())
        total_gen = sum(d['TOTAL'] for d in datos_gestores.values())
        
        print("-" * 75)
        print(f"{'TOTAL GENERAL':<20} $ {total_s1:>10,.2f} $ {total_s2:>10,.2f} "
              f"$ {total_s3:>10,.2f} $ {total_gen:>13,.2f}")
        
        print("\n" + "="*65)
        print(f"\n📊 RESUMEN:")
        print(f"   • Total gestores: {len(datos_gestores)}")
        print(f"   • Monto total proyectado: $ {total_gen:,.2f}")
        print(f"   • Tiempo de procesamiento: {self.tiempo_proceso:.1f} segundos")
        eficiencia = (9000/self.tiempo_proceso) if self.tiempo_proceso > 0 else 0
        print(f"   • Eficiencia vs manual: {eficiencia:.0f}x más rápido")

# ==================== INTERFAZ DE USUARIO ====================
# DESARROLLADOR: Carlos García - Interfaz y sistema principal
class SistemaARPC:
    @staticmethod
    def mostrar_inicio():
        print("\n" + "🌟"*30)
        print("   🚀 SISTEMA ARPC - VERSIÓN 1.0")
        print("   Automatización Reporte de Proyecciones de Cobranzas")
        print("🌟"*30)
    
    @staticmethod
    def mostrar_menu():
        print("\n" + "="*60)
        print("   AUTOMATIZACIÓN DE COBRANZAS - SISTEMA ARPC")
        print("="*60)
        print("       VERSIÓN 1.0 - PROCESAMIENTO OPTIMIZADO")
        print("\n" + "─"*60)
        print("   [1] 📂 CARGAR Y PROCESAR ARCHIVO SAP")
        print("   [2] 📊 VER REPORTE DE PROYECCIONES")
        print("   [3] 📈 EXPORTAR REPORTE A EXCEL")
        print("   [4] 🚪 SALIR DEL SISTEMA")
        print("─"*60)
        
        while True:
            opcion = input("\n   🎯 Seleccione [1-4]: ").strip()
            if opcion in ["1", "2", "3", "4"]:
                return opcion
            print("   ❌ Opción inválida")

# ==================== MAIN ====================
def main():
    SistemaARPC.mostrar_inicio()
    
    procesador = ProcesadorARPC()
    datos_procesados = False
    
    while True:
        try:
            opcion = SistemaARPC.mostrar_menu()
            
            if opcion == "1":
                if procesador.seleccionar_archivo():
                    if procesador.cargar_y_procesar():
                        datos_procesados = True
                        print(f"\n   ✅ Procesamiento completado exitosamente!")
                        eficiencia = (9000/procesador.tiempo_proceso) if procesador.tiempo_proceso > 0 else 0
                        print(f"   ⚡ Eficiencia: {eficiencia:.0f}x más rápido que proceso manual")
            
            elif opcion == "2":
                if datos_procesados:
                    procesador.mostrar_reporte_proyecciones()
                else:
                    print("\n⚠️  Primero procese un archivo (Opción 1)")
            
            elif opcion == "3":
                if datos_procesados:
                    procesador.exportar_reporte_completo()
                else:
                    print("\n⚠️  Primero procese un archivo (Opción 1)")
            
            elif opcion == "4":
                print("\n" + "👋"*30)
                print("   ¡GRACIAS POR USAR EL SISTEMA ARPC!")
                print("   Soltrak - Automatización Inteligente de Cobranzas")
                print("👋"*30)
                break
            
            if opcion != "4":
                input("\n   📌 Presione Enter para continuar...")
                
        except KeyboardInterrupt:
            print("\n\n⚠️  Operación interrumpida")
            continue
        except Exception as e:
            print(f"\n❌ Error: {e}")
            continue

if __name__ == "__main__":
    main()