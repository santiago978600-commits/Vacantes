import pandas as pd
import json
import os
import datetime
import re

EXCEL_PATH = os.environ.get("EXCEL_PATH", "Informe vacantes Auto.xlsx")
HTML_PATH = "index.html"
API_KEY = os.environ.get("GEMINI_API_KEY")

def clean_string(val):
    if pd.isna(val):
        return "N/A"
    return str(val).strip()

def process():
    if not os.path.exists(EXCEL_PATH):
        print(f"Error: No se encontro el archivo Excel en {EXCEL_PATH}")
        return

    print("Leyendo Excel...")
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
    except Exception as e:
        print(f"Error al abrir el Excel: {e}")
        return
    
    # Procesar Hoja 1 (Lista de vacantes)
    try:
        df1 = xl.parse("Hoja1", header=None)
        # Buscar la fila de encabezados (la que tiene 'CARGO')
        header_idx = 0
        for i, row in df1.iterrows():
            if 'CARGO' in row.values:
                header_idx = i
                break
        
        df1.columns = df1.iloc[header_idx]
        df1 = df1[header_idx + 1:].copy()
        df1 = df1.dropna(subset=['CARGO', 'ESTADO'])
    except Exception as e:
        print(f"Error al procesar Hoja1: {e}")
        return
    
    vacantes = []
    estados_counter = {}
    procesos_counter = {}
    
    total_activas = 0
    total_congeladas = 0
    total_vacantes_generales = 0
    
    # Categorías para el Funnel
    funnel_counts = {"Reclutamiento": 0, "Entrevista": 0, "Seleccion": 0}
    
    for _, row in df1.iterrows():
        v_cargo = str(row.get('CARGO', '')).strip()
        v_proceso = str(row.get('Proceso', '')).strip()
        v_psicologa = str(row.get('psicologa', '')).strip().title()
        v_estado = str(row.get('ESTADO', '')).strip()
        v_obs = str(row.get('OBSERVACIONES', 'N/A')).strip()
        v_planta = row.get('Planta Autorizada', 0)
        v_vinculacion = str(row.get('TIPO DE VINCULACION', 'N/A')).strip()
        
        try:
            tiempo = float(row.get('VACANTES /  TIEMPOS', 0))
            if pd.isna(tiempo): tiempo = 0.0
        except:
            tiempo = 0.0
        
        try:
            v_planta = float(row.get('Planta Autorizada', 0))
            if pd.isna(v_planta): v_planta = 0.0
        except:
            v_planta = 0.0

        total_vacantes_generales += tiempo

        vacantes.append({
            "proceso": v_proceso,
            "cargo": v_cargo,
            "estado": v_estado,
            "psicologa": v_psicologa,
            "obs": v_obs,
            "tiempo": tiempo,
            "planta": v_planta,
            "vinculacion": v_vinculacion
        })
        
        estado_low = v_estado.lower()
        if estado_low == 'congelado':
            total_congeladas += tiempo
        else:
            total_activas += tiempo
            estados_counter[v_estado] = estados_counter.get(v_estado, 0) + tiempo
            if v_proceso != 'N/A':
                procesos_counter[v_proceso] = procesos_counter.get(v_proceso, 0) + tiempo
                
            # Asignación al Funnel
            if 'reclutamiento' in estado_low or 'convocatoria' in estado_low:
                funnel_counts['Reclutamiento'] += tiempo
            elif 'entrevista' in estado_low or 'revisión' in estado_low or 'revision' in estado_low:
                funnel_counts['Entrevista'] += tiempo
            else: # En proceso, Seleccionado, etc.
                funnel_counts['Seleccion'] += tiempo

    # Procesar Alertas FPP (Normalmente Hoja3 o FPP)
    fpp_list = []
    sheet_fpp = "FPP" if "FPP" in xl.sheet_names else ("Hoja3" if "Hoja3" in xl.sheet_names else None)
    
    if sheet_fpp:
        try:
            df3 = xl.parse(sheet_fpp)
            # Normalizar nombres de columnas (quitar espacios raros, mayúsculas/minúsculas)
            df3.columns = [str(c).strip() for c in df3.columns]
            
            for _, row in df3.iterrows():
                nombre = clean_string(row.get('Nombre'))
                if nombre == 'N/A' or nombre == '': continue
                
                fpp_val = row.get('FPP')
                if pd.isna(fpp_val): continue
                
                if isinstance(fpp_val, pd.Timestamp):
                    fpp_str = fpp_val.strftime('%Y-%m-%d')
                else:
                    fpp_str = str(fpp_val)
                    
                fpp_list.append({
                    "nombre": nombre,
                    "cargo": clean_string(row.get('Cargo')),
                    "proceso": clean_string(row.get('Proceso')),
                    "fpp": fpp_str
                })
        except Exception as e:
            print(f"Error al procesar Hoja FPP: {e}")
    else:
        print("Aviso: No se encontró la hoja de FPP (Hoja3 o FPP).")
    
    # Extraer insights anteriores para preservarlos si no hay API_KEY
    old_ai_insight = "Análisis pendiente."
    try:
        if os.path.exists(HTML_PATH):
            with open(HTML_PATH, 'r', encoding='utf-8') as f:
                content = f.read()
                m = re.search(r'"aiInsight":\s*"(.*?)"', content)
                if m:
                    old_ai_insight = m.group(1).encode().decode('unicode_escape')
    except:
        pass

    ai_insight = old_ai_insight
    if API_KEY:
        try:
            import google.generativeai as genai
            print("Consultando a Gemini para Insights estratégicos...")
            genai.configure(api_key=API_KEY)
            model = genai.GenerativeModel('gemini-flash-latest') # Usando flash para velocidad
            
            # Resumen de observaciones críticas para la IA
            obs_criticas = [f"- {v['cargo']} ({v['proceso']}): {v['obs']}" 
                           for v in vacantes if v['obs'] != '' and v['estado'].lower() != 'congelado'][:15]
            
            prompt = f"""
            Como experto en Reclutamiento y Selección, analiza este resumen de vacantes activas y sus observaciones:
            {chr(10).join(obs_criticas)}
            
            Proporciona un único "Insight Ejecutivo" (máximo 250 caracteres) que identifique un riesgo crítico, 
            cuello de botella o una recomendación estratégica inmediata. Sé directo y profesional.
            """
            
            response = model.generate_content(prompt)
            ai_insight = response.text.strip().replace('\n', ' ')
        except Exception as e:
            ai_insight = f"Error en análisis AI: {str(e)}"

    top_procesos = sorted(procesos_counter.items(), key=lambda x: x[1], reverse=True)[:5]
    
    # KPIs Gerenciales
    indice_congelamiento = (total_congeladas / total_vacantes_generales * 100) if total_vacantes_generales > 0 else 0
    eficiencia_general = (funnel_counts['Seleccion'] / total_activas * 100) if total_activas > 0 else 0
    
    # Tendencia mock (distribuyendo el total activo en 6 meses ascendentes para el widget)
    base_trend = [0.1, 0.25, 0.4, 0.6, 0.85, 1.0]
    trend_data = [round(total_activas * factor) for factor in base_trend]

    kpis_gerenciales = {
        "indiceCongelamiento": round(indice_congelamiento, 1),
        "volumenActivo": total_activas,
        "capacidadTotal": total_vacantes_generales,
        "vacantesCongeladas": total_congeladas,
        "eficienciaGeneral": round(eficiencia_general, 1),
        "funnel": {
            "labels": ["Reclutamiento", "Entrevista", "Selección"],
            "data": [funnel_counts['Reclutamiento'], funnel_counts['Entrevista'], funnel_counts['Seleccion']]
        },
        "distribucionProcesos": {
            "names": [x[0][:20] + ".." if len(x[0]) > 20 else x[0] for x in top_procesos],
            "counts": [x[1] for x in top_procesos]
        },
        "tendencia": trend_data
    }

    # Procesar Aprendices
    kpis_aprendices = None
    sheet_aprendices = "Aprendices" if "Aprendices" in xl.sheet_names else None
    
    if sheet_aprendices:
        try:
            print("Procesando datos de Aprendices...")
            df_apr = xl.parse(sheet_aprendices)
            today = pd.Timestamp(datetime.datetime.now().date())
            
            # Cargar festivos para cálculo exacto de días hábiles
            holidays = []
            if "Hoja4" in xl.sheet_names:
                df_h = xl.parse("Hoja4")
                holidays = pd.to_datetime(df_h['Fecha'], errors='coerce').dropna().tolist()
            
            from pandas.tseries.offsets import CustomBusinessDay
            co_bday = CustomBusinessDay(holidays=holidays)
            
            df_apr['Fecha fin'] = pd.to_datetime(df_apr['Fecha fin'], errors='coerce')
            
            # Calculamos la Fecha Límite sumando exactamente 20 días hábiles a la Fecha fin
            df_apr['Limit_Date'] = df_apr['Fecha fin'].apply(lambda x: x + 20 * co_bday if pd.notnull(x) else pd.NaT)
            
            vigentes = len(df_apr[df_apr['Estado del contrato'] == 'VIGENTE'])
            # Un contrato es válido si es VIGENTE o si no ha pasado su Fecha Límite calculada
            validos_df = df_apr[(df_apr['Estado del contrato'] == 'VIGENTE') | 
                                (pd.notnull(df_apr['Limit_Date']) & (df_apr['Limit_Date'] >= today))]
            validos = len(validos_df)
            
            # Calculo Riesgo de Multa: Contratos no vigentes, cuyo limite no ha pasado, y falta poco (menos de 5 dias habiles exactos)
            riesgo_multa = len(validos_df[(validos_df['Estado del contrato'] != 'VIGENTE') & 
                                          (validos_df['Limit_Date'] >= today) & 
                                          (validos_df['Limit_Date'] <= today + 5 * co_bday)])
                                          
            next_3_months = today + pd.DateOffset(months=3)
            # Lista de próximos vencimientos (fechas limites)
            vencimientos = validos_df[(validos_df['Limit_Date'] >= today) & (validos_df['Limit_Date'] <= next_3_months)]['Limit_Date'].dt.strftime('%Y-%m-%d').tolist()
            
            old_ai_insight_apr = "Análisis pendiente."
            old_forecast_data = []
            try:
                if os.path.exists(HTML_PATH):
                    with open(HTML_PATH, 'r', encoding='utf-8') as f:
                        content = f.read()
                        m1 = re.search(r'"kpisAprendices":\s*\{.*?"insight":\s*"(.*?)",.*?"forecast":\s*(\[.*?\])\s*\}', content, re.DOTALL)
                        if m1:
                            old_ai_insight_apr = m1.group(1).encode().decode('unicode_escape')
                            import ast
                            # A simple way to get the old forecast if possible, else empty
                            try:
                                old_forecast_data = json.loads(m1.group(2))
                            except:
                                pass
            except:
                pass

            ai_insight_apr = old_ai_insight_apr
            forecast_data = old_forecast_data
            
            if API_KEY:
                try:
                    import google.generativeai as genai
                    genai.configure(api_key=API_KEY)
                    model_pro = genai.GenerativeModel('gemini-flash-latest')
                    
                    prompt_apr = f"""
                    Eres un planificador estratégico de talento.
                    Cuota objetivo: Mínimo 166 contratos de aprendizaje hábiles.
                    Estado actual: {vigentes} vigentes, {validos} válidos para cumplimiento.
                    Regla de Oro: Cuando un contrato termina, el cupo sigue siendo 'hábil' por 20 días hábiles. Si no se contrata un reemplazo en ese periodo, el cupo se pierde.
                    Calendario de Contratación: Solo se realizan ingresos los viernes de las primeras tres semanas del mes.
                    
                    Aquí tienes la lista de fechas límite de reemplazo (Fecha fin + 20 días hábiles) para los próximos meses:
                    {vencimientos}
                    
                    Basado en estos {validos} cupos y sus fechas de vencimiento, indica cuántos aprendices exactos debo contratar en cada uno de los próximos viernes hábiles para nunca bajar de 166 contratos válidos.
                    
                    Devuelve el resultado estrictamente en este formato JSON, sin markdown, sin texto adicional:
                    {{
                        "insight": "Un mensaje gerencial corto y directivo de max 150 caracteres sobre la estrategia de contratacion.",
                        "forecast": [
                            {{"fecha": "YYYY-MM-DD", "cantidad": N}},
                            {{"fecha": "YYYY-MM-DD", "cantidad": N}},
                            {{"fecha": "YYYY-MM-DD", "cantidad": N}}
                        ]
                    }}
                    """
                    
                    response_apr = model_pro.generate_content(prompt_apr)
                    
                    # Extraer JSON limpio
                    json_str = response_apr.text.strip()
                    if json_str.startswith("```json"): json_str = json_str[7:]
                    if json_str.endswith("```"): json_str = json_str[:-3]
                    
                    res_json = json.loads(json_str.strip())
                    ai_insight_apr = res_json.get("insight", "Análisis completado.")
                    forecast_data = res_json.get("forecast", [])
                except Exception as e:
                    print(f"Error en IA de Aprendices: {e}")
                    ai_insight_apr = f"Error IA: {e}"

            kpis_aprendices = {
                "vigentes": vigentes,
                "validos": validos,
                "meta": 166,
                "riesgoMulta": riesgo_multa,
                "insight": ai_insight_apr,
                "forecast": forecast_data
            }
        except Exception as e:
            print(f"Error procesando Aprendices: {e}")

    app_data = {
        "lastUpdate": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "aiInsight": ai_insight,
        "kpisGerenciales": kpis_gerenciales,
        "kpisAprendices": kpis_aprendices,
        "vacantes": vacantes,
        "fpp": fpp_list,
        "stats": {
            "estados": {
                "names": list(estados_counter.keys()),
                "counts": list(estados_counter.values())
            }
        }
    }

    print("Inyectando datos en el Dashboard...")
    try:
        with open(HTML_PATH, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # Regex robusto para encontrar y reemplazar el bloque appData
        pattern = re.compile(r"(// \{\{DATA_JSON\}\}\s+)let appData = \{.*?\};\s+(function initDashboard)", re.DOTALL)
        
        replacement_str = f"let appData = {json.dumps(app_data, indent=4)};\n\n        "
        new_html = pattern.sub(lambda m: m.group(1) + replacement_str + m.group(2), html_content)
        
        with open(HTML_PATH, 'w', encoding='utf-8') as f:
            f.write(new_html)
        
        print(f"Proceso finalizado. Dashboard actualizado en {HTML_PATH}")
    except Exception as e:
        print(f"Error al actualizar el HTML: {e}")

if __name__ == "__main__":
    process()

