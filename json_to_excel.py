import json
import pandas as pd
import os
from datetime import datetime
import glob

def json_to_excel(json_file_path,
                  excel_file_path=None,
                  inclusion_criteria=None,
                  stopping_criteria=None,
                  study_type=None,
                  primary_topic_area=None,
                  exclusion_criteria=None):
    """
    Convierte un archivo JSON de resultados de PubMed a Excel
    
    Args:
        json_file_path (str): Ruta al archivo JSON
        excel_file_path (str): Ruta del archivo Excel de salida (opcional)
        inclusion_criteria (str): Criterios de inclusión (opcional)
        stopping_criteria (str): Criterios de parada (opcional)
        study_type (str): Tipo de estudio (opcional)
        primary_topic_area (str): Área temática principal (opcional)
    """
    
    # Verificar que el archivo JSON existe
    if not os.path.exists(json_file_path):
        print(f"❌ Error: El archivo {json_file_path} no existe")
        return
    
    # Cargar datos JSON
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"✅ Archivo JSON cargado exitosamente: {json_file_path}")
    except Exception as e:
        print(f"❌ Error al cargar el archivo JSON: {e}")
        return
    
    # Extraer información del resumen
    print("\n" + "="*60)
    print("📊 RESUMEN DE DATOS")
    print("="*60)
    print(f"🔍 Término de búsqueda: {data.get('termino_busqueda', 'No disponible')}")
    print(f"📈 Total encontrados: {data.get('total_encontrados', 0)}")
    print(f"✅ Papers procesados: {data.get('papers_procesados', 0)}")
    print(f"❌ Papers con errores: {data.get('papers_con_errores', 0)}")
    
    # Crear DataFrame con los resultados
    if 'resultados' in data and data['resultados']:
        # DataFrame para los papers con las columnas específicas solicitadas
        df_papers = pd.DataFrame(data['resultados'])
        
        # Asegurar que tenemos las columnas correctas
        required_columns = ['paper_name', 'paper_year', 'paper_authors', 'journal', 'publisher', 'resumen']
        for col in required_columns:
            if col not in df_papers.columns:
                df_papers[col] = 'N/A'
        
        # Seleccionar solo las columnas solicitadas
        df_papers = df_papers[required_columns]
        
        print(f"📋 DataFrame de papers creado con {len(df_papers)} filas")
        
        # Crear DataFrame para información de búsqueda
        search_info = data.get('search_info', {})
        
        # Usar los parámetros pasados o valores por defecto del JSON
        final_inclusion_criteria = inclusion_criteria or search_info.get('inclusion_criteria', 'No especificado')
        final_stopping_criteria = stopping_criteria or search_info.get('stopping_criteria', 'No especificado')
        final_study_type = study_type or search_info.get('study_type', 'No especificado')
        final_primary_topic_area = primary_topic_area or search_info.get('primary_topic_area', 'No especificado')
        
        search_data = {
            'Campo': ['Search Engine', 'Date', 'Search Terms', 'Hits', 'Inclusion Criteria', 
                     'Stopping Criteria', 'Study Type', 'Primary Topic Area'],
            'Valor': [
                search_info.get('search_engine', 'PubMed'),
                search_info.get('date', datetime.now().strftime('%Y-%m-%d')),
                search_info.get('search_terms', data.get('termino_busqueda', '')),
                search_info.get('hits', data.get('total_encontrados', 0)),
                final_inclusion_criteria,
                final_stopping_criteria,
                final_study_type,
                final_primary_topic_area
            ]
        }
        df_search = pd.DataFrame(search_data)
        
        print(f"📋 DataFrame de búsqueda creado con {len(df_search)} filas")
        
    else:
        print("❌ No se encontraron resultados en el archivo JSON")
        return
    
    # Generar nombre del archivo Excel si no se proporciona
    if excel_file_path is None:
        base_name = os.path.splitext(os.path.basename(json_file_path))[0]
        excel_file_path = f"{base_name}.xlsx"
    
    # Crear archivo Excel con formato
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            # Crear una sola hoja con el formato específico
            workbook = writer.book
            worksheet = workbook.create_sheet('Search Results')
            
            # Remover la hoja por defecto
            if 'Sheet' in workbook.sheetnames:
                workbook.remove(workbook['Sheet'])
            
            # Escribir la información de búsqueda en la parte superior
            search_info = data.get('search_info', {})
            
            # Usar los parámetros pasados o valores por defecto del JSON
            final_inclusion_criteria = inclusion_criteria or search_info.get('inclusion_criteria', 'No especificado')
            final_stopping_criteria = stopping_criteria or search_info.get('stopping_criteria', 'No especificado')
            final_study_type = study_type or search_info.get('study_type', 'No especificado')
            final_primary_topic_area = primary_topic_area or search_info.get('primary_topic_area', 'No especificado')
            
            # Fila 1: Search Engine
            worksheet['A1'] = 'Search Engine'
            worksheet['B1'] = search_info.get('search_engine', 'PubMed')
            
            # Fila 2: Date
            worksheet['A2'] = 'Date'
            worksheet['B2'] = search_info.get('date', datetime.now().strftime('%Y-%m-%d'))
            
            # Fila 3: Search Terms
            worksheet['A3'] = 'Search Terms'
            worksheet['B3'] = search_info.get('search_terms', data.get('termino_busqueda', ''))
            
            # Fila 4: Hits
            worksheet['A4'] = 'Hits'
            worksheet['B4'] = search_info.get('hits', data.get('total_encontrados', 0))
            
            # Fila 5: Inclusion Criteria
            worksheet['A5'] = 'Inclusion Criteria'
            worksheet['B5'] = final_inclusion_criteria
            
            # Fila 6: Stopping Criteria
            worksheet['A6'] = 'Stopping Criteria'
            worksheet['B6'] = final_stopping_criteria
            
            # Fila 7: Study Type
            worksheet['A7'] = 'Study Type'
            worksheet['B7'] = final_study_type
            
            # Fila 8: Primary Topic Area
            worksheet['A8'] = 'Primary Topic Area'
            worksheet['B8'] = final_primary_topic_area
            
            # Fila 9: Exclusion Criteria
            worksheet['A9'] = 'Exclusion Criteria'
            worksheet['B9'] = exclusion_criteria
            
            # Fila 11: Encabezados de la tabla de papers
            headers = ['#', 'Paper Name', 'Paper Year', 'Paper Authors', 'Journal', 'Publisher', 'Sum Up', 'Brief Reason For Rejection', 'Brief Reason For Acceptance']
            for col_idx, header in enumerate(headers, 1):
                worksheet.cell(row=11, column=col_idx, value=header)
            
            # Escribir los datos de los papers
            for row_idx, paper in enumerate(df_papers.itertuples(index=False), 12):
                counter = row_idx - 11  # Contador que empieza en 1
                worksheet.cell(row=row_idx, column=1, value=counter)  # Contador en columna A
                worksheet.cell(row=row_idx, column=2, value=paper.paper_name)
                worksheet.cell(row=row_idx, column=3, value=paper.paper_year)
                worksheet.cell(row=row_idx, column=4, value=paper.paper_authors)
                worksheet.cell(row=row_idx, column=5, value=paper.journal)
                worksheet.cell(row=row_idx, column=6, value=paper.publisher)
                # Columna 7 (Sum Up) contiene el resumen del paper
                worksheet.cell(row=row_idx, column=7, value=paper.resumen)
                worksheet.cell(row=row_idx, column=8, value='')
                worksheet.cell(row=row_idx, column=9, value='')
            
            # Formatear las columnas
            worksheet.column_dimensions['A'].width = 25  # Etiquetas de búsqueda / Contador
            worksheet.column_dimensions['B'].width = 80  # Valores de búsqueda / Paper Name
            worksheet.column_dimensions['C'].width = 15  # Paper Year
            worksheet.column_dimensions['D'].width = 50  # Paper Authors
            worksheet.column_dimensions['E'].width = 40  # Journal
            worksheet.column_dimensions['F'].width = 30  # Publisher
            worksheet.column_dimensions['G'].width = 60  # Sum Up (resumen)
            worksheet.column_dimensions['H'].width = 30  # Brief Reason For Rejection
            worksheet.column_dimensions['I'].width = 30  # Brief Reason For Acceptance
            
            # Aplicar filtros automáticos a la tabla de papers
            if len(df_papers) > 0:
                worksheet.auto_filter.ref = f'A11:I{11 + len(df_papers)}'
        
        print(f"✅ Archivo Excel creado exitosamente: {excel_file_path}")
        print(f"📊 Datos exportados: {len(df_papers)} papers")
        print(f"📋 Información de búsqueda incluida en la parte superior")
        
    except Exception as e:
        print(f"❌ Error al crear el archivo Excel: {e}")
        return
    
    return excel_file_path

def buscar_archivos_json():
    """
    Busca todos los archivos JSON en el directorio actual
    """
    archivos_json = glob.glob("*_resultados.json")
    return archivos_json

def convertir_todos_json_a_excel():
    """
    Convierte todos los archivos JSON de resultados a Excel
    """
    archivos_json = buscar_archivos_json()
    
    if not archivos_json:
        print("❌ No se encontraron archivos JSON de resultados en el directorio actual")
        return
    
    print(f"📁 Encontrados {len(archivos_json)} archivos JSON:")
    for archivo in archivos_json:
        print(f"   - {archivo}")
    
    print("\n🔄 Iniciando conversión a Excel...")
    print("="*60)
    
    archivos_creados = []
    for archivo_json in archivos_json:
        print(f"\n📄 Procesando: {archivo_json}")
        # Llamar sin parámetros adicionales para usar los valores del JSON
        excel_file = json_to_excel(archivo_json)
        if excel_file:
            archivos_creados.append(excel_file)
    
    print("\n" + "="*60)
    print("✅ CONVERSIÓN COMPLETADA")
    print("="*60)
    print(f"📊 Archivos Excel creados: {len(archivos_creados)}")
    for archivo in archivos_creados:
        print(f"   - {archivo}")

# Ejemplo de uso
if __name__ == "__main__":
    print("="*60)
    print("🚀 CONVERSOR JSON A EXCEL")
    print("="*60)
    
    # Opción 1: Convertir un archivo específico
    # json_to_excel("leptin_and_cognitive_decline_review_or_systematic_resultados.json")
    
    # Opción 2: Convertir todos los archivos JSON de resultados
    convertir_todos_json_a_excel()
