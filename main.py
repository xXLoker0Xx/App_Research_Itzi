import json
import re

# Importar nuestros módulos
from extraccion_datos import buscar_pubmed
from json_to_excel import json_to_excel

def proceso_completo(query,
                     inclusion_criteria="No especificado",
                     stopping_criteria="Todos los resultados disponibles",
                     study_type="No especificado",
                     primary_topic_area="No especificado",
                     exclusion_criteria="No especificado",
                     max_papers=None):
    """
    Ejecuta el proceso completo: extracción + conversión
    
    Args:
        max_papers (int, optional): Límite máximo de papers a procesar. Si es None, no hay límite.
    """
    # Extracción
    resultados = buscar_pubmed(
        termino_busqueda=query,
        max_papers=max_papers
    )
    
    if resultados and resultados.get('papers_procesados', 0) > 0:
        # Generar nombre de archivo
        nombre_archivo = re.sub(r'[^\w\s-]', '', query.replace('"', ''))
        nombre_archivo = re.sub(r'[-\s]+', '_', nombre_archivo)
        nombre_archivo = nombre_archivo.lower()[:50]
        nombre_archivo = f"{nombre_archivo}_resultados.json"
        
        # Guardar JSON
        with open(nombre_archivo, "w", encoding="utf-8") as f:
            json.dump(resultados, f, ensure_ascii=False, indent=2)
        
        # Convertir a Excel pasando los parámetros
        excel_file = json_to_excel(
            json_file_path=nombre_archivo,
            inclusion_criteria=inclusion_criteria,
            stopping_criteria=stopping_criteria,
            study_type=study_type,
            primary_topic_area=primary_topic_area,
            exclusion_criteria=exclusion_criteria
        )
        
        return nombre_archivo, excel_file
    
    return None, None

if __name__ == "__main__":
    # Configuración de la búsqueda
    # query = 'Leptin AND Alzheimer\'s disease AND ("systematic review"[Publication Type] OR "review"[Publication Type])'

    query = 'hippocampus and ("neurogenesis" OR "neurons" OR "LTP") and leptin'
    inclusion_criteria = """
    • Studies published in English, peer-reviewed journals
    • About leptin and Alzheimer’s
    • Relevant papers available as full text
    • Randomized control trials 
    """
    stopping_criteria = "40% of total quota selected for tranche"
    study_type = "Randomized control trials"
    primary_topic_area = "Neuroscience, Alzheimer’s and Leptin"
    exclusion_criteria = """
    • Any other studies
    • Non relevant topic
    • Full-text not available
    • Paper not available in English
    """
    
    # Límite de papers (opcional)
    max_papers = None  # Si quieres limitar, pon un número como: max_papers = 50

    # Ejecutar proceso completo
    json_file, excel_file = proceso_completo(
        query=query,
        inclusion_criteria=inclusion_criteria,
        stopping_criteria=stopping_criteria,
        study_type=study_type,
        primary_topic_area=primary_topic_area,
        exclusion_criteria=exclusion_criteria,
        max_papers=max_papers
    )