from Bio import Entrez
import json
from dotenv import load_dotenv
import os
from datetime import datetime

# Cargar variables de entorno
load_dotenv()

Entrez.email = os.getenv("NCBI_EMAIL")
Entrez.api_key = os.getenv("NCBI_API_KEY")

def buscar_pubmed(termino_busqueda, max_papers=None):
    if max_papers:
        print(f"🔍 Buscando: {termino_busqueda}")
        print(f"📊 Límite de papers: {max_papers}")
    else:
        print(f"🔍 Buscando: {termino_busqueda}")
        print(f"📊 Obteniendo TODOS los resultados disponibles")
    
    # Información de la búsqueda
    search_info = {
        "search_engine": "PubMed",
        "date": datetime.now().strftime('%Y-%m-%d'),
        "search_terms": termino_busqueda
    }
    
    # Determinar el número máximo de resultados a obtener
    retmax_value = max_papers if max_papers else 100000
    
    handle = Entrez.esearch(
        db="pubmed",
        term=termino_busqueda,
        retmax=retmax_value
    )
    record = Entrez.read(handle)
    
    # Mostrar información de la búsqueda
    total_encontrados = record.get("Count", "0") if isinstance(record, dict) else "0"
    ids = record["IdList"] if isinstance(record, dict) and "IdList" in record else []
    
    print(f"📈 Total de papers encontrados en PubMed: {total_encontrados}")
    print(f"📥 Papers descargados para procesar: {len(ids)}")
    print(f"🆔 IDs de los papers: {ids[:10]}{'...' if len(ids) > 10 else ''}")
    
    resultados = []

    if not ids:
        print("❌ No se encontraron resultados")
        return {
            "search_info": search_info,
            "total_encontrados": int(total_encontrados),
            "papers_procesados": 0,
            "termino_busqueda": termino_busqueda,
            "resultados": []
        }

    # Buscar detalles para cada artículo
    print(f"📚 Obteniendo detalles de {len(ids)} papers...")
    handle = Entrez.efetch(db="pubmed", id=",".join(ids), rettype="xml", retmode="xml")
    articles = Entrez.read(handle)

    # Entrez.read(handle) returns a list, so access the first element
    if isinstance(articles, list) and articles:
        articles_dict = articles[0]
    else:
        articles_dict = articles

    if isinstance(articles_dict, dict):
        pubmed_articles = articles_dict.get('PubmedArticle', [])
    elif isinstance(articles_dict, list):
        pubmed_articles = articles_dict
    else:
        pubmed_articles = []

    print(f"✅ Articles procesables encontrados: {len(pubmed_articles)}")
    papers_procesados_exitosamente = 0

    for i, article in enumerate(pubmed_articles, 1):
        try:
            print(f"📄 Procesando paper {i}/{len(pubmed_articles)}...")
            datos = article['MedlineCitation']['Article']
            titulo = datos.get('ArticleTitle', 'No title')
            autores = [f"{a['LastName']} {a['Initials']}" for a in datos.get('AuthorList', []) if 'LastName' in a]
            resumen = datos.get('Abstract', {}).get('AbstractText', [''])[0]
            journal = datos['Journal']['Title']
            año = datos['Journal']['JournalIssue']['PubDate'].get('Year', 'No year')
            
            # Extraer publisher (editorial)
            publisher = 'No publisher'
            try:
                if 'Journal' in datos and 'ISOAbbreviation' in datos['Journal']:
                    publisher = datos['Journal'].get('ISOAbbreviation', 'No publisher')
                elif 'MedlineJournalInfo' in article['MedlineCitation']:
                    publisher = article['MedlineCitation']['MedlineJournalInfo'].get('MedlineTA', 'No publisher')
            except:
                publisher = 'No publisher'
            
            papers_procesados_exitosamente += 1
        except Exception as e:
            print(f"❌ Error procesando artículo {i}: {e}")
            continue

        resultados.append({
            "paper_name": titulo,
            "paper_year": año,
            "paper_authors": ', '.join(autores),
            "journal": journal,
            "publisher": publisher,
            "resumen": resumen
        })

    # Actualizar hits en search_info
    search_info["hits"] = int(total_encontrados)
    
    print(f"✅ Procesamiento completado:")
    print(f"   📊 Papers procesados exitosamente: {papers_procesados_exitosamente}")
    print(f"   ❌ Papers con errores: {len(pubmed_articles) - papers_procesados_exitosamente}")

    return {
        "search_info": search_info,
        "total_encontrados": int(total_encontrados),
        "papers_procesados": papers_procesados_exitosamente,
        "papers_con_errores": len(pubmed_articles) - papers_procesados_exitosamente,
        "termino_busqueda": termino_busqueda,
        "resultados": resultados
    }
