import openpyxl
import openai
import json
from collections import defaultdict

# Configurar la API key de OpenAI
openai.api_key = "apigpt"

def leer_excel(archivo):
    wb = openpyxl.load_workbook(archivo)
    sheet = wb.active
    datos = defaultdict(lambda: defaultdict(dict))
    
    for row in sheet.iter_rows(min_row=2, values_only=True):
        id_producto, barcode, image, country, field, original_value, found_value = row
        datos[barcode][field] = {'original': original_value, 'found': found_value}
    
    return datos

def dividir_datos(datos, chunk_size=50):
    items = list(datos.items())
    for i in range(0, len(items), chunk_size):
        yield dict(items[i:i + chunk_size])

def analizar_con_gpt(datos_chunk):
    prompt = f"""
    Analiza los siguientes resultados de productos:

    {json.dumps(datos_chunk, indent=2)}

    Por favor, proporciona un análisis detallado que incluya:
    1. Resumen general de los datos (número total de productos, campos más comúnmente encontrados, etc.)
    2. Discrepancias notables entre los valores originales y los encontrados
    3. Campos que tienden a faltar o estar incompletos
    4. Patrones o tendencias interesantes en los datos
    5. Sugerencias para mejorar la calidad de los datos o el proceso de recopilación

    Presenta tu análisis de manera clara y concisa.
    """

    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo-16k",
        messages=[
            {"role": "system", "content": "Eres un analista de datos experto especializado en información de productos."},
            {"role": "user", "content": prompt}
        ]
    )

    return response.choices[0].message['content']

def main():
    archivo_excel = 'Resultados_todos_barcodes.xlsx'
    print(f"Leyendo el archivo: {archivo_excel}")
    datos = leer_excel(archivo_excel)
    
    print("Analizando los datos con GPT...")
    analisis_completo = ""
    for i, chunk in enumerate(dividir_datos(datos)):
        print(f"Analizando chunk {i+1}...")
        analisis_chunk = analizar_con_gpt(chunk)
        analisis_completo += f"\n\nAnálisis del chunk {i+1}:\n{analisis_chunk}\n"
        
        # Guardar el análisis parcial después de cada chunk
        with open('analisis_resultados_parcial.txt', 'a', encoding='utf-8') as f:
            f.write(f"\n\nAnálisis del chunk {i+1}:\n{analisis_chunk}\n")
    
    print("\nAnálisis completo de los resultados:")
    print(analisis_completo)
    
    # Guardar el análisis completo en un archivo de texto
    with open('analisis_resultados_completo.txt', 'w', encoding='utf-8') as f:
        f.write(analisis_completo)
    print("\nEl análisis completo se ha guardado en 'analisis_resultados_completo.txt'")

if __name__ == "__main__":
    main()
