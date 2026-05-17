import os

def nuclear_fix(directory):
    # Diccionario de secuencias de bytes mal interpretadas
    # Convertimos los caracteres de la pantalla a lo que realmente son
    replacements = {
        'Ã rea': 'Área',
        'Ã\x81rea': 'Área',
        'Ã³': 'ó',
        'Ã¡': 'á',
        'Ã©': 'é',
        'Ã\xad': 'í',
        'Ã±': 'ñ',
        'Ã\x93': 'Ó',
        'Ã\x8d': 'Í',
        'DINÃ\x81MICO': 'DINÁMICO',
        'PÃ\x81GINAS': 'PÁGINAS',
        'AUDIOMETRÃ\x81A': 'AUDIOMETRÍA',
        'ESPIROMETRÃ\x81A': 'ESPIROMETRÍA'
    }
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(('.html', '.js')):
                path = os.path.join(root, file)
                # Leer como bytes para no tener problemas de decodificación
                with open(path, 'rb') as f:
                    content = f.read()
                
                # Reemplazar secuencias de bytes
                # Ã rea en UTF-8 mal interpretado suele ser C3 81 72 65 61
                # Pero si el archivo ya está en UTF-8 y tiene los caracteres literales C3 83 C2 81...
                
                # Lo más seguro es tratarlo como texto pero con una codificación que no falle
                text = content.decode('latin-1')
                original = text
                
                # Mapeo manual basado en lo que veo en el grep
                text = text.replace('Ã³', 'ó')
                text = text.replace('Ã¡', 'á')
                text = text.replace('Ã©', 'é')
                text = text.replace('Ã\xad', 'í')
                text = text.replace('Ã\xad', 'í')
                text = text.replace('Ã±', 'ñ')
                text = text.replace('Ã\x81rea', 'Área')
                text = text.replace('Ã rea', 'Área')
                text = text.replace('Ã\x8d', 'Í')
                text = text.replace('Ã\x93', 'Ó')
                text = text.replace('AUDIOMETRÃ\x8d', 'AUDIOMETRÍ')
                text = text.replace('ESPIROMETRÃ\x8d', 'ESPIROMETRÍ')
                text = text.replace('DINÃ\x81MICO', 'DINÁMICO')
                text = text.replace('PÃ\x81GINAS', 'PÁGINAS')
                
                if text != original:
                    print(f"Nuclearly fixed {path}")
                    with open(path, 'w', encoding='utf-8') as f:
                        f.write(text)

if __name__ == "__main__":
    nuclear_fix('frontend')
    nuclear_fix('static')
