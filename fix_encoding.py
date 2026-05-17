import os
import re

def fix_html_files(directory):
    replacements = {
        'Ã¡': 'á', 'Ã©': 'é', 'Ã­': 'í', 'Ã³': 'ó', 'Ãº': 'ú', 'Ã±': 'ñ',
        'Ã ': 'Á', 'Ã‰': 'É', 'Ã ': 'Í', 'Ã“': 'Ó', 'Ãš': 'Ú', 'Ã‘': 'Ñ',
        'â€“': '–', 'Â': ''
    }
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.html'):
                path = os.path.join(root, file)
                print(f"Fixing {path}...")
                
                # Leer como bytes para detectar y arreglar
                with open(path, 'rb') as f:
                    content = f.read().decode('utf-8', errors='ignore')
                
                # Aplicar reemplazos de tildes rotas
                for search, replace in replacements.items():
                    content = content.replace(search, replace)
                
                # Asegurar link al CSS offline si no existe
                if 'offline-fonts.css' not in content:
                    content = content.replace('</head>', '<link rel="stylesheet" href="/static/js/offline-fonts.css">\n</head>')
                
                # Escribir como UTF-8 puro (sin BOM)
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(content)

if __name__ == "__main__":
    fix_html_files('frontend')
