import os

def final_polish(directory):
    replacements = {
        'Ã rea': 'Área',
        'Ã‰': 'É',
        'Ã ': 'Í',
        'Ã“': 'Ó',
        'Ãš': 'Ú',
        'Ã‘': 'Ñ',
        'Ã¡': 'á',
        'Ã©': 'é',
        'Ã­': 'í',
        'Ã³': 'ó',
        'Ãº': 'ú',
        'Ã±': 'ñ',
        'DINÃ MICO': 'DINÁMICO',
        'AUDIOMETRÃ A': 'AUDIOMETRÍA',
        'ESPIROMETRÃ A': 'ESPIROMETRÍA',
        'PÃ GINAS': 'PÁGINAS',
        'TÃ‰CNICA': 'TÉCNICA',
        'INFORMACIÃ“N': 'INFORMACIÓN',
        'PRESENTACIÃ“N': 'PRESENTACIÓN',
        'Ã³n': 'ón'
    }
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.html') or file.endswith('.js'):
                path = os.path.join(root, file)
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                
                original = content
                for search, replace in replacements.items():
                    content = content.replace(search, replace)
                
                if content != original:
                    print(f"Fixed encoding in {path}")
                    with open(path, 'w', encoding='utf-8') as f:
                        f.write(content)

if __name__ == "__main__":
    final_polish('frontend')
    final_polish('static')
