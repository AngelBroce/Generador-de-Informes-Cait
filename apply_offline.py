import os

OFFLINE_CSS_CONTENT = """/* Fuentes locales para modo Offline */
@font-face {
    font-family: 'Inter';
    src: url('/static/js/inter.woff2') format('woff2');
    font-weight: 100 900;
    font-display: swap;
    font-style: normal;
}
@font-face {
    font-family: 'Material Icons';
    src: url('/static/js/MaterialIcons-Regular.ttf') format('truetype');
    font-weight: normal;
    font-style: normal;
}
body { font-family: 'Inter', sans-serif !important; }
.material-symbols-outlined, .material-icons {
    font-family: 'Material Icons' !important;
    font-weight: normal; font-style: normal; font-size: 24px;
    line-height: 1; letter-spacing: normal; text-transform: none;
    display: inline-block; white-space: nowrap; word-wrap: normal;
    direction: ltr; -webkit-font-smoothing: antialiased;
}
"""

def make_offline(directory):
    # Crear el CSS offline
    css_path = 'static/js/offline-fonts.css'
    os.makedirs(os.path.dirname(css_path), exist_ok=True)
    with open(css_path, 'w', encoding='utf-8') as f:
        f.write(OFFLINE_CSS_CONTENT)

    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.html'):
                path = os.path.join(root, file)
                print(f"Patching {path}...")
                
                # Leer con detección de codificación cuidadosa
                with open(path, 'rb') as f:
                    content_bytes = f.read()
                
                try:
                    content = content_bytes.decode('utf-8')
                except UnicodeDecodeError:
                    content = content_bytes.decode('latin-1')
                
                # Reemplazos de CDN
                content = content.replace('https://cdn.tailwindcss.com', '/static/js/tailwind.js')
                content = content.replace('https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css', '/static/js/flatpickr.min.css')
                content = content.replace('https://cdn.jsdelivr.net/npm/flatpickr/dist/l10n/es.js', '/static/js/flatpickr_es.js')
                content = content.replace('https://cdn.jsdelivr.net/npm/flatpickr', '/static/js/flatpickr.js')
                
                # Desactivar Google Fonts e Inyectar CSS local
                import re
                content = re.sub(r'<link[^>]+fonts\.googleapis\.com[^>]+>', '<!-- Offline Fonts Enabled -->', content)
                
                if 'offline-fonts.css' not in content:
                    content = content.replace('</head>', '<link rel="stylesheet" href="/static/js/offline-fonts.css">\n</head>')
                
                # Guardar como UTF-8 puro
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(content)

if __name__ == "__main__":
    make_offline('frontend')
