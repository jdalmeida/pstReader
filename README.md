<!--
@author João Gbriel de Almeida
-->

### Leitor de PST em Python (Tkinter)

Aplicativo desktop simples para abrir arquivos `.pst`, navegar pastas, visualizar mensagens e exportar `.eml`/`.txt` no Windows.

### Requisitos
- Python 3.9+
- Tentar instalar `pypff`/`libpff-python` (pode exigir wheel pré-compilado no Windows)
- Opcional: `readpst` (libpst) no PATH para fallback

### Instalação
```bash
pip install -r requirements.txt
# Tente também instalar pypff/libpff conforme seu ambiente Windows
```

### Execução
```bash
python -m src.main
```

### Empacotamento (opcional)
```bash
pip install pyinstaller
pyinstaller --noconsole --onefile src/main.py
```

Observações:
- Se `pypff` não estiver disponível, o app tentará usar `readpst` se encontrado no PATH.
- Renderização de HTML é básica; por padrão converte HTML para texto simples. `tkhtmlview` é opcional.
