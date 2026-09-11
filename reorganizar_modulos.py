#!/usr/bin/env python3
import re

# Leer el archivo actual
with open('frontend/index.html', 'r', encoding='utf-8') as f:
    contenido = f.read()

# Paso 1: Renumerar módulos operacionales
cambios = [
    ('data-vista="presupuestal"><span class="nav-numero">04</span>', 'data-vista="presupuestal"><span class="nav-numero">04</span>'),  # Sin cambio
    ('data-vista="contratacion"><span class="nav-numero">05</span>', 'data-vista="contratacion"><span class="nav-numero">05</span>'),  # Sin cambio
    ('data-vista="planeacion"><span class="nav-numero">06</span>', 'data-vista="planeacion"><span class="nav-numero">06</span>'),  # Sin cambio
    ('data-vista="oxp"><span class="nav-numero">09</span>', 'data-vista="oxp"><span class="nav-numero">07</span>'),  # 09 → 07
    ('data-vista="exogena"><span class="nav-numero">10</span>', 'data-vista="exogena"><span class="nav-numero">08</span>'),  # 10 → 08
]

for viejo, nuevo in cambios:
    contenido = contenido.replace(viejo, nuevo)

# Paso 2: Remover los botones de Ciberseguridad y Licencias del menú principal
contenido = re.sub(
    r'<button class="nav-item" data-vista="ciberseguridad">.*?</button>\s*',
    '',
    contenido,
    flags=re.DOTALL
)
contenido = re.sub(
    r'<button class="nav-item" data-vista="licencias">.*?</button>\s*',
    '',
    contenido,
    flags=re.DOTALL
)

# Paso 3: Insertar sección de Administración antes de </nav>
admin_seccion = '''    </nav>

    <nav class="nav-modulos" style="margin-top: 20px; border-top: 1px solid #ddd; padding-top: 15px;">
      <span class="nav-grupo-titulo">Administración</span>
      <button class="nav-item" data-vista="licencias"><span class="nav-numero">⚙</span><svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M15 7a4 4 0 11-8 0 4 4 0 018 0z"/><path d="M11 11l7 7 2-2-2-2 2-2-3-3"/></svg> Licencias</button>
      <button class="nav-item" data-vista="ciberseguridad"><span class="nav-numero">⚙</span><svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M12 2l7 4v6c0 5-3 8-7 10-4-2-7-5-7-10V6z"/><path d="M9 12l2 2 4-4"/></svg> Ciberseguridad</button>
    </nav>'''

contenido = contenido.replace('    </nav>', admin_seccion, 1)  # Solo la primera ocurrencia

# Guardar el archivo actualizado
with open('frontend/index.html', 'w', encoding='utf-8') as f:
    f.write(contenido)

print("✓ Cambios realizados exitosamente:")
print("  - Renumerado: OXP (09→07), Exógena (10→08)")
print("  - Removido: Ciberseguridad y Licencias del menú principal")
print("  - Creada: Nueva sección Administración")
