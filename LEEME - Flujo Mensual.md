# 📋 PMO Dashboard I+D – Instrucciones de uso mensual

## 📁 Archivos de esta carpeta

| Archivo | Qué es |
|---|---|
| `pmo-dashboard-v2.html` | El dashboard ejecutivo (abrirlo en el navegador) |
| `proyectos.json` | Los datos del portfolio (se genera automáticamente) |
| `generar_json.py` | Script Python que lee el Excel y genera el JSON |
| `LEEME - Flujo Mensual.md` | Este archivo de instrucciones |

---

## 🔄 Flujo mensual (solo 3 pasos)

### PASO 1 – Actualizar el Excel (como siempre)
Trabajá en el archivo `.xlsm` con normalidad. No hay que hacer nada especial.

### PASO 2 – Regenerar los datos (1 vez por mes)
Abrí una terminal y ejecutá:
```
cd "C:\Users\user\OneDrive - REFINERIA DEL CENTRO S.A\CLAUDE\PROYECTOS I+D\GESTION DE PROYECTOS I+D\Instrucciones"
python generar_json.py
```
Esto actualiza `proyectos.json` con los datos más recientes del Excel.

### PASO 3 – Ver el dashboard
```
python -m http.server 8080
```
Luego abrí en el navegador:
```
http://localhost:8080/pmo-dashboard-v2.html
```

---

## ❓ Por qué necesito un servidor local (python -m http.server)?

Cuando abrís un archivo HTML haciendo doble clic, el navegador usa el protocolo `file://`.
Por seguridad, los navegadores **no permiten que un archivo HTML cargue otros archivos** con `file://`.

El servidor local (`python -m http.server`) usa el protocolo `http://`, que sí lo permite.
Es 100% local: no sale a internet, solo funciona en tu computadora.

---

## 🔧 Requisitos (instalar una sola vez)

```
pip install openpyxl
```

---

## 📌 ¿Qué hace cada archivo?

### proyectos.json
Es un archivo de texto con los datos de todos los proyectos en formato estructurado.
El dashboard lo lee automáticamente al abrirse.
Ejemplo de cómo se ve por dentro:
```json
{
  "_metadata": { "corte": "2026-05-07", "score_gpo": 0.963 },
  "proyectos": [
    { "id": "2025.B60", "n": "MARGARINAS CM...", "est": "EJECUCIÓN", ... }
  ]
}
```

### generar_json.py
Lee el archivo Excel y genera/actualiza `proyectos.json` automáticamente.
- Detecta los proyectos de la hoja REPORTE
- Calcula desvíos y semáforos
- Guarda el JSON listo para el dashboard

### pmo-dashboard-v2.html
El dashboard visual. Usa `fetch()` para pedir los datos al JSON.
- NO tiene datos adentro del código
- Se actualiza solo cuando cambia el JSON
- Misma interfaz que v1: filtros, semáforos, gráficos, alertas

---

## 🚀 Evolución futura (Etapa 2)

Una vez que esto funciona bien, el siguiente paso será:
- Reemplazar el servidor local por un servidor en la nube
- Que el JSON se genere automáticamente sin tener que correr el script a mano
- Agregar login de usuarios
- Permitir editar proyectos desde el propio dashboard
