# 📄 Formateador automático APA 7

Aplicación local con interfaz web sencilla (Streamlit) que formatea documentos de Word (`.docx`) según la **norma APA 7.ª edición**. Subes tu archivo, rellenas los datos de la portada y descargas el documento ya formateado.

> **Autor:** Alejandro Gutiérrez Villalba

---

## ✨ Qué hace

| Elemento | Formato aplicado |
|---|---|
| **Portada** | Título (negrita), autor, institución, asignatura, docente (opcional) y fecha, centrados y con interlineado doble |
| **Índice** | Campo nativo `TOC` de Word (niveles 1 a 3), actualizable con un clic |
| **Márgenes** | 2,54 cm en todo el contorno |
| **Tipografía** | Times New Roman, 12 pt en todo el documento |
| **Párrafos** | Interlineado doble, sin espaciado antes ni después, sangría de primera línea de 1,27 cm |
| **Títulos (5 niveles)** | N1 centrado en negrita · N2 a la izquierda en negrita · N3 negrita y cursiva · N4 y N5 con sangría y punto final, en la misma línea del texto |
| **Tablas** | Sin bordes verticales, solo líneas horizontales finas, texto a espacio sencillo |
| **Figuras** | Imágenes centradas con etiqueta «Figura X», título en cursiva y nota |
| **Referencias** | Orden alfabético y sangría francesa de 1,27 cm |
| **Citas** | Normalización básica autor-fecha: `(Autor 2020)` → `(Autor, 2020)` |
| **Paginación** | Número de página a la derecha en el encabezado de todas las páginas |

---

## 🧰 Requisitos

- Python 3.9 o superior
- [`python-docx`](https://pypi.org/project/python-docx/)
- [`streamlit`](https://streamlit.io/)

## 🚀 Instalación

```bash
git clone https://github.com/TU_USUARIO/TU_REPO.git
cd TU_REPO
pip install -r requirements.txt
```

O, sin `requirements.txt`:

```bash
pip install python-docx streamlit
```

## ▶️ Uso

1. Abre la carpeta del proyecto en una terminal (o en VS Code).
2. Ejecuta:

   ```bash
   python -m streamlit run app.py
   ```

3. Se abrirá el navegador en `http://localhost:8501`. Si no se abre solo, copia esa dirección.
4. Sube tu `.docx`, completa los datos de la portada y pulsa **Formatear documento**.
5. Descarga el resultado (`nombre_APA7.docx`).
6. Al abrirlo en Word, acepta la actualización de campos o haz clic derecho sobre el índice → **Actualizar campo** → **Actualizar toda la tabla**.

---

## ⚠️ Consideraciones

- **Títulos:** se detectan por los estilos de Word *Título 1–5* (Heading 1–5). Si tu documento usa negritas manuales en lugar de estilos, no se reconocerán los niveles.
- **Índice:** incluye los niveles 1 a 3. Los niveles 4 y 5 son títulos en línea (van dentro del párrafo), por lo que no aparecen en el índice.
- **Etiquetas de figuras y tablas:** siguiendo APA 7, «Figura 1» va en negrita y el título en cursiva. Para poner también la etiqueta en cursiva, cambia `ETIQUETAS_EN_CURSIVA = True` al inicio de `app.py`.
- **Marcadores:** si faltan etiquetas o notas en figuras y tablas, se insertan textos entre corchetes para que los rellenes (se puede desactivar en *Opciones*).
- **Citas:** la normalización es básica y solo actúa si la cita está dentro de un mismo fragmento de texto.
- **No se procesan:** notas al pie, cuadros de texto y comentarios.
- Revisa siempre el resultado final: ninguna herramienta automática sustituye a una revisión manual.

---

## 📁 Estructura del proyecto

```
.
├── app.py              # Aplicación completa (Streamlit + python-docx)
├── requirements.txt    # Dependencias
├── .gitignore
└── README.md
```

---

## 👤 Autor

**Alejandro Gutiérrez Villalba**

© 2026 Alejandro Gutiérrez Villalba. Todos los derechos reservados.