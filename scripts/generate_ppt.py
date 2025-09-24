#!/usr/bin/env python3
import locale
from datetime import datetime
from pptx import Presentation

# Configurar la localización en español
try:
    locale.setlocale(locale.LC_TIME, "es_ES.UTF-8")  # Linux/Mac
except:
    try:
        locale.setlocale(locale.LC_TIME, "Spanish_Spain")  # Windows
    except:
        print("⚠️ No se pudo establecer locale en español, se usará inglés.")

def parse_slides(md_path):
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    raw_slides = content.split("---")
    slides = []

    for raw in raw_slides:
        raw_lines = raw.strip().splitlines()
        if not raw_lines:
            continue

        title = None
        bullets = []
        images = []
        layout = "contenido"

        for line in raw_lines:
            stripped = line.strip()
            if not stripped:
                continue

            if stripped.startswith("#"):
                if "[layout:" in stripped:
                    parts = stripped.split("[layout:")
                    title = parts[0].lstrip("#").strip()
                    layout = parts[1].rstrip("]").strip()
                else:
                    title = stripped.lstrip("#").strip()
            elif stripped.endswith((".png", ".jpg", ".jpeg")):
                images.append(stripped)
            else:
                bullets.append(stripped.lstrip("-").strip())

        if title:
            slides.append((title, bullets, images, layout))
    return slides


def add_slide(prs, layout_map, title, bullets, images, layout_key):
    layout_index = layout_map.get(layout_key, layout_map["contenido"])
    slide = prs.slides.add_slide(prs.slide_layouts[layout_index])

    # Para portada → título + salto + fecha
    if layout_key == "portada":
        if slide.shapes.title:
            tf = slide.shapes.title.text_frame
            tf.clear()

            # Párrafo con título
            p1 = tf.add_paragraph()
            p1.text = title

            # Párrafo vacío (doble salto visual)
            p2 = tf.add_paragraph()
            p2.text = ""

            # Mes y año en español
            fecha = datetime.now().strftime("%B %Y").capitalize()
            p3 = tf.add_paragraph()
            p3.text = fecha
        return

    # Título normal
    slide.shapes.title.text = title

    # Si hay texto (bullets)
    if len(slide.placeholders) > 1 and bullets:
        tf = slide.placeholders[1].text_frame
        tf.clear()
        for i, bullet in enumerate(bullets):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = bullet

    # Si el layout es "imagen" → usar el placeholder de imagen
    if layout_key == "imagen" and images:
        for shape in slide.placeholders:
            if "Picture" in shape.name or "Imagen" in shape.name:  # PowerPoint en inglés/español
                try:
                    shape.insert_picture(images[0])  # Usa la primera imagen
                    print(f"✅ Imagen insertada en placeholder: {images[0]}")
                except Exception as e:
                    print(f"⚠️ No se pudo insertar imagen {images[0]}: {e}")
                break


def main():
    prs = Presentation("templates/template.pptx")
    slides = parse_slides("slides.md")

    layout_map = {
        "portada": 0,     # Tu layout de portada en la plantilla
        "contenido": 1,   # Layout normal
        "imagen": 2,      # Layout de título + imagen
    }

    for title, bullets, images, layout in slides:
        add_slide(prs, layout_map, title, bullets, images, layout)

    prs.save("presentacion.pptx")
    print("✅ Presentación generada con portada multilínea y fecha en español: presentacion.pptx")


if __name__ == "__main__":
    main()
