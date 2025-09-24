#!/usr/bin/env python3
from pptx import Presentation

def parse_slides(md_path):
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    raw_slides = content.split("---")
    slides = []

    for raw in raw_slides:
        lines = [line.strip() for line in raw.strip().splitlines() if line.strip()]
        if not lines:
            continue

        title = None
        bullets = []
        images = []
        layout = "contenido"

        for line in lines:
            if line.startswith("#"):
                if "[layout:" in line:
                    parts = line.split("[layout:")
                    title = parts[0].lstrip("#").strip()
                    layout = parts[1].rstrip("]").strip()
                else:
                    title = line.lstrip("#").strip()
            elif line.endswith((".png", ".jpg", ".jpeg")):
                images.append(line)
            else:
                bullets.append(line.lstrip("-").strip())

        if title:
            slides.append((title, bullets, images, layout))
    return slides


def add_slide(prs, layout_map, title, bullets, images, layout_key):
    layout_index = layout_map.get(layout_key, layout_map["contenido"])
    slide = prs.slides.add_slide(prs.slide_layouts[layout_index])
    slide.shapes.title.text = title

    # Si hay texto
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
        "portada": 0,
        "contenido": 1,
        "imagen": 2,  # Layout de título + imagen
    }

    for title, bullets, images, layout in slides:
        add_slide(prs, layout_map, title, bullets, images, layout)

    prs.save("presentacion.pptx")
    print("✅ Presentación generada con imágenes en placeholders: presentacion.pptx")


if __name__ == "__main__":
    main()
