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
        layout = "contenido"  # por defecto

        for line in lines:
            if line.startswith("#"):
                # Separar etiqueta de layout si existe
                if "[layout:" in line:
                    parts = line.split("[layout:")
                    title = parts[0].lstrip("#").strip()
                    layout = parts[1].rstrip("]").strip()
                else:
                    title = line.lstrip("#").strip()
            else:
                bullets.append(line.lstrip("-").strip())

        if title:
            slides.append((title, bullets, layout))
    return slides


def add_slide(prs, layout_map, title, bullets, layout_key):
    # Obtener índice de layout desde el mapa
    layout_index = layout_map.get(layout_key, layout_map["contenido"])
    layout = prs.slide_layouts[layout_index]

    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title

    # Manejo de contenido en placeholders
    if len(slide.placeholders) > 1 and bullets:
        tf = slide.placeholders[1].text_frame
        tf.clear()
        for i, bullet in enumerate(bullets):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = bullet


def main():
    prs = Presentation("templates/template.pptx")
    slides = parse_slides("slides.md")

    # Mapeo de tipos de layout en la plantilla
    layout_map = {
        "portada": 0,    # primer layout en tu plantilla
        "contenido": 1,  # título + viñetas
        "imagen": 2,     # layout con espacio para imagen
    }

    for title, bullets, layout in slides:
        add_slide(prs, layout_map, title, bullets, layout)

    prs.save("presentacion.pptx")
    print("✅ Presentación generada: presentacion.pptx")


if __name__ == "__main__":
    main()