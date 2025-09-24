#!/usr/bin/env python3
import sys
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
        for line in lines:
            if line.startswith("#"):
                title = line.lstrip("#").strip()
            else:
                bullets.append(line.lstrip("-").strip())
        if title:
            slides.append((title, bullets))
    return slides

def add_slide(prs, title, bullets):
    layout = prs.slide_layouts[1]  # Title + Content
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title
    tf = slide.shapes.placeholders[1].text_frame
    tf.clear()
    for i, bullet in enumerate(bullets):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = bullet
        p.level = 0

def main():
    md_path = "slides.md"
    prs = Presentation()

    for i, (title, bullets) in enumerate(parse_slides(md_path)):
        add_slide(prs, title, bullets)

    out_file = "presentacion.pptx"
    prs.save(out_file)
    print(f"Presentación generada: {out_file}")

if __name__ == "__main__":
    main()