#!/usr/bin/env python3
import sys
from pptx import Presentation
from pptx.util import Inches, Pt

def add_bullet_slide(prs, title_text, bullet_lines):
    slide_layout = prs.slide_layouts[1]  # Title and Content
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    body = slide.shapes.placeholders[1].text_frame
    title.text = title_text
    body.clear()
    for i, line in enumerate(bullet_lines):
        p = body.add_paragraph() if i>0 else body.paragraphs[0]
        p.text = line
        p.level = 0

def main():
    # Expecting: generate_ppt.py "<titulo>" "<autor>"
    if len(sys.argv) < 3:
        print("Uso: generate_ppt.py \"<titulo>\" \"<autor>\"")
        sys.exit(1)

    titulo = sys.argv[1]
    autor = sys.argv[2]

    prs = Presentation()

    # Slide 1 - Portada
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = titulo
    try:
        slide.placeholders[1].text = f"Autor: {autor}"
    except Exception:
        pass

    # Slide 2 - ¿Qué es GitHub Actions?
    add_bullet_slide(prs, "¿Qué es GitHub Actions?", [
        "Plataforma de automatización integrada en GitHub.",
        "Permite ejecutar workflows cuando ocurren eventos en el repo.",
        "Útil para CI/CD, testing y despliegues."
    ])

    # Slide 3 - Conceptos clave
    add_bullet_slide(prs, "Conceptos clave", [
        "Workflow: archivo .yml que define el flujo.",
        "Jobs: conjunto de steps que se ejecutan en un runner.",
        "Steps: comandos o actions dentro de un job.",
        "Actions: bloques reutilizables (Marketplace o personalizadas)."
    ])

    # Slide 4 - Primeros pasos
    add_bullet_slide(prs, "Primeros pasos", [
        "Crear o abrir un repositorio en GitHub.",
        "Ir a la pestaña Actions y elegir plantilla o configurar.",
        "Agregar archivo .github/workflows/mi-workflow.yml."
    ])

    # Slide 5 - Estructura de un workflow
    add_bullet_slide(prs, "Estructura de un workflow", [
        "Archivo: .github/workflows/mi-workflow.yml",
        "Ejemplo: on: [push], jobs: { ... }"
    ])

    # Slide 6 - Ejemplo práctico (Node.js)
    add_bullet_slide(prs, "Ejemplo: pruebas en Node.js", [
        "name: CI Node.js",
        "on: [push, pull_request]",
        "jobs: test -> runs-on: ubuntu-latest",
        "steps: checkout, npm install, npm test"
    ])

    # Slide 7 - Ejecutando y verificando
    add_bullet_slide(prs, "Ejecutando y verificando", [
        "Workflows se ejecutan al ocurrir eventos o manualmente.",
        "Ver estado y logs en la pestaña Actions.",
        "Opción de reintentar jobs fallidos."
    ])

    # Slide 8 - Casos de uso
    add_bullet_slide(prs, "Casos de uso comunes", [
        "CI/CD para aplicaciones.",
        "Construcción y publicación de contenedores Docker.",
        "Despliegue a GitHub Pages o nubes (AWS, Azure, GCP)."
    ])

    # Slide 9 - Mejores prácticas
    add_bullet_slide(prs, "Mejores prácticas", [
        "Usar secrets para credenciales.",
        "Aprovechar cache para acelerar builds.",
        "Reutilizar workflows y mantener YAML simple."
    ])

    # Slide 10 - Conclusión
    add_bullet_slide(prs, "Conclusión", [
        "GitHub Actions simplifica la automatización y CI/CD.",
        "Explorar Marketplace y docs oficiales."
    ])

    # Guardar
    out_file = "presentacion.pptx"
    prs.save(out_file)
    print(f"Presentación guardada como {out_file}")

if __name__ == "__main__":
    main()