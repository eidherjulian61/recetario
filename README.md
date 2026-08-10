# Recetario

Recetario Comfama en formato de página web autocontenida.

## Uso

Disponible en línea en https://eidherjulian61.github.io/recetario/ (se publica automáticamente con cada push a `main`). Incluye:

- Buscador de recetas por texto
- Filtro por libro
- Secciones plegables para navegar cómodamente

## Estructura

- [`index.html`](index.html): la página con el buscador y el visor de recetas.
- [`recetas.json`](recetas.json): los datos de todos los libros y recetas. Para agregar o editar recetas, modifica este archivo.
- [`imagenes/`](imagenes): las fotos de las recetas.

Para ejecutarlo localmente sirve el directorio por HTTP (por ejemplo `python3 -m http.server`) y abre `http://localhost:8000`, ya que la página carga `recetas.json` con `fetch`.
