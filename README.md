# Recetario

Recetario Comfama en formato de página web autocontenida.

## Uso

Disponible en línea en https://eidherjulian61.github.io/recetario/ (se publica automáticamente con cada push a `main`). Incluye:

- Buscador de recetas por texto
- Filtro por libro
- Secciones plegables para navegar cómodamente

## Prueba del Estado Colombiano

[`test-estado-colombiano.html`](test-estado-colombiano.html) es una aplicación aparte, autocontenida:
un test interactivo de 10 preguntas de alta dificultad sobre presupuesto público, estructura del
Estado, ramas del poder y administración pública, pensado para líderes sociales.

- Enlace para compartir: https://eidherjulian61.github.io/recetario/test-estado-colombiano.html
- Panel de la coordinación: el mismo enlace con `#panel` al final. Contraseña inicial `estado2026`
  (cámbiela desde la pestaña «Configuración»).
- Respuesta anónima: no se pide nombre, correo ni ningún dato personal. Cada intento se
  identifica con un código aleatorio generado en el propio dispositivo.
- Los resultados se guardan en el navegador de cada participante. Para reunirlos en un solo lugar,
  configure el repositorio central (Google Apps Script) desde el panel; el código listo para pegar
  está allí mismo.

## Estructura

- [`index.html`](index.html): la página con el buscador y el visor de recetas.
- [`recetas.json`](recetas.json): los datos de todos los libros y recetas. Para agregar o editar recetas, modifica este archivo.
- [`imagenes/`](imagenes): las fotos de las recetas.

Para ejecutarlo localmente sirve el directorio por HTTP (por ejemplo `python3 -m http.server`) y abre `http://localhost:8000`, ya que la página carga `recetas.json` con `fetch`.
