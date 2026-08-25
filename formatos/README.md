# Formato de cuenta de cobro mensual — ESAP

Formato de captura de información para los contratistas de la **Dirección de
Procesos de Selección** de la Escuela Superior de Administración Pública.

| Archivo | Descripción |
| --- | --- |
| `Formato-cuenta-de-cobro-mensual-ESAP.docx` | Documento final, listo para imprimir o diligenciar en Word (carta, 4 páginas). |
| `generar-formato.js` | Script que construye el `.docx` con la librería [`docx`](https://www.npmjs.com/package/docx). |
| `recursos/logo-esap.png` | Logotipo institucional usado en el encabezado. |

## Regenerar el documento

```bash
npm install docx
node formatos/generar-formato.js
```

## Paleta institucional

| Uso | Color |
| --- | --- |
| Azul ESAP (bandas de sección, logotipo) | `#25478C` |
| Azul profundo (numerales, rótulos) | `#1A3567` |
| Azul tenue (fondo de etiquetas) | `#E7ECF6` |
| Azul medio (encabezados de tabla) | `#CBD6EA` |
| Gris de líneas | `#B9C2D4` |
| Ámbar (campos obligatorios y aviso inicial) | `#B8860B` |

Tipografía: Arial en toda la pieza — 9 pt para los datos, 7,5 pt en versalita
espaciada para las etiquetas y 6,5–7 pt para notas y pie de página.
