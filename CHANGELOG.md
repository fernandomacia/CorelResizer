# Changelog

Todos los cambios relevantes de este proyecto se documentan en este archivo.

El formato sigue [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y el proyecto usa [Versionado Semántico](https://semver.org/lang/es/) desde la
versión 1.1.0. Las versiones anteriores se numeraban 1.03 y 1.04; aquí se
recogen como 1.0.3 y 1.0.4.

## [1.2.0] - 2026-11-02

### Added
- Además de círculos, la macro ajusta elipses, cuadrados, rectángulos, rombos y cometas: un diámetro o lado para círculos y cuadrados, y dos medidas (larga y corta) para el resto.
- Corrige formas deformadas: cada forma recupera su forma correcta con las medidas indicadas, manteniendo su centro y su orientación aunque estuviera estirada, girada o inclinada. Si solo se estiró, recupera exactamente su giro original.
- Formulario único (`frmCorelResizer`) que sustituye a los cuadros de diálogo: tipo de forma, medidas, qué hacer al terminar y avisos de error dentro del propio formulario.
- Comprobación previa de que todas las formas seleccionadas encajan con el tipo elegido; si no, se avisa sin modificar el documento.
- Al terminar se puede elegir entre no hacer nada, agrupar o combinar las formas.
- El formulario recuerda la última configuración usada.

### Changed
- Los rectángulos, elipses y polígonos nativos de CorelDRAW se convierten a curvas antes de ajustarlos.
- El formulario cuenta formas en lugar de objetos: una curva combinada con 500 cristales cuenta como 500.
- La instalación requiere importar también `frmCorelResizer.frm` (con su `.frx`).
- Nuevas capturas en el README.

### Removed
- El cuadro de diálogo para el tamaño y la pregunta final para combinar los objetos.

## [1.1.0] - 2026-10-29

### Added
- Se acepta tanto el punto como la coma como separador decimal (`2.5` o `2,5`).
- Todo el proceso (desagrupar, redimensionar y combinar) se deshace con un solo `Ctrl + Z`.
- Aviso cuando el tamaño introducido es 0 o negativo.
- Si ocurre un error durante el proceso se muestra un mensaje y se restaura la configuración del documento.
- Instrucciones de instalación, uso y creación del botón en el README, y créditos del icono.

### Changed
- La licencia pasa de MIT a GPL-3.0, con el archivo `LICENSE` y el aviso en la cabecera de la macro.
- El archivo de la macro pasa de `CorelResizer.vb` a `CorelResizer.bas`, en Windows-1252 + CRLF, y el módulo se llama `modCorelResizer`.
- El redibujado se desactiva mientras se procesan los objetos, lo que acelera los diseños grandes.
- El cuadro de diálogo indica que el tamaño es en milímetros.

### Fixed
- Error en tiempo de ejecución al lanzar la macro sin ningún documento abierto.
- La selección se vuelve a leer tras desagrupar, para procesar los objetos que estaban dentro de grupos.
- La unidad de medida y el punto de referencia del documento ya no quedan modificados tras ejecutar la macro.
- Ortografía de los mensajes (`¿Quieres combinar los objetos?`).

## [1.0.4] - 2024-07-09

### Added
- Aviso al cancelar el cuadro de diálogo del tamaño.

### Changed
- La macro pasa a llamarse `CorelResizer` (antes `SS06`).

## [1.0.3] - 2021-10-27

### Added
- Primera versión: redimensiona los objetos seleccionados a un tamaño dado manteniendo su posición, y ofrece combinarlos.
