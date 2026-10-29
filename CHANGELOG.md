# Changelog

Todos los cambios relevantes de este proyecto se documentan en este archivo.

El formato sigue [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y el proyecto usa [Versionado Semántico](https://semver.org/lang/es/) desde la
versión 1.1.0. Las versiones anteriores se numeraban 1.03 y 1.04; aquí se
recogen como 1.0.3 y 1.0.4.

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
