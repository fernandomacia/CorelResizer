# CorelResizer

Macro VBA para CorelDRAW (probada en X6, X7 y X8) que redimensiona múltiples
objetos manteniendo su posición. Pensada para diseños de strass.

## Forma de trabajo

- Cada cambio se propone por separado, con el código antes/después, y no se
  aplica hasta que se aprueba.
- Nunca se hacen commits, etiquetas ni Releases por iniciativa propia: solo
  cuando se piden (ver "Flujo de commits").
- Todo el proyecto está en español: código, comentarios, mensajes de la macro,
  documentación y commits. A diferencia de otros proyectos, aquí los nombres de
  variables también son en español (`entrada`, `Pregunta`), y se mantiene así
  para no mezclar idiomas dentro del mismo archivo.

## El archivo de la macro

- `CorelResizer.bas` está en disco en **Windows-1252 + CRLF**, porque el editor
  VBA de CorelDRAW no entiende UTF-8 y las tildes saldrían rotas al importar.
  Git lo guarda en UTF-8 gracias a `.gitattributes`
  (`working-tree-encoding=CP1252`), para que GitHub lo muestre bien.
- Nunca editarlo con herramientas que asuman UTF-8: leerlo y escribirlo como
  cp1252 y conservar CRLF. Comprobar después que no queda ningún LF suelto.
- El módulo se llama `modCorelResizer` (`Attribute VB_Name`) y la macro
  `CorelResizer`. Mismo nombre en ambos provoca ambigüedades en VBA.

## Formato de commits

    type(scope): Descripción corta

    - Viñeta que describe el cambio principal
    - Viñeta adicional (si hace falta)

El mensaje va en español. El prefijo `type(scope):` y `BREAKING CHANGE` se
quedan en inglés: son vocabulario de Conventional Commits.

Los commits anteriores al 2026-10-29 no siguen este formato y se quedan como
están: reescribirlos cambiaría todos los hashes a cambio de un log más bonito.

### Tipos

- `feat`: funcionalidad nueva → sube MINOR
- `fix`: corrección de error → sube PATCH
- `refactor`, `docs`, `style`, `perf`, `chore` → no suben versión

### Scopes

- `macro`: `CorelResizer.bas`
- `readme`: `README.md`
- `config`: `.gitattributes`, `.gitignore`, `.vscode/`
- `assets`: imágenes e iconos de `static/`

## Versionado

Versionado semántico (MAYOR.MENOR.PARCHE). Única fuente de verdad: la cabecera
de `CorelResizer.bas`, `'@version: X.Y.Z DD/MM/AAAA`. El README no repite el
número.

| Subida | Regla |
|---|---|
| PATCH | Corrección de error (`fix:`) |
| MINOR | Funcionalidad nueva compatible con el uso actual (`feat:`) |
| MAJOR | Cambio que obliga al usuario a usar la macro de otra forma |

La subida de versión va dentro del commit `feat:`/`fix:`, nunca en un commit
aparte.

## CHANGELOG

`CHANGELOG.md` sigue Keep a Changelog. Cada commit que sube versión añade su
entrada, con la fecha del commit, en la sección que corresponda (Added,
Changed, Fixed, Removed). Los títulos de sección en inglés, como en Keep a
Changelog; las entradas en español.

## Releases y etiquetas

Cuando se pida publicar una versión, en este orden:

1. Revisar que `README.md` y `CHANGELOG.md` describen el estado actual, no solo
   el último cambio. Si no, proponer la corrección y hacer el commit antes.
2. Crear una etiqueta anotada `vX.Y.Z` que coincida exactamente con la cabecera
   del `.bas`, con un mensaje que resuma los cambios desde la anterior.
3. Subir solo esa etiqueta: `git push origin vX.Y.Z` (nunca `git push --tags`).
   Comprobar con `git ls-remote --tags origin`.
4. Crear la Release de GitHub sobre la etiqueta y adjuntar el `.bas` del disco
   (ya en Windows-1252 + CRLF): quien lo descargue suelto desde la web lo
   recibiría en UTF-8.

## Flujo de commits

Cuando se pidan "los commits", nunca se ejecutan automáticamente:

1. Revisar los cambios: listar archivos modificados y nuevos y sus scopes.
2. Analizar la estrategia: decidir si forman un único cambio o varios, y
   ordenarlos empezando por los menos acoplados.
3. Un commit cada vez:
   - Si es `feat:` o `fix:`: subir la versión en la cabecera del `.bas` y
     añadir la entrada al `CHANGELOG.md`, e incluirlo en la lista de archivos.
   - Si es `feat:`: actualizar el `README.md` si cambia algo visible para el
     usuario e incluirlo en la lista.
   - `CLAUDE.md` va dentro del `feat:`/`fix:` cuando la regla que recoge nace
     de ese cambio; `docs:` queda para documentación sin código asociado.
   - Presentar el mensaje y la lista de archivos listos para copiar y pegar,
     sin trailer `Co-Authored-By:`.
   - Esperar confirmación antes de pasar al siguiente.

## Revisión del código

- [ ] `Option Explicit` y todas las variables declaradas con tipo cuando sea
      posible
- [ ] Comprobaciones de entrada antes de modificar el documento
- [ ] Toda modificación del documento dentro de
      `BeginCommandGroup`/`EndCommandGroup` (un solo Ctrl+Z)
- [ ] `Optimization`/`EventsEnabled` y la configuración del documento
      (unidad, punto de referencia) restaurados también si hay error
- [ ] Archivo en Windows-1252 + CRLF, sin LF sueltos
- [ ] Mensajes al usuario en español con ortografía correcta (`¿…?`, tildes)
