# Import de precios desde documento Word — Diseño

Fecha: 2026-08-20
Estado: aprobado, pendiente de implementación
Archivo afectado: `index.html` (app de un solo archivo, sin build ni backend)

## 1. Objetivo

Agregar un botón en la pestaña **Lista de Precios** que lea un documento Word con la
lista de precios vigente y actualice las cuatro listas de la app, mostrando antes un
preview editable que se puede cancelar sin efecto.

Hoy los precios se cargan a mano, uno por uno, o vía import/export de JSON que la app
misma genera. Los precios de verdad llegan como documentos Word desde el proveedor.

## 2. Alcance

Incluye:

- Lectura de `.doc` (Word 97-2003 binario, OLE2/CFB), `.docx` (OOXML) y texto pegado.
- Selección de **un solo archivo** (`input type=file` sin `multiple`).
- Parseo de filas de precio y ruteo automático a las cuatro listas.
- Matching contra las entradas existentes, con normalización y coincidencia difusa.
- Preview editable con cancelar/aceptar.
- Aplicación de los cambios aceptados y persistencia.

No incluye (decidido explícitamente):

- OCR. Los documentos no son imágenes escaneadas; el texto está en el archivo.
- Borrado de entradas de la app ausentes del documento (ver política en §6).
- Cambio de nombre o categoría de entradas existentes: el import solo toca `precio`.
- Import de más de un archivo por vez.

## 3. Datos de entrada reales

Los dos documentos que se usaron para diseñar y que sirven de caso de prueba:

### `GRANITOS, MARMOLES Y PURASTONE PRIMA.doc`

72 líneas con precio. Estructura por línea: `NOMBRE \t [ESPESOR] \t PRECIO`,
con encabezados de categoría en líneas sueltas sin precio.

- 59 líneas con espesor (`2 CM`, `12 MM`) → **Materiales**, en 5 categorías.
- 12 líneas sin espesor (`TRAFORO BACHA`, `ARMADO DE PATA`, …) → **Agregados**.
- 1 línea `METRO LINEAL DE FRENTE DOBLE - INGLETE` → **Frente doble**.

Encabezados de categoría de material, y su equivalente en la app:

| En el documento | Categoría en la app |
|---|---|
| `GRANITOS NACIONALES` | `Granitos Nacionales` |
| `GRANITOS IMPORTADOS` | `Granitos Importados` |
| `MÁRMOLES IMPORTADOS` | `Marmoles Importados` |
| `PURASTONE PRIMA` | `Pura Stone Prima` |
| `PURASTONE` (inline, primera columna de la fila) | `Purastone` |

El caso `PURASTONE` es especial: no es un encabezado sino la primera columna de la
fila (`PURASTONE \t BLANCO CANA \t 2 CM \t USD 385`). El parser tiene que detectarlo,
usarlo como categoría y quitarlo del nombre.

Además el bloque de Agregados aparece **entre** categorías de material (después de
`MARRON BOSQUE`, antes de las filas `PURASTONE`). Una fila sin espesor no debe
resetear la categoría de material vigente.

### `PILETAS ACERO JOHNSON.doc`

97 líneas con precio. Estructura: `CODIGO \t MEDIDAS \t DESCRIPCION \t PRECIO`,
agrupadas por secciones. Todas van a **Bachas**.

El nombre en la app se reconstruye como `PREFIJO + CODIGO + " (" + MEDIDAS + ")"`:

| Encabezado de sección | Prefijo | Ejemplo en la app |
|---|---|---|
| `LAVADERO ACERO 304` | `LAVADERO` | `LAVADERO LN 50 (50 x 40 x 25)` |
| `SIMPLES ACERO 304` | `SIMPLE` | `SIMPLE SIGNATURE ENKEL SE 45 (45 x 42 x 22)` |
| `DOBLES ACERO 304` | `DOBLE` | `DOBLE QUADRA Q 084 A (83,8 x 55,9 x 17,5)` |
| `ESQUINERAS DOBLES ACERO 304` | `ESQUINERA` | `ESQUINERA X 28 (61,5 x 61,5 x 15)` |
| `TRIPLE c/ accesorios incluidos (…) ACERO 304` | `TRIPLE` | `TRIPLE HYDRA J 107 A (1,07 x 43,5 x 18)` |
| `VANITORY ACERO 304` | `VANITORY` | `VANITORY SIGN. AURA 38 ROSE GOLD (38 dm x 11)` |
| `ACCESORIOS OPCIONALES` | ninguno | `APIDO - DOSIFICADOR` |

`ACCESORIOS OPCIONALES` usa otro formato: `CODIGO \t DESCRIPCION \t PRECIO`, sin
medidas, y el nombre en la app es `CODIGO + " - " + DESCRIPCION`.

Notar que el encabezado de `TRIPLE` no está todo en mayúsculas: la detección de
encabezados no puede depender de eso.

Cada sección repite una fila de rótulos `MEDIDAS / DESCRIPCION / PRECIO`. Se descarta
sola porque no tiene dígitos donde va el precio.

### Aserción de calibración

Los defaults de la app (`DEFAULT_MATERIALES` 59, `DEFAULT_ITEMS` 12,
`DEFAULT_FRENTES` 1, `DEFAULT_OTROS` 97 — total **169**) fueron transcritos a mano de
estos mismos documentos.

Por lo tanto, sobre una app con los precios por defecto, importar los dos documentos
debe producir **169 filas matcheadas, 0 filas nuevas y 0 sin matchear**. Cualquier
fila que caiga como "nueva" es un defecto del matcher. Esta es la prueba de
aceptación principal.

## 4. Arquitectura

```
 .doc  ─→ readDocBinary()  ┐
.docx  ─→ readDocxText()   ├─→ parsePriceRows() ─→ matchRows() ─→ [preview] ─→ applyImport()
pegar  ─→ (identidad)      ┘        texto→filas      filas→dif    editable      muta + guarda
```

Los tres orígenes convergen en **texto plano**. Todo lo que sigue es común. Cada
unidad recibe y devuelve datos simples, sin tocar el DOM ni el estado global salvo
`applyImport`.

### 4.1 `readDocBinary(arrayBuffer) → string`

Parseo del contenedor CFB (OLE2) y extracción del texto del documento Word.

1. **Header** en offset 0. Firma `D0 CF 11 E0 A1 B1 1A E1`. Campos: `sectorShift`
   (0x1E), `miniSectorShift` (0x20), `numFatSectors` (0x2C), `firstDirSector` (0x30),
   `firstMiniFatSector` (0x3C), `numMiniFatSectors` (0x40), `firstDifatSector` (0x44),
   `numDifatSectors` (0x48).
2. **FAT**: primeros 109 punteros en el header (offset 0x4C), el resto siguiendo la
   cadena DIFAT.
3. **Directorio**: entradas de 128 bytes. Nombre en UTF-16LE (64 bytes), tipo en 0x42,
   sector inicial en 0x74, tamaño en 0x78.
4. **Streams necesarios**: `WordDocument`, y `1Table` o `0Table` según el bit
   `fWhichTblStm` (bit 9 de los flags del FIB en offset 0x0A del `WordDocument`).
   Streams menores a 4096 bytes viven en el miniFAT.
5. **FIB**: `fcClx` en offset 0x01A2, `lcbClx` en 0x01A6. **Estos dos offsets son el
   único punto del diseño que hay que validar empíricamente contra los archivos
   reales antes de seguir.**
6. **CLX** en el Table stream, en `fcClx..fcClx+lcbClx`: saltar los bloques `Prc`
   (byte `0x01` seguido de un `int16` con el largo), hasta el byte `0x02` seguido de
   un `int32` con el largo del `PlcPcd`.
7. **PlcPcd**: `n+1` CPs (`int32`) y después `n` PCDs de 8 bytes. En cada PCD, el `fc`
   es un `int32` en offset 2. Si el bit 30 (`0x40000000`) está prendido, el tramo es
   texto de 8 bits CP1252 y el offset real es `(fc & ~0x40000000) / 2`; si está
   apagado, es UTF-16LE en `fc`.
8. **Decodificación** de cada tramo y concatenación.

Mapeo de caracteres de control, necesario para preservar las columnas:

| Byte | Se convierte en | Por qué |
|---|---|---|
| `0x07` | tab | fin de celda / fila de tabla |
| `0x0D`, `0x0B` | salto de línea | fin de párrafo / salto de línea |
| `0x1E`, `0x1F` | `-` | guión de no separación / opcional |
| `0x13`, `0x14`, `0x15` | se descartan | delimitadores de campo |

Si el CLX no se puede parsear, la función lanza un error con un mensaje que apunta a
la opción de pegar texto. No hay extracción por fuerza bruta: es preferible fallar
claro que entregar filas basura.

### 4.2 `readDocxText(arrayBuffer) → Promise<string>`

1. Buscar el EOCD del ZIP (`50 4B 05 06`) desde el final.
2. Recorrer el directorio central hasta `word/document.xml`.
3. Inflar con `DecompressionStream('deflate-raw')` (nativo). Método 0 = almacenado,
   se usa tal cual.
4. XML a texto: `</w:p>` → salto de línea, `<w:tab/>` → tab, contenido de los `<w:t>`,
   quitar el resto de las etiquetas, decodificar entidades.

Es asíncrona, así que el pipeline entero es asíncrono.

### 4.3 `parsePriceRows(texto) → fila[]`

Recorre línea por línea manteniendo dos variables de estado: `catMaterial` y
`seccionBacha`.

Reconocedores:

- Precio: `/(USD|U\$S|\$)\s*([\d.,]+)\s*$/` — moneda y monto. Parseo es-AR: el punto
  es separador de miles, la coma es decimal.
- Espesor: `/(\d+(?:[.,]\d+)?)\s*(CM|MM)\b/i`.
- Medidas: `/\d+(?:[.,]\d+)?\s*(?:dm\s*)?x\s*\d+(?:[.,]\d+)?(?:\s*x\s*\d+(?:[.,]\d+)?)?/i`.
- Encabezado: línea sin precio, con a lo sumo un tab, de 2 a 8 palabras.

Ruteo (en este orden):

1. Coincide con `/frente doble/i` → `frentes`.
2. Tiene espesor → `materiales`, categoría = `catMaterial` traducida por la tabla de
   §3, o la primera columna si es una categoría inline conocida.
3. Estamos dentro de un documento de piletas (hay `seccionBacha` vigente) → `otros`,
   con el nombre reconstruido según la tabla de §3.
4. Resto → `items`.

Convención de mayúsculas al insertar, para respetar cómo guarda hoy cada lista:
`materiales` y `otros` en MAYÚSCULAS, `items` y `frentes` en formato oración.
Solo aplica a inserciones; las actualizaciones no tocan el nombre.

Devuelve `{destino, nombre, espesor, medidas, precio, moneda, lineaOriginal}`.

### 4.4 `matchRows(filas) → filaConDif[]`

`normKey(s)`: mayúsculas, sin tildes (NFD y descarte de diacríticos), guiones
unificados (`–`, `—` → `-`), sin contenido entre paréntesis, sin caracteres no
alfanuméricos salvo espacios y dígitos, espacios colapsados.

1. Coincidencia exacta sobre `normKey` dentro de la lista de destino.
2. Para `otros`, intento adicional por código, que es la parte estable del nombre
   (prefijo de sección y medidas removidos).
3. Difuso, con la restricción de abajo.

### Por qué el difuso no puede ser un umbral de similitud

La versión original de este spec aceptaba cualquier par con Levenshtein ≥ 0.86 y
marcaba `dudoso` por debajo de 0.95. **Se implementó, se probó y falló.** Los nombres
de esta lista son códigos densos donde cada token es significativo:

| Par | Ratio | ¿Mismo artículo? |
|---|---|---|
| `SIMPLE SIGNATURE ENKEL SE 45` / `... SE 55` | 0.96 | **No** |
| `SIMPLE SIGNATURE ENKEL SE 45` / `... SE 45 GB` | 0.90 | **No** |
| `SIMPLE O 37 A` / `SIMPLE O 37 AT` | 0.96 | **No** |
| `TRAFORO BACHA` / `TRASFORO BACHA` | 0.92 | Sí |

Con el umbral, borrar `SE 45` de la app hacía que la fila `SE 45` del documento
matcheara `SE 55` — un artículo distinto — como `actualiza` y **marcada por defecto,
sin aviso de dudoso**, porque 0.96 supera 0.95. Aceptar el preview le habría puesto a
`SE 55` el precio de `SE 45`. Es el falso positivo que la política conservadora de §6
no alcanza a cubrir, porque el daño ocurre en una fila que se presenta como segura.

### Regla efectiva

Un match difuso se acepta **solo** si:

1. Los dos nombres tienen la **misma cantidad de tokens**.
2. Tienen la **misma secuencia de dígitos** (`45` ≠ `55`).
3. Difieren en **exactamente un token**.
4. Ese token difiere pero sigue siendo reconocible: ambos de **4 caracteres o más** y
   con similitud ≥ 0.75 entre sí.

Las condiciones 1 y 3 descartan `SE 45` / `SE 45 GB`. La 2 descarta `SE 45` / `SE 55`.
La 4 descarta `O 37 A` / `O 37 AT`, donde el token que cambia es demasiado corto para
que un carácter de diferencia signifique un error de tipeo. `TRAFORO` / `TRASFORO`
cumple las cuatro y sigue matcheando.

**Todo match no exacto se marca `dudoso`**, sin umbral: si hubo que recurrir al difuso,
vale mirarlo. Lo que la regla rechaza cae como `nuevo` y desmarcado, que es el lado
seguro: el peor caso es cargar un precio a mano.

Efecto colateral útil: la regla destapó que el documento dice
`TABLA PICAR QUADRA VIDRIO TEMPADO` donde la app dice `TEMPLADO`. Uno de los dos tiene
un error de tipeo, y antes pasaba invisible.

Estado de cada fila:

- `actualiza` — matchea y el precio difiere.
- `sin cambios` — matchea y el precio es igual.
- `nuevo` — no matcheó nada.

Casos borde, definidos para que no queden a criterio de la implementación:

- **Dos filas del documento que matchean la misma entrada.** Gana la primera; la
  segunda queda como `nuevo` con un badge "duplicado: ya matcheó la fila N". Evita
  que dos filas escriban el mismo precio pisándose.
- **El matching es siempre dentro de la lista de destino.** Una fila ruteada a
  `items` no matchea contra `otros`, aunque el nombre exista allá.
- **Cambiar el destino en el preview vuelve a correr el match** de esa fila contra la
  lista nueva, y recalcula su estado. Es el mecanismo por el que se corrige un ruteo
  equivocado.

### 4.5 Preview

Modal con el patrón que ya usa la app (`.modal-overlay` / `.modal`).

Encabezado con el resumen: cuántas actualizan, cuántas sin cambios, cuántas nuevas,
y cuántas entradas de la app no aparecen en el documento.

Una fila por item, todas editables:
`[✓] destino(select) · nombre(input) · espesor(input) · precio(input) · moneda(select) · estado`

- `actualiza`: marcada, muestra `$195.200 → $210.000` con el viejo tachado.
- `sin cambios`: desmarcada y atenuada.
- `nuevo`: **desmarcada** (política conservadora), con badge "nuevo".
- `dudoso`: badge de aviso indicando con qué entrada matcheó.

Barra de herramientas: marcar todos / desmarcar todos, y un filtro por texto.

Al pie, una sección de solo lectura "no están en el documento", sin acción posible.

Botones: **Cancelar** (cierra sin efecto) y **Aceptar (N cambios)**.

### 4.6 `applyImport(filas)`

Procesa solo las filas marcadas. `actualiza` asigna `precio` por índice. `nuevo`
inserta, y en `materiales` lo hace al final de su categoría, con la misma lógica de
inserción que ya usa `addNewMaterial`. Después `savePrecios()`, `renderPrecios()`,
`populateItemSelect()` y un resumen de lo aplicado.

## 5. Ubicación en la interfaz

Card propia como **primer elemento** de `#tab-precios`, arriba de "Lista de Precios
de Materiales":

- Título: "Actualizar precios desde documento".
- `input type=file` con `accept=".doc,.docx"` y **sin** `multiple`.
- Link "o pegar texto" que despliega un textarea y un botón para procesarlo.
- Nota breve de qué hace y que muestra un preview antes de aplicar.

## 6. Política de cambios

Decidida explícitamente (variante conservadora):

| Situación | Comportamiento |
|---|---|
| Precio distinto en algo que ya existe | Se actualiza, marcado por defecto |
| Fila del documento que no matchea nada | Se ofrece insertar, **desmarcada** por defecto |
| Entrada de la app ausente del documento | **No se borra nunca**; se lista como aviso |
| Nombre o categoría de algo existente | No se modifica: el import solo toca `precio` |

El motivo de no borrar: el matching difuso puede fallar. Un falso negativo en un
esquema de sincronización total duplicaría el item y borraría el original en la misma
pasada. Acá el peor caso es que no pase nada y se corrija a mano.

## 7. Manejo de errores

- Firma no válida (no es CFB ni ZIP) → mensaje claro que apunta a pegar texto.
- CLX ilegible → ídem.
- Cero filas con precio → "no encontré filas con precio; revisá el formato o pegá el
  texto".
- Todo el parseo va dentro de `try/catch`.
- **Nada muta hasta Aceptar.** El parseo y el matching son funciones puras sobre
  copias; el único punto de mutación es `applyImport`.

## 8. Pruebas

### Harness en Node (scratchpad, descartable)

Corre el parser contra los dos `.doc` reales y verifica:

- Conteos: 59 materiales + 12 agregados + 1 frente doble desde granitos; 97 bachas
  desde piletas.
- Valores puntuales: `ROSA DEL SALTO` = 195200 ARS, `NEGRO BRASIL` = 320 USD,
  `MARRON BOSQUE` = 640 USD, `CALACATTA VAGLI - P` = 610 USD,
  `TRAFORO BACHA` = 72000, `LN 50` = 199900, `KG` = 8900.
- Categorías: las 5 de materiales, correctamente traducidas.
- Nombres reconstruidos de bachas idénticos a los de `DEFAULT_OTROS`.

### Aserción de aceptación

Sobre precios por defecto, importar los dos documentos da **169 matches, 0 nuevas,
0 sin matchear** (§3). Es la prueba que cubre parser, ruteo y matcher a la vez.

### Navegador

- Las mismas aserciones vía `evaluate_script`.
- Corrida completa: elegir archivo → preview → editar una fila → aceptar → verificar
  `localStorage`.
- Cancelar no deja rastro: snapshot de `localStorage` antes y después, idénticos.
- Camino `.docx`: generar uno con `textutil` desde el `.doc` y verificar que da las
  mismas filas.
- Camino pegar texto: pegar el `.txt` extraído y verificar lo mismo.
- Consola sin errores.

## 9. Riesgo declarado — resuelto

El parser de `.doc` binario era la única parte que no se podía garantizar sin
ejecutarla. **Validado:** los offsets del FIB (§4.1, paso 5) resultaron correctos y
los dos documentos parsean limpio, incluyendo las tildes de `MÁRMOLES` y la
conversión de fin-de-celda a tabs que preserva las columnas.

## 10. Resultado de la implementación

Todo lo de §8 pasa. Además de la aserción de aceptación, la implementación encontró
tres defectos que el spec original no anticipaba:

1. **Columnas separadas por espacios, no por tabs.** Siete filas `SIGNATURE ENKEL`
   traen espacios entre el código y las medidas, así que dividir por tabs las metía
   en una sola columna. Las medidas se buscan ahora en la línea entera y el código es
   lo que va antes.
2. **Descripciones que imitan medidas.** `REJILLA ENROLLABLE RECT 32X42` matchea el
   regex de medidas. El formato lo decide ahora la sección (`ACCESORIOS` nunca busca
   medidas), como decía §3, en lugar de inferirse de si aparecen medidas.
3. **El umbral de similitud era inseguro.** Ver §4.4.

Verificado en el navegador: los tres orígenes (`.doc`, `.docx`, pegar texto) producen
las mismas filas, Cancelar deja `localStorage` byte-idéntico, un precio editado a mano
en el preview gana sobre el del documento, una fila desmarcada no se toca, y las
inserciones de material caen al final de su categoría. Cero errores de consola.
