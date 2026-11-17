# Formato de `marca.json`

El más pequeño para copiar es `ejemplos/lumen-cafe.json`. El completo es `ejemplos/north-face.json`. `generar.py` valida todo lo marcado con ✔ y no genera nada si hay errores.

## Datos del plan

| Campo | Qué es | ✔ Validación |
|---|---|---|
| `marca`, `mercado`, `slug` | Nombre, país y prefijo de los archivos | |
| `contexto` | Antetítulo de la portada (máster, módulo, profesor) | |
| `inicio`, `semanas` | Lunes de la semana 1 (AAAA-MM-DD) y número de semanas | ✔ tiene que ser lunes |
| `supuestos.base` | Suscriptores activos (supuesto si no hay CRM) | |
| `supuestos.control` | Grupo de control: no recibe los programados | |
| `supuestos.tope_semanal` | Máximo de programados por perfil y semana | ✔ ningún perfil lo supera |
| `supuestos.nuevos` | Altas nuevas en el periodo (reciben la bienvenida) | |
| `perfiles[]` | De 1 a 4: `k` (una letra mayúscula distinta de S), `nombre`, `corto`, `pct`, `max_ofertas` (opcional), `refutacion` y `tarjeta` | ✔ códigos únicos; ✔ ofertas ≤ `max_ofertas`; aviso si falta `refutacion`, si el titular no se envía o si un asunto suena a oferta para un perfil con `max_ofertas` |
| `perfiles[].tarjeta` | `territorio`, `decide`, `barrera`, `atributo`, `no` y `titular` (diapositiva 2) | |
| `sin_perfil` | `nombre`, `pct` y `recibe` (el código del perfil cuyo programa recibe) | ✔ los % suman 100 |
| `tipos[]` | Los 6 tipos de clase con `que`, `variables`, `kpi`, `uso`, `cuando_no`, `ejemplo`, `deck_que` y `deck_variable` | ✔ exactamente los 6 |
| `eventos[]` | Un texto por semana, con fecha verificada | ✔ uno por semana |
| `calendario[]` | `[semana, perfil o "Todos", tipo, día, asunto, objetivo, test A/B o ""]` | ✔ semana, perfil, tipo programable y día (Lun…Dom) |
| `triggers[]` | `nombre`, `perfiles` ("A,B" o "Nuevos"), `tasa`, `emails`, `plantillas` y `tope` (opcional) | ✔ perfiles existentes; tasa entre 0 y 1 |
| `textos` | Notas del Excel: `fuente_supuestos`, `nota_reparto`, `tipos_intro`, `reglas_calendario`, `lectura_matriz`, `variable_discrimina`, `resumen_contenidos`, `resumen_eventos` y `resumen_medicion` | |
| `medicion[]` | Filas `[concepto, texto]` de la hoja Medicion | |

## Datos reales (los escribe `importar.py`, no se tocan a mano)

| Campo | Qué es |
|---|---|
| `familias` | Solo para la forma B: `{"<familia de producto>": "<código de perfil>"}`. Lo escribes tú, antes de importar |
| `datos_reales` | `fuente`, `fecha`, `regla` (cómo se asignó el perfil), `conteo` (clientes por perfil y S) y `antes` (base, % y tasas supuestas antes de importar) |
| `triggers[].real` | `true` si la tasa viene de los datos de la marca |

Si existe `datos_reales`, el Excel añade la hoja `Datos_reales`: la base y los % de Supuestos pasan a ser fórmulas que leen de ella. El deck cambia «supuesto» por «dato real», y la defensa cuenta qué corrigió el dato.

## Cómo funcionan los triggers

- Con `perfiles: "Nuevos"`, el trigger lo recibe cada alta y la tasa no se usa. Es el caso de la bienvenida.
- `tasa` es el % de la audiencia del perfil que lo activa en todo el periodo. Es un supuesto de orden de magnitud y hay que decirlo así.
- Con `tope: true`, el trigger manda un email menos en las semanas en que el perfil ya llega al tope de programados (por ejemplo, el carrito en Black Friday).
- `plantillas` son piezas que hay que producir: una por perfil al que aplica, no por envío.

## Textos del deck (`deck`)

| Campo | Diapositiva |
|---|---|
| `estilo` | `oscuro`, `oscuro2`, `oscuro3` (opcional), `papel`, `acento`, `fuente_titulo`, `fuente_texto` y `tipos` (color por tipo programable) |
| `pie` | El pie de todas las diapositivas |
| `portada` | `titulo` (admite \n), `subtitulo`, `autores`, `antetitulo` (opcional) y `montanas: false` para quitar el motivo de montaña |
| `personas_titulo`, `personas_pie` | 2 |
| `tipos_lectura` | 3 |
| `meses[]` | 4: un texto por mes del periodo (el motor dice cuántos hacen falta) ✔ |
| `momentos_titulo`, `momentos[]`, `momentos_pie` | 5: cada momento con `tag`, `titulo` y `puntos[]` (2 o 3 momentos) |
| `numeros_lectura` | 6: la frase debajo de «emails por persona» |
| `medicion_titulo`, `tests_texto`, `revision{fecha, texto}`, `cierre` | 7 |
| `notas[]` (en la raíz) | Opcional: las 7 notas del orador ✔ |

Todos los números del deck (impactos, piezas, emails por persona, tabla, número de tests, % de control) salen de `cifras.json`. No los escribas en los textos.
