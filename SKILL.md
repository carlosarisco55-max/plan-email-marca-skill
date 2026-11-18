---
name: plan-email-marca
description: Use when a brand of any sector needs an email-only communication plan by buyer persona over a fixed number of weeks - newsletter content proposal, week-by-week calendar with commercial events, impacts per persona and total pieces to produce - as a class assignment, final project, agency proposal or pitch; also when someone asks to repeat "el plan de email de The North Face" for another brand, or to defend the numbers of such a plan.
---

# Plan de email por buyer persona, para cualquier marca

**Principio:** el criterio va en un archivo (`marca.json`) y la mecánica en un motor. Las cifras de impactos y piezas **solo** salen del motor, nunca de una cuenta a mano: así el Excel, el deck y el guion dicen lo mismo, y cambiar un supuesto lo recalcula todo.

Caso de origen: The North Face, 16 semanas, 3 perfiles. El motor lo reproduce exacto (50 emails · 1.661.240 impactos · 67 piezas).

## Qué se entrega y qué pregunta responde

| Pregunta típica del brief | Excel (hoja) | Deck (diapositiva) |
|---|---|---|
| ¿Quién? Buyer personas | Supuestos | 2 · tarjetas con barrera, quién decide y titular |
| 1 · Propuesta de contenidos | Tipos_Email | 3 · tipos × perfiles |
| 2 · Programa semana a semana | Calendario · Matriz_Semanal | 4 · rueda · 5 · momentos clave |
| 3 · Impactos por persona | Impactos | 6 · números |
| 4 · Total de piezas | Piezas | 6 · números |
| ¿Cómo sabremos si funciona? | Medicion | 7 · control, tests y revisión |

Además se entregan:
- el guion de los ponentes en PDF;
- la defensa en PDF, con el ejemplo trabajado y las preguntas trampa;
- el deck también en PDF, más las vistas previas en PNG (todo lo genera `entrega.py`);
- dos documentos a mano: `verificacion.md` (dato · fuente · fecha) y el documento de personas.

Ejemplos generados, para ver cómo queda: `ejemplos/muestras/` (The North Face, Lumen Café y Rodalia Bikes).

## Proceso (no saltar fases)

1. **Brief.** Copia las preguntas literales. Fija el lunes de inicio, el número de semanas, los ponentes y la duración. Pregunta solo lo que cambie el plan.
2. **Investigación con fuente y fecha.** Busca en la web, no de memoria:
   - la segmentación oficial de la marca (presentaciones a inversores, memoria anual);
   - el programa de fidelización real;
   - las fechas reales de sus promociones (no las del evento genérico);
   - garantía, devoluciones y servicios **en ese país**.

   Resúmelo en una tabla de dato · fuente · fecha. Lo que no esté verificado no entra en ninguna pieza. Si la marca es ficticia, o si el dato no se puede verificar, márcalo como «supuesto a confirmar» en el JSON y en la diapositiva; nunca lo presentes como un hecho.
3. **Personas.** **REQUIRED SUB-SKILL:** usa `buyer-persona-hipotesis`. Necesitas de 2 a 4 perfiles más un grupo sin perfil, con su %. Para cada perfil:
   - la variable que lo discrimina;
   - la barrera, escrita antes que el beneficio;
   - quién decide (no quién usa);
   - el titular;
   - la condición de refutación (campo `refutacion`) y la fecha de revisión.

   Si la barrera del perfil prohíbe descuentos, ponle `max_ofertas` en el JSON.

   **Si la marca da sus números** (un Excel o CSV de clientes, segmentos o ventas), usa `importar.py` en lugar de suponer la base y el reparto. Mira las dos formas de tabla en `python M/importar.py -h`. El motor guarda el supuesto anterior, y el Excel muestra los dos en la hoja `Datos_reales`. La diferencia entre lo que suponíamos y el dato es un hallazgo: cuéntalo en la defensa. Un export real casi nunca viene limpio: `--saltar N` salta filas de título por encima de la cabecera, y una fila con cliente sin id, importe o fecha rotos se descarta con un AVISO en vez de tumbar todo el import. Lo que no arregla solo: celdas combinadas (una familia repartida por varias filas) hay que deshacerlas en el Excel antes de importar.
4. **Calendario en `marca.json`.** Copia `ejemplos/lumen-cafe.json`. El formato está en `references/formato-marca-json.md`. Las reglas:
   - Por norma, 1 email por perfil y semana. Los 2 de los picos son solo para los perfiles que compran en ese pico. Un perfil fuera de su temporada puede quedarse sin email (es un AVISO que se justifica).
   - `Todos` solo para un mensaje de marca.
   - Cada test A/B pone a prueba una hipótesis de perfil y cambia solo la primera línea (`libro-ch14-afinidad-al-mensaje`).
   - Los triggers no van en el calendario: van en `triggers`, con su tasa marcada como supuesto.
5. **Motor.** Ejecuta `generar.py`. Si da un ERROR, se corrige el JSON, nunca el Excel. Cada AVISO se decide y se justifica (por ejemplo, una semana sin email puesta a propósito).
6. **Auditoría con subagente**, con el contexto limpio. Pásale `references/checklist-auditoria.md`, el documento de personas, `marca.json` y la tabla de datos verificados. Corrige y vuelve a la fase 5.
7. **Guion.** Escribe `guion.md` junto a `marca.json`. Cuenta una historia con un personaje por perfil cuya inicial sea la del código (Álex = A). Los ponentes se alternan y cada uno pasa la palabra con un gancho. El formato de cabecera está en `ejemplos/north-face-guion.md`. Las notas del orador van en `notas`, una por diapositiva.
8. **Entrega visible (siempre).** Ejecuta `entrega.py`. Genera la carpeta `entrega_<SLUG>/` con el Excel, el deck en .pptx y en PDF, `vista_previa/*.png`, la defensa en PDF y el guion en PDF. Mira las vistas previas, una a una (QA de `deck-pptx`): si algo se sale, acorta el JSON y repite. Con 4 perfiles, las tarjetas admiten poco texto. **Termina siempre enseñando al usuario la vista previa y la lista de archivos con su ruta.** Que el trabajo existe se demuestra con la carpeta, no con un resumen.

## Comandos

`M` es la carpeta `motor/` de esta skill. Lánzalos desde PowerShell: desde Git Bash, Edge no imprime los PDF y se usa LibreOffice, con peor resultado.

```
python M/entrega.py marca.json [--abrir]          # TODO: valida, Excel, deck, PDF, vistas previas, defensa y guion
python M/importar.py marca.json datos.xlsx ...    # datos reales de la marca (ver -h)
python M/generar.py marca.json salida/            # solo Excel + cifras.json + defensa (para iterar el calendario)
python M/test_motor.py                            # después de tocar el motor
```

Cada pieza por separado: `deck.js`, `render.py` y `md_pdf.py`.

## Reglas que el motor ya aplica (no las rehagas a mano)

| Sin la skill se hace así | Con el motor |
|---|---|
| Impactos = audiencia × emails, y el control se resta al final («serían ×0,9») | Los programados se calculan sobre la audiencia sin el control. Los triggers, sobre la audiencia total, porque también llegan al control |
| Dos totales de piezas distintos («32… o 28 únicas») | Una sola definición: emails únicos + variantes A/B + comunes + plantillas de trigger |
| Ningún test A/B | Las variantes cuentan como piezas y la diapositiva 7 enseña el número de tests |
| Una oferta al perfil «sin descuentos» porque «no es la suya» | `max_ofertas` la convierte en ERROR. Si el descuento se cuela en un email de otro tipo, da AVISO |
| Un titular de perfil que nunca se envía | AVISO: ese perfil no se llega a probar |
| Más de 2 emails en una semana pico | Tope semanal. El trigger con `tope` manda un email menos en esas semanas |
| Fechas de memoria | Las semanas y los meses salen del lunes de inicio; los eventos los pones tú, con fuente |

## Errores comunes

Están en `references/checklist-auditoria.md` y salieron del caso de The North Face:
- inventar un dato técnico o una fecha de promoción;
- asignar la recompra por segmento en vez de a quien tiene el producto;
- cifras distintas entre documentos;
- atribuir a un producto un aniversario que es de la marca;
- ofrecer un servicio que no existe en ese país;
- regenerar el PPTX encima de la versión que ha editado un compañero.

**No vale para:** un plan multicanal (SMS, redes o paid). En ese caso, amplía el motor solo cuando el brief lo pida.
