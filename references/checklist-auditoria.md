# Checklist de auditoría (fase 6, con subagente y contexto limpio)

Cada punto es un error que se cometió de verdad en el caso de The North Face y se corrigió antes de presentar. El subagente recibe:
- el documento de personas;
- `marca.json`;
- la tabla de datos verificados (dato · fuente · fecha).

Devuelve una lista de problemas con su propuesta de arreglo. No los arregla él.

## Personas frente a calendario
- [ ] **La barrera manda.** Si un perfil «compra a precio completo» o «no quiere descuentos», ¿recibe ofertas? (`max_ofertas` lo bloquea, pero revisa también los asuntos que suenan a oferta en emails de otro tipo).
- [ ] **El titular del perfil aparece en el calendario.** Si no se envía, el perfil no se ha probado.
- [ ] **Prueba del mensaje único.** Si intercambias los asuntos de dos perfiles y no suena raro, son el mismo perfil.
- [ ] **Quién decide y quién paga.** En las fechas de regalo, ¿hay un mensaje para quien paga y está el test que separa los dos discursos?
- [ ] **Las familias de producto son exclusivas.** Un SKU pertenece a una sola familia; si no, la variable que discrimina cuenta doble.
- [ ] **El grupo sin perfil** tiene asignado el programa de otro perfil y la bienvenida le asigna perfil por el clic.

## Segmentación frente a individuo
- [ ] **Recompra, renovación y reposición van a quien tiene el producto, no a todo el segmento.** En The North Face se habían asignado por segmento y llegaban a 515 personas que no lo tenían.
- [ ] Quien compra en el pico sale de la oferta siguiente («sin quien ya compró»).

## Datos externos
- [ ] Las **fechas de promoción son las de la marca** (en The North Face, la Black Week iba del 24-nov al 1-dic, con acceso anticipado el 19), no las del evento genérico.
- [ ] **Urgencia verdadera.** «Último día» solo si lo es; si no, «acaba mañana».
- [ ] Los **servicios existen en ese país**: reparaciones, recogida de usado, ticket regalo frente a cambio de talla gratis. Mira dónde y en qué tiendas.
- [ ] **Ninguna cifra técnica inventada.** Una temperatura, un gramaje o un % sin fuente se quita.
- [ ] **El aniversario es de quien es.** Era la marca la que cumplía 60, no el producto.
- [ ] Cada dato del deck tiene fila en la tabla de verificación, con su fecha de consulta.

## Números
- [ ] Ninguna cifra escrita a mano en el deck, el guion o el documento de personas: todas salen de `cifras.json` o de `defensa_<slug>.md`.
- [ ] Los mismos números en todos los documentos. En The North Face convivieron 501/90 y 515/85 hasta la auditoría.
- [ ] Los porcentajes del reparto están etiquetados como **supuesto** donde aparecen.
- [ ] Emails por persona = programados + triggers esperados / audiencia total (no la que recibe).
- [ ] Si se pide una fecha de revisión, que tenga corte provisional (tests tempranos) y revisión completa.

## Entrega
- [ ] Las fuentes del deck existen en el equipo que proyecta. Bahnschrift viene con Windows y no está en Mac: si el ordenador es un Mac, manda el PDF o cambia `estilo`.
- [ ] Si un compañero ya ha editado el PPTX, **no regeneres encima**: los cambios de texto se le pasan o se aplican a mano en su versión.
- [ ] Al guion le quedan nombres reales en lugar de los marcadores ([Nombre 1]) y su reparto coincide con las diapositivas.
