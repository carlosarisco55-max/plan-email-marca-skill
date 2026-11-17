# Cómo defender el Excel · Rodalia Bikes

## La idea en una frase
Todo sale de dos sitios: la hoja **Supuestos** (celdas amarillas) y la hoja **Calendario** (una fila por email). El resto son fórmulas.

## Ejemplo trabajado: C · Ciclista de carretera exigente
- Recibe **10 emails programados**. Su audiencia es el 15% de 40.000 = 6.000; sin el 10% de control quedan **5.400**.
- Programados: 10 × 5.400 = **54.000 impactos**.
- Triggers (solo a quien hace la acción, por eso van con tasa): Carrito abandonado (sin código de descuento en la plantilla de C) 8%, Reposición de consumibles (solo a quien compró cadena, cubiertas o pastillas) 15% → **1.860**.
- Total del perfil: **55.860**. Sumando todos los perfiles y la bienvenida (3.200): **377.356 impactos**.

## Piezas: 49
Pieza = creatividad distinta, no envío. 35 emails únicos + 4 variantes A/B + 1 común ('Todos') + 9 plantillas de trigger.

## Preguntas trampa
- **¿De dónde sale la base de 40.000?** Es un supuesto; sin el CRM de la marca no hay dato público. Se cambia una celda (o se importan los datos con motor/importar.py) y todo se recalcula.
- **¿Por qué ese reparto entre perfiles?** También es supuesto, y está marcado como tal. El método no cambia con el dato real.
- **¿Por qué 8,4–11,3 emails por persona si hay 8–11 programados?** Se suman los triggers que le tocan de media.
- **¿El control no recibe nada?** No recibe programados; sí triggers y transaccionales, porque responden a lo que hace el cliente. Lo que se mide es el efecto del plan.
- **¿Y la saturación en los picos?** Tope de 2 programados por semana y perfil; en esas semanas el trigger con tope manda un email menos, y la fórmula lo aplica.

## Demostración en directo
Cambia la base de 40.000 a 20.000 en Supuestos: los impactos cambian y las piezas siguen en 49, porque producir no depende de cuánta gente lo recibe.
