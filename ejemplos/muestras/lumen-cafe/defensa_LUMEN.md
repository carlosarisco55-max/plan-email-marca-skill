# Cómo defender el Excel · Lumen Café

## La idea en una frase
Todo sale de dos sitios: la hoja **Supuestos** (celdas amarillas) y la hoja **Calendario** (una fila por email). El resto son fórmulas.

## Ejemplo trabajado: R · Recurrente de suscripción
- Recibe **12 emails programados**. Su audiencia es el 55% de 20.000 = 11.000; sin el 10% de control quedan **9.900**.
- Programados: 12 × 9.900 = **118.800 impactos**.
- Triggers (solo a quien hace la acción, por eso van con tasa): Valoración del envío mensual 100% → **33.000**.
- Total del perfil: **151.800**. Sumando todos los perfiles y la bienvenida (3.000): **270.917 impactos**.

## Piezas: 30
Pieza = creatividad distinta, no envío. 23 emails únicos + 2 variantes A/B + 1 común ('Todos') + 4 plantillas de trigger.

## Preguntas trampa
- **¿De dónde sale la base de 20.000?** Es un supuesto; sin el CRM de la marca no hay dato público. Se cambia una celda (o se importan los datos con motor/importar.py) y todo se recalcula.
- **¿Por qué ese reparto entre perfiles?** También es supuesto, y está marcado como tal. El método no cambia con el dato real.
- **¿Por qué 13,5–15,0 emails por persona si hay 12–13 programados?** Se suman los triggers que le tocan de media.
- **¿El control no recibe nada?** No recibe programados; sí triggers y transaccionales, porque responden a lo que hace el cliente. Lo que se mide es el efecto del plan.
- **¿Y la saturación en los picos?** Tope de 2 programados por semana y perfil; en esas semanas el trigger con tope manda un email menos, y la fórmula lo aplica.

## Demostración en directo
Cambia la base de 20.000 a 10.000 en Supuestos: los impactos cambian y las piezas siguen en 30, porque producir no depende de cuánta gente lo recibe.
