# Cómo defender el Excel · The North Face

## La idea en una frase
Todo sale de dos sitios: la hoja **Supuestos** (celdas amarillas) y la hoja **Calendario** (una fila por email). El resto son fórmulas.

## Ejemplo trabajado: A · Alpinista técnico
- Recibe **16 emails programados**. Su audiencia es el 12% de 100.000 = 12.000; sin el 10% de control quedan **10.800**.
- Programados: 16 × 10.800 = **172.800 impactos**.
- Triggers (solo a quien hace la acción, por eso van con tasa): Carrito abandonado 8%, Navegación sin compra 15%, Post-compra: talla, cuidado e invitación XPLR Pass 12%, Recompra de temporada / complemento de cesta 10% → **6.360**.
- Total del perfil: **179.160**. Sumando todos los perfiles y la bienvenida (12.000): **1.661.240 impactos**.

## Piezas: 67
Pieza = creatividad distinta, no envío. 49 emails únicos + 4 variantes A/B + 1 común ('Todos') + 13 plantillas de trigger.

## Preguntas trampa
- **¿De dónde sale la base de 100.000?** Es un supuesto; sin el CRM de la marca no hay dato público. Se cambia una celda (o se importan los datos con motor/importar.py) y todo se recalcula.
- **¿Por qué ese reparto entre perfiles?** También es supuesto, y está marcado como tal. El método no cambia con el dato real.
- **¿Por qué 16,5–18,5 emails por persona si hay 16–18 programados?** Se suman los triggers que le tocan de media.
- **¿El control no recibe nada?** No recibe programados; sí triggers y transaccionales, porque responden a lo que hace el cliente. Lo que se mide es el efecto del plan.
- **¿Y la saturación en los picos?** Tope de 2 programados por semana y perfil; en esas semanas el trigger con tope manda un email menos, y la fórmula lo aplica.

## Demostración en directo
Cambia la base de 100.000 a 50.000 en Supuestos: los impactos cambian y las piezas siguen en 67, porque producir no depende de cuánta gente lo recibe.
