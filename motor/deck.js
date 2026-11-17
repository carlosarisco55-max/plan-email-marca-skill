// Deck de 7 diapositivas que responde a las 4 preguntas: marca.json + cifras.json -> .pptx
// Uso: node deck.js <marca.json> <cifras.json> <salida.pptx>
// 1 portada · 2 buyer personas · 3 P1 contenidos · 4 P2 rueda · 5 P2 momentos clave · 6 P3+P4 números · 7 medición
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const [, , fMarca, fCifras, OUT] = process.argv;
if (!OUT) { console.error("Uso: node deck.js <marca.json> <cifras.json> <salida.pptx>"); process.exit(1); }
const m = JSON.parse(fs.readFileSync(fMarca, "utf8")), c = JSON.parse(fs.readFileSync(fCifras, "utf8"));
const d = m.deck, E = d.estilo, PF = m.perfiles, N = PF.length, W = m.semanas, cal = m.calendario, DR = m.datos_reales;

const SW = 13.333, M = 0.6, CW = SW - 2 * M;
const C = { INK: E.oscuro, INK2: E.oscuro2, PAPER: E.papel, RED: E.acento, TEXT: "1A1A1A", MUTED: "5E5B55", ONDARK: "EDEBE6", ONDARK_M: "A8A49C", LINE: "CFCAC0" };
const TIPO = E.tipos, H = E.fuente_titulo, B = E.fuente_texto;
const NUM = ["", "un", "dos", "tres", "cuatro"];
const es = (x, dec = 0) => Number(x).toLocaleString("es-ES", { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: "always" });
const mill = (x) => (x >= 1e6 ? `${es(x / 1e6, 2)} M` : es(x));
const fmtD = (iso, add) => { const t = new Date(iso + "T12:00:00"); t.setDate(t.getDate() + add);
  return `${t.getDate()}-${["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"][t.getMonth()]}-${t.getFullYear()}`; };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = `${m.marca} · Plan de email ${fmtD(m.inicio, 0)} a ${fmtD(m.inicio, 7 * W - 1)}`;
pres.theme = { headFontFace: H, bodyFontFace: B };
const foot = (color) => ({ x: 12.13, y: 6.97, w: 0.6, h: 0.3, fontFace: B, fontSize: 10, color, align: "right" });
const pie = (color) => [{ text: { text: d.pie, options: { x: M, y: 6.97, w: 9, h: 0.3, fontFace: B, fontSize: 9, color, charSpacing: 1 } } }];
pres.defineSlideMaster({ title: "LIGHT", background: { color: C.PAPER }, slideNumber: foot(C.MUTED), objects: pie(C.MUTED) });
pres.defineSlideMaster({ title: "DARK", background: { color: C.INK }, slideNumber: foot(C.ONDARK_M), objects: pie(C.ONDARK_M) });
pres.defineSlideMaster({ title: "COVER", background: { color: C.INK }, objects: [] });

const ALL = [];
const slide = (mm) => { const s = pres.addSlide({ masterName: mm }); ALL.push(s); return s; };
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontFace: B, valign: "top" }, o));
const rect = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 } }, o));
const run = (text, o = {}) => ({ text, options: o });
const kicker = (s, t, dark) => { rect(s, M, 0.5, 0.42, 0.06, C.RED);
  T(s, t.toUpperCase(), { x: M + 0.55, y: 0.38, w: 11, h: 0.3, fontSize: 12, fontFace: H, charSpacing: 3, color: dark ? C.ONDARK_M : C.MUTED }); };
const title = (s, t, dark, size = 34) => T(s, t.toUpperCase(), { x: M, y: 0.75, w: CW, h: 0.9, fontSize: size, fontFace: H, color: dark ? C.ONDARK : C.TEXT });
const mountains = (s, y0, h, c1, c2) => [[0, 2.6, 1.0], [1.9, 3.4, 0.8], [4.6, 2.8, 1.0], [6.6, 3.8, 0.75], [9.6, 2.9, 1.0], [11.2, 2.6, 0.8]].forEach(([x, w, k], i) =>
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x, y: y0 + h * (1 - k), w, h: h * k, fill: { color: i % 2 ? c2 : c1 }, line: { color: i % 2 ? c2 : c1, width: 0 } }));
const cnt = (k, t) => cal.filter((r) => (r[1] === k || r[1] === "Todos") && r[2] === t).length;

// 1 · Portada
let s = slide("COVER");
if (d.portada.montanas !== false) mountains(s, 4.9, 2.6, C.INK2, E.oscuro3 || "1C1C1C");
rect(s, 0, 7.38, SW, 0.12, C.RED);
T(s, (d.portada.antetitulo || m.contexto || "").toUpperCase(), { x: M, y: 0.6, w: CW, h: 0.3, fontSize: 12, fontFace: H, charSpacing: 3, color: C.ONDARK_M });
T(s, d.portada.titulo, { x: M, y: 1.35, w: 11, h: 2.4, fontSize: 66, fontFace: H, color: C.ONDARK, lineSpacingMultiple: 0.9 });
rect(s, M, 3.95, 1.2, 0.08, C.RED);
T(s, d.portada.subtitulo, { x: M, y: 4.2, w: 12.1, h: 0.5, fontSize: 19, color: C.ONDARK });
T(s, d.portada.autores, { x: M, y: 4.85, w: 10, h: 0.4, fontSize: 14, color: C.ONDARK_M });

// 2 · Buyer personas (una tarjeta por perfil)
s = slide("LIGHT");
kicker(s, `Quién · ${NUM[N]} buyer persona${N > 1 ? "s" : ""}`);
title(s, d.personas_titulo);
const gap = 0.2, cw = (CW - gap * (N - 1)) / N, small = N === 4;
PF.forEach((p, i) => {
  const t = p.tarjeta, x = M + i * (cw + gap), y = 1.75, h = 4.75;
  rect(s, x, y, cw, h, "FFFFFF", { line: { color: C.LINE, width: 0.75 } });
  rect(s, x, y, cw, 0.08, C.RED);
  s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.3, w: 0.62, h: 0.62, fill: { color: C.INK }, line: { color: C.INK, width: 0 } });
  T(s, p.k, { x: x + 0.25, y: y + 0.3, w: 0.62, h: 0.62, fontSize: 24, fontFace: H, color: C.ONDARK, align: "center", valign: "middle" });
  T(s, p.nombre.toUpperCase(), { x: x + 1.0, y: y + 0.28, w: cw - 1.15, h: 0.4, fontSize: small ? 14 : 17, fontFace: H, color: C.TEXT, fit: "shrink" });
  T(s, `${t.territorio ? t.territorio + " · " : ""}${es(p.pct * 100)}% de la base (${DR ? "dato real" : "supuesto"})`,
    { x: x + 1.0, y: y + (small ? 0.74 : 0.66), w: cw - 1.1, h: 0.3, fontSize: small ? 10 : 11, color: C.MUTED });
  T(s, [run("BARRERA", { fontFace: H, fontSize: 11, color: C.RED, charSpacing: 2, breakLine: true }), run(t.barrera, { fontSize: small ? 12.5 : 14, bold: true, color: C.TEXT })],
    { x: x + 0.25, y: y + 1.15, w: cw - 0.5, h: 0.95 });
  [["QUIÉN DECIDE", t.decide], ["VA DELANTE", t.atributo], ["NO DECIRLE", t.no]].forEach(([lab, val], j) =>
    T(s, [run(lab, { fontFace: H, fontSize: 11, color: C.MUTED, charSpacing: 2, breakLine: true }), run(val, { fontSize: small ? 11.5 : 12.5, color: C.TEXT })],
      { x: x + 0.25, y: y + 2.15 + j * 0.66, w: cw - 0.5, h: 0.62 }));
  rect(s, x + 0.25, y + 4.15, cw - 0.5, 0.02, C.LINE);
  T(s, `«${t.titular}»`, { x: x + 0.25, y: y + 4.22, w: cw - 0.5, h: 0.5, fontSize: 11.5, italic: true, color: C.MUTED });
});
T(s, d.personas_pie, { x: M, y: 6.6, w: CW, h: 0.3, fontSize: 11, color: C.MUTED });

// 3 · Pregunta 1: contenidos (tabla tipos × perfiles)
s = slide("LIGHT");
kicker(s, "Pregunta 1 · Propuesta de contenidos");
title(s, `Seis tipos de email, usados distinto en cada perfil`);
const hd = (t) => ({ text: t, options: { fontFace: H, fontSize: 13, color: C.ONDARK, fill: { color: C.INK }, charSpacing: 1 } });
const tbl = [[hd(""), hd("TIPO"), hd("QUÉ HACE"), hd("VARIABLE CLAVE"), ...PF.map((p) => hd(p.k))]];
const tdef = Object.fromEntries(m.tipos.map((t) => [t.tipo, t]));
["Oferta", "Producto", "Inspiracional", "Contenido", "Fidelización"].forEach((t) => {
  const n = PF.map((p) => cnt(p.k, t)), max = Math.max(...n);
  tbl.push([{ text: "", options: { fill: { color: TIPO[t] } } }, { text: t.toUpperCase(), options: { fontFace: H, fontSize: 14, color: C.TEXT } },
    { text: tdef[t].deck_que, options: { fontSize: 12.5 } }, { text: tdef[t].deck_variable, options: { fontSize: 12, color: C.MUTED } },
    ...n.map((x) => ({ text: String(x), options: { align: "center", fontFace: H, fontSize: x === max && max > 0 ? 20 : 16, color: x === max && max > 0 ? C.RED : C.TEXT } }))]);
});
tbl.push([{ text: "", options: { fill: { color: "8A867E" } } }, { text: "TRIGGER", options: { fontFace: H, fontSize: 14, color: C.TEXT } },
  { text: tdef.Trigger.deck_que, options: { fontSize: 12.5 } }, { text: tdef.Trigger.deck_variable, options: { fontSize: 12, color: C.MUTED } },
  { text: "siempre activos", options: { colspan: N, align: "center", fontSize: 12, color: C.MUTED } }]);
const pc = 0.85, vw = CW - 0.18 - 1.9 - 3.3 - pc * N;
s.addTable(tbl, { x: M, y: 1.75, w: CW, colW: [0.18, 1.9, 3.3, vw, ...PF.map(() => pc)], fontFace: B, color: C.TEXT, valign: "middle",
  border: { type: "solid", pt: 0.75, color: C.LINE }, fill: { color: "FFFFFF" }, rowH: 0.56, margin: [0.05, 0.12, 0.05, 0.12] });
T(s, [run("Cómo se lee: ", { bold: true, color: C.TEXT }), run(d.tipos_lectura, { color: C.MUTED })], { x: M, y: 6.05, w: CW, h: 0.6, fontSize: 13 });

// 4 · Pregunta 2: la rueda (un anillo por perfil, una porción por semana)
s = slide("DARK");
kicker(s, `Pregunta 2 · Programa de ${W} semanas`, true);
title(s, `La rueda de contenidos: ${c.emails_programados} emails, semana a semana`, true);
const cx = 3.85, cy = 4.33, R0 = 0.88, R1 = 2.23, RG = 0.05, th = (R1 - R0 - RG * (N - 1)) / N, step = 360 / W;
const RING = Object.fromEntries(PF.map((p, i) => [p.k, [R0 + i * (th + RG), R0 + i * (th + RG) + th]]));
const arc = (r0, r1, a0, a1, color) => s.addShape(pres.shapes.BLOCK_ARC, { x: cx - r1, y: cy - r1, w: 2 * r1, h: 2 * r1,
  angleRange: [a0, a1], arcThicknessRatio: (r1 - r0) / r1, fill: { color }, line: { color, width: 0 } });
const pad = Math.min(0.7, step * 0.03);
for (let w = 1; w <= W; w++) {
  const a0 = (270 + step * (w - 1) + pad) % 360, a1 = (270 + step * w - pad) % 360;
  for (const p of PF) {
    const ts = cal.filter((r) => r[0] === w && (r[1] === p.k || r[1] === "Todos")).map((r) => r[2]);
    const [r0, r1] = RING[p.k];
    if (ts.length === 1) arc(r0, r1, a0, a1, TIPO[ts[0]]);
    else if (ts.length > 1) { const rm = (r0 + r1) / 2; arc(r0, rm - 0.015, a0, a1, TIPO[ts[0]]); arc(rm + 0.015, r1, a0, a1, TIPO[ts[1]]); }
  }
  const am = ((-90 + step * (w - 0.5)) * Math.PI) / 180;
  T(s, String(w), { x: cx + 2.5 * Math.cos(am) - 0.2, y: cy + 2.5 * Math.sin(am) - 0.15, w: 0.4, h: 0.3, fontSize: 12, fontFace: H, color: C.ONDARK_M, align: "center", valign: "middle" });
}
s.addShape(pres.shapes.OVAL, { x: cx - 2.31, y: cy - 2.31, w: 4.62, h: 4.62, fill: { type: "none" }, line: { color: "6F6B64", width: 1.5, dashType: "dash" } });
c.meses.forEach(([, w]) => { const a = ((-90 + step * (w - 1)) * Math.PI) / 180, x1 = cx + 0.83 * Math.cos(a), y1 = cy + 0.83 * Math.sin(a), x2 = cx + 2.33 * Math.cos(a), y2 = cy + 2.33 * Math.sin(a);
  s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001), flipH: (x2 - x1) * (y2 - y1) < 0, line: { color: C.ONDARK, width: 2.5 } }); });
T(s, [run(String(W), { fontFace: H, fontSize: 34, color: C.ONDARK, breakLine: true }), run("SEMANAS", { fontFace: H, fontSize: 11, color: C.ONDARK_M, charSpacing: 2 })],
  { x: cx - 0.8, y: cy - 0.45, w: 1.6, h: 0.9, align: "center", valign: "middle" });
const LX = 7.55;
T(s, N > 1 ? "ANILLOS (DE DENTRO A FUERA)" : "ANILLO", { x: LX, y: 1.8, w: 5, h: 0.3, fontSize: 12, fontFace: H, color: C.ONDARK_M, charSpacing: 2 });
T(s, PF.flatMap((p, i) => [run(p.k + " ", { fontFace: H, color: C.ONDARK }), run(p.corto + (i < N - 1 ? "  ·  " : ""), { color: C.ONDARK_M })]), { x: LX, y: 2.12, w: 5.2, h: 0.35, fontSize: 14 });
T(s, "Línea discontinua: triggers siempre activos · celda partida: 2 emails esa semana", { x: LX, y: 2.48, w: 5.2, h: 0.3, fontSize: 11, color: C.ONDARK_M });
Object.entries(TIPO).forEach(([t, col], i) => { const x = LX + (i % 3) * 1.75, y = 2.95 + Math.floor(i / 3) * 0.42;
  rect(s, x, y + 0.05, 0.22, 0.22, col); T(s, t, { x: x + 0.32, y, w: 1.4, h: 0.32, fontSize: 13, color: C.ONDARK, valign: "middle" }); });
T(s, "MOMENTOS DEL CALENDARIO", { x: LX, y: 3.95, w: 5, h: 0.3, fontSize: 12, fontFace: H, color: C.ONDARK_M, charSpacing: 2 });
const mh = Math.min(0.62, 2.55 / c.meses.length);
c.meses.forEach(([lab], i) => T(s, [run(lab, { fontFace: H, fontSize: 14, color: C.RED, breakLine: true }), run(d.meses[i], { fontSize: 12, color: C.ONDARK })],
  { x: LX, y: 4.3 + i * mh, w: 5.2, h: mh }));

// 5 · Momentos clave (los que deciden el plan)
s = slide("LIGHT");
kicker(s, `Pregunta 2 · Los ${NUM[d.momentos.length] === "un" ? "" : NUM[d.momentos.length] + " "}momentos que deciden el plan`);
title(s, d.momentos_titulo);
const nm = d.momentos.length, mw = (CW - gap * (nm - 1)) / nm;
d.momentos.forEach((mo, i) => {
  const x = M + i * (mw + gap), y = 1.8;
  rect(s, x, y, mw, 0.55, C.INK);
  T(s, mo.tag, { x: x + 0.2, y, w: mw - 0.4, h: 0.55, fontSize: 13, fontFace: H, color: C.ONDARK, charSpacing: 2, valign: "middle" });
  rect(s, x, y + 0.55, mw, 4.1, "FFFFFF", { line: { color: C.LINE, width: 0.75 } });
  T(s, mo.titulo, { x: x + 0.25, y: y + 0.8, w: mw - 0.5, h: 0.75, fontSize: 19, fontFace: H, color: C.RED });
  T(s, mo.puntos.map((b, j) => run(b, { bullet: { code: "25A0" }, breakLine: j < mo.puntos.length - 1 })),
    { x: x + 0.25, y: y + 1.6, w: mw - 0.5, h: 2.9, fontSize: 13.5, color: C.TEXT, paraSpaceAfter: 9 });
});
T(s, d.momentos_pie, { x: M, y: 6.6, w: CW, h: 0.3, fontSize: 11, color: C.MUTED });

// 6 · Preguntas 3 y 4: números (todo sale de cifras.json)
s = slide("DARK");
kicker(s, "Preguntas 3 y 4 · Impactos y piezas", true);
title(s, "Lo que recibe la audiencia y lo que hay que producir", true);
const big = (x, n, lab, sub) => { T(s, n, { x, y: 1.85, w: 3.9, h: 1.1, fontSize: 64, fontFace: H, color: C.ONDARK, fit: "shrink" });
  rect(s, x, 3.0, 0.7, 0.06, C.RED);
  T(s, [run(lab, { fontFace: H, fontSize: 16, color: C.ONDARK, breakLine: true, charSpacing: 1 }), run(sub, { fontSize: 12, color: C.ONDARK_M })], { x, y: 3.15, w: 3.85, h: 0.9 }); };
const I = c.impactos, Z = c.piezas, pp = c.perfiles.map((p) => p.emails_por_persona);
const lo = es(Math.min(...pp), 1), hi = es(Math.max(...pp), 1);
big(M, mill(I.total), `IMPACTOS EN ${W} SEMANAS`, `${mill(I.programados)} programados + ${es(I.triggers)} por triggers`);
big(M + 4.1, lo === hi ? lo : `${lo}–${hi}`, "EMAILS POR PERSONA", d.numeros_lectura);
big(M + 8.2, String(Z.total), "PIEZAS A PRODUCIR", `${Z.unicos} emails únicos + ${Z.variantes} variantes A/B + ${Z.comunes} común${Z.comunes === 1 ? "" : "es"} + ${Z.plantillas} plantillas de trigger`);
const nh = (t, i) => ({ text: t, options: { fontFace: H, fontSize: 13, color: C.ONDARK, fill: { color: "2E2E2E" }, align: i ? "right" : "left" } });
const nr = (cells, b) => cells.map((x, i) => ({ text: x, options: { align: i ? "right" : "left", color: C.ONDARK, fontFace: b ? H : B, fontSize: b ? 14 : 13, fill: { color: b ? "2E2E2E" : C.INK2 } } }));
const rows = [["PERFIL", "AUDIENCIA QUE RECIBE", "EMAILS PROGRAMADOS", "IMPACTOS TOTALES"].map(nh)];
c.perfiles.forEach((p) => rows.push(nr([p.k === "S" ? `Sin perfil (programa de ${p.programa_de})` : `${p.k} · ${p.nombre}`, es(p.recibe), String(p.programados), es(p.imp_total)])));
rows.push(nr([`Bienvenida (${es(c.bienvenida.altas)} altas nuevas)`, es(c.bienvenida.altas), `${c.bienvenida.emails} (trigger)`, es(c.bienvenida.impactos)]));
rows.push(nr(["TOTAL", "", "", es(I.total)], true));
s.addTable(rows, { x: M, y: 4.2, w: CW, colW: [4.5, 2.6, 2.6, CW - 9.7], border: { type: "solid", pt: 0.5, color: "3A3A3A" },
  rowH: Math.min(0.33, 2.3 / rows.length), margin: [0.03, 0.12, 0.03, 0.12], valign: "middle", fontSize: N === 4 ? 12 : 13 });
const S_ = m.supuestos;
T(s, `${DR ? `Datos reales de ${DR.fuente}` : "Supuestos editables en el Excel"}: base de ${es(S_.base)} suscriptores, reparto ${[...PF.map((p) => p.pct), m.sin_perfil.pct].map((x) => es(x * 100)).join("/")} %, control del ${es(S_.control * 100)} % que no recibe los programados.`,
  { x: M, y: 6.6, w: CW, h: 0.3, fontSize: 11, color: C.ONDARK_M });

// 7 · Medición
s = slide("LIGHT");
kicker(s, "Cómo sabremos si funciona");
title(s, d.medicion_titulo);
[{ t: "GRUPO DE CONTROL", n: `${es(S_.control * 100)} %`, b: "Se aparta antes de segmentar y no recibe ningún email programado. Lo que vende el plan es la diferencia con este grupo, no las ventas brutas de los picos." },
 { t: `${c.tests.length} TEST${c.tests.length === 1 ? "" : "S"} A/B`, n: "50/50", b: d.tests_texto },
 { t: "REVISIÓN", n: d.revision.fecha, b: d.revision.texto }].forEach((k, i) => { const x = M + i * 4.1;
  T(s, k.n, { x, y: 1.85, w: 3.9, h: 0.95, fontSize: 56, fontFace: H, color: C.RED });
  T(s, k.t, { x, y: 2.85, w: 3.9, h: 0.4, fontSize: 18, fontFace: H, color: C.TEXT, charSpacing: 1 });
  T(s, k.b, { x, y: 3.3, w: 3.8, h: 1.9, fontSize: 13.5, color: C.MUTED }); });
rect(s, M, 5.45, CW, 1.0, C.INK);
rect(s, M, 5.45, 0.12, 1.0, C.RED);
T(s, d.cierre, { x: M + 0.4, y: 5.45, w: CW - 0.6, h: 1.0, fontSize: 20, fontFace: H, color: C.ONDARK, valign: "middle" });

const notas = m.notas || [];
if (notas.length && notas.length !== ALL.length) throw new Error(`notas ${notas.length} != diapositivas ${ALL.length}`);
notas.forEach((n, i) => n && ALL[i].addNotes(n));
pres.writeFile({ fileName: OUT }).then((f) => console.log("OK", f));
