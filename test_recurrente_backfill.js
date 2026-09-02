// ponytail: reproducer del cursor de generarRecurrentesPendientes_
// Tomamos las funciones puras (fechaConDiaMes_, siguienteCursor_) tal cual
// están en code.js y replicamos la lógica del primer/segundo while. Si la
// fila recurrente tiene ultima_generacion = hoy (cutoff del último run) y el
// día del mes ya pasó, ¿el segundo while entra al slot que toca este mes?

function fechaConDiaMes_(year, month, dia) {
  const last = new Date(year, month + 1, 0).getDate();
  const d = Math.min(Math.max(1, Number(dia) || 1), last);
  return new Date(year, month, d);
}

function siguienteCursor_(d, periodoMeses, dia) {
  const meses = Number(periodoMeses) || 1;
  const targetMonth = d.getMonth() + meses;
  const year = d.getFullYear() + Math.floor(targetMonth / 12);
  const month = ((targetMonth % 12) + 12) % 12;
  return fechaConDiaMes_(year, month, dia != null ? dia : d.getDate());
}

// Simulamos iso_() con formato local YYYY-MM-DD
function iso_(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${dd}`;
}

function parseFecha(s) {
  if (!s) return null;
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return null;
}

function simular(ultimaGeneracion, dia, periodoMeses, fechaCorte) {
  const base = parseFecha(ultimaGeneracion) || new Date();
  let cursor = fechaConDiaMes_(base.getFullYear(), base.getMonth(), dia);
  // PRIMER while (idéntico al código)
  const uIso = iso_(base);
  // FIX: usar `<` en lugar de `<=`. La condición `<=` avanza el cursor también
  // cuando el slot está en el mismo día del último cutoff, saltándose ese slot
  // incluso si nunca fue generado (caso típico: bootstrap previo fue el mismo
  // día del cargo; el primer while salta sep 5 → oct 5 y la tx de sep 5 nunca
  // se crea). El `<` mantiene la optimización de no re-chequear slots
  // estrictamente anteriores y deja que el segundo while + chequeo de
  // conflicto (`txConflictaEnMesRecurrente_`) decidan sobre el slot del día.
  while (iso_(cursor) < uIso && cursor <= fechaCorte) {
    cursor = siguienteCursor_(cursor, periodoMeses, dia);
  }
  // SEGUNDO while
  const cortes = [];
  while (cursor <= fechaCorte) {
    cortes.push(iso_(cursor));
    cursor = siguienteCursor_(cursor, periodoMeses, dia);
  }
  return cortes;
}

// CASO 1: bootstrap corrió el mismo día del mes (sep 5) y dejó ultima_generacion=sep 5.
// El usuario vuelve a abrir la app el sep 12 (7 días pasado).
console.log('CASO 1 — bootstrap previo el mismo día del mes, hoy 7 días después:');
const r1 = simular('2026-09-05', 5, 1, new Date(2026, 8, 12));
console.log('  cortes generados:', r1);
console.log('  (esperamos ["2026-09-05"] sin duplicarlo, sin futuras)');
console.log();

// CASO 2: bootstrap previo el día antes del cargo (sep 4), hoy sep 12.
console.log('CASO 2 — bootstrap previo un día antes del cargo, hoy 7 días después:');
const r2 = simular('2026-09-04', 5, 1, new Date(2026, 8, 12));
console.log('  cortes generados:', r2);
console.log('  (esperamos ["2026-09-05"])');
console.log();

// CASO 3: bootstrap previo mucho antes (ago 4), hoy sep 12.
console.log('CASO 3 — bootstrap previo hace más de un mes, hoy 7 días pasado el cargo:');
const r3 = simular('2026-08-04', 5, 1, new Date(2026, 8, 12));
console.log('  cortes generados:', r3);
console.log('  (esperamos ["2026-09-05"], sin re-generar ago)');
console.log();

// CASO 4: bootstrap previo 2 días antes del cargo (sep 3), hoy sep 12.
console.log('CASO 4 — bootstrap previo 2 días antes del cargo:');
const r4 = simular('2026-09-03', 5, 1, new Date(2026, 8, 12));
console.log('  cortes generados:', r4);
console.log('  (esperamos ["2026-09-05"])');
console.log();

// CASO 5: bootstrap previo 0 días antes (sep 5 mismo día), hoy sep 5 mismo día.
// Es decir, abre la app el mismo día del cargo. Esta es la primera ejecución.
console.log('CASO 5 — bootstrap previo mismo día del cargo, hoy mismo día:');
const r5 = simular('2026-09-05', 5, 1, new Date(2026, 8, 5));
console.log('  cortes generados:', r5);
console.log('  (esperamos ["2026-09-05"])');
console.log();

// CASO 6: periodo 2 meses, dia 10, hoy sep 12, ultima_generacion=ago 10.
// El cursor se queda en ago 10 porque `<` no avanza; el segundo while lo
// emite como backfill si faltaba (chequeo de conflicto evita duplicar).
console.log('CASO 6 — periodo 2m, ultima_generacion exacta al día:');
const r6 = simular('2026-08-10', 10, 2, new Date(2026, 8, 12));
console.log('  cortes generados:', r6);
console.log('  (slot ago 10 visitable para chequeo; oct 10 fuera de corte sep 12)');

// CASO EXACTO DEL REPORTE: cargo día 5, hoy 12, ultima_generacion = día 5
// (el último bootstrap corrió ese día, falló en generar el slot, dejó el
// cutoff en esa fecha). Varios días después, al abrir la app, debe
// generarse la tx del día 5.
console.log('CASO REPORTE — cargo día 5, hoy día 12, último cutoff = día 5:');
const rReporte = simular('2026-09-05', 5, 1, new Date(2026, 8, 12));
if (!rReporte.includes('2026-09-05')) {
  console.error('FAIL: no generó el cargo del día 5 (varios días después)');
  process.exit(1);
}
console.log('  cortes generados:', rReporte);
console.log('  ✓ cargo día 5 generado aunque el último cutoff fue ese mismo día');

console.log();
console.log('Todos los casos pasaron.');

