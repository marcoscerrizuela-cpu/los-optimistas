import React, { useState, useEffect, useCallback, useMemo } from "react";
import {
  Flag, Trophy, PlusCircle, Trash2, Settings, ChevronRight, Loader2,
  Camera, X, ImageOff, Lock, UserPlus, Check, Ban, KeyRound, Plus, RefreshCw, Pencil, Wallet,
} from "lucide-react";

const FONT_LINK_ID = "golf-league-fonts";

// Se actualiza en cada entrega para que sea visible en pantalla (carátula) qué versión
// de código está corriendo un link publicado puntual, sin depender del nombre del archivo.
const APP_VERSION = "v25";


function useFonts() {
  useEffect(() => {
    if (document.getElementById(FONT_LINK_ID)) return;
    const link = document.createElement("link");
    link.id = FONT_LINK_ID;
    link.rel = "stylesheet";
    link.href =
      "https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=IBM+Plex+Mono:wght@500;600&family=Inter:wght@400;500;600&display=swap";
    document.head.appendChild(link);
  }, []);
}

const COLORS = {
  green900: "#16302A",
  green700: "#234F42",
  green500: "#2F6B54",
  paper: "#F7F5EF",
  paperDim: "#EDEAE1",
  ink: "#1B2420",
  brass: "#A9824F",
  brassLight: "#C8A76B",
  danger: "#8C3B2E",
};

// La cantidad mínima de tarjetas y la modalidad (suma / mejor) ahora se configuran
// por cancha y por torneo — ver courseConfig en cada torneo.

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

const MONTHS_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];

// Convierte "2026-07-05" (formato interno, ISO) a "05-jul-2026" (formato de visualización).
// Solo afecta cómo se muestran fechas ya guardadas; los selectores nativos de fecha
// (<input type="date">) siguen mostrando el formato propio del navegador/dispositivo.
function formatDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  const monthIdx = parseInt(m, 10) - 1;
  if (isNaN(monthIdx) || monthIdx < 0 || monthIdx > 11) return iso;
  return `${d}-${MONTHS_ES[monthIdx]}-${y}`;
}

function uid() {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

function csvEscape(value) {
  const s = value === null || value === undefined ? "" : String(value);
  if (s.includes(",") || s.includes('"') || s.includes("\n")) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function downloadCSV(filename, rows) {
  const csv = rows.map((row) => row.map(csvEscape).join(",")).join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" }); // BOM para que Excel abra bien los acentos
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}


function safeFilename(name) {
  return name.trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

async function sha256(text) {
  const enc = new TextEncoder().encode(text);
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

// ---- Capa de datos: habla con /api/sheets (función intermedia de Vercel), que a su
// vez reenvía a Google Apps Script desde el servidor. Se llama directo desde el
// navegador a Apps Script fallaba por un problema de redirección/CORS — ya lo vivimos
// en la migración de Control DJ, así que acá lo evitamos desde el principio.
const API_TOKEN = import.meta.env.VITE_API_TOKEN;

async function apiGet(key) {
  try {
    const res = await fetch(`/api/sheets?action=get&key=${encodeURIComponent(key)}&token=${encodeURIComponent(API_TOKEN)}`);
    const data = await res.json();
    return data && data.ok ? data.value : null;
  } catch (e) {
    return null;
  }
}

async function apiPost(body) {
  try {
    const res = await fetch("/api/sheets", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...body, token: API_TOKEN }),
    });
    return await res.json();
  } catch (e) {
    return { ok: false, msg: String(e) };
  }
}

async function storageGet(key) {
  return apiGet(key);
}

async function storageSet(key, value) {
  const res = await apiPost({ action: "set", key, value });
  return !!(res && res.ok);
}

// Google Sheets + Apps Script no tiene el mismo riesgo de "éxito falso" que
// window.storage (la escritura es síncrona del lado del servidor), pero se
// mantiene la relectura de verificación por consistencia con el resto del código.
async function storageSetVerified(key, value) {
  const ok = await storageSet(key, value);
  if (!ok) return false;
  const check = await storageGet(key);
  return JSON.stringify(check) === JSON.stringify(value);
}

async function storageDelete(key) {
  // las claves de foto (photo:ID) necesitan borrar también el archivo real en
  // Drive, no solo la fila de la hoja — si no, quedan archivos huérfanos.
  if (key.startsWith("photo:")) {
    const roundId = key.slice("photo:".length);
    const res = await apiPost({ action: "deletePhoto", roundId });
    return !!(res && res.ok);
  }
  const res = await apiPost({ action: "delete", key });
  return !!(res && res.ok);
}

async function safeSetPhoto(id, dataUrl, meta) {
  const res = await apiPost({ action: "uploadPhoto", roundId: id, dataUrl, ...meta });
  if (!res || !res.ok) {
    return { ok: false, reason: (res && res.msg) || "fallo desconocido al subir a Drive", size: dataUrl.length };
  }
  return { ok: true };
}

function compressImage(file, maxDim = 1400, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          const scale = maxDim / Math.max(width, height);
          width = Math.round(width * scale);
          height = Math.round(height * scale);
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = () => reject(new Error("No se pudo leer la imagen"));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
    reader.readAsDataURL(file);
  });
}

// Red de contención: si algo se rompe en el render, muestra un mensaje en vez de
// dejar la pantalla completamente en blanco sin ninguna pista de qué pasó.
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, message: "" };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, message: (error && error.message) || String(error) };
  }
  render() {
    if (this.state.hasError) {
      return (
        <div
          style={{
            minHeight: "100vh",
            background: "#16302A",
            color: "#F7F5EF",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
            padding: 24,
            textAlign: "center",
            fontFamily: "sans-serif",
          }}
        >
          <div style={{ fontSize: 40, marginBottom: 12 }}>⛳💥</div>
          <h1 style={{ fontFamily: "serif", fontSize: 22, margin: "0 0 10px" }}>Algo se rompió</h1>
          <p style={{ fontSize: 14, opacity: 0.8, maxWidth: 400, marginBottom: 6 }}>
            La app tuvo un error y no puede seguir. Recargá la página. Si persiste, avisale al administrador
            con este mensaje:
          </p>
          <code style={{ background: "rgba(0,0,0,0.3)", padding: "8px 12px", borderRadius: 8, fontSize: 12, maxWidth: 400, wordBreak: "break-word" }}>
            {this.state.message}
          </code>
        </div>
      );
    }
    return this.props.children;
  }
}

function GolfLeagueInner() {
  useFonts();

  const [users, setUsers] = useState([]); // [{name,pinHash,status}]
  const [config, setConfig] = useState({});
  const [courses, setCourses] = useState([]);
  const [tournaments, setTournaments] = useState([]); // [{name, enabled}]
  const [rounds, setRounds] = useState([]);
  const [tesoreria, setTesoreria] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("posiciones");
  const [tournament, setTournament] = useState("");
  const [showAdmin, setShowAdmin] = useState(false);
  const [showTreasury, setShowTreasury] = useState(false);
  const [treasuryAuthed, setTreasuryAuthed] = useState(false);
  const [adminAuthed, setAdminAuthed] = useState(false);
  const [showRegister, setShowRegister] = useState(false);
  const [toast, setToast] = useState(null);
  const [viewingPhoto, setViewingPhoto] = useState(null);

  const showToast = useCallback((msg, duration = 2600) => {
    setToast(msg);
    setTimeout(() => setToast(null), duration);
  }, []);

  const loadAllData = useCallback(async () => {
    const [u, cfg, c, t, r, legacyPlayers, tes] = await Promise.all([
      storageGet("users"),
      storageGet("config"),
      storageGet("courses"),
      storageGet("tournaments"),
      storageGet("rounds"),
      storageGet("players"),
      storageGet("tesoreria"),
    ]);

    let finalUsers = u;
    if (!finalUsers && legacyPlayers) {
      finalUsers = legacyPlayers.map((name) => ({
        name,
        pinHash: null,
        status: "aprobado",
      }));
      await storageSet("users", finalUsers);
    }

    setUsers(finalUsers || []);
    setConfig(cfg || {});
    setCourses(c || []);
    setTournaments(t || []);
    setTesoreria(tes || []);

    // migración: las rondas viejas usaban "season" (semestre fijo) en vez de "tournament"
    const migrated = (r || []).map((rd) => {
      if (rd.tournament) return rd;
      if (rd.season) return { ...rd, tournament: rd.season };
      return rd;
    });
    setRounds(migrated);

    return t || [];
  }, []);

  useEffect(() => {
    (async () => {
      const t = await loadAllData();
      const firstEnabled = t.find((tt) => tt.enabled);
      if (firstEnabled) setTournament(firstEnabled.name);
      setLoading(false);
    })();
  }, [loadAllData]);

  async function refreshData() {
    await loadAllData();
    showToast("Datos actualizados desde el almacenamiento");
  }

  const enabledTournaments = useMemo(() => tournaments.filter((t) => t.enabled), [tournaments]);

  const activeTournamentObj = useMemo(
    () => tournaments.find((t) => t.name === tournament) || null,
    [tournaments, tournament]
  );

  useEffect(() => {
    if (loading) return;
    const stillValid = enabledTournaments.some((t) => t.name === tournament);
    if (!stillValid) {
      setTournament(enabledTournaments.length > 0 ? enabledTournaments[0].name : "");
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [enabledTournaments, loading]);

  const approvedUsers = useMemo(() => users.filter((u) => u.status === "aprobado"), [users]);
  const pendingUsers = useMemo(() => users.filter((u) => u.status === "pendiente"), [users]);
  const approvedNames = useMemo(() => approvedUsers.map((u) => u.name), [approvedUsers]);

  const pendingParticipantsCount = useMemo(
    () => tournaments.reduce((acc, t) => acc + (t.participants || []).filter((p) => p.status === "pendiente").length, 0),
    [tournaments]
  );
  const totalPendingCount = pendingUsers.length + pendingParticipantsCount;

  const acceptedParticipantUsers = useMemo(() => {
    if (!activeTournamentObj) return [];
    const acceptedNames = (activeTournamentObj.participants || [])
      .filter((p) => p.status === "aceptado")
      .map((p) => p.name);
    return approvedUsers.filter((u) => acceptedNames.includes(u.name));
  }, [activeTournamentObj, approvedUsers]);

  const roundsInTournament = useMemo(() => rounds.filter((r) => r.tournament === tournament), [rounds, tournament]);

  const tournamentCourseConfig = useMemo(
    () => (activeTournamentObj && activeTournamentObj.courseConfig) || [],
    [activeTournamentObj]
  );

  const courseNames = useMemo(() => tournamentCourseConfig.map((cc) => cc.course), [tournamentCourseConfig]);

  const courseRuleByName = useMemo(() => {
    const map = {};
    for (const cc of tournamentCourseConfig) map[cc.course] = cc;
    return map;
  }, [tournamentCourseConfig]);

  const parByCourse = useMemo(() => {
    const map = {};
    for (const c of courses) map[c.name] = c.par;
    return map;
  }, [courses]);

  const roundsCountByCourse = useMemo(() => {
    const map = {};
    for (const r of rounds) map[r.course] = (map[r.course] || 0) + 1;
    return map;
  }, [rounds]);

  const countingRoundIds = useMemo(() => {
    const counting = new Set();
    for (const course of courseNames) {
      const rule = courseRuleByName[course];
      if (!rule) continue;
      const takeCount = rule.mode === "mejor" ? 1 : rule.minRequired;
      const byPlayerRounds = {};
      for (const r of roundsInTournament.filter((r) => r.course === course)) {
        if (!byPlayerRounds[r.player]) byPlayerRounds[r.player] = [];
        byPlayerRounds[r.player].push(r);
      }
      for (const playerRounds of Object.values(byPlayerRounds)) {
        const sorted = [...playerRounds].sort((a, b) => a.score - b.score || a.id.localeCompare(b.id));
        for (const r of sorted.slice(0, takeCount)) counting.add(r.id);
      }
    }
    return counting;
  }, [roundsInTournament, courseNames, courseRuleByName]);

  const acceptedParticipantNames = useMemo(
    () => acceptedParticipantUsers.map((u) => u.name),
    [acceptedParticipantUsers]
  );

  const verifiedStatusByPlayer = useMemo(() => {
    const map = {};
    for (const player of acceptedParticipantNames) {
      const theirCounting = roundsInTournament.filter((r) => r.player === player && countingRoundIds.has(r.id));
      map[player] = theirCounting.length > 0 && theirCounting.every((r) => r.verified);
    }
    return map;
  }, [acceptedParticipantNames, roundsInTournament, countingRoundIds]);

  const standings = useMemo(() => {
    const byPlayer = {};
    // arrancan TODOS los aceptados al torneo, aunque todavía no hayan cargado ninguna tarjeta
    for (const name of acceptedParticipantNames) byPlayer[name] = { player: name, courseData: {}, total: 0, roundsCount: 0 };

    for (const course of courseNames) {
      const rule = courseRuleByName[course];
      if (!rule) continue;
      const takeCount = rule.mode === "mejor" ? 1 : rule.minRequired;
      const scoresByPlayer = {};
      for (const r of roundsInTournament.filter((r) => r.course === course)) {
        if (!scoresByPlayer[r.player]) scoresByPlayer[r.player] = [];
        scoresByPlayer[r.player].push(r.score);
      }
      for (const [player, scores] of Object.entries(scoresByPlayer)) {
        if (!byPlayer[player]) byPlayer[player] = { player, courseData: {}, total: 0, roundsCount: 0 };
        const sorted = [...scores].sort((a, b) => a - b);
        const bestN = sorted.slice(0, takeCount);
        const subtotal = bestN.reduce((a, b) => a + b, 0);
        byPlayer[player].courseData[course] = {
          subtotal,
          count: sorted.length,
          used: bestN.length,
          complete: sorted.length >= rule.minRequired,
          minRequired: rule.minRequired,
          mode: rule.mode,
        };
        byPlayer[player].total += subtotal;
        byPlayer[player].roundsCount += sorted.length;
      }
    }

    return Object.values(byPlayer)
      .map((p) => ({ ...p, eligible: p.roundsCount > 0 && Object.values(p.courseData).every((d) => d.complete) }))
      .sort((a, b) => {
        // quien no presentó ninguna tarjeta nunca puede figurar antes ni en mejor
        // posición que alguien que sí jugó, sea cual sea el total (que empieza en 0)
        const aPlayed = a.roundsCount > 0;
        const bPlayed = b.roundsCount > 0;
        if (aPlayed !== bPlayed) return aPlayed ? -1 : 1;
        if (aPlayed) return a.total - b.total;
        return a.player.localeCompare(b.player);
      });
  }, [acceptedParticipantNames, roundsInTournament, courseNames, courseRuleByName]);

  // ---- Usuarios ----
  async function registerUser(name, pin) {
    const trimmed = name.trim();
    const freshUsers = (await storageGet("users")) || users;
    if (!trimmed || freshUsers.some((u) => u.name === trimmed)) return { ok: false, msg: "Ese nombre ya existe." };
    const pinHash = await sha256(pin);
    const next = [...freshUsers, { name: trimmed, pinHash, status: "pendiente" }];
    setUsers(next);
    const saved = await storageSetVerified("users", next);
    if (!saved) return { ok: false, msg: "No se pudo guardar el registro. Revisá tu conexión y probá de nuevo." };
    return { ok: true };
  }

  async function addUserDirect(name, pin) {
    const trimmed = name.trim();
    const freshUsers = (await storageGet("users")) || users;
    if (!trimmed || freshUsers.some((u) => u.name === trimmed)) return { ok: false, msg: "Ese nombre ya existe." };
    if (!/^\d{4}$/.test(pin)) return { ok: false, msg: "El PIN debe tener 4 dígitos." };
    const pinHash = await sha256(pin);
    const next = [...freshUsers, { name: trimmed, pinHash, status: "aprobado" }];
    setUsers(next);
    const saved = await storageSetVerified("users", next);
    if (!saved) return { ok: false, msg: "No se pudo guardar. Probá de nuevo." };
    return { ok: true };
  }

  async function approveUser(name) {
    const freshUsers = (await storageGet("users")) || users;
    const next = freshUsers.map((u) => (u.name === name ? { ...u, status: "aprobado" } : u));
    setUsers(next);
    await storageSet("users", next);
  }

  async function rejectUser(name) {
    const freshUsers = (await storageGet("users")) || users;
    const next = freshUsers.filter((u) => u.name !== name);
    setUsers(next);
    await storageSet("users", next);
  }

  async function removeUser(name) {
    const freshUsers = (await storageGet("users")) || users;
    const next = freshUsers.filter((u) => u.name !== name);
    setUsers(next);
    await storageSet("users", next);
  }

  async function setUserPin(name, pin) {
    if (!/^\d{4}$/.test(pin)) return { ok: false, msg: "El PIN debe tener 4 dígitos." };
    const pinHash = await sha256(pin);
    const freshUsers = (await storageGet("users")) || users;
    const next = freshUsers.map((u) => (u.name === name ? { ...u, pinHash } : u));
    setUsers(next);
    await storageSet("users", next);
    showToast(`PIN de ${name} actualizado`);
    return { ok: true };
  }

  async function verifyPin(name, pin) {
    const freshUsers = (await storageGet("users")) || users;
    const user = freshUsers.find((u) => u.name === name);
    if (!user || !user.pinHash) return false;
    const hash = await sha256(pin);
    return hash === user.pinHash;
  }

  // ---- Admin ----
  async function setupAdminPin(pin, adminName) {
    const hash = await sha256(pin);
    const freshConfig = (await storageGet("config")) || config;
    const next = { ...freshConfig, adminPinHash: hash, adminName: (adminName || "").trim() };
    setConfig(next);
    const saved = await storageSetVerified("config", next);
    if (!saved) return { ok: false, msg: "No se pudo guardar. Revisá tu conexión y probá de nuevo." };
    return { ok: true };
  }

  async function verifyAdminPin(pin) {
    const freshConfig = (await storageGet("config")) || config;
    if (!freshConfig.adminPinHash) return false;
    const hash = await sha256(pin);
    return hash === freshConfig.adminPinHash;
  }

  // ---- Canchas ----
  async function addCourse(name, par) {
    const trimmed = name.trim();
    const freshCourses = (await storageGet("courses")) || courses;
    if (!trimmed || freshCourses.some((c) => c.name === trimmed)) return;
    const next = [...freshCourses, { name: trimmed, par: par ? Number(par) : null }];
    setCourses(next);
    await storageSet("courses", next);
  }

  async function removeCourseCascade(name, adminPin) {
    const ok = await verifyAdminPin(adminPin);
    if (!ok) return { ok: false, msg: "PIN de administrador incorrecto." };

    const freshRounds = (await storageGet("rounds")) || rounds;
    const affectedRounds = freshRounds.filter((r) => r.course === name);
    for (const r of affectedRounds) {
      if (r.hasPhoto) await storageDelete(`photo:${r.id}`);
    }
    const nextRounds = freshRounds.filter((r) => r.course !== name);
    setRounds(nextRounds);
    await storageSet("rounds", nextRounds);

    const freshCourses = (await storageGet("courses")) || courses;
    const nextCourses = freshCourses.filter((c) => c.name !== name);
    setCourses(nextCourses);
    await storageSet("courses", nextCourses);

    // saca la cancha borrada del courseConfig de TODOS los torneos, para que no queden
    // columnas huérfanas en Posiciones sin cancha ni tarjetas detrás
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const nextTournaments = freshTournaments.map((t) => ({
      ...t,
      courseConfig: (t.courseConfig || []).filter((cc) => cc.course !== name),
    }));
    setTournaments(nextTournaments);
    await storageSet("tournaments", nextTournaments);

    showToast(`Cancha "${name}" borrada junto con ${affectedRounds.length} tarjeta(s) cargada(s) en ella.`, 5000);
    return { ok: true };
  }

  async function editCoursePar(name, par, adminPin) {
    const ok = await verifyAdminPin(adminPin);
    if (!ok) return { ok: false, msg: "PIN de administrador incorrecto." };
    const freshCourses = (await storageGet("courses")) || courses;
    const next = freshCourses.map((c) => (c.name === name ? { ...c, par: par ? Number(par) : null } : c));
    setCourses(next);
    await storageSet("courses", next);
    showToast(`Cancha ${name} actualizada`);
    return { ok: true };
  }

  // ---- Torneos ----
  async function addTournament(name, startDate, endDate, controller1, controller2, courseConfig) {
    const trimmed = name.trim();
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    if (!trimmed || freshTournaments.some((t) => t.name === trimmed)) return { ok: false, msg: "Ese nombre ya existe." };
    if (!startDate || !endDate) return { ok: false, msg: "Ingresá fecha desde y hasta." };
    if (startDate > endDate) return { ok: false, msg: "La fecha desde no puede ser posterior a la fecha hasta." };
    if (!controller1 || !controller2) return { ok: false, msg: "Elegí los 2 controllers." };
    if (controller1 === controller2) return { ok: false, msg: "Los 2 controllers deben ser jugadores distintos." };
    if (!courseConfig || courseConfig.length === 0) return { ok: false, msg: "Elegí al menos una cancha para el torneo." };
    for (const cc of courseConfig) {
      if (!cc.minRequired || cc.minRequired < 1) return { ok: false, msg: `Ingresá un mínimo válido para ${cc.course}.` };
      if (cc.mode !== "suma" && cc.mode !== "mejor") return { ok: false, msg: `Elegí modalidad para ${cc.course}.` };
    }
    const next = [
      ...freshTournaments,
      {
        name: trimmed,
        startDate,
        endDate,
        enabled: true,
        controllers: [controller1, controller2],
        participants: [],
        courseConfig,
      },
    ];
    setTournaments(next);
    await storageSet("tournaments", next);
    return { ok: true };
  }

  async function toggleTournament(name) {
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const next = freshTournaments.map((t) => (t.name === name ? { ...t, enabled: !t.enabled } : t));
    setTournaments(next);
    await storageSet("tournaments", next);
  }

  async function removeTournament(name) {
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const next = freshTournaments.filter((t) => t.name !== name);
    setTournaments(next);
    await storageSet("tournaments", next);
  }

  async function editTournamentDates(name, startDate, endDate, adminPin) {
    const ok = await verifyAdminPin(adminPin);
    if (!ok) return { ok: false, msg: "PIN de administrador incorrecto." };
    if (!startDate || !endDate) return { ok: false, msg: "Ingresá fecha desde y hasta." };
    if (startDate > endDate) return { ok: false, msg: "La fecha desde no puede ser posterior a la fecha hasta." };
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const next = freshTournaments.map((t) => (t.name === name ? { ...t, startDate, endDate } : t));
    setTournaments(next);
    await storageSet("tournaments", next);
    showToast(`Torneo ${name} actualizado`);
    return { ok: true };
  }

  // ---- Postulación por torneo ----
  async function requestJoinTournament(tournamentName, playerName, pin) {
    const okPin = await verifyPin(playerName, pin);
    if (!okPin) return { ok: false, msg: "PIN incorrecto para ese nombre." };
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const t = freshTournaments.find((tt) => tt.name === tournamentName);
    if (!t) return { ok: false, msg: "Torneo no encontrado." };
    const already = (t.participants || []).find((p) => p.name === playerName);
    if (already) return { ok: false, msg: `Ya tenés una postulación (${already.status}) en ese torneo.` };
    const next = freshTournaments.map((tt) =>
      tt.name === tournamentName
        ? { ...tt, participants: [...(tt.participants || []), { name: playerName, status: "pendiente" }] }
        : tt
    );
    setTournaments(next);
    const saved = await storageSetVerified("tournaments", next);
    if (!saved) return { ok: false, msg: "No se pudo guardar la postulación. Probá de nuevo." };
    return { ok: true };
  }

  async function approveParticipant(tournamentName, playerName) {
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const next = freshTournaments.map((t) =>
      t.name === tournamentName
        ? { ...t, participants: (t.participants || []).map((p) => (p.name === playerName ? { ...p, status: "aceptado" } : p)) }
        : t
    );
    setTournaments(next);
    await storageSet("tournaments", next);
  }

  async function rejectParticipant(tournamentName, playerName) {
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const next = freshTournaments.map((t) =>
      t.name === tournamentName
        ? { ...t, participants: (t.participants || []).filter((p) => p.name !== playerName) }
        : t
    );
    setTournaments(next);
    await storageSet("tournaments", next);
  }

  // ---- Verificación de tarjetas (controllers + admin) ----
  async function verifyRound(roundId, enteredPin) {
    const freshRounds = (await storageGet("rounds")) || rounds;
    const round = freshRounds.find((r) => r.id === roundId);
    if (!round) return { ok: false, msg: "Tarjeta no encontrada." };
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const t = freshTournaments.find((tt) => tt.name === round.tournament);
    const controllerNames = (t && t.controllers) || [];
    const hash = await sha256(enteredPin);
    const freshUsers = (await storageGet("users")) || users;

    const freshConfig = (await storageGet("config")) || config;
    const isAdmin = freshConfig.adminPinHash && hash === freshConfig.adminPinHash;
    const matchingController = freshUsers.find((u) => controllerNames.includes(u.name) && u.pinHash === hash);

    if (!isAdmin && !matchingController) {
      return { ok: false, msg: "PIN incorrecto, o no sos controller de este torneo." };
    }
    if (matchingController && matchingController.name === round.player) {
      return { ok: false, msg: "No podés verificar tu propia tarjeta." };
    }
    if (isAdmin && freshConfig.adminName && freshConfig.adminName === round.player) {
      return { ok: false, msg: "No podés verificar tu propia tarjeta (sos el administrador de este torneo)." };
    }

    const verifierName = matchingController ? matchingController.name : "Administrador";
    const next = freshRounds.map((r) => (r.id === roundId ? { ...r, verified: true, verifiedBy: verifierName } : r));
    setRounds(next);
    await storageSet("rounds", next);
    showToast(`Tarjeta verificada por ${verifierName}`);
    return { ok: true };
  }

  async function editRoundScore(roundId, newScore, enteredPin) {
    const freshRounds = (await storageGet("rounds")) || rounds;
    const round = freshRounds.find((r) => r.id === roundId);
    if (!round) return { ok: false, msg: "Tarjeta no encontrada." };
    const scoreNum = Number(newScore);
    if (!newScore || isNaN(scoreNum) || scoreNum <= 0) return { ok: false, msg: "Ingresá un score válido." };

    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const t = freshTournaments.find((tt) => tt.name === round.tournament);
    const controllerNames = (t && t.controllers) || [];
    const hash = await sha256(enteredPin);
    const freshUsers = (await storageGet("users")) || users;

    const freshConfig = (await storageGet("config")) || config;
    const isAdmin = freshConfig.adminPinHash && hash === freshConfig.adminPinHash;
    const matchingController = freshUsers.find((u) => controllerNames.includes(u.name) && u.pinHash === hash);

    if (!isAdmin && !matchingController) {
      return { ok: false, msg: "PIN incorrecto, o no sos controller de este torneo." };
    }
    if (matchingController && matchingController.name === round.player) {
      return { ok: false, msg: "No podés editar tu propia tarjeta." };
    }
    if (isAdmin && freshConfig.adminName && freshConfig.adminName === round.player) {
      return { ok: false, msg: "No podés editar tu propia tarjeta (sos el administrador de este torneo)." };
    }

    const editorName = matchingController ? matchingController.name : "Administrador";
    const next = freshRounds.map((r) =>
      r.id === roundId ? { ...r, score: scoreNum, verified: true, verifiedBy: editorName } : r
    );
    setRounds(next);
    await storageSet("rounds", next);
    showToast(`Score corregido a ${scoreNum} por ${editorName} — queda verificada`);
    return { ok: true };
  }

  // ---- Rondas ----
  async function addRound(round, photoDataUrl) {
    const id = uid();
    let hasPhoto = false;
    let photoDiag = null;
    if (photoDataUrl) {
      const result = await safeSetPhoto(id, photoDataUrl, { player: round.player, course: round.course, date: round.date });
      hasPhoto = result.ok;
      if (!result.ok) photoDiag = result;
    }

    const record = { ...round, id, tournament, hasPhoto, verified: false, verifiedBy: null };
    const freshRounds = (await storageGet("rounds")) || rounds;
    const next = [...freshRounds, record];
    setRounds(next);
    await storageSet("rounds", next);

    if (photoDataUrl && !hasPhoto) {
      showToast(`Foto NO guardada — ${photoDiag.reason} (tamaño: ${Math.round(photoDiag.size / 1024)} KB)`, 8000);
    } else {
      showToast("Tarjeta cargada");
    }
  }

  // ---- Tesorería (admin + controllers de cualquier torneo) ----
  async function verifyTreasuryAccess(pin) {
    const hash = await sha256(pin);
    const freshConfig = (await storageGet("config")) || config;
    if (freshConfig.adminPinHash && hash === freshConfig.adminPinHash) return true;

    const freshUsers = (await storageGet("users")) || users;
    const freshTournaments = (await storageGet("tournaments")) || tournaments;
    const allControllerNames = new Set();
    freshTournaments.forEach((t) => (t.controllers || []).forEach((c) => allControllerNames.add(c)));
    const match = freshUsers.find((u) => allControllerNames.has(u.name) && u.pinHash === hash);
    return !!match;
  }

  async function addTreasuryMovement(entry) {
    const fresh = (await storageGet("tesoreria")) || tesoreria;
    const record = { ...entry, id: uid() };
    const next = [...fresh, record];
    setTesoreria(next);
    const saved = await storageSetVerified("tesoreria", next);
    if (!saved) return { ok: false, msg: "No se pudo guardar. Probá de nuevo." };
    showToast(entry.type === "ingreso" ? "Ingreso registrado" : "Egreso registrado");
    return { ok: true };
  }

  async function removeTreasuryMovement(id) {
    const fresh = (await storageGet("tesoreria")) || tesoreria;
    const next = fresh.filter((m) => m.id !== id);
    setTesoreria(next);
    await storageSet("tesoreria", next);
    showToast("Movimiento eliminado");
  }

  function exportTreasuryCSV() {
    const header = ["Fecha", "Tipo", "Categoría", "Concepto", "Monto", "Registrado por"];
    const rows = [
      header,
      ...[...tesoreria].sort((a, b) => (a.date < b.date ? 1 : -1)).map((m) => [
        formatDate(m.date),
        m.type === "ingreso" ? "Ingreso" : "Egreso",
        m.category,
        m.concept,
        m.amount,
        m.registeredBy || "",
      ]),
    ];
    downloadCSV("tesoreria_los_optimistas.csv", rows);
  }

  function exportHistorialCSV() {
    const header = ["Jugador", "Cancha", "Fecha", "Score", "Handicap", "Verificada", "Verificada por", "Suma al total ahora"];
    const rows = [
      header,
      ...roundsInTournament.map((r) => [
        r.player,
        r.course,
        formatDate(r.date),
        r.score,
        r.handicap != null ? r.handicap : "",
        r.verified ? "Sí" : "No",
        r.verifiedBy || "",
        countingRoundIds.has(r.id) ? "Sí" : "No",
      ]),
    ];
    downloadCSV(`historial_${safeFilename(tournament)}.csv`, rows);
  }

  function exportPosicionesCSV() {
    const header = ["Jugador", ...courseNames, "Total", "Elegible (cumple mínimos)", "Todo verificado"];
    const rows = [
      header,
      ...standings.map((s) => [
        s.player,
        ...courseNames.map((c) => (s.courseData[c] ? s.courseData[c].subtotal : "")),
        s.total,
        s.eligible ? "Sí" : "No",
        verifiedStatusByPlayer[s.player] ? "Sí" : "No",
      ]),
    ];
    downloadCSV(`posiciones_${safeFilename(tournament)}.csv`, rows);
  }

  async function resetAllData(adminPin) {
    const ok = await verifyAdminPin(adminPin);
    if (!ok) return { ok: false, msg: "PIN de administrador incorrecto." };

    // borrar fotos asociadas a las rondas actuales antes de borrar las rondas
    for (const r of rounds) {
      if (r.hasPhoto) await storageDelete(`photo:${r.id}`);
    }

    await storageDelete("users");
    await storageDelete("courses");
    await storageDelete("tournaments");
    await storageDelete("rounds");
    await storageDelete("config");
    await storageDelete("players"); // clave vieja de migración, por las dudas

    setUsers([]);
    setCourses([]);
    setTournaments([]);
    setRounds([]);
    setConfig({});
    setAdminAuthed(false);
    setShowAdmin(false);
    showToast("Se borraron todos los datos. La app queda como recién instalada.", 6000);
    return { ok: true };
  }


  async function performDeleteRound(id) {
    const round = rounds.find((r) => r.id === id);
    const next = rounds.filter((r) => r.id !== id);
    setRounds(next);
    await storageSet("rounds", next);
    if (round && round.hasPhoto) await storageDelete(`photo:${id}`);
  }

  async function requestDeleteRound(id, pin) {
    const round = rounds.find((r) => r.id === id);
    if (!round) return false;
    const hash = await sha256(pin);
    const freshConfig = (await storageGet("config")) || config;
    const okAdmin = freshConfig.adminPinHash && hash === freshConfig.adminPinHash;
    if (!okAdmin) return false;
    await performDeleteRound(id);
    return true;
  }

  async function clearBrokenPhoto(id) {
    const next = rounds.map((r) => (r.id === id ? { ...r, hasPhoto: false } : r));
    setRounds(next);
    await storageSet("rounds", next);
    showToast("Se quitó la referencia a la foto");
  }

  async function openPhoto(roundId) {
    setViewingPhoto({ roundId, url: null, loading: true, error: false });
    const rec = await storageGet(`photo:${roundId}`);
    const url = rec ? rec.url : null;
    setViewingPhoto({ roundId, url, loading: false, error: !url });
  }

  return (
    <div
      style={{
        minHeight: "100vh",
        background: COLORS.green900,
        fontFamily: "'Inter', sans-serif",
        color: COLORS.paper,
        display: "flex",
        flexDirection: "column",
      }}
    >
      <style>{`
        * { box-sizing: border-box; }
        ::selection { background: ${COLORS.brass}; color: ${COLORS.green900}; }
        input, select { font-family: 'Inter', sans-serif; }
        input:focus, select:focus, button:focus-visible {
          outline: 2px solid ${COLORS.brassLight};
          outline-offset: 2px;
        }
        @media (prefers-reduced-motion: reduce) {
          * { transition: none !important; animation: none !important; }
        }
        .glTab { transition: color 0.15s, border-color 0.15s; }
        .glRow { transition: background 0.15s; }
        .glRow:hover { background: rgba(255,255,255,0.03); }
        .glBtn { transition: transform 0.1s, background 0.15s; }
        .glBtn:active { transform: scale(0.97); }
        @keyframes spin { to { transform: rotate(360deg); } }
      `}</style>

      <Header
        onAdmin={() => setShowAdmin(true)}
        onRegister={() => setShowRegister(true)}
        onTreasury={() => setShowTreasury(true)}
        tournament={tournament}
        setTournament={setTournament}
        enabledTournaments={enabledTournaments}
        pendingCount={totalPendingCount}
      />

      <div style={{ maxWidth: 720, margin: "0 auto", width: "100%", flex: 1, padding: "0 20px 60px" }}>
        <Tabs tab={tab} setTab={setTab} />

        {loading ? (
          <div style={{ display: "flex", justifyContent: "center", padding: "60px 0", opacity: 0.6 }}>
            <Loader2 size={22} style={{ animation: "spin 1s linear infinite" }} />
          </div>
        ) : enabledTournaments.length === 0 ? (
          <Card>
            <p style={{ margin: 0, fontSize: 14 }}>
              Todavía no hay torneos habilitados. Entrá al panel de administrador (ícono del engranaje) para
              crear uno y habilitarlo.
            </p>
          </Card>
        ) : (
          <>
            {tab === "cargar" && (
              <CargarTarjeta
                users={acceptedParticipantUsers}
                courseNames={courseNames}
                onSubmit={addRound}
                onVerifyPin={verifyPin}
                tournament={tournament}
                tournamentWindow={activeTournamentObj}
              />
            )}
            {tab === "posiciones" && (
              <Posiciones
                standings={standings}
                courses={courseNames}
                parByCourse={parByCourse}
                courseRuleByName={courseRuleByName}
                roundsInTournament={roundsInTournament}
                tournament={tournament}
                tournamentInfo={activeTournamentObj}
                adminName={config.adminName}
                verifiedStatusByPlayer={verifiedStatusByPlayer}
                onVerifyRound={verifyRound}
              />
            )}
            {tab === "historial" && (
              <Historial
                rounds={roundsInTournament}
                onRequestDelete={requestDeleteRound}
                onViewPhoto={openPhoto}
                onVerifyRound={verifyRound}
                onEditScore={editRoundScore}
                tournament={tournament}
                countingRoundIds={countingRoundIds}
              />
            )}
            {tab === "estadisticas" && (
              <Estadisticas
                rounds={roundsInTournament}
                courses={courseNames}
                parByCourse={parByCourse}
                tournament={tournament}
              />
            )}
          </>
        )}
      </div>

      {showAdmin && (
        <AdminGate
          hasAdminPin={!!config.adminPinHash}
          authed={adminAuthed}
          onSetupPin={async (pin, adminName) => {
            const res = await setupAdminPin(pin, adminName);
            if (res.ok) setAdminAuthed(true);
            return res;
          }}
          onVerifyPin={async (pin) => {
            const ok = await verifyAdminPin(pin);
            if (ok) setAdminAuthed(true);
            return ok;
          }}
          onClose={() => {
            setShowAdmin(false);
            setAdminAuthed(false);
          }}
        >
          <AdminPanel
            pendingUsers={pendingUsers}
            approvedUsers={approvedUsers}
            courses={courses}
            tournaments={tournaments}
            onApprove={approveUser}
            onReject={rejectUser}
            onRemoveUser={removeUser}
            onSetUserPin={setUserPin}
            onAddUserDirect={addUserDirect}
            onAddCourse={addCourse}
            onRemoveCourse={removeCourseCascade}
            roundsCountByCourse={roundsCountByCourse}
            onEditCoursePar={editCoursePar}
            onAddTournament={addTournament}
            onToggleTournament={toggleTournament}
            onRemoveTournament={removeTournament}
            onEditTournamentDates={editTournamentDates}
            onApproveParticipant={approveParticipant}
            onRejectParticipant={rejectParticipant}
            onResetAllData={resetAllData}
            onRefreshData={refreshData}
            onExportHistorial={exportHistorialCSV}
            onExportPosiciones={exportPosicionesCSV}
            activeTournamentName={tournament}
            onClose={() => {
              setShowAdmin(false);
              setAdminAuthed(false);
            }}
          />
        </AdminGate>
      )}

      {showTreasury && (
        <TreasuryGate
          authed={treasuryAuthed}
          onVerifyPin={async (pin) => {
            const ok = await verifyTreasuryAccess(pin);
            if (ok) setTreasuryAuthed(true);
            return ok;
          }}
          onClose={() => {
            setShowTreasury(false);
            setTreasuryAuthed(false);
          }}
        >
          <TreasuryPanel
            movements={tesoreria}
            onAdd={addTreasuryMovement}
            onRemove={removeTreasuryMovement}
            onExport={exportTreasuryCSV}
            onClose={() => {
              setShowTreasury(false);
              setTreasuryAuthed(false);
            }}
          />
        </TreasuryGate>
      )}

      {showRegister && (
        <RegisterModal
          onRegister={registerUser}
          onClose={() => setShowRegister(false)}
          approvedUserNames={approvedNames}
          enabledTournaments={enabledTournaments}
          onJoinTournament={requestJoinTournament}
        />
      )}

      {viewingPhoto && (
        <PhotoModal
          state={viewingPhoto}
          onClose={() => setViewingPhoto(null)}
          onClearBroken={async () => {
            await clearBrokenPhoto(viewingPhoto.roundId);
            setViewingPhoto(null);
          }}
        />
      )}

      {toast && <Toast msg={toast} />}
    </div>
  );
}

export default function GolfLeague() {
  return (
    <ErrorBoundary>
      <GolfLeagueInner />
    </ErrorBoundary>
  );
}

function Header({ onAdmin, onRegister, onTreasury, tournament, setTournament, enabledTournaments, pendingCount }) {
  return (
    <div
      style={{
        background: `linear-gradient(180deg, ${COLORS.green900} 0%, ${COLORS.green700} 100%)`,
        borderBottom: `1px solid rgba(200,167,107,0.25)`,
        padding: "28px 20px 18px",
      }}
    >
      <div style={{ maxWidth: 720, margin: "0 auto" }}>
        <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between" }}>
          <div>
            <div
              style={{
                display: "flex", alignItems: "center", gap: 8, color: COLORS.brassLight,
                fontSize: 12, letterSpacing: "0.14em", textTransform: "uppercase",
                fontFamily: "'IBM Plex Mono', monospace", marginBottom: 6,
              }}
            >
              <Flag size={14} strokeWidth={2.2} />
              Liga · Score neto · {APP_VERSION}
            </div>
            <h1 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 34, lineHeight: 1.05, color: COLORS.paper }}>
              Los Optimistas
            </h1>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <button
              onClick={onAdmin}
              className="glBtn"
              aria-label={pendingCount > 0 ? `Administrar — ${pendingCount} solicitud(es) pendiente(s)` : "Administrar"}
              style={{
                position: "relative",
                background: "transparent", border: `1px solid rgba(200,167,107,0.4)`, borderRadius: 8,
                width: 38, height: 38, display: "flex", alignItems: "center", justifyContent: "center",
                color: COLORS.brassLight, cursor: "pointer",
              }}
            >
              <Settings size={17} />
              {pendingCount > 0 && (
                <span
                  title={`${pendingCount} solicitud(es) pendiente(s)`}
                  style={{
                    position: "absolute", top: -6, right: -6,
                    minWidth: 18, height: 18, padding: "0 4px",
                    borderRadius: 9, background: "#C0392B", color: "#fff",
                    fontSize: 10.5, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace",
                    display: "flex", alignItems: "center", justifyContent: "center",
                    border: `2px solid ${COLORS.green900}`, lineHeight: 1,
                  }}
                >
                  {pendingCount > 9 ? "9+" : pendingCount}
                </span>
              )}
            </button>
            <button
              onClick={onTreasury}
              className="glBtn"
              aria-label="Tesorería"
              title="Tesorería (admin y controllers)"
              style={{
                background: "transparent", border: `1px solid rgba(200,167,107,0.4)`, borderRadius: 8,
                width: 38, height: 38, display: "flex", alignItems: "center", justifyContent: "center",
                color: COLORS.brassLight, cursor: "pointer",
              }}
            >
              <Wallet size={17} />
            </button>
            <button
              onClick={onRegister}
              className="glBtn"
              aria-label="Sumarme al grupo"
              title="Sumarme al grupo"
              style={{
                background: "transparent", border: `1px solid rgba(200,167,107,0.4)`, borderRadius: 8,
                width: 38, height: 38, display: "flex", alignItems: "center", justifyContent: "center",
                color: COLORS.brassLight, cursor: "pointer",
              }}
            >
              <Plus size={17} />
            </button>
          </div>
        </div>

        {enabledTournaments.length > 0 && (
          <div style={{ display: "flex", gap: 6, marginTop: 18, overflowX: "auto" }}>
            {enabledTournaments.map((t) => (
              <button
                key={t.name}
                onClick={() => setTournament(t.name)}
                className="glBtn"
                style={{
                  flex: enabledTournaments.length <= 3 ? 1 : "0 0 auto",
                  background: tournament === t.name ? COLORS.brass : "rgba(255,255,255,0.06)",
                  color: tournament === t.name ? COLORS.green900 : "rgba(247,245,239,0.7)", border: "none",
                  borderRadius: 8, padding: "9px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer",
                  whiteSpace: "nowrap",
                }}
              >
                {t.name}
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function Tabs({ tab, setTab }) {
  const items = [
    { id: "posiciones", label: "Posiciones" },
    { id: "historial", label: "Historial" },
    { id: "estadisticas", label: "Estadísticas" },
    { id: "cargar", label: "Cargar" },
  ];
  return (
    <div style={{ display: "flex", gap: 4, marginTop: 24, marginBottom: 20, borderBottom: `1px solid rgba(255,255,255,0.1)`, overflowX: "auto" }}>
      {items.map((it) => (
        <button
          key={it.id}
          className="glTab"
          onClick={() => setTab(it.id)}
          style={{
            background: "none", border: "none",
            borderBottom: `2px solid ${tab === it.id ? COLORS.brass : "transparent"}`,
            color: tab === it.id ? COLORS.paper : "rgba(247,245,239,0.5)",
            fontFamily: "'Inter', sans-serif", fontWeight: 600, fontSize: 14,
            padding: "10px 14px", cursor: "pointer", whiteSpace: "nowrap",
          }}
        >
          {it.label}
        </button>
      ))}
    </div>
  );
}

function Switch({ on, onClick }) {
  return (
    <button
      onClick={onClick}
      role="switch"
      aria-checked={on}
      style={{
        width: 40, height: 22, borderRadius: 999, border: "none", cursor: "pointer",
        background: on ? COLORS.green700 : "rgba(27,36,32,0.2)",
        position: "relative", flexShrink: 0, transition: "background 0.15s",
      }}
    >
      <span
        style={{
          position: "absolute", top: 2, left: on ? 20 : 2, width: 18, height: 18, borderRadius: "50%",
          background: COLORS.paper, transition: "left 0.15s", boxShadow: "0 1px 3px rgba(0,0,0,0.3)",
        }}
      />
    </button>
  );
}

function Card({ children, style }) {
  return (
    <div style={{ background: COLORS.paper, color: COLORS.ink, borderRadius: 14, padding: 22, boxShadow: "0 8px 24px rgba(0,0,0,0.25)", ...style }}>
      {children}
    </div>
  );
}

function FieldLabel({ children, style }) {
  return (
    <label style={{ display: "block", fontSize: 11, letterSpacing: "0.08em", textTransform: "uppercase", fontWeight: 600, color: "rgba(27,36,32,0.55)", marginBottom: 6, fontFamily: "'IBM Plex Mono', monospace", ...style }}>
      {children}
    </label>
  );
}

function Select({ value, onChange, placeholder, children }) {
  return (
    <select value={value} onChange={(e) => onChange(e.target.value)} style={inputStyle}>
      <option value="" disabled>{placeholder}</option>
      {children}
    </select>
  );
}

const inputStyle = {
  width: "100%", padding: "11px 12px", borderRadius: 8, border: `1px solid ${COLORS.paperDim}`,
  background: "#fff", fontSize: 14, color: COLORS.ink,
};

const primaryBtnStyle = {
  background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 10,
  padding: "13px 0", fontSize: 15, fontWeight: 600, cursor: "pointer",
  display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
};

// ---------------- Cargar tarjeta ----------------

function CargarTarjeta({ users, courseNames, onSubmit, onVerifyPin, tournament, tournamentWindow }) {
  const [player, setPlayer] = useState("");
  const [course, setCourse] = useState("");
  const [score, setScore] = useState("");
  const [handicap, setHandicap] = useState("");
  const [date, setDate] = useState(todayStr());
  const [photoFile, setPhotoFile] = useState(null);
  const [photoPreview, setPhotoPreview] = useState(null);
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [submitting, setSubmitting] = useState(false);

  const selectedUser = users.find((u) => u.name === player);
  const pinNotAssignedYet = selectedUser && !selectedUser.pinHash;

  function handlePhotoChange(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setPhotoFile(file);
    const reader = new FileReader();
    reader.onload = (ev) => setPhotoPreview(ev.target.result);
    reader.readAsDataURL(file);
  }

  async function handleSubmit() {
    setError("");
    if (!player) return setError("Elegí tu nombre.");
    if (pinNotAssignedYet) return setError("Este jugador todavía no tiene PIN asignado. Pedile al administrador que te lo asigne desde el panel.");
    if (!course) return setError("Elegí una cancha.");
    if (!date) return setError("Ingresá la fecha.");
    if (date > todayStr()) return setError("La fecha no puede ser futura.");
    if (tournamentWindow) {
      if (tournamentWindow.startDate && date < tournamentWindow.startDate) {
        return setError(`La fecha no puede ser anterior al inicio del torneo (${formatDate(tournamentWindow.startDate)}).`);
      }
      if (tournamentWindow.endDate && date > tournamentWindow.endDate) {
        return setError(`La fecha no puede ser posterior al cierre del torneo (${formatDate(tournamentWindow.endDate)}).`);
      }
    }
    const scoreNum = Number(score);
    if (!score || isNaN(scoreNum) || scoreNum <= 0) return setError("Ingresá un score neto válido.");
    const handicapNum = handicap === "" ? null : Number(handicap);
    if (handicap !== "" && (isNaN(handicapNum) || handicapNum < 0)) return setError("Ingresá un handicap válido.");
    if (!/^\d{4}$/.test(pin)) return setError("Ingresá tu PIN de 4 dígitos.");

    setSubmitting(true);
    try {
      const ok = await onVerifyPin(player, pin);
      if (!ok) {
        setError("PIN incorrecto.");
        setSubmitting(false);
        return;
      }

      let photoDataUrl = null;
      if (photoFile) photoDataUrl = await compressImage(photoFile);

      await onSubmit({ player, course, score: scoreNum, handicap: handicapNum, date }, photoDataUrl);

      setScore("");
      setHandicap("");
      setPin("");
      setPhotoFile(null);
      setPhotoPreview(null);
    } catch (e) {
      setError("No se pudo procesar la foto. Probá con otra imagen.");
    } finally {
      setSubmitting(false);
    }
  }

  if (users.length === 0) {
    return (
      <Card>
        <p style={{ margin: 0, fontSize: 14 }}>
          Todavía no hay jugadores aceptados en este torneo. Si ya sos parte del grupo, tocá el botón "+" y
          postulate a este torneo. Si sos nuevo, primero sumate al grupo desde ese mismo botón.
        </p>
      </Card>
    );
  }

  return (
    <Card>
      <div style={{ fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace", color: COLORS.brass, fontWeight: 600, marginBottom: 16, textTransform: "uppercase", letterSpacing: "0.05em" }}>
        Cargando para: {tournament}
        {tournamentWindow && (
          <span style={{ display: "block", marginTop: 2, fontWeight: 500, letterSpacing: "normal", textTransform: "none" }}>
            Tarjetas válidas del {formatDate(tournamentWindow.startDate)} al {formatDate(tournamentWindow.endDate)}
          </span>
        )}
      </div>

      <FieldLabel>Jugador</FieldLabel>
      <Select
        value={player}
        onChange={(v) => {
          setPlayer(v);
          setPin("");
          setPinConfirm("");
          setError("");
        }}
        placeholder="Elegí tu nombre"
      >
        {users.map((u) => (
          <option key={u.name} value={u.name}>{u.name}</option>
        ))}
      </Select>

      <FieldLabel style={{ marginTop: 16 }}>Cancha</FieldLabel>
      <Select value={course} onChange={setCourse} placeholder="Elegí una cancha">
        {courseNames.map((c) => <option key={c} value={c}>{c}</option>)}
      </Select>
      {courseNames.length === 0 && (
        <p style={{ fontSize: 12, color: COLORS.danger, marginTop: 6 }}>
          Este torneo no tiene canchas configuradas. Se definen al crear el torneo, desde el panel de administrador.
        </p>
      )}

      <FieldLabel style={{ marginTop: 16 }}>Score neto</FieldLabel>
      <input type="number" inputMode="numeric" value={score} onChange={(e) => setScore(e.target.value)} placeholder="Ej: 72" style={inputStyle} />

      <FieldLabel style={{ marginTop: 16 }}>Handicap jugado del día</FieldLabel>
      <input type="number" step="0.1" inputMode="decimal" value={handicap} onChange={(e) => setHandicap(e.target.value)} placeholder="Ej: 14.2" style={inputStyle} />

      <FieldLabel style={{ marginTop: 16 }}>Fecha</FieldLabel>
      <input
        type="date"
        value={date}
        min={tournamentWindow && tournamentWindow.startDate ? tournamentWindow.startDate : undefined}
        max={tournamentWindow && tournamentWindow.endDate && tournamentWindow.endDate < todayStr() ? tournamentWindow.endDate : todayStr()}
        onChange={(e) => setDate(e.target.value)}
        style={inputStyle}
      />

      <FieldLabel style={{ marginTop: 16 }}>Foto de la tarjeta (opcional)</FieldLabel>
      <label className="glBtn" style={{ display: "flex", alignItems: "center", gap: 8, border: `1px dashed ${COLORS.green700}`, borderRadius: 8, padding: "11px 12px", fontSize: 13.5, color: COLORS.green700, cursor: "pointer", fontWeight: 600 }}>
        <Camera size={16} />
        {photoFile ? "Cambiar foto" : "Subir foto de la tarjeta"}
        <input type="file" accept="image/*" onChange={handlePhotoChange} style={{ display: "none" }} />
      </label>
      {photoPreview && (
        <img src={photoPreview} alt="Vista previa de la tarjeta" style={{ marginTop: 10, maxHeight: 160, borderRadius: 8, border: `1px solid ${COLORS.paperDim}` }} />
      )}

      {player && pinNotAssignedYet && (
        <div style={{ marginTop: 16, background: "rgba(140,59,46,0.08)", border: `1px solid rgba(140,59,46,0.25)`, borderRadius: 8, padding: 14 }}>
          <p style={{ margin: 0, fontSize: 13, color: COLORS.danger, fontWeight: 500 }}>
            Este jugador todavía no tiene PIN asignado. El administrador tiene que asignarlo desde el panel
            antes de que puedas cargar tarjetas con este nombre.
          </p>
        </div>
      )}

      {player && !pinNotAssignedYet && (
        <div style={{ marginTop: 16, background: COLORS.paperDim, borderRadius: 8, padding: 14 }}>
          <FieldLabel style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <Lock size={12} /> Tu PIN
          </FieldLabel>
          <input
            type="password"
            inputMode="numeric"
            maxLength={4}
            value={pin}
            onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))}
            placeholder="4 dígitos"
            style={inputStyle}
          />
        </div>
      )}

      {error && <p style={{ color: COLORS.danger, fontSize: 13, marginTop: 12, marginBottom: 0, fontWeight: 500 }}>{error}</p>}

      <button className="glBtn" onClick={handleSubmit} disabled={submitting} style={{ ...primaryBtnStyle, marginTop: 20, width: "100%", opacity: submitting ? 0.7 : 1, cursor: submitting ? "default" : "pointer" }}>
        {submitting ? "Guardando..." : "Cargar tarjeta"}
        {!submitting && <ChevronRight size={16} />}
      </button>
    </Card>
  );
}

// ---------------- Sumarme (registro) ----------------

function RegisterModal({ onRegister, onClose, approvedUserNames, enabledTournaments, onJoinTournament }) {
  const [mode, setMode] = useState("sumarme"); // "sumarme" | "postularme"

  // --- Sumarme ---
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [pinConfirm, setPinConfirm] = useState("");
  const [error, setError] = useState("");
  const [done, setDone] = useState(false);

  // --- Postularme ---
  const [joinName, setJoinName] = useState("");
  const [joinPin, setJoinPin] = useState("");
  const [joinTournament, setJoinTournament] = useState("");
  const [joinError, setJoinError] = useState("");
  const [joinDone, setJoinDone] = useState(false);

  async function handleSubmit() {
    setError("");
    if (!name.trim()) return setError("Ingresá tu nombre.");
    if (!/^\d{4}$/.test(pin)) return setError("El PIN debe tener 4 dígitos.");
    if (pin !== pinConfirm) return setError("Los dos PIN no coinciden.");

    const res = await onRegister(name, pin);
    if (!res.ok) return setError(res.msg || "No se pudo registrar.");
    setDone(true);
  }

  async function handleJoin() {
    setJoinError("");
    if (!joinName) return setJoinError("Elegí tu nombre.");
    if (!joinTournament) return setJoinError("Elegí un torneo.");
    if (!/^\d{4}$/.test(joinPin)) return setJoinError("Ingresá tu PIN de 4 dígitos.");

    const res = await onJoinTournament(joinTournament, joinName, joinPin);
    if (!res.ok) return setJoinError(res.msg || "No se pudo postular.");
    setJoinDone(true);
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 50 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.paper, color: COLORS.ink, width: "100%", maxWidth: 420, borderRadius: "16px 16px 0 0", padding: 24, maxHeight: "85vh", overflowY: "auto" }}>
        {!done && !joinDone && (
          <div style={{ display: "flex", gap: 6, marginBottom: 18 }}>
            <button
              onClick={() => setMode("sumarme")}
              style={{
                flex: 1, padding: "9px 0", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
                background: mode === "sumarme" ? COLORS.green700 : COLORS.paperDim,
                color: mode === "sumarme" ? COLORS.paper : COLORS.ink,
              }}
            >
              Soy nuevo
            </button>
            <button
              onClick={() => setMode("postularme")}
              style={{
                flex: 1, padding: "9px 0", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
                background: mode === "postularme" ? COLORS.green700 : COLORS.paperDim,
                color: mode === "postularme" ? COLORS.paper : COLORS.ink,
              }}
            >
              Ya soy del grupo
            </button>
          </div>
        )}

        {mode === "sumarme" && (
          done ? (
            <>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <Check size={20} color={COLORS.green700} />
                <h2 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontSize: 19 }}>Listo</h2>
              </div>
              <p style={{ margin: "0 0 18px", fontSize: 14 }}>
                Quedaste registrado en la nómina, pendiente de aprobación. Avisale al administrador del grupo
                para que te habilite — una vez aprobado, vas a poder postularte a los torneos abiertos.
              </p>
              <button onClick={onClose} style={{ width: "100%", background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "10px 0", fontWeight: 600, cursor: "pointer" }}>
                Cerrar
              </button>
            </>
          ) : (
            <>
              <h2 style={{ margin: "0 0 4px", fontFamily: "'Fraunces', serif", fontSize: 20 }}>Sumarme al grupo</h2>
              <p style={{ margin: "0 0 18px", fontSize: 13, color: "rgba(27,36,32,0.6)" }}>
                Esto te suma a la nómina general. Después vas a tener que postularte a cada torneo por separado.
              </p>

              <FieldLabel>Nombre</FieldLabel>
              <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Tu nombre" style={inputStyle} />

              <FieldLabel style={{ marginTop: 14 }}>Elegí un PIN de 4 dígitos</FieldLabel>
              <input type="password" inputMode="numeric" maxLength={4} value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} placeholder="4 dígitos" style={inputStyle} />

              <FieldLabel style={{ marginTop: 14 }}>Repetí el PIN</FieldLabel>
              <input type="password" inputMode="numeric" maxLength={4} value={pinConfirm} onChange={(e) => setPinConfirm(e.target.value.replace(/\D/g, ""))} placeholder="4 dígitos" style={inputStyle} />

              {error && <p style={{ color: COLORS.danger, fontSize: 13, marginTop: 12, marginBottom: 0, fontWeight: 500 }}>{error}</p>}

              <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
                <button className="glBtn" onClick={handleSubmit} style={{ ...primaryBtnStyle, flex: 1 }}>
                  <UserPlus size={16} /> Registrarme
                </button>
                <button onClick={onClose} style={{ background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 10, padding: "0 16px", fontWeight: 600, cursor: "pointer" }}>
                  Cerrar
                </button>
              </div>
            </>
          )
        )}

        {mode === "postularme" && (
          joinDone ? (
            <>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <Check size={20} color={COLORS.green700} />
                <h2 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontSize: 19 }}>Listo</h2>
              </div>
              <p style={{ margin: "0 0 18px", fontSize: 14 }}>
                Quedaste postulado a ese torneo, pendiente de aceptación. Avisale al administrador para que
                te acepte — recién ahí tu nombre aparece en Cargar tarjeta para ese torneo.
              </p>
              <button onClick={onClose} style={{ width: "100%", background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "10px 0", fontWeight: 600, cursor: "pointer" }}>
                Cerrar
              </button>
            </>
          ) : (
            <>
              <h2 style={{ margin: "0 0 4px", fontFamily: "'Fraunces', serif", fontSize: 20 }}>Postularme a un torneo</h2>
              <p style={{ margin: "0 0 18px", fontSize: 13, color: "rgba(27,36,32,0.6)" }}>
                Solo si ya estás en la nómina general del grupo.
              </p>

              {approvedUserNames.length === 0 ? (
                <p style={{ fontSize: 13.5 }}>Todavía no hay nadie en la nómina general.</p>
              ) : (
                <>
                  <FieldLabel>Tu nombre</FieldLabel>
                  <Select value={joinName} onChange={setJoinName} placeholder="Elegí tu nombre">
                    {approvedUserNames.map((n) => <option key={n} value={n}>{n}</option>)}
                  </Select>

                  <FieldLabel style={{ marginTop: 14 }}>Torneo</FieldLabel>
                  {enabledTournaments.length === 0 ? (
                    <p style={{ fontSize: 13.5 }}>No hay torneos habilitados en este momento.</p>
                  ) : (
                    <Select value={joinTournament} onChange={setJoinTournament} placeholder="Elegí un torneo">
                      {enabledTournaments.map((t) => <option key={t.name} value={t.name}>{t.name}</option>)}
                    </Select>
                  )}

                  <FieldLabel style={{ marginTop: 14 }}>Tu PIN</FieldLabel>
                  <input type="password" inputMode="numeric" maxLength={4} value={joinPin} onChange={(e) => setJoinPin(e.target.value.replace(/\D/g, ""))} placeholder="4 dígitos" style={inputStyle} />

                  {joinError && <p style={{ color: COLORS.danger, fontSize: 13, marginTop: 12, marginBottom: 0, fontWeight: 500 }}>{joinError}</p>}

                  <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
                    <button className="glBtn" onClick={handleJoin} style={{ ...primaryBtnStyle, flex: 1 }}>
                      Postularme
                    </button>
                    <button onClick={onClose} style={{ background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 10, padding: "0 16px", fontWeight: 600, cursor: "pointer" }}>
                      Cerrar
                    </button>
                  </div>
                </>
              )}
            </>
          )
        )}
      </div>
    </div>
  );
}

// ---------------- Posiciones ----------------

function Posiciones({ standings, courses, parByCourse, courseRuleByName, roundsInTournament, tournament, tournamentInfo, adminName, verifiedStatusByPlayer, onVerifyRound }) {
  const [detail, setDetail] = useState(null); // { player, course } | null
  const [sortKey, setSortKey] = useState("total"); // "total" | nombre de cancha
  const [sortDir, setSortDir] = useState("asc"); // "asc" | "desc"

  const banner = (
    <div style={{ fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace", color: COLORS.brassLight, fontWeight: 600, marginBottom: 12, textTransform: "uppercase", letterSpacing: "0.05em" }}>
      Viendo: {tournament}
    </div>
  );

  if (standings.length === 0) {
    return (
      <>
        {banner}
        <Card><p style={{ margin: 0, fontSize: 14 }}>Todavía no hay tarjetas cargadas en este torneo.</p></Card>
      </>
    );
  }

  function handleSortClick(key) {
    if (sortKey === key) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortKey(key);
      setSortDir("asc");
    }
  }

  function valueFor(s, key) {
    if (key === "total") return s.total;
    const d = s.courseData[key];
    return d ? d.subtotal : null;
  }

  const sortedStandings = [...standings].sort((a, b) => {
    const va = valueFor(a, sortKey);
    const vb = valueFor(b, sortKey);
    // los que no jugaron esa cancha quedan siempre al final, sea cual sea la dirección
    if (va == null && vb == null) return 0;
    if (va == null) return 1;
    if (vb == null) return -1;
    return sortDir === "asc" ? va - vb : vb - va;
  });

  function sortArrow(key) {
    if (sortKey !== key) return "";
    return sortDir === "asc" ? " ▲" : " ▼";
  }

  return (
    <>
      {banner}
      <Card style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13.5 }}>
          <thead>
            <tr style={{ background: COLORS.green900, color: COLORS.paper }}>
              <th style={thStyle}>#</th>
              <th style={{ ...thStyle, textAlign: "left" }}>Jugador</th>
              {courses.map((c) => {
                const rule = courseRuleByName[c];
                const ruleLabel = rule ? (rule.mode === "mejor" ? `mejor de ${rule.minRequired}` : `mejores ${rule.minRequired}`) : "";
                return (
                  <th
                    key={c}
                    style={{ ...thStyle, cursor: "pointer" }}
                    title={`${parByCourse[c] ? `Par ${parByCourse[c]} · ` : ""}${ruleLabel} — tocá para ordenar`}
                    onClick={() => handleSortClick(c)}
                  >
                    {c.length > 12 ? c.slice(0, 11) + "…" : c}{sortArrow(c)}
                    <div style={{ fontSize: 9.5, opacity: 0.6, fontWeight: 500, marginTop: 1 }}>
                      {parByCourse[c] ? `par ${parByCourse[c]} · ` : ""}{ruleLabel}
                    </div>
                  </th>
                );
              })}
              <th style={{ ...thStyle, color: COLORS.brassLight, cursor: "pointer" }} title="Tocá para ordenar" onClick={() => handleSortClick("total")}>
                Total{sortArrow("total")}
              </th>
            </tr>
          </thead>
          <tbody>
            {sortedStandings.map((s, i) => (
              <tr key={s.player} className="glRow" style={{ borderBottom: `1px solid ${COLORS.paperDim}` }}>
                <td style={{ ...tdStyle, fontFamily: "'IBM Plex Mono', monospace", color: i === 0 && sortKey === "total" && sortDir === "asc" ? COLORS.brass : COLORS.ink }}>
                  {i === 0 && sortKey === "total" && sortDir === "asc" ? <Trophy size={14} style={{ verticalAlign: "-2px" }} /> : i + 1}
                </td>
                <td style={{ ...tdStyle, textAlign: "left", fontWeight: 600 }}>
                  {s.player}
                  {!s.eligible && <span title="Todavía no llega al mínimo exigido en alguna cancha" style={{ color: COLORS.brass, marginLeft: 4 }}>⚠</span>}
                  {verifiedStatusByPlayer[s.player] && <span title="Todo lo que le suma está verificado" style={{ marginLeft: 4 }}>✅</span>}
                </td>
                {courses.map((c) => {
                  const d = s.courseData[c];
                  return (
                    <td
                      key={c}
                      style={{ ...tdStyle, fontFamily: "'IBM Plex Mono', monospace", cursor: d ? "pointer" : "default" }}
                      onClick={() => d && setDetail({ player: s.player, course: c })}
                    >
                      {d ? (
                        <span title={`${d.count} tarjeta(s) cargada(s), se usan ${d.used} — tocá para ver el detalle`}>
                          {d.subtotal}{!d.complete && <sup style={{ color: COLORS.brass, marginLeft: 2 }}>*</sup>}
                        </span>
                      ) : <span style={{ opacity: 0.25 }}>—</span>}
                    </td>
                  );
                })}
                <td style={{ ...tdStyle, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: COLORS.green700 }}>{s.total}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>

    {detail && (
      <ScoreDetailModal
        detail={detail}
        rounds={roundsInTournament}
        rule={courseRuleByName[detail.course]}
        onVerifyRound={onVerifyRound}
        onClose={() => setDetail(null)}
      />
    )}

    <Card style={{ marginTop: 12 }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 13 }}>
        <div>
          <span style={{ fontWeight: 700, color: COLORS.green700 }}>Vigencia: </span>
          {tournamentInfo && tournamentInfo.startDate
            ? `${formatDate(tournamentInfo.startDate)} al ${formatDate(tournamentInfo.endDate)}`
            : "sin fechas definidas"}
        </div>
        <div>
          <span style={{ fontWeight: 700, color: COLORS.green700 }}>Controllers: </span>
          {tournamentInfo && (tournamentInfo.controllers || []).length > 0
            ? tournamentInfo.controllers.join(" y ")
            : "sin definir"}
        </div>
        <div>
          <span style={{ fontWeight: 700, color: COLORS.green700 }}>Administrador: </span>
          {adminName || "sin nombre configurado"}
        </div>
      </div>
    </Card>

    <Card style={{ marginTop: 12, padding: "14px 16px" }}>
      <div style={{ fontSize: 11.5, color: "rgba(27,36,32,0.55)", lineHeight: 1.5 }}>
        Cada cancha tiene su propia regla (mínimo de tarjetas y si suma las mejores o toma solo la mejor),
        definida al crear el torneo — mirá el encabezado de cada columna. * = todavía no llegó al mínimo
        exigido en esa cancha. ⚠ junto al nombre = no cumple el mínimo en al menos una cancha jugada, o
        todavía no cargó ninguna tarjeta. ✅ junto al nombre = todas las tarjetas que le suman ahora mismo
        fueron verificadas por un controller o el administrador. Tocá el encabezado de cualquier columna
        para ordenar ascendente/descendente por esa cancha o por el total. Tocá cualquier número de la
        tabla para ver el detalle de las tarjetas que lo componen.
      </div>
    </Card>
    </>
  );
}

function DetailRoundRow({ round, isLast, counts, onVerifyRound }) {
  const [verifying, setVerifying] = useState(false);
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function confirmVerify() {
    if (!/^\d{4}$/.test(pin)) return setError("PIN de 4 dígitos.");
    setBusy(true);
    const res = await onVerifyRound(round.id, pin);
    setBusy(false);
    if (!res.ok) return setError(res.msg);
    setVerifying(false);
    setPin("");
  }

  return (
    <div style={{ borderBottom: isLast ? "none" : `1px solid ${COLORS.paperDim}`, opacity: counts ? 1 : 0.5 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0" }}>
        <div>
          <div style={{ fontSize: 13.5, fontWeight: counts ? 700 : 500 }}>{formatDate(round.date)}</div>
          {counts && (
            <div style={{ fontSize: 10.5, color: COLORS.brass, fontWeight: 700, textTransform: "uppercase" }}>Suma al total</div>
          )}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 15, color: COLORS.green700 }}>{round.score}</span>
          {round.verified ? (
            <span title={`Verificada por ${round.verifiedBy}`}>✅</span>
          ) : (
            <button
              onClick={() => { setVerifying((v) => !v); setError(""); setPin(""); }}
              aria-label="Verificar tarjeta"
              title="Verificar tarjeta"
              style={{ background: "none", border: "none", cursor: "pointer", padding: 2, display: "flex", alignItems: "center", fontSize: 15, lineHeight: 1 }}
            >
              🌀
            </button>
          )}
        </div>
      </div>
      {verifying && (
        <div style={{ paddingBottom: 10, display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <input
            type="password" inputMode="numeric" maxLength={4} value={pin}
            onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))}
            placeholder="PIN de controller o admin"
            style={{ ...inputStyle, width: 170 }}
          />
          <button className="glBtn" onClick={confirmVerify} disabled={busy} style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "8px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
            Confirmar
          </button>
          <button onClick={() => setVerifying(false)} style={{ background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "8px 12px", fontSize: 12, cursor: "pointer" }}>
            Cancelar
          </button>
          {error && <p style={{ color: COLORS.danger, fontSize: 11.5, margin: 0, width: "100%" }}>{error}</p>}
        </div>
      )}
    </div>
  );
}

function ScoreDetailModal({ detail, rounds, rule, onVerifyRound, onClose }) {
  const { player, course } = detail;
  const playerRounds = rounds
    .filter((r) => r.player === player && r.course === course)
    .sort((a, b) => a.score - b.score || a.id.localeCompare(b.id));

  const takeCount = rule ? (rule.mode === "mejor" ? 1 : rule.minRequired) : 0;
  const countingIds = new Set(playerRounds.slice(0, takeCount).map((r) => r.id));

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 55 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.paper, color: COLORS.ink, width: "100%", maxWidth: 420, borderRadius: "16px 16px 0 0", padding: 22, maxHeight: "80vh", overflowY: "auto" }}>
        <h2 style={{ margin: "0 0 4px", fontFamily: "'Fraunces', serif", fontSize: 19 }}>{player}</h2>
        <p style={{ margin: "0 0 4px", fontSize: 13, color: "rgba(27,36,32,0.6)" }}>
          {course}
          {rule && (
            <> · {rule.mode === "mejor" ? `toma solo la mejor (mínimo ${rule.minRequired})` : `suma las mejores ${rule.minRequired}`}</>
          )}
        </p>

        {playerRounds.length === 0 ? (
          <p style={{ fontSize: 13.5, marginTop: 12 }}>No hay tarjetas cargadas.</p>
        ) : (
          <div style={{ marginTop: 12 }}>
            {playerRounds.map((r, idx) => (
              <DetailRoundRow
                key={r.id}
                round={r}
                isLast={idx === playerRounds.length - 1}
                counts={countingIds.has(r.id)}
                onVerifyRound={onVerifyRound}
              />
            ))}
          </div>
        )}

        <button onClick={onClose} style={{ marginTop: 18, width: "100%", background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "10px 0", fontWeight: 600, cursor: "pointer" }}>
          Cerrar
        </button>
      </div>
    </div>
  );
}

const thStyle = { padding: "12px 10px", fontSize: 11, letterSpacing: "0.06em", textTransform: "uppercase", fontWeight: 600, textAlign: "center", fontFamily: "'IBM Plex Mono', monospace" };
const tdStyle = { padding: "11px 10px", textAlign: "center" };

// ---------------- Historial ----------------

function Historial({ rounds, onRequestDelete, onViewPhoto, onVerifyRound, onEditScore, tournament, countingRoundIds }) {
  const [filterPlayer, setFilterPlayer] = useState("");
  const [filterCourse, setFilterCourse] = useState("");
  const [scoreSort, setScoreSort] = useState("fecha"); // "fecha" | "asc" | "desc"

  const playerOptions = useMemo(() => Array.from(new Set(rounds.map((r) => r.player))).sort(), [rounds]);
  const courseOptions = useMemo(() => Array.from(new Set(rounds.map((r) => r.course))).sort(), [rounds]);

  const filtered = rounds.filter(
    (r) => (!filterPlayer || r.player === filterPlayer) && (!filterCourse || r.course === filterCourse)
  );

  const sorted = [...filtered].sort((a, b) => {
    if (scoreSort === "asc") return a.score - b.score;
    if (scoreSort === "desc") return b.score - a.score;
    return a.date < b.date ? 1 : -1; // fecha, más reciente primero
  });

  const banner = (
    <div style={{ fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace", color: COLORS.brassLight, fontWeight: 600, marginBottom: 12, textTransform: "uppercase", letterSpacing: "0.05em" }}>
      Viendo: {tournament}
    </div>
  );

  const glossary = (
    <Card style={{ marginBottom: 12, padding: "12px 16px" }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: COLORS.green700, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>
        Qué significa cada símbolo
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 5, fontSize: 12.5 }}>
        <div><strong>◀️</strong> — esta tarjeta es una de las que suman al total ahora mismo en Posiciones.</div>
        <div><strong>✅</strong> — tarjeta ya verificada por un controller o el administrador.</div>
        <div><strong>🌀</strong> — botón para verificar esta tarjeta (solo visible si todavía no está verificada).</div>
        <div><strong>✏️</strong> — corregir el score (solo controller o administrador, no sobre la propia tarjeta; al guardar queda verificada).</div>
      </div>
    </Card>
  );

  const filters = (
    <Card style={{ marginBottom: 12, padding: "12px 16px" }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 110 }}>
          <FieldLabel style={{ marginBottom: 4 }}>Jugador</FieldLabel>
          <Select value={filterPlayer} onChange={setFilterPlayer} placeholder="Todos">
            <option value="">Todos</option>
            {playerOptions.map((p) => <option key={p} value={p}>{p}</option>)}
          </Select>
        </div>
        <div style={{ flex: 1, minWidth: 110 }}>
          <FieldLabel style={{ marginBottom: 4 }}>Cancha</FieldLabel>
          <Select value={filterCourse} onChange={setFilterCourse} placeholder="Todas">
            <option value="">Todas</option>
            {courseOptions.map((c) => <option key={c} value={c}>{c}</option>)}
          </Select>
        </div>
        <div style={{ flex: 1, minWidth: 130 }}>
          <FieldLabel style={{ marginBottom: 4 }}>Orden</FieldLabel>
          <Select value={scoreSort} onChange={setScoreSort} placeholder="Orden">
            <option value="fecha">Más recientes primero</option>
            <option value="asc">Score: menor a mayor</option>
            <option value="desc">Score: mayor a menor</option>
          </Select>
        </div>
      </div>
    </Card>
  );

  if (sorted.length === 0) {
    return (
      <>
        {banner}
        {glossary}
        {filters}
        <Card><p style={{ margin: 0, fontSize: 14 }}>No hay tarjetas que coincidan con el filtro.</p></Card>
      </>
    );
  }

  return (
    <>
      {banner}
      {glossary}
      {filters}
      <Card style={{ padding: 0 }}>
        {sorted.map((r, idx) => (
          <RoundRow
            key={r.id}
            round={r}
            isLast={idx === sorted.length - 1}
            onRequestDelete={onRequestDelete}
            onViewPhoto={onViewPhoto}
            onVerifyRound={onVerifyRound}
            onEditScore={onEditScore}
            counts={countingRoundIds.has(r.id)}
          />
        ))}
      </Card>
      <p style={{ fontSize: 11.5, color: "rgba(247,245,239,0.5)", marginTop: 10, textAlign: "center", lineHeight: 1.6 }}>
        ◀️ = tarjeta que suma al total ahora mismo · ✅ = verificada por un controller o el administrador
      </p>
    </>
  );
}

function RoundRow({ round, isLast, onRequestDelete, onViewPhoto, onVerifyRound, onEditScore, counts }) {
  const [deleting, setDeleting] = useState(false);
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const [verifying, setVerifying] = useState(false);
  const [verifyPinValue, setVerifyPinValue] = useState("");
  const [verifyError, setVerifyError] = useState("");
  const [verifyBusy, setVerifyBusy] = useState(false);

  const [editing, setEditing] = useState(false);
  const [editScoreValue, setEditScoreValue] = useState(String(round.score));
  const [editPinValue, setEditPinValue] = useState("");
  const [editError, setEditError] = useState("");
  const [editBusy, setEditBusy] = useState(false);

  async function confirmDelete() {
    if (!/^\d{4}$/.test(pin)) return setError("PIN de 4 dígitos.");
    setBusy(true);
    const ok = await onRequestDelete(round.id, pin);
    setBusy(false);
    if (!ok) {
      setError("PIN de administrador incorrecto.");
      return;
    }
  }

  async function confirmVerify() {
    if (!/^\d{4}$/.test(verifyPinValue)) return setVerifyError("PIN de 4 dígitos.");
    setVerifyBusy(true);
    const res = await onVerifyRound(round.id, verifyPinValue);
    setVerifyBusy(false);
    if (!res.ok) {
      setVerifyError(res.msg);
      return;
    }
    setVerifying(false);
    setVerifyPinValue("");
  }

  async function confirmEdit() {
    if (!/^\d{4}$/.test(editPinValue)) return setEditError("PIN de 4 dígitos.");
    setEditBusy(true);
    const res = await onEditScore(round.id, editScoreValue, editPinValue);
    setEditBusy(false);
    if (!res.ok) {
      setEditError(res.msg);
      return;
    }
    setEditing(false);
    setEditPinValue("");
  }

  return (
    <div style={{ borderBottom: isLast ? "none" : `1px solid ${COLORS.paperDim}` }}>
      <div className="glRow" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 18px" }}>
        <div>
          <div style={{ fontWeight: 600, fontSize: 14 }}>{round.player}</div>
          <div style={{ fontSize: 12.5, color: "rgba(27,36,32,0.55)", marginTop: 2 }}>
            {round.course} · {formatDate(round.date)}
            {round.handicap != null && ` · hcp ${round.handicap}`}
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 16, color: COLORS.green700 }}>{round.score}</div>
            {counts && <span title="Esta tarjeta suma al total ahora mismo" style={{ fontSize: 14 }}>◀️</span>}
            {round.verified && <span title={`Verificada por ${round.verifiedBy}`} style={{ fontSize: 14 }}>✅</span>}
          </div>
          {round.hasPhoto && (
            <button onClick={() => onViewPhoto(round.id)} aria-label="Ver foto de la tarjeta" style={{ background: "none", border: "none", color: COLORS.green700, cursor: "pointer", padding: 4, display: "flex" }}>
              <Camera size={16} />
            </button>
          )}
          {!round.verified && (
            <button
              onClick={() => { setVerifying((v) => !v); setVerifyError(""); setVerifyPinValue(""); setDeleting(false); setEditing(false); }}
              aria-label="Verificar tarjeta"
              title="Verificar tarjeta"
              style={{ background: "none", border: "none", cursor: "pointer", padding: 4, display: "flex", alignItems: "center", fontSize: 16, lineHeight: 1 }}
            >
              🌀
            </button>
          )}
          <button
            onClick={() => { setEditing((v) => !v); setEditError(""); setEditPinValue(""); setEditScoreValue(String(round.score)); setDeleting(false); setVerifying(false); }}
            aria-label="Corregir score"
            title="Corregir score (solo controller o administrador)"
            style={{ background: "none", border: "none", color: COLORS.brass, cursor: "pointer", padding: 4, display: "flex" }}
          >
            <Pencil size={15} />
          </button>
          <button
            onClick={() => { setDeleting((d) => !d); setError(""); setPin(""); setVerifying(false); setEditing(false); }}
            aria-label="Eliminar tarjeta"
            style={{ background: "none", border: "none", color: "rgba(140,59,46,0.6)", cursor: "pointer", padding: 4 }}
          >
            <Trash2 size={15} />
          </button>
        </div>
      </div>
      {deleting && (
        <div style={{ padding: "0 18px 14px", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <input
            type="password" inputMode="numeric" maxLength={4} value={pin}
            onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))}
            placeholder="PIN de administrador"
            style={{ ...inputStyle, width: 160 }}
          />
          <button className="glBtn" onClick={confirmDelete} disabled={busy} style={{ background: COLORS.danger, color: COLORS.paper, border: "none", borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>
            Confirmar borrado
          </button>
          <button className="glBtn" onClick={() => setDeleting(false)} style={{ background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "9px 14px", fontSize: 12.5, cursor: "pointer" }}>
            Cancelar
          </button>
          {error && <p style={{ color: COLORS.danger, fontSize: 12.5, margin: 0, width: "100%" }}>{error}</p>}
        </div>
      )}
      {verifying && (
        <div style={{ padding: "0 18px 14px", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <input
            type="password" inputMode="numeric" maxLength={4} value={verifyPinValue}
            onChange={(e) => setVerifyPinValue(e.target.value.replace(/\D/g, ""))}
            placeholder="PIN de controller o admin"
            style={{ ...inputStyle, width: 180 }}
          />
          <button className="glBtn" onClick={confirmVerify} disabled={verifyBusy} style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>
            Confirmar verificación
          </button>
          <button className="glBtn" onClick={() => setVerifying(false)} style={{ background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "9px 14px", fontSize: 12.5, cursor: "pointer" }}>
            Cancelar
          </button>
          {verifyError && <p style={{ color: COLORS.danger, fontSize: 12.5, margin: 0, width: "100%" }}>{verifyError}</p>}
        </div>
      )}
      {editing && (
        <div style={{ padding: "0 18px 14px", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <input
            type="number" inputMode="numeric" value={editScoreValue}
            onChange={(e) => setEditScoreValue(e.target.value)}
            placeholder="Score correcto"
            style={{ ...inputStyle, width: 110 }}
          />
          <input
            type="password" inputMode="numeric" maxLength={4} value={editPinValue}
            onChange={(e) => setEditPinValue(e.target.value.replace(/\D/g, ""))}
            placeholder="PIN de controller o admin"
            style={{ ...inputStyle, width: 180 }}
          />
          <button className="glBtn" onClick={confirmEdit} disabled={editBusy} style={{ background: COLORS.brass, color: COLORS.green900, border: "none", borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>
            Guardar corrección
          </button>
          <button className="glBtn" onClick={() => setEditing(false)} style={{ background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "9px 14px", fontSize: 12.5, cursor: "pointer" }}>
            Cancelar
          </button>
          <p style={{ fontSize: 11, color: "rgba(27,36,32,0.5)", margin: "2px 0 0", width: "100%" }}>
            Solo un controller de este torneo o el administrador pueden corregir, y no sobre su propia tarjeta.
            Al guardar, queda marcada como verificada.
          </p>
          {editError && <p style={{ color: COLORS.danger, fontSize: 12.5, margin: 0, width: "100%" }}>{editError}</p>}
        </div>
      )}
    </div>
  );
}

// ---------------- Foto modal ----------------

function PhotoModal({ state, onClose, onClearBroken }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.75)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60, padding: 20 }} onClick={onClose}>
      <div style={{ position: "relative", maxWidth: "100%", maxHeight: "100%" }} onClick={(e) => e.stopPropagation()}>
        <button onClick={onClose} aria-label="Cerrar" style={{ position: "absolute", top: -14, right: -14, background: COLORS.paper, border: "none", borderRadius: "50%", width: 30, height: 30, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", boxShadow: "0 2px 8px rgba(0,0,0,0.4)" }}>
          <X size={16} color={COLORS.ink} />
        </button>
        {state.loading && <div style={{ color: COLORS.paper, display: "flex", alignItems: "center", gap: 8 }}><Loader2 size={18} style={{ animation: "spin 1s linear infinite" }} /> Cargando foto...</div>}
        {!state.loading && state.error && (
          <div style={{ color: COLORS.paper, display: "flex", flexDirection: "column", alignItems: "center", gap: 10, background: "rgba(0,0,0,0.3)", padding: 20, borderRadius: 10, maxWidth: 280 }}>
            <ImageOff size={20} />
            <p style={{ margin: 0, fontSize: 13.5, textAlign: "center" }}>
              Esta foto no quedó guardada correctamente. Podés quitar la referencia para que no vuelva a mostrar
              este error, y volver a cargarla desde la tarjeta si la necesitás.
            </p>
            <button onClick={onClearBroken} className="glBtn" style={{ background: COLORS.brass, color: COLORS.green900, border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>
              Quitar referencia a la foto
            </button>
          </div>
        )}
        {!state.loading && !state.error && (
          <img src={state.url} alt="Tarjeta de golf" style={{ maxWidth: "88vw", maxHeight: "85vh", borderRadius: 10, display: "block" }} />
        )}
      </div>
    </div>
  );
}

// ---------------- Admin gate + panel ----------------

function Estadisticas({ rounds, courses, parByCourse, tournament }) {
  const banner = (
    <div style={{ fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace", color: COLORS.brassLight, fontWeight: 600, marginBottom: 12, textTransform: "uppercase", letterSpacing: "0.05em" }}>
      Viendo: {tournament}
    </div>
  );

  const bestByCourse = courses.map((c) => {
    const inCourse = rounds.filter((r) => r.course === c);
    if (inCourse.length === 0) return { course: c, best: null };
    const best = [...inCourse].sort((a, b) => a.score - b.score || a.id.localeCompare(b.id))[0];
    return { course: c, best };
  });

  return (
    <>
      {banner}
      <Card style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 16px 8px" }}>
          <h3 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontSize: 17 }}>Mejor ronda de cada cancha</h3>
        </div>
        {courses.length === 0 ? (
          <p style={{ padding: "0 16px 16px", margin: 0, fontSize: 14 }}>Este torneo no tiene canchas configuradas.</p>
        ) : (
          <div>
            {bestByCourse.map(({ course, best }, idx) => (
              <div
                key={course}
                style={{
                  display: "flex", justifyContent: "space-between", alignItems: "center",
                  padding: "12px 16px", borderTop: `1px solid ${COLORS.paperDim}`,
                }}
              >
                <div>
                  <div style={{ fontWeight: 600, fontSize: 14 }}>
                    {course}
                    {parByCourse[course] ? <span style={{ opacity: 0.5, fontWeight: 400 }}> · par {parByCourse[course]}</span> : null}
                  </div>
                  {!best && (
                    <div style={{ fontSize: 12.5, color: "rgba(27,36,32,0.4)", marginTop: 2 }}>Sin tarjetas cargadas</div>
                  )}
                </div>
                {best && (
                  <div style={{ textAlign: "right" }}>
                    <div style={{ display: "flex", alignItems: "baseline", gap: 10, justifyContent: "flex-end" }}>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 18, color: COLORS.ink }}>
                        {best.player}
                      </span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 18, color: COLORS.green700 }}>
                        {best.score}
                      </span>
                    </div>
                    {best.handicap != null && (
                      <div style={{ fontSize: 11, color: "rgba(27,36,32,0.5)", marginTop: 2 }}>hcp {best.handicap}</div>
                    )}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </Card>
    </>
  );
}

function TreasuryGate({ authed, onVerifyPin, onClose, children }) {
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  if (authed) return children;

  async function handleLogin() {
    setError("");
    if (!/^\d{4}$/.test(pin)) return setError("Ingresá el PIN de 4 dígitos.");
    setBusy(true);
    const ok = await onVerifyPin(pin);
    setBusy(false);
    if (!ok) return setError("PIN incorrecto — tiene que ser el del administrador o el de un controller de algún torneo.");
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 50 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.paper, color: COLORS.ink, width: "100%", maxWidth: 420, borderRadius: "16px 16px 0 0", padding: 24 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
          <Wallet size={20} color={COLORS.brass} />
          <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, margin: 0 }}>Tesorería</h2>
        </div>
        <p style={{ fontSize: 12.5, color: "rgba(27,36,32,0.6)", margin: "4px 0 16px" }}>
          Acceso restringido: administrador o controller de cualquier torneo.
        </p>
        <input
          type="password" inputMode="numeric" maxLength={4} value={pin}
          onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))}
          placeholder="PIN"
          style={{ ...inputStyle, marginBottom: 10 }}
        />
        {error && <p style={{ color: COLORS.danger, fontSize: 12.5, margin: "0 0 10px" }}>{error}</p>}
        <button className="glBtn" onClick={handleLogin} disabled={busy} style={{ width: "100%", background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "12px 0", fontWeight: 600, cursor: "pointer" }}>
          Entrar
        </button>
        <button onClick={onClose} style={{ marginTop: 10, width: "100%", background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "10px 0", fontWeight: 600, cursor: "pointer" }}>
          Cerrar
        </button>
      </div>
    </div>
  );
}

const TREASURY_CATEGORIES = {
  ingreso: ["Cuota", "Aporte extra", "Otro"],
  egreso: ["Cancha", "Torneo", "Evento", "Otro"],
};

function TreasuryPanel({ movements, onAdd, onRemove, onExport, onClose }) {
  const [type, setType] = useState("ingreso");
  const [date, setDate] = useState(todayStr());
  const [concept, setConcept] = useState("");
  const [category, setCategory] = useState(TREASURY_CATEGORIES.ingreso[0]);
  const [amount, setAmount] = useState("");
  const [registeredBy, setRegisteredBy] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [deleting, setDeleting] = useState(null);

  const totalIngresos = movements.filter((m) => m.type === "ingreso").reduce((a, m) => a + Number(m.amount || 0), 0);
  const totalEgresos = movements.filter((m) => m.type === "egreso").reduce((a, m) => a + Number(m.amount || 0), 0);
  const saldo = totalIngresos - totalEgresos;
  const sorted = [...movements].sort((a, b) => (a.date < b.date ? 1 : -1));

  function fmtMoney(n) {
    return "$" + Number(n || 0).toLocaleString("es-AR");
  }

  function handleTypeChange(t) {
    setType(t);
    setCategory(TREASURY_CATEGORIES[t][0]);
  }

  async function handleAdd() {
    setError("");
    if (!concept.trim()) return setError("Ingresá un concepto.");
    const amountNum = Number(amount);
    if (!amount || isNaN(amountNum) || amountNum <= 0) return setError("Ingresá un monto válido.");
    if (!registeredBy.trim()) return setError("Ingresá tu nombre (quién registra el movimiento).");
    setBusy(true);
    const res = await onAdd({ type, date, concept: concept.trim(), category, amount: amountNum, registeredBy: registeredBy.trim() });
    setBusy(false);
    if (!res.ok) return setError(res.msg);
    setConcept("");
    setAmount("");
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 50 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.paper, color: COLORS.ink, width: "100%", maxWidth: 480, borderRadius: "16px 16px 0 0", padding: 22, maxHeight: "88vh", overflowY: "auto" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
          <Wallet size={20} color={COLORS.brass} />
          <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, margin: 0 }}>Tesorería</h2>
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 18 }}>
          <div style={{ flex: 1, background: COLORS.paperDim, borderRadius: 10, padding: "10px 12px" }}>
            <div style={{ fontSize: 10.5, textTransform: "uppercase", opacity: 0.6, fontFamily: "'IBM Plex Mono', monospace" }}>Ingresos</div>
            <div style={{ fontSize: 16, fontWeight: 700, color: COLORS.green700 }}>{fmtMoney(totalIngresos)}</div>
          </div>
          <div style={{ flex: 1, background: COLORS.paperDim, borderRadius: 10, padding: "10px 12px" }}>
            <div style={{ fontSize: 10.5, textTransform: "uppercase", opacity: 0.6, fontFamily: "'IBM Plex Mono', monospace" }}>Egresos</div>
            <div style={{ fontSize: 16, fontWeight: 700, color: COLORS.danger }}>{fmtMoney(totalEgresos)}</div>
          </div>
          <div style={{ flex: 1, background: COLORS.green700, borderRadius: 10, padding: "10px 12px" }}>
            <div style={{ fontSize: 10.5, textTransform: "uppercase", opacity: 0.75, color: COLORS.paper, fontFamily: "'IBM Plex Mono', monospace" }}>Saldo</div>
            <div style={{ fontSize: 16, fontWeight: 700, color: COLORS.paper }}>{fmtMoney(saldo)}</div>
          </div>
        </div>

        <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 15, margin: "0 0 10px" }}>Registrar movimiento</h3>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button
            onClick={() => handleTypeChange("ingreso")}
            style={{
              flex: 1, padding: "9px 0", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13,
              border: type === "ingreso" ? "none" : `1px solid ${COLORS.paperDim}`,
              background: type === "ingreso" ? COLORS.green700 : "none",
              color: type === "ingreso" ? COLORS.paper : COLORS.ink,
            }}
          >
            Ingreso
          </button>
          <button
            onClick={() => handleTypeChange("egreso")}
            style={{
              flex: 1, padding: "9px 0", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13,
              border: type === "egreso" ? "none" : `1px solid ${COLORS.paperDim}`,
              background: type === "egreso" ? COLORS.danger : "none",
              color: type === "egreso" ? COLORS.paper : COLORS.ink,
            }}
          >
            Egreso
          </button>
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 130 }}>
            <FieldLabel style={{ marginBottom: 4 }}>Fecha</FieldLabel>
            <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={inputStyle} />
          </div>
          <div style={{ flex: 1, minWidth: 130 }}>
            <FieldLabel style={{ marginBottom: 4 }}>Categoría</FieldLabel>
            <Select value={category} onChange={setCategory}>
              {TREASURY_CATEGORIES[type].map((c) => <option key={c} value={c}>{c}</option>)}
            </Select>
          </div>
        </div>

        <FieldLabel style={{ marginBottom: 4 }}>Concepto</FieldLabel>
        <input value={concept} onChange={(e) => setConcept(e.target.value)} placeholder="Ej: cuota agosto, alquiler cancha Norte..." style={{ ...inputStyle, marginBottom: 10 }} />

        <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 100 }}>
            <FieldLabel style={{ marginBottom: 4 }}>Monto</FieldLabel>
            <input type="number" inputMode="decimal" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" style={inputStyle} />
          </div>
          <div style={{ flex: 1, minWidth: 140 }}>
            <FieldLabel style={{ marginBottom: 4 }}>Registrado por (vos)</FieldLabel>
            <input value={registeredBy} onChange={(e) => setRegisteredBy(e.target.value)} placeholder="Tu nombre" style={inputStyle} />
          </div>
        </div>

        {error && <p style={{ color: COLORS.danger, fontSize: 12.5, margin: "0 0 10px" }}>{error}</p>}
        <button className="glBtn" onClick={handleAdd} disabled={busy} style={{ width: "100%", background: COLORS.brass, color: COLORS.green900, border: "none", borderRadius: 8, padding: "12px 0", fontWeight: 700, cursor: "pointer", marginBottom: 18 }}>
          Registrar {type === "ingreso" ? "ingreso" : "egreso"}
        </button>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 15, margin: 0 }}>Movimientos ({movements.length})</h3>
          <button onClick={onExport} style={{ background: "none", border: `1px solid ${COLORS.green700}`, color: COLORS.green700, borderRadius: 6, padding: "5px 10px", fontSize: 11.5, fontWeight: 600, cursor: "pointer" }}>
            Exportar CSV
          </button>
        </div>

        {sorted.length === 0 ? (
          <p style={{ fontSize: 13, opacity: 0.6 }}>Todavía no hay movimientos cargados.</p>
        ) : (
          <div>
            {sorted.map((m, idx) => (
              <div key={m.id} style={{ borderBottom: idx < sorted.length - 1 ? `1px solid ${COLORS.paperDim}` : "none", padding: "10px 0" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontSize: 13.5, fontWeight: 600 }}>{m.concept}</div>
                    <div style={{ fontSize: 11.5, opacity: 0.6 }}>
                      {formatDate(m.date)} · {m.category} · {m.registeredBy}
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 14, color: m.type === "ingreso" ? COLORS.green700 : COLORS.danger }}>
                      {m.type === "ingreso" ? "+" : "−"}{fmtMoney(m.amount)}
                    </div>
                    <button onClick={() => setDeleting(deleting === m.id ? null : m.id)} style={{ background: "none", border: "none", color: "rgba(140,59,46,0.7)", cursor: "pointer" }}>
                      <Trash2 size={14} />
                    </button>
                  </div>
                </div>
                {deleting === m.id && (
                  <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
                    <button
                      onClick={() => { onRemove(m.id); setDeleting(null); }}
                      style={{ flex: 1, background: COLORS.danger, color: "#fff", border: "none", borderRadius: 6, padding: "7px 0", fontSize: 12, fontWeight: 600, cursor: "pointer" }}
                    >
                      Confirmar borrado
                    </button>
                    <button
                      onClick={() => setDeleting(null)}
                      style={{ flex: 1, background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 6, padding: "7px 0", fontSize: 12, cursor: "pointer" }}
                    >
                      Cancelar
                    </button>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        <button onClick={onClose} style={{ marginTop: 18, width: "100%", background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "10px 0", fontWeight: 600, cursor: "pointer" }}>
          Cerrar
        </button>
      </div>
    </div>
  );
}

function AdminGate({ hasAdminPin, authed, onSetupPin, onVerifyPin, onClose, children }) {
  const [pin, setPin] = useState("");
  const [pinConfirm, setPinConfirm] = useState("");
  const [adminName, setAdminName] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  if (authed) return children;

  async function handleSetup() {
    setError("");
    if (!adminName.trim()) return setError("Ingresá tu nombre.");
    if (!/^\d{4}$/.test(pin)) return setError("El PIN debe tener 4 dígitos.");
    if (pin !== pinConfirm) return setError("Los PIN no coinciden.");
    setBusy(true);
    try {
      const res = await onSetupPin(pin, adminName);
      if (!res.ok) setError(res.msg || "No se pudo configurar el PIN. Probá de nuevo.");
    } catch (e) {
      setError("No se pudo configurar el PIN. Probá de nuevo.");
    } finally {
      setBusy(false);
    }
  }

  async function handleLogin() {
    setError("");
    if (!/^\d{4}$/.test(pin)) return setError("Ingresá el PIN de 4 dígitos.");
    setBusy(true);
    const ok = await onVerifyPin(pin);
    setBusy(false);
    if (!ok) setError("PIN incorrecto.");
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 50 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.paper, color: COLORS.ink, width: "100%", maxWidth: 420, borderRadius: "16px 16px 0 0", padding: 24 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
          <KeyRound size={18} color={COLORS.green700} />
          <h2 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontSize: 19 }}>
            {hasAdminPin ? "PIN de administrador" : "Configurá el PIN de administrador"}
          </h2>
        </div>
        <p style={{ margin: "0 0 16px", fontSize: 13, color: "rgba(27,36,32,0.6)" }}>
          {hasAdminPin
            ? "Protege aprobaciones, canchas y borrado de tarjetas de otros."
            : "Es la primera vez que entrás. Este PIN va a proteger el panel de administrador de acá en adelante — guardalo bien."}
        </p>

        {!hasAdminPin && (
          <>
            <FieldLabel>Tu nombre (aparece como administrador en Posiciones)</FieldLabel>
            <input value={adminName} onChange={(e) => setAdminName(e.target.value)} placeholder="Tu nombre" style={inputStyle} />
          </>
        )}

        <FieldLabel style={{ marginTop: hasAdminPin ? 0 : 14 }}>PIN</FieldLabel>
        <input type="password" inputMode="numeric" maxLength={4} value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} placeholder="4 dígitos" style={inputStyle} />

        {!hasAdminPin && (
          <>
            <FieldLabel style={{ marginTop: 14 }}>Repetí el PIN</FieldLabel>
            <input type="password" inputMode="numeric" maxLength={4} value={pinConfirm} onChange={(e) => setPinConfirm(e.target.value.replace(/\D/g, ""))} placeholder="4 dígitos" style={inputStyle} />
          </>
        )}

        {error && <p style={{ color: COLORS.danger, fontSize: 13, marginTop: 12, marginBottom: 0, fontWeight: 500 }}>{error}</p>}

        <div style={{ display: "flex", gap: 8, marginTop: 18 }}>
          <button className="glBtn" onClick={hasAdminPin ? handleLogin : handleSetup} disabled={busy} style={{ ...primaryBtnStyle, flex: 1 }}>
            {hasAdminPin ? "Entrar" : "Configurar"}
          </button>
          <button onClick={onClose} style={{ background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 10, padding: "0 16px", fontWeight: 600, cursor: "pointer" }}>
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );
}

function AdminPanel({
  pendingUsers, approvedUsers, courses, tournaments, roundsCountByCourse,
  onApprove, onReject, onRemoveUser, onSetUserPin,
  onAddUserDirect, onAddCourse, onRemoveCourse, onEditCoursePar,
  onAddTournament, onToggleTournament, onRemoveTournament, onEditTournamentDates,
  onApproveParticipant, onRejectParticipant, onResetAllData, onRefreshData,
  onExportHistorial, onExportPosiciones, activeTournamentName, onClose,
}) {
  const [name, setName] = useState("");
  const [newUserPin, setNewUserPin] = useState("");
  const [addError, setAddError] = useState("");
  const [pinEditing, setPinEditing] = useState(null);
  const [pinValue, setPinValue] = useState("");
  const [pinError, setPinError] = useState("");

  const [courseName, setCourseName] = useState("");
  const [coursePar, setCoursePar] = useState("");

  const [courseEditing, setCourseEditing] = useState(null);
  const [courseParValue, setCourseParValue] = useState("");
  const [courseAdminPin, setCourseAdminPin] = useState("");
  const [courseError, setCourseError] = useState("");

  const [courseDeleting, setCourseDeleting] = useState(null);
  const [deleteCoursePin, setDeleteCoursePin] = useState("");
  const [deleteCourseError, setDeleteCourseError] = useState("");
  const [deleteCourseBusy, setDeleteCourseBusy] = useState(false);

  const [tournamentName, setTournamentName] = useState("");
  const [tournamentStart, setTournamentStart] = useState("");
  const [tournamentEnd, setTournamentEnd] = useState("");
  const [controller1, setController1] = useState("");
  const [controller2, setController2] = useState("");
  const [newTournamentCourses, setNewTournamentCourses] = useState({}); // { [courseName]: {checked, minRequired, mode} }
  const [tournamentAddError, setTournamentAddError] = useState("");

  const [tournamentEditing, setTournamentEditing] = useState(null);
  const [editStart, setEditStart] = useState("");
  const [editEnd, setEditEnd] = useState("");
  const [tournamentAdminPin, setTournamentAdminPin] = useState("");
  const [tournamentEditError, setTournamentEditError] = useState("");

  async function handleAddDirect() {
    setAddError("");
    const res = await onAddUserDirect(name, newUserPin);
    if (!res.ok) return setAddError(res.msg);
    setName(""); setNewUserPin("");
  }

  async function handleSetPin() {
    setPinError("");
    const res = await onSetUserPin(pinEditing, pinValue);
    if (!res.ok) return setPinError(res.msg);
    setPinEditing(null);
    setPinValue("");
  }

  async function handleAddCourse() {
    if (!courseName.trim() || !coursePar) return;
    await onAddCourse(courseName, coursePar);
    setCourseName(""); setCoursePar("");
  }

  async function handleSaveCoursePar(courseNm) {
    setCourseError("");
    const res = await onEditCoursePar(courseNm, courseParValue, courseAdminPin);
    if (!res.ok) return setCourseError(res.msg);
    setCourseEditing(null); setCourseParValue(""); setCourseAdminPin("");
  }

  async function handleDeleteCourse(courseNm) {
    setDeleteCourseError("");
    setDeleteCourseBusy(true);
    const res = await onRemoveCourse(courseNm, deleteCoursePin);
    setDeleteCourseBusy(false);
    if (!res.ok) return setDeleteCourseError(res.msg);
    setCourseDeleting(null); setDeleteCoursePin("");
  }

  async function handleAddTournament() {
    setTournamentAddError("");
    const courseConfig = Object.entries(newTournamentCourses)
      .filter(([, v]) => v.checked)
      .map(([courseName, v]) => ({
        course: courseName,
        minRequired: Number(v.minRequired) || 0,
        mode: v.mode || "suma",
      }));
    const res = await onAddTournament(tournamentName, tournamentStart, tournamentEnd, controller1, controller2, courseConfig);
    if (!res.ok) return setTournamentAddError(res.msg);
    setTournamentName(""); setTournamentStart(""); setTournamentEnd(""); setController1(""); setController2("");
    setNewTournamentCourses({});
  }

  async function handleSaveTournamentDates(tName) {
    setTournamentEditError("");
    const res = await onEditTournamentDates(tName, editStart, editEnd, tournamentAdminPin);
    if (!res.ok) return setTournamentEditError(res.msg);
    setTournamentEditing(null); setEditStart(""); setEditEnd(""); setTournamentAdminPin("");
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 50 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: COLORS.paper, color: COLORS.ink, width: "100%", maxWidth: 480, borderRadius: "16px 16px 0 0", padding: 22, maxHeight: "85vh", overflowY: "auto" }}>

        <button
          onClick={onRefreshData}
          className="glBtn"
          style={{
            display: "flex", alignItems: "center", justifyContent: "center", gap: 8, width: "100%",
            background: COLORS.paperDim, color: COLORS.green700, border: "none", borderRadius: 8,
            padding: "10px 0", fontWeight: 700, cursor: "pointer", fontSize: 13, marginBottom: 18,
          }}
        >
          <RefreshCw size={15} /> Actualizar datos (por si alguien cargó algo desde otro dispositivo)
        </button>

        {pendingUsers.length > 0 && (
          <>
            <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, margin: "0 0 14px" }}>Solicitudes pendientes</h2>
            <div style={{ marginBottom: 24 }}>
              {pendingUsers.map((u) => (
                <div key={u.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: `1px solid ${COLORS.paperDim}`, gap: 10 }}>
                  <div>
                    <div style={{ fontSize: 14, fontWeight: 600 }}>{u.name}</div>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button onClick={() => onApprove(u.name)} style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 6, padding: "6px 10px", cursor: "pointer" }} aria-label="Aprobar">
                      <Check size={14} />
                    </button>
                    <button onClick={() => onReject(u.name)} style={{ background: "none", border: `1px solid ${COLORS.danger}`, color: COLORS.danger, borderRadius: 6, padding: "6px 10px", cursor: "pointer" }} aria-label="Rechazar">
                      <Ban size={14} />
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </>
        )}

        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, margin: "0 0 6px" }}>Jugadores aprobados</h2>
        <p style={{ fontSize: 12, color: "rgba(27,36,32,0.55)", margin: "0 0 10px" }}>
          Al agregar directo, vos elegís el PIN inicial y se lo pasás a esa persona por fuera de la app
          (WhatsApp, etc.) — así nadie más puede apropiarse de su nombre antes que ella.
        </p>
        <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
          <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Nombre" style={{ ...inputStyle, flex: 2, minWidth: 120 }} />
          <input
            type="password" inputMode="numeric" maxLength={4} value={newUserPin}
            onChange={(e) => setNewUserPin(e.target.value.replace(/\D/g, ""))}
            placeholder="PIN inicial"
            style={{ ...inputStyle, flex: 1, minWidth: 90 }}
          />
        </div>
        {addError && <p style={{ color: COLORS.danger, fontSize: 12.5, margin: "0 0 8px" }}>{addError}</p>}
        <button
          className="glBtn"
          onClick={handleAddDirect}
          style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "9px 16px", fontWeight: 600, cursor: "pointer", marginBottom: 16 }}
        >
          Agregar con este PIN
        </button>

        {approvedUsers.length === 0 ? (
          <p style={{ fontSize: 13.5, opacity: 0.6 }}>Sin jugadores todavía.</p>
        ) : (
          <div style={{ marginBottom: 24 }}>
            {approvedUsers.map((u) => (
              <div key={u.name}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: pinEditing === u.name ? "none" : `1px solid ${COLORS.paperDim}`, gap: 8 }}>
                  <div>
                    <span style={{ fontSize: 14, fontWeight: 500 }}>{u.name}</span>
                    {!u.pinHash && <span style={{ fontSize: 11, color: COLORS.danger, marginLeft: 6 }}>sin PIN — no puede cargar</span>}
                  </div>
                  <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                    <button
                      onClick={() => { setPinEditing(pinEditing === u.name ? null : u.name); setPinValue(""); setPinError(""); }}
                      title="Asignar / cambiar PIN"
                      style={{ background: "none", border: "none", color: u.pinHash ? COLORS.brass : COLORS.danger, cursor: "pointer", display: "flex" }}
                    >
                      <KeyRound size={15} />
                    </button>
                    <button onClick={() => onRemoveUser(u.name)} style={{ background: "none", border: "none", color: "rgba(140,59,46,0.7)", cursor: "pointer" }}>
                      <Trash2 size={15} />
                    </button>
                  </div>
                </div>
                {pinEditing === u.name && (
                  <div style={{ padding: "0 0 14px", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", borderBottom: `1px solid ${COLORS.paperDim}` }}>
                    <input
                      type="password" inputMode="numeric" maxLength={4} value={pinValue}
                      onChange={(e) => setPinValue(e.target.value.replace(/\D/g, ""))}
                      placeholder="Nuevo PIN (4 dígitos)"
                      style={{ ...inputStyle, width: 160 }}
                    />
                    <button className="glBtn" onClick={handleSetPin} style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>
                      Guardar PIN
                    </button>
                    {pinError && <p style={{ color: COLORS.danger, fontSize: 12, margin: 0, width: "100%" }}>{pinError}</p>}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, margin: "0 0 14px" }}>Canchas</h2>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          <input value={courseName} onChange={(e) => setCourseName(e.target.value)} placeholder="Nombre de la cancha" style={{ ...inputStyle, flex: 2 }} />
          <input type="number" inputMode="numeric" value={coursePar} onChange={(e) => setCoursePar(e.target.value)} placeholder="Par" style={{ ...inputStyle, flex: 1 }} />
          <button
            className="glBtn"
            onClick={handleAddCourse}
            style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "0 16px", fontWeight: 600, cursor: "pointer" }}
          >
            Agregar
          </button>
        </div>

        {courses.length === 0 ? (
          <p style={{ fontSize: 13.5, opacity: 0.6 }}>Sin canchas todavía.</p>
        ) : (
          <div>
            {courses.map((c) => (
              <div key={c.name}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: (courseEditing === c.name || courseDeleting === c.name) ? "none" : `1px solid ${COLORS.paperDim}` }}>
                  <span style={{ fontSize: 14, fontWeight: 500 }}>{c.name} <span style={{ opacity: 0.5, fontWeight: 400 }}>· par {c.par}</span></span>
                  <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                    <button
                      onClick={() => { setCourseEditing(courseEditing === c.name ? null : c.name); setCourseParValue(String(c.par || "")); setCourseAdminPin(""); setCourseError(""); setCourseDeleting(null); }}
                      style={{ background: "none", border: "none", color: COLORS.green700, cursor: "pointer", fontSize: 11.5, fontWeight: 700 }}
                    >
                      Editar
                    </button>
                    <button
                      onClick={() => { setCourseDeleting(courseDeleting === c.name ? null : c.name); setDeleteCoursePin(""); setDeleteCourseError(""); setCourseEditing(null); }}
                      style={{ background: "none", border: "none", color: "rgba(140,59,46,0.7)", cursor: "pointer" }}
                    >
                      <Trash2 size={15} />
                    </button>
                  </div>
                </div>
                {courseEditing === c.name && (
                  <div style={{ padding: "0 0 14px", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", borderBottom: `1px solid ${COLORS.paperDim}` }}>
                    <input
                      type="number" inputMode="numeric" value={courseParValue}
                      onChange={(e) => setCourseParValue(e.target.value)}
                      placeholder="Nuevo par"
                      style={{ ...inputStyle, width: 100 }}
                    />
                    <input
                      type="password" inputMode="numeric" maxLength={4} value={courseAdminPin}
                      onChange={(e) => setCourseAdminPin(e.target.value.replace(/\D/g, ""))}
                      placeholder="PIN admin"
                      style={{ ...inputStyle, width: 110 }}
                    />
                    <button className="glBtn" onClick={() => handleSaveCoursePar(c.name)} style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>
                      Guardar
                    </button>
                    {courseError && <p style={{ color: COLORS.danger, fontSize: 12, margin: 0, width: "100%" }}>{courseError}</p>}
                  </div>
                )}
                {courseDeleting === c.name && (
                  <div style={{ padding: "0 0 14px", borderBottom: `1px solid ${COLORS.paperDim}` }}>
                    <p style={{ fontSize: 12.5, color: COLORS.danger, margin: "0 0 8px" }}>
                      {(roundsCountByCourse[c.name] || 0) > 0
                        ? `Esto borra la cancha, las ${roundsCountByCourse[c.name]} tarjeta(s) cargada(s) en ella (de todos los torneos), y la saca de la configuración de cualquier torneo que la tuviera asignada. No se puede deshacer.`
                        : "Esta cancha no tiene tarjetas cargadas. Se borra sin más."}
                    </p>
                    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                      <input
                        type="password" inputMode="numeric" maxLength={4} value={deleteCoursePin}
                        onChange={(e) => setDeleteCoursePin(e.target.value.replace(/\D/g, ""))}
                        placeholder="PIN admin"
                        style={{ ...inputStyle, width: 110 }}
                      />
                      <button
                        className="glBtn"
                        onClick={() => handleDeleteCourse(c.name)}
                        disabled={deleteCourseBusy}
                        style={{ background: COLORS.danger, color: COLORS.paper, border: "none", borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}
                      >
                        Confirmar borrado
                      </button>
                      <button
                        onClick={() => { setCourseDeleting(null); setDeleteCoursePin(""); setDeleteCourseError(""); }}
                        style={{ background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "9px 14px", fontSize: 12.5, cursor: "pointer" }}
                      >
                        Cancelar
                      </button>
                    </div>
                    {deleteCourseError && <p style={{ color: COLORS.danger, fontSize: 12, margin: "8px 0 0" }}>{deleteCourseError}</p>}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, margin: "24px 0 6px" }}>Torneos</h2>
        <p style={{ fontSize: 12, color: "rgba(27,36,32,0.55)", margin: "0 0 10px" }}>
          Solo los torneos habilitados aparecen en el panel principal. El nombre no se puede editar después de
          creado (evita romper el historial ya cargado con ese nombre) — si te equivocaste, borralo y creá uno
          nuevo antes de que tenga tarjetas cargadas.
        </p>
        <div style={{ display: "flex", gap: 8, marginBottom: 4, flexWrap: "wrap" }}>
          <input
            value={tournamentName}
            onChange={(e) => setTournamentName(e.target.value)}
            placeholder="Nombre del torneo"
            style={{ ...inputStyle, flex: 1, minWidth: 140 }}
          />
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 130 }}>
            <FieldLabel style={{ marginBottom: 4 }}>Tarjetas válidas desde</FieldLabel>
            <input type="date" value={tournamentStart} onChange={(e) => setTournamentStart(e.target.value)} style={inputStyle} />
          </div>
          <div style={{ flex: 1, minWidth: 130 }}>
            <FieldLabel style={{ marginBottom: 4 }}>Tarjetas válidas hasta</FieldLabel>
            <input type="date" value={tournamentEnd} onChange={(e) => setTournamentEnd(e.target.value)} style={inputStyle} />
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 130 }}>
            <FieldLabel style={{ marginBottom: 4 }}>Controller 1</FieldLabel>
            <Select value={controller1} onChange={setController1} placeholder="Elegí un jugador">
              {approvedUsers.map((u) => <option key={u.name} value={u.name}>{u.name}</option>)}
            </Select>
          </div>
          <div style={{ flex: 1, minWidth: 130 }}>
            <FieldLabel style={{ marginBottom: 4 }}>Controller 2</FieldLabel>
            <Select value={controller2} onChange={setController2} placeholder="Elegí un jugador">
              {approvedUsers.map((u) => <option key={u.name} value={u.name}>{u.name}</option>)}
            </Select>
          </div>
        </div>

        <FieldLabel style={{ marginBottom: 4 }}>Canchas de este torneo</FieldLabel>
        <p style={{ fontSize: 11.5, color: "rgba(27,36,32,0.55)", margin: "0 0 8px" }}>
          Elegí cuáles se juegan, cuántas veces como mínimo cada una, y si se suman las mejores o se toma solo la mejor.
        </p>
        {courses.length === 0 ? (
          <p style={{ fontSize: 13, color: COLORS.danger, marginBottom: 12 }}>
            Todavía no hay canchas cargadas — agregá al menos una en la sección Canchas antes de crear un torneo.
          </p>
        ) : (
          <div style={{ marginBottom: 12 }}>
            {courses.map((c) => {
              const cfg = newTournamentCourses[c.name] || { checked: false, minRequired: "2", mode: "suma" };
              return (
                <div key={c.name} style={{ border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "8px 10px", marginBottom: 6 }}>
                  <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, fontWeight: 500, cursor: "pointer" }}>
                    <input
                      type="checkbox"
                      checked={!!cfg.checked}
                      onChange={(e) =>
                        setNewTournamentCourses((prev) => ({ ...prev, [c.name]: { ...cfg, checked: e.target.checked } }))
                      }
                    />
                    {c.name} <span style={{ opacity: 0.5, fontWeight: 400 }}>· par {c.par}</span>
                  </label>
                  {cfg.checked && (
                    <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
                      <div style={{ flex: 1, minWidth: 110 }}>
                        <FieldLabel style={{ marginBottom: 4 }}>Mínimo a jugar</FieldLabel>
                        <input
                          type="number" inputMode="numeric" min={1} value={cfg.minRequired}
                          onChange={(e) =>
                            setNewTournamentCourses((prev) => ({ ...prev, [c.name]: { ...cfg, minRequired: e.target.value } }))
                          }
                          style={inputStyle}
                        />
                      </div>
                      <div style={{ flex: 1, minWidth: 140 }}>
                        <FieldLabel style={{ marginBottom: 4 }}>Modalidad</FieldLabel>
                        <Select
                          value={cfg.mode}
                          onChange={(v) => setNewTournamentCourses((prev) => ({ ...prev, [c.name]: { ...cfg, mode: v } }))}
                          placeholder="Elegí"
                        >
                          <option value="suma">Suma las mejores</option>
                          <option value="mejor">Toma solo la mejor</option>
                        </Select>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {tournamentAddError && <p style={{ color: COLORS.danger, fontSize: 12.5, margin: "0 0 8px" }}>{tournamentAddError}</p>}
        <button
          className="glBtn"
          onClick={handleAddTournament}
          style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "9px 16px", fontWeight: 600, cursor: "pointer", marginBottom: 16 }}
        >
          Crear torneo
        </button>

        {tournaments.length === 0 ? (
          <p style={{ fontSize: 13.5, opacity: 0.6 }}>Sin torneos todavía.</p>
        ) : (
          <div>
            {tournaments.map((t) => (
              <div key={t.name}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: tournamentEditing === t.name ? "none" : `1px solid ${COLORS.paperDim}` }}>
                  <div>
                    <span style={{ fontSize: 14, fontWeight: 500 }}>
                      {t.name}
                      <span style={{ marginLeft: 8, fontSize: 11, fontWeight: 700, color: t.enabled ? COLORS.green700 : "rgba(27,36,32,0.4)" }}>
                        {t.enabled ? "ON" : "OFF"}
                      </span>
                    </span>
                    <div style={{ fontSize: 11.5, opacity: 0.55 }}>
                      {t.startDate ? `${formatDate(t.startDate)} → ${formatDate(t.endDate)}` : "sin fechas"}
                    </div>
                    <div style={{ fontSize: 11.5, opacity: 0.55 }}>
                      Controllers: {(t.controllers || []).join(" · ") || "sin definir"}
                    </div>
                    <div style={{ fontSize: 11.5, opacity: 0.55 }}>
                      Canchas: {(t.courseConfig || []).length > 0
                        ? t.courseConfig.map((cc) => `${cc.course} (${cc.mode === "mejor" ? "mejor de" : "mejores"} ${cc.minRequired})`).join(" · ")
                        : "sin definir"}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
                    <button
                      onClick={() => {
                        setTournamentEditing(tournamentEditing === t.name ? null : t.name);
                        setEditStart(t.startDate || "");
                        setEditEnd(t.endDate || "");
                        setTournamentAdminPin("");
                        setTournamentEditError("");
                      }}
                      style={{ background: "none", border: "none", color: COLORS.green700, cursor: "pointer", fontSize: 11.5, fontWeight: 700 }}
                    >
                      Editar
                    </button>
                    <Switch on={t.enabled} onClick={() => onToggleTournament(t.name)} />
                    <button onClick={() => onRemoveTournament(t.name)} style={{ background: "none", border: "none", color: "rgba(140,59,46,0.7)", cursor: "pointer" }}>
                      <Trash2 size={15} />
                    </button>
                  </div>
                </div>
                {tournamentEditing === t.name && (
                  <div style={{ padding: "0 0 14px", borderBottom: `1px solid ${COLORS.paperDim}` }}>
                    <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
                      <div style={{ flex: 1, minWidth: 120 }}>
                        <FieldLabel style={{ marginBottom: 4 }}>Desde</FieldLabel>
                        <input type="date" value={editStart} onChange={(e) => setEditStart(e.target.value)} style={inputStyle} />
                      </div>
                      <div style={{ flex: 1, minWidth: 120 }}>
                        <FieldLabel style={{ marginBottom: 4 }}>Hasta</FieldLabel>
                        <input type="date" value={editEnd} onChange={(e) => setEditEnd(e.target.value)} style={inputStyle} />
                      </div>
                    </div>
                    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                      <input
                        type="password" inputMode="numeric" maxLength={4} value={tournamentAdminPin}
                        onChange={(e) => setTournamentAdminPin(e.target.value.replace(/\D/g, ""))}
                        placeholder="PIN admin"
                        style={{ ...inputStyle, width: 110 }}
                      />
                      <button className="glBtn" onClick={() => handleSaveTournamentDates(t.name)} style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>
                        Guardar fechas
                      </button>
                    </div>
                    {tournamentEditError && <p style={{ color: COLORS.danger, fontSize: 12, margin: "8px 0 0" }}>{tournamentEditError}</p>}
                  </div>
                )}
                {(t.participants || []).some((p) => p.status === "pendiente") && (
                  <div style={{ padding: "8px 0 14px", borderBottom: `1px solid ${COLORS.paperDim}` }}>
                    <div style={{ fontSize: 11.5, fontWeight: 700, color: COLORS.brass, marginBottom: 6 }}>
                      Postulaciones pendientes a "{t.name}"
                    </div>
                    {(t.participants || [])
                      .filter((p) => p.status === "pendiente")
                      .map((p) => (
                        <div key={p.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0" }}>
                          <span style={{ fontSize: 13.5 }}>{p.name}</span>
                          <div style={{ display: "flex", gap: 6 }}>
                            <button
                              onClick={() => onApproveParticipant(t.name, p.name)}
                              style={{ background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 6, padding: "5px 9px", cursor: "pointer" }}
                              aria-label="Aceptar"
                            >
                              <Check size={13} />
                            </button>
                            <button
                              onClick={() => onRejectParticipant(t.name, p.name)}
                              style={{ background: "none", border: `1px solid ${COLORS.danger}`, color: COLORS.danger, borderRadius: 6, padding: "5px 9px", cursor: "pointer" }}
                              aria-label="Rechazar"
                            >
                              <Ban size={13} />
                            </button>
                          </div>
                        </div>
                      ))}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        <div style={{ marginTop: 24, paddingTop: 16, borderTop: `1px dashed ${COLORS.paperDim}` }}>
          <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 18, margin: "0 0 6px" }}>Exportar datos</h2>
          <p style={{ fontSize: 11.5, color: "rgba(27,36,32,0.55)", margin: "0 0 10px" }}>
            Descarga un CSV del torneo activo ("{activeTournamentName || "ninguno seleccionado"}"). Abrilo con
            Excel, Google Sheets o lo que uses — la app no sincroniza automáticamente con nada externo.
          </p>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button
              onClick={onExportHistorial}
              disabled={!activeTournamentName}
              style={{ flex: 1, minWidth: 140, background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "10px 0", fontWeight: 600, cursor: activeTournamentName ? "pointer" : "not-allowed", opacity: activeTournamentName ? 1 : 0.5, fontSize: 13 }}
            >
              Exportar Historial (CSV)
            </button>
            <button
              onClick={onExportPosiciones}
              disabled={!activeTournamentName}
              style={{ flex: 1, minWidth: 140, background: COLORS.green700, color: COLORS.paper, border: "none", borderRadius: 8, padding: "10px 0", fontWeight: 600, cursor: activeTournamentName ? "pointer" : "not-allowed", opacity: activeTournamentName ? 1 : 0.5, fontSize: 13 }}
            >
              Exportar Posiciones (CSV)
            </button>
          </div>
        </div>


        <ResetAllSection onResetAllData={onResetAllData} />

        <button onClick={onClose} style={{ marginTop: 18, width: "100%", background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "10px 0", fontWeight: 600, cursor: "pointer" }}>
          Cerrar
        </button>
      </div>
    </div>
  );
}

function ResetAllSection({ onResetAllData }) {
  const [open, setOpen] = useState(false);
  const [confirmText, setConfirmText] = useState("");
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function handleReset() {
    setError("");
    if (confirmText !== "BORRAR") return setError('Escribí exactamente "BORRAR" para confirmar.');
    if (!/^\d{4}$/.test(pin)) return setError("Ingresá el PIN de administrador.");
    setBusy(true);
    const res = await onResetAllData(pin);
    setBusy(false);
    if (!res.ok) return setError(res.msg);
    // si salió bien, el panel se cierra solo (el admin queda deslogueado)
  }

  return (
    <div style={{ marginTop: 16, paddingTop: 16, borderTop: `1px solid ${COLORS.danger}` }}>
      <p style={{ fontSize: 11.5, color: COLORS.danger, fontWeight: 600, margin: "0 0 8px" }}>
        Zona de riesgo — borra jugadores, canchas, torneos, tarjetas, fotos y hasta tu propio PIN de
        administrador. No hay forma de deshacerlo.
      </p>
      {!open ? (
        <button
          onClick={() => setOpen(true)}
          style={{ width: "100%", background: "none", border: `1px solid ${COLORS.danger}`, color: COLORS.danger, borderRadius: 8, padding: "10px 0", fontWeight: 600, cursor: "pointer", fontSize: 13 }}
        >
          Borrar todos los datos
        </button>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <input
            value={confirmText}
            onChange={(e) => setConfirmText(e.target.value)}
            placeholder='Escribí BORRAR para confirmar'
            style={inputStyle}
          />
          <input
            type="password" inputMode="numeric" maxLength={4} value={pin}
            onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))}
            placeholder="PIN de administrador"
            style={inputStyle}
          />
          {error && <p style={{ color: COLORS.danger, fontSize: 12.5, margin: 0 }}>{error}</p>}
          <div style={{ display: "flex", gap: 8 }}>
            <button
              onClick={handleReset}
              disabled={busy}
              style={{ flex: 1, background: COLORS.danger, color: COLORS.paper, border: "none", borderRadius: 8, padding: "10px 0", fontWeight: 700, cursor: "pointer" }}
            >
              Confirmar borrado total
            </button>
            <button
              onClick={() => { setOpen(false); setConfirmText(""); setPin(""); setError(""); }}
              style={{ background: "none", border: `1px solid ${COLORS.paperDim}`, borderRadius: 8, padding: "0 14px", cursor: "pointer" }}
            >
              Cancelar
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function Toast({ msg }) {
  return (
    <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: COLORS.green900, color: COLORS.paper, border: `1px solid ${COLORS.brass}`, padding: "10px 18px", borderRadius: 10, fontSize: 13.5, fontWeight: 500, zIndex: 60, maxWidth: "90vw", textAlign: "center" }}>
      {msg}
    </div>
  );
}
