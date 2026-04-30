const AUTH_API_URL = "https://script.google.com/macros/s/AKfycbyTqHC1as7ahGbG14_r5WJq3Dkg10SHGYT65Kr58adf-3wnvCokYPelSy4YejoVANUT/exec";

const AuthGate = {
    tokenKey: "pm_auth_token",
    emailKey: "pm_auth_email",
    getToken() { return localStorage.getItem(this.tokenKey) || ""; },
    getEmail() { return localStorage.getItem(this.emailKey) || ""; },
    setSession(token, email) {
        localStorage.setItem(this.tokenKey, token || "");
        localStorage.setItem(this.emailKey, email || "");
    },
    clearSession() {
        localStorage.removeItem(this.tokenKey);
        localStorage.removeItem(this.emailKey);
    },
    async post(payload) {
        const res = await fetch(AUTH_API_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            credentials: "include",
            body: JSON.stringify(payload || {})
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return res.json();
    },
    redirectToLogin() {
        if (!location.pathname.toLowerCase().endsWith("/login.html")) {
            window.location.href = "login.html";
        }
    },
    requireLocalSession() {
        const token = this.getToken();
        if (!token) {
            this.redirectToLogin();
            return false;
        }
        return true;
    },
    async verifyRemoteOrLogout() {
        const token = this.getToken();
        if (!token) return false;
        try {
            const r = await this.post({ action: "verifySession", token });
            if (!r?.ok) {
                this.clearSession();
                this.redirectToLogin();
                return false;
            }
            return true;
        } catch (_e) {
            return true;
        }
    },
    async login(email, password) {
        return this.post({ action: "login", email, password });
    },
    async logout() {
        const token = this.getToken();
        try {
            if (token) await this.post({ action: "logout", token });
        } catch (_e) {}
        this.clearSession();
        this.redirectToLogin();
    },
    mountLogoutButton(selector, label = "Cerrar sesión") {
        const host = document.querySelector(selector);
        if (!host) return;
        if (host.querySelector(".auth-logout-btn")) return;
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "btn auth-logout-btn";
        btn.textContent = label;
        btn.addEventListener("click", () => this.logout());
        host.appendChild(btn);
    }
};


(function(){ if(!document.body.classList.contains('page-procesador')) return;
if (!AuthGate.requireLocalSession()) return;
AuthGate.verifyRemoteOrLogout();
/* ----------------------------------------------------------
    FUNCIONES AUXILIARES
---------------------------------------------------------- */
const mesesES = { ene:'01', enero:'01', feb:'02', febrero:'02', mar:'03', marzo:'03', abr:'04', abril:'04',
may:'05', mayo:'05', jun:'06', junio:'06', jul:'07', julio:'07', ago:'08', agosto:'08', sep:'09',
sept:'09', septiembre:'09', oct:'10', octubre:'10', nov:'11', noviembre:'11', dic:'12', diciembre:'12' };
const diasES = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];

function excelSerialToYMD(serial){
    if(!serial) return null;
    const n = Number(serial);
    const serialDateOnly = Math.floor(n);
    if(isNaN(serialDateOnly)) return null;
    const ms = (serialDateOnly - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if(isNaN(d)) return null;
    return `${d.getUTCFullYear()}${String(d.getUTCMonth()+1).padStart(2,'0')}${String(d.getUTCDate()).padStart(2,'0')}`;
}

// NUEVA FUNCIÓN: Convierte YMD directamente a Serial Excel (Evita errores de parsing)
function ymdToExcelSerial(ymd) {
    if (!ymd || ymd.length !== 8) return 0;
    const y = parseInt(ymd.substring(0, 4), 10);
    const m = parseInt(ymd.substring(4, 6), 10) - 1; // Meses 0-11
    const d = parseInt(ymd.substring(6, 8), 10);
    // Usamos UTC para evitar problemas de zona horaria al calcular días puros
    const utcMs = Date.UTC(y, m, d);
    // 25569 es la diferencia en días entre 1970-01-01 y 1899-12-30
    const serial = (utcMs / 86400000) + 25569;
    return Math.floor(serial);
}

function toYMDFromAny(v){
    if(!v) return null;
    const s=String(v).trim();
    if(/^\d{4}-\d{2}-\d{2}$/i.test(s)) return s.replace(/-/g,"");
    const slash=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if(slash) return slash[3]+String(slash[2]).padStart(2,'0')+String(slash[1]).padStart(2,'0');
    if(/^\d+(\.\d+)?$/.test(s)){ return excelSerialToYMD(v); }
    const coma=s.split(',')[0].trim();
    const p=coma.split(/\s+/);
    if(p.length>=3){
        const d=String(p[0]).padStart(2,'0');
        const mm=mesesES[p[1].toLowerCase()]||mesesES[p[1].toLowerCase().substring(0,3)];
        const yy=p[2];
        if(mm && /^\d{4}$/.test(yy)) return yy+mm+d;
    }
    const dt=new Date(s);
    if(!isNaN(dt))
        return `${dt.getFullYear()}${String(dt.getMonth()+1).padStart(2,'0')}${String(dt.getDate()).padStart(2,'0')}`;
    return null;
}

function toHHMMFromAny(v){
    if (v === null || v === undefined || v === "") return null;
    const n = Number(v);
    if(typeof v === 'number' && !isNaN(n)) {
        let frac = n - Math.floor(n);
        if (frac < 0.000001 && frac > -0.000001) {
            const s = String(v).trim();
            const timeMatch = s.match(/(\d{1,2}):(\d{2})$/);
            if (timeMatch) return `${timeMatch[1].padStart(2, '0')}:${timeMatch[2]}`;
            return null;
        }
        let totalSeconds = Math.round(frac * 86400);
        let hours = Math.floor(totalSeconds / 3600) % 24;
        let minutes = Math.floor((totalSeconds % 3600) / 60);
        const pad = (num) => String(num).padStart(2, '0');
        return `${pad(hours)}:${pad(minutes)}`;
    }
    const s = String(v).trim();
    if (/^\d{1,2}:\d{2}$/.test(s)) {
        return s.padStart(5, '0');
    }
    const timeMatch = s.match(/(\d{1,2}):(\d{2})$/); 
    if (timeMatch) {
         return `${timeMatch[1].padStart(2, '0')}:${timeMatch[2]}`;
    }
    return null;
}

function formatYMDtoDDMMYYYY(y){
    if(!y || y.length!==8) return "";
    return `${y.slice(6,8)}/${y.slice(4,6)}/${y.slice(0,4)}`;
}

// Formateador para las columnas nuevas (simulando Format(date, "yyyy-mm-dd HH:nn:ss"))
function formatExcelSerialToMacroString(v) {
    if (v === null || v === undefined || v === "" || v === "-") return "";
    let ms = Math.round((v - 25569) * 86400 * 1000); // Math.round corrige el segundo perdido
    const d = new Date(ms);
    if (isNaN(d)) return "";
    
    // Ajuste zona horaria UTC para que coincida con el serial
    const pad = (num) => String(num).padStart(2, '0');
    const Y = d.getUTCFullYear();
    const M = pad(d.getUTCMonth() + 1);
    const D = pad(d.getUTCDate());
    const h = pad(d.getUTCHours());
    const m = pad(d.getUTCMinutes());
    const s = pad(d.getUTCSeconds());
    return `${Y}-${M}-${D} ${h}:${m}:${s}`;
}

function formatDateFromActivos(v) {
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === 'string') return v.trim();
    if (typeof v === 'number' && !isNaN(v)) {
        let ms = (v - 25569) * 86400 * 1000;
        const d = new Date(ms);
        if (isNaN(d)) return String(v);
        const pad = (num) => String(num).padStart(2, '0');
        const year = d.getUTCFullYear();
        const month = pad(d.getUTCMonth() + 1);
        const day = pad(d.getUTCDate());
        const frac = v - Math.floor(v);
        if (frac > 0.000001) {
            const totalSeconds = Math.round(frac * 86400);
            const hours = pad(Math.floor(totalSeconds / 3600));
            const minutes = pad(Math.floor((totalSeconds % 3600) / 60));
            return `${day}/${month}/${year} ${hours}:${minutes}`;
        }
        return `${day}/${month}/${year}`;
    }
    return String(v);
}

function nomPropio(t){
    if(!t) return "";
    return String(t).toLowerCase().replace(/\b\w/g,c=>c.toUpperCase());
}

function safeNum(v){
    if(!v) return 0;
    return parseFloat(String(v).replace(',','.')) || 0;
}

// CONTROL DEL SPINNER
function setLoadingText(message) {
    document.getElementById('loadingText').textContent = message || 'Procesando datos, por favor espere...';
}
function showLoading(message) {
    setLoadingText(message);
    document.getElementById('loadingOverlay').style.display = 'flex';
}
function hideLoading(delayMs = 0) {
    const close = () => {
        document.getElementById('loadingOverlay').style.display = 'none';
        setLoadingText('Procesando datos, por favor espere...');
    };
    if (delayMs > 0) {
        setTimeout(close, delayMs);
        return;
    }
    close();
}
function notifyWithSpinner(message, delayMs = 1800) {
    showLoading(message);
    hideLoading(delayMs);
}

async function readSheetAOA(file){
    const ext=file.name.split('.').pop().toLowerCase();
    
    if(ext==="csv"){
        const txt=await file.text();
        const lines = txt.split(/\r?\n/).filter(l => l.trim() !== "");
        if (lines.length === 0) return [];
        
        let delimiter = ',';
        const firstLine = lines[0];

        if (file.name.toLowerCase().includes("programacion") || file.name.toLowerCase().includes("marcaciones")) {
            if (firstLine.includes(';')) {
                delimiter = ';';
            } else if (!firstLine.includes(',')) {
                try {
                     const ab=await file.arrayBuffer();
                     const wb=XLSX.read(ab,{type:'array', cellDates: false, raw: true}); 
                     const sh=wb.SheetNames[0];
                     return XLSX.utils.sheet_to_json(wb.Sheets[sh],{header:1,defval:""});
                } catch (e) {
                     delimiter = ',';
                }
            }
            if (firstLine.includes(delimiter)) {
                 const aoa = [];
                 for (const l of lines) {
                    const row=[]; let cur="",q=false;
                    for(const c of l){
                        if(c=='"') q=!q;
                        else if(c===delimiter&&!q){ row.push(cur); cur=""; }
                        else cur+=c;
                    }
                    row.push(cur);
                    aoa.push(row);
                }
                return aoa;
            }
        } 
        
        try {
             const ab=await file.arrayBuffer();
             const wb=XLSX.read(ab,{type:'array', cellDates: false, raw: true}); 
             const sh=wb.SheetNames[0];
             return XLSX.utils.sheet_to_json(wb.Sheets[sh],{header:1,defval:""});
        } catch (e) {
             const aoa = [];
             for (const l of lines) {
                 aoa.push(l.split(',').map(c => c.trim()));
             }
             return aoa;
        }
    }
    
    const ab=await file.arrayBuffer();
    const wb=XLSX.read(ab,{type:'array', cellDates: false, raw: true}); 
    const sh=wb.SheetNames[0];
    return XLSX.utils.sheet_to_json(wb.Sheets[sh],{header:1,defval:""});
}

function getExcelSerialDate(v) {
    if (v === null || v === undefined || v === "" || v === "-") return null;
    if (typeof v === 'number' && !isNaN(v)) return v;
    let s = String(v).trim().toLowerCase();
    
    if (/^\d+(\.\d+)?$/.test(s)) {
        return parseFloat(s);
    }
    
    s = s.replace(/ene/g, 'jan').replace(/feb/g, 'feb').replace(/mar/g, 'mar')
         .replace(/abr/g, 'apr').replace(/may/g, 'may').replace(/jun/g, 'jun')
         .replace(/jul/g, 'jul').replace(/ago/g, 'aug').replace(/sep/g, 'sep')
         .replace(/oct/g, 'oct').replace(/nov/g, 'nov').replace(/dic/g, 'dec');

    s = s.replace(/,/, ' ').replace(/-/g, '/');
    const dt = new Date(s);
    if (!isNaN(dt) && dt.getFullYear() > 1900) {
        let serial = (dt.getTime() / (86400 * 1000)) + 25569;
        const timezoneOffsetDays = dt.getTimezoneOffset() / (60 * 24);
        serial -= timezoneOffsetDays;
        return serial;
    }
    return null;
}

function getPdvAgrupado(pdv) {
    if (!pdv) return "";
    const p = String(pdv).toUpperCase().trim();
    if (p === "PPP2" || p === "PPP1") return "PPP";
    else if (p === "PPP3") return "PPH";
    else if (p === "LOG2"|| p === "LOG1") return "LOG";
    else return pdv;
}

function getWeekNumber(ymd) {
    if (!ymd || ymd.length !== 8) return "";
    const year = parseInt(ymd.substring(0, 4), 10);
    const month = parseInt(ymd.substring(4, 6), 10) - 1;
    const day = parseInt(ymd.substring(6, 8), 10);
    const date = new Date(Date.UTC(year, month, day));
    date.setUTCDate(date.getUTCDate() + 4 - (date.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
    return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
}

function getDayOfWeekName(ymd) {
    if (!ymd || ymd.length !== 8) return "";
    const year = parseInt(ymd.substring(0, 4), 10);
    const month = parseInt(ymd.substring(4, 6), 10) - 1;
    const day = parseInt(ymd.substring(6, 8), 10);
    const date = new Date(year, month, day);
    return diasES[date.getDay()];
}

function getDateNMonthsAgoYMD(months) {
    const d = new Date();
    d.setMonth(d.getMonth() - months);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    return `${year}${month}01`;
}

function getDateNextMonthYMD() {
    const d = new Date();
    d.setMonth(d.getMonth() + 1);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    return `${year}${month}01`;
}

function shiftYmdDays(ymd, days) {
    if (!ymd || ymd.length !== 8) return null;
    const d = new Date(ymd.slice(0,4), ymd.slice(4,6)-1, ymd.slice(6,8));
    d.setDate(d.getDate() + days);
    return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

function getIsoWeekBounds(ymd) {
    if (!ymd || ymd.length !== 8) return { start: "", end: "" };
    const date = new Date(Date.UTC(
        parseInt(ymd.substring(0, 4), 10),
        parseInt(ymd.substring(4, 6), 10) - 1,
        parseInt(ymd.substring(6, 8), 10)
    ));
    const isoDay = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() - isoDay + 1);
    const monday = `${date.getUTCFullYear()}${String(date.getUTCMonth() + 1).padStart(2, "0")}${String(date.getUTCDate()).padStart(2, "0")}`;
    const sundayDate = new Date(date);
    sundayDate.setUTCDate(sundayDate.getUTCDate() + 6);
    const sunday = `${sundayDate.getUTCFullYear()}${String(sundayDate.getUTCMonth() + 1).padStart(2, "0")}${String(sundayDate.getUTCDate()).padStart(2, "0")}`;
    return { start: monday, end: sunday };
}

function textBeforeSpace(value) {
    return String(value || "").trim().split(/\s+/)[0] || "";
}

function normalizeSheetValue(value) {
    return value === undefined || value === null ? "" : value;
}

function normalizeComparable(value) {
    return String(value || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

function normalizeDocumentId(value) {
    const raw = String(value ?? "")
        .trim()
        .replace(/^'+/, "")
        .replace(/\s+/g, "")
        .replace(/\.0+$/, "");
    if (!raw) return "";
    if (!/[A-Za-z]/.test(raw)) {
        const digitsOnly = raw.replace(/[^\d]/g, "");
        return digitsOnly || raw;
    }
    return raw.toUpperCase();
}

function hasDataInRange(row, start, end){
    for(let i=start; i<=end; i++){
        if(String(row?.[i] ?? "").trim() !== "") return true;
    }
    return false;
}

const COL_ESTADO_CONTRATO = 38;
const COL_FECHA_RETIRO = 39;
const ROW_HEADER_ACTIVOS = 3;
const INDICES_ACTIVOS_EXPORT = [0, 1, 23, 21, 22, 29, 38, 39, 41, 115];
const REV_STATIC_COLS_END = 6;
const REV_DATES_START_IDX = 7;
const HEADERS_DETALLE = [
    "LLAVE", "PDV", "DESCRIPCION", "NOMBRES", "APELLIDOS", "CEDULA",
    "ENTRA COLAB", "SALIDA COLAB", "HRS COLAB", "ENTRA GRTE", "SALIDA GRTE", "HRS GRTE",
    "MARCAC.COMP", "MARCAC.PUB", "CONTRATO", "AUSENCIA", "HRS NOMINA",
    "OBSERVACION", "ENTRADA_REAL", "SALIDA_REAL", "TOMADA_DE", "NOVEDAD"
];
const HEADERS_CONSOL = [
    "LLAVE", "NOVEDAD", "FECHA", "DOCUMENTO", "NOMBRE", "APELLIDO",
    "PUNTO DE VENTA ASIGNADO", "DESCRIPCION ASIGNADA", "CENTRO DE COSTOS", "CONTRATO", "CARGO",
    "FECHA DE CONTRATACION", "FECHA DE TERMINACION", "PUNTO DE VENTA MARCACION", "DESCRIPCION MARCACION",
    "MARCACION COMPLETA", "MARCACION PUBLICADA", "HORAS NOMINA", "OBSERVACION", "ESTADO DEL CONTRATO",
    "NÚMERO DE SEMANA", "DÍA DE LA SEMANA", "PROGRAMACION", "EJECUCION", "AUSENTISMO", "NOVEDADES"
];
const HEADERS_DESCANSOS = [
    "LLAVE", "NOVEDAD", "FECHA", "DOCUMENTO", "NOMBRE", "APELLIDO",
    "PUNTO DE VENTA ASIGNADO", "DESCRIPCION ASIGNADA"
];
const HEADERS_AUS = [
    "LLAVE", "TIPO DE AUSENTISMO", "NOMBRE EMPLEADO", "CÉDULA", "SUCURSAL",
    "FECHA INICIO", "FECHA FIN", "DÍAS SOLICITADOS", "HORAS TOTALES", "ESTADO", "PASO ACTUAL", "OBSERVACION"
];

function isEstadoActivo(value){
    const estado = String(value || "").trim().toUpperCase();
    return estado === "A" || estado === "ACTIVO";
}

function compareActivosRows(a, b){
    const activeA = isEstadoActivo(a?.[COL_ESTADO_CONTRATO]);
    const activeB = isEstadoActivo(b?.[COL_ESTADO_CONTRATO]);
    if (activeA !== activeB) return activeA ? -1 : 1;
    if (!activeA) {
        const retiroA = toYMDFromAny(a?.[COL_FECHA_RETIRO]) || "00000000";
        const retiroB = toYMDFromAny(b?.[COL_FECHA_RETIRO]) || "00000000";
        if (retiroA !== retiroB) return retiroB.localeCompare(retiroA);
    }
    return String(a?.[0] || "").localeCompare(String(b?.[0] || ""));
}

function appendSheetIfData(wb, name, aoa) {
    if (!aoa || aoa.length <= 1) return false;
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
    return true;
}

function reorderWorkbookSheets(wb) {
    if (!wb || !Array.isArray(wb.SheetNames)) return;
    const getSheetOrder = (name) => {
        const match = String(name || "").match(/^(\d+)/);
        return match ? parseInt(match[1], 10) : 999;
    };
    wb.SheetNames.sort((a, b) => getSheetOrder(b) - getSheetOrder(a));
}

function diffDaysInclusive(ymdStart, ymdEnd) {
    if (!ymdStart || !ymdEnd) return 0;
    const start = new Date(ymdStart.slice(0,4), ymdStart.slice(4,6)-1, ymdStart.slice(6,8));
    const end = new Date(ymdEnd.slice(0,4), ymdEnd.slice(4,6)-1, ymdEnd.slice(6,8));
    return Math.floor((end - start) / 86400000) + 1;
}

function maxYmd(a, b) {
    return !a ? b : !b ? a : (a > b ? a : b);
}

function minYmd(a, b) {
    return !a ? b : !b ? a : (a < b ? a : b);
}

function overlapDaysInclusive(startA, endA, startB, endB) {
    const start = maxYmd(startA, startB);
    const end = minYmd(endA, endB);
    if (!start || !end || start > end) return 0;
    return diffDaysInclusive(start, end);
}

function getRetiroYmdFromRow(row) {
    return toYMDFromAny(row?.[COL_FECHA_RETIRO]) || null;
}

function getIngresoYmdFromActRow(actRow, fallbackRow) {
    return toYMDFromAny(actRow?.[29]) || toYMDFromAny(fallbackRow?.[5]) || null;
}

function getTerminacionYmdFromActRow(actRow, fallbackRow) {
    return getRetiroYmdFromRow(actRow) || toYMDFromAny(fallbackRow?.[6]) || null;
}

function wasRetiredLongAgo(actRow, cutoffYmd) {
    const retiro = getRetiroYmdFromRow(actRow);
    return !!(retiro && retiro < cutoffYmd);
}

function isAdministrativoException(row) {
    const estadoContrato = normalizeComparable(row?.[COL_ESTADO_CONTRATO]);
    const docRaw = String(row?.[0] || "").trim();
    const descripcionFinal = normalizeComparable(row?.[INDICES_ACTIVOS_EXPORT[2]] || row?.[23]);
    const centroCostos = normalizeComparable(row?.[INDICES_ACTIVOS_EXPORT[3]] || row?.[21]);
    const colI = normalizeComparable(row?.[41]);
    const cargoUpper = normalizeComparable(row?.[115]);
    const cargosMantenimientoAdministrativos = new Set([
        "LIDER DE MANTENIMIENTO PDV",
        "AUXILIAR ADMINISTRATIVO DE MANTENIMIENTO",
        "LIDER DE MANTENIMIENTO DE PPP"
    ]);
    const administrativos = new Set([
        "GERENCIA GENERAL", "ANALITICA DE DATOS", "TESORERIA", "CONTABILIDAD", "NOMINA",
        "PLANEACION Y ANALISIS FINANCIERO", "COSTOS", "I+D+I", "COMUNICARTE",
        "DIRECCION ADMINISTRATIVA", "SERVICIOS GENERALES", "SERVICIOS ADMINISTRATIVOS",
        "DESARROLLO HUMANO", "SELECCION", "ACADEMIA DE ARTES Y FORMACION",
        "SEGURIDAD Y SALUD EN EL TRABAJO", "TECNOLOGIA", "DIRECCION DE CALIDAD",
        "CALIDAD", "CONTROL Y MEJORA CONTINUA", "GESTION AMBIENTAL", "COMPRAS",
        "BIENESTAR Y CULTURA ORGANIZACIONAL", "GESTION DE ACTIVOS FIJOS",
        "DIRECCION OPERATIVA", "VINCULOS Y RELACIONES PERSONALES"
    ]);

    if (estadoContrato === "R") return "Ok (Retirado)";
    if (!docRaw || /^0+$/.test(docRaw)) return "Sin Datos";
    if (centroCostos === "SENA" || descripcionFinal === "SENA") return "Ok (Aprendiz)";
    if (descripcionFinal === "MANTENIMIENTO Y OBRAS" && cargosMantenimientoAdministrativos.has(cargoUpper)) return "Ok (Administrativo)";
    if (descripcionFinal === "MANTENIMIENTO Y OBRAS" && textBeforeSpace(cargoUpper) === "LIDER") return "Ok (Administrativo)";
    if (centroCostos === "PPH" && colI === "R") return "Ok (Administrativo)";
    if (centroCostos === "PPP" && colI === "R") return "Ok (Administrativo)";
    if (centroCostos === "LOG" && colI === "R") return "Ok (Administrativo)";
    if (administrativos.has(descripcionFinal)) return "Ok (Administrativo)";
    return "";
}

function getEstadoContratoEnFecha(actRow, fallbackRow, ymd) {
    const ingreso = getIngresoYmdFromActRow(actRow, fallbackRow);
    const retiro = getTerminacionYmdFromActRow(actRow, fallbackRow);
    if (ingreso && ymd < ingreso) return "R";
    if (retiro && ymd > retiro) return "R";
    return "A";
}

function buildActivosRevisionStatus(row, reviewCount, expectedDays, minReviewYmd, maxReviewYmd) {
    const fechaIngreso = toYMDFromAny(row?.[29]);
    const exceptionResult = isAdministrativoException(row);
    if (exceptionResult) return exceptionResult;
    if (fechaIngreso && maxReviewYmd && fechaIngreso > maxReviewYmd) {
        return "Ok (Personal Nuevo)";
    }
    if (reviewCount === Number(expectedDays || 0)) return `Ok (${reviewCount} Días En Revisión)`;
    return `Revisar (${reviewCount} Días En Revisión)`;
}

function drawSummaryChart(summary){
    const canvas = document.getElementById("summaryCanvas");
    const legend = document.getElementById("chartLegend");
    const pdvList = document.getElementById("pdvList");
    if (!canvas || !legend) return;

    const ctx = canvas.getContext("2d");
    const metrics = [
        { label: "Activos", value: summary.activos || 0, color: "#2b7cff" },
        { label: "Retirados", value: summary.retirados || 0, color: "#5593ff" },
        { label: "Novedades", value: summary.novedades || 0, color: "#1d5edc" },
        { label: "Marcaciones", value: summary.marcaciones || 0, color: "#6da3ff" },
        { label: "Ausentismos", value: summary.ausentismos || 0, color: "#3d7ef2" },
        { label: "Descansos", value: summary.descansos || 0, color: "#8bb7ff" }
    ];
    const maxValue = Math.max(1, ...metrics.map(item => item.value));

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#f7faff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#2b7cff";
    ctx.font = "bold 20px Segoe UI";
    ctx.fillText("Resumen de novedades procesadas", 24, 34);
    ctx.fillStyle = "#667085";
    ctx.font = "14px Segoe UI";
    ctx.fillText(`Hojas generadas: ${summary.hojas || 0}`, 24, 58);

    const chartTop = 96;
    const chartLeft = 70;
    const chartHeight = 248;
    const chartWidth = canvas.width - 108;
    const slotWidth = chartWidth / metrics.length;
    const barWidth = Math.max(28, slotWidth - 18);

    ctx.strokeStyle = "#d9e3f1";
    ctx.beginPath();
    ctx.moveTo(chartLeft, chartTop);
    ctx.lineTo(chartLeft, chartTop + chartHeight);
    ctx.lineTo(chartLeft + chartWidth, chartTop + chartHeight);
    ctx.stroke();

    metrics.forEach((item, index) => {
        const x = chartLeft + index * slotWidth + (slotWidth - barWidth) / 2;
        const barHeight = (item.value / maxValue) * (chartHeight - 18);
        const y = chartTop + chartHeight - barHeight;

        ctx.fillStyle = item.color;
        ctx.fillRect(x, y, barWidth, barHeight);
        ctx.fillStyle = "#344054";
        ctx.textAlign = "center";
        ctx.font = "12px Segoe UI";
        ctx.fillText(String(item.value), x + barWidth / 2, y - 8);
        ctx.fillText(item.label, x + barWidth / 2, chartTop + chartHeight + 18);
    });
    ctx.textAlign = "start";

    legend.innerHTML = metrics.map(item => `
        <div class="legend-item">
            <span>${item.label}</span>
            <strong>${item.value}</strong>
        </div>
    `).join("");

    const pdvEntries = (summary.pdvCounts || []).slice(0, 8);
    const maxPdv = Math.max(1, ...pdvEntries.map(item => item.count || 0));
    if (pdvList) {
        pdvList.innerHTML = pdvEntries.length ? pdvEntries.map(item => `
            <div class="pdv-row">
                <strong>${item.pdv}</strong>
                <div class="pdv-bar"><div class="pdv-bar-fill" style="width:${(item.count / maxPdv) * 100}%"></div></div>
                <span>${item.count}</span>
            </div>
        `).join("") : `<div class="metric-sub">Sin datos de PDV todavía.</div>`;
    }
    const metricNovedades = document.getElementById("metricNovedades");
    const metricActivos = document.getElementById("metricActivos");
    const metricRetirados = document.getElementById("metricRetirados");
    const metricTopPdv = document.getElementById("metricTopPdv");
    const metricTopPdvCount = document.getElementById("metricTopPdvCount");
    if (metricNovedades) metricNovedades.textContent = String(summary.novedades || 0);
    if (metricActivos) metricActivos.textContent = String(summary.activos || 0);
    if (metricRetirados) metricRetirados.textContent = `${summary.retirados || 0} retirados en ventana`;
    if (metricTopPdv) metricTopPdv.textContent = pdvEntries[0]?.pdv || "-";
    if (metricTopPdvCount) metricTopPdvCount.textContent = `${pdvEntries[0]?.count || 0} novedades`;
}

const APP_DB_NAME = "ProcesadorMarcacionesDB";
const APP_DB_VERSION = 1;
const APP_STORE_NAME = "appState";
const APP_SUMMARY_KEY = "summary";
const APP_UI_KEY = "uiState";

function openAppDb() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(APP_DB_NAME, APP_DB_VERSION);
        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(APP_STORE_NAME)) {
                db.createObjectStore(APP_STORE_NAME);
            }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function saveSummaryData(summary) {
    try {
        const db = await openAppDb();
        await new Promise((resolve, reject) => {
            const tx = db.transaction(APP_STORE_NAME, "readwrite");
            tx.objectStore(APP_STORE_NAME).put(summary, APP_SUMMARY_KEY);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
        db.close();
    } catch (error) {
        console.warn("No se pudo guardar el resumen del procesamiento.", error);
    }
}

function saveUiState() {
    try {
        const state = {
            dateStart: document.getElementById("dateStart")?.value || "",
            dateEnd: document.getElementById("dateEnd")?.value || "",
            includeAllContracts: !!document.getElementById("includeAllContracts")?.checked,
            chkOnlyReviewMain: !!document.getElementById("chkOnlyReviewMain")?.checked,
            chkAct: !!document.getElementById("chkAct")?.checked,
            chkRev: !!document.getElementById("chkRev")?.checked,
            chkMarc: !!document.getElementById("chkMarc")?.checked,
            chkAus: !!document.getElementById("chkAus")?.checked,
            chkProg: !!document.getElementById("chkProg")?.checked
        };
        localStorage.setItem(APP_UI_KEY, JSON.stringify(state));
    } catch (error) {
        console.warn("No se pudo guardar el estado de la interfaz.", error);
    }
}

function restoreUiState() {
    try {
        const raw = localStorage.getItem(APP_UI_KEY);
        if (!raw) return;
        const state = JSON.parse(raw);
        if (!state || typeof state !== "object") return;
        if (typeof state.dateStart === "string") document.getElementById("dateStart").value = state.dateStart;
        if (typeof state.dateEnd === "string") document.getElementById("dateEnd").value = state.dateEnd;
        if (typeof state.includeAllContracts === "boolean") document.getElementById("includeAllContracts").checked = state.includeAllContracts;
        if (typeof state.chkOnlyReviewMain === "boolean") document.getElementById("chkOnlyReviewMain").checked = state.chkOnlyReviewMain;
        if (typeof state.chkAct === "boolean") document.getElementById("chkAct").checked = state.chkAct;
        if (typeof state.chkRev === "boolean") document.getElementById("chkRev").checked = state.chkRev;
        if (typeof state.chkMarc === "boolean") document.getElementById("chkMarc").checked = state.chkMarc;
        if (typeof state.chkAus === "boolean") document.getElementById("chkAus").checked = state.chkAus;
        if (typeof state.chkProg === "boolean") document.getElementById("chkProg").checked = state.chkProg;
    } catch (error) {
        console.warn("No se pudo restaurar el estado de la interfaz.", error);
    }
}

drawSummaryChart({ activos: 0, retirados: 0, novedades: 0, marcaciones: 0, ausentismos: 0, descansos: 0, hojas: 0 });
restoreUiState();
["dateStart", "dateEnd", "includeAllContracts", "chkOnlyReviewMain", "chkAct", "chkRev", "chkMarc", "chkAus", "chkProg"].forEach(id => {
    const element = document.getElementById(id);
    if (element) element.addEventListener("change", saveUiState);
});
const toggleSidebarBtn = document.getElementById("toggleSidebar");
if (toggleSidebarBtn) {
    toggleSidebarBtn.addEventListener("click", () => {
        document.body.classList.toggle("sidebar-collapsed");
    });
}

function syncProgramacionWithRevision() {
    if (document.getElementById("chkRev").checked) {
        document.getElementById("chkProg").checked = true;
        saveUiState();
    }
}

document.getElementById("chkRev").addEventListener("change", syncProgramacionWithRevision);
syncProgramacionWithRevision();
const btnSummary = document.getElementById("btnSummary");
if (btnSummary) {
    btnSummary.addEventListener("click", () => {
        window.location.href = "dashboard.html";
    });
}

function normalizeProgramacionTurno(value) {
    const cleaned = String(value || "").replace(/\s+/g, " ").trim();
    if (!cleaned || cleaned === "A") return "";
    return cleaned;
}

function buildUltimoPdvLookup(marcMap) {
    const lookup = new Map();
    for (const [llave, items] of marcMap.entries()) {
        const doc = String(llave).replace(/\d{8}$/, "");
        const ymd = llave.slice(doc.length);
        const first = items[0];
        const pdvRaw = first?.rawRow?.[0] || "";
        const pdv = getPdvAgrupado(pdvRaw);
        const descripcion = nomPropio(first?.rawRow?.[1] || "");
        if (!pdv) continue;
        if (!lookup.has(doc)) lookup.set(doc, []);
        lookup.get(doc).push({ ymd, pdv, descripcion });
    }
    for (const values of lookup.values()) {
        values.sort((a, b) => a.ymd.localeCompare(b.ymd));
    }
    return lookup;
}

function getMarcacionReferencia(doc, ymd, itemsDia, actRow, ultimoPdvLookup) {
    if (itemsDia.length > 0) {
        return {
            pdv: getPdvAgrupado(itemsDia[0].rawRow?.[0] || actRow?.[22] || ""),
            descripcion: nomPropio(itemsDia[0].rawRow?.[1] || actRow?.[23] || "")
        };
    }

    const history = ultimoPdvLookup.get(doc) || [];
    let previous = null;
    let next = null;
    for (const item of history) {
        if (item.ymd <= ymd) previous = item;
        if (!next && item.ymd > ymd) next = item;
    }
    const referencia = previous || next;
    return {
        pdv: referencia?.pdv || getPdvAgrupado(actRow?.[22] || ""),
        descripcion: referencia?.descripcion || nomPropio(actRow?.[23] || "")
    };
}

function getTipoDiaRevision(row) {
    const ausentismo = String(row["AUSENTISMO"] || "").trim();
    const descanso = String(row["NOVEDAD"] || "").toUpperCase().includes("DESCANSO");
    const ejecucion = String(row["EJECUCION"] || "").trim();
    const horas = safeNum(row["HORAS NOMINA"]);
    const tieneMarcacion = horas > 0 || (ejecucion !== "" && ejecucion !== "No ejecutado" && ejecucion !== "Descanso");

    if (tieneMarcacion && ausentismo) return "marcación y ausentismo";
    if (descanso && ausentismo) return "descanso y ausentismo";
    if (tieneMarcacion && descanso) return "marcación y descanso";
    if (tieneMarcacion) return "marcacion";
    if (ausentismo) return "ausentismo";
    if (descanso) return "descanso";
    return "no identificado";
}

function formatHoursForMessage(value) {
    const num = safeNum(value);
    if (!isFinite(num)) return "0.0";
    return (Math.round(num * 10) / 10).toFixed(1);
}

function shouldApplyNoDescansoSemana(row, descansoSemana, ymdStart, ymdEnd) {
    const contrato = String(row["CONTRATO"] || "").trim();
    const centroCostosNorm = normalizeComparable(row["CENTRO DE COSTOS"]);
    const descripcionAsignadaNorm = normalizeComparable(row["DESCRIPCION ASIGNADA"]);
    const programacionNorm = normalizeComparable(row["PROGRAMACION"]);
    const ausentismo = String(row["AUSENTISMO"] || "").trim();
    const fechaYmd = String(row["_FECHA_YMD"] || "").trim();
    const fechaIngresoYmd = String(row["_FECHA_INGRESO_YMD"] || "").trim();
    const fechaRetiroYmd = String(row["_FECHA_RETIRO_YMD"] || "").trim();
    const weekBounds = getIsoWeekBounds(fechaYmd);
    const contratosAdministrativos = new Set([
        "C44ADM - Jornada laboral máxima legal Admon",
        "CMCADM - jornada personal dirección, confianza o manejo (admon)"
    ]);

    if (descansoSemana !== 0) return false;
    if (ausentismo) return false;
    if (programacionNorm.includes("CAPACITACION")) return false;
    if (contrato === "CFXPDV - jornada laboral de común acuerdo entre las partes") return false;
    if (contratosAdministrativos.has(contrato)) return false;
    if (centroCostosNorm === "MANTENIMIENTO Y OBRAS" || descripcionAsignadaNorm === "MANTENIMIENTO Y OBRAS") return false;
    if (!weekBounds.start || !weekBounds.end) return false;
    if (ymdStart && weekBounds.start < ymdStart) return false;
    if (ymdEnd && weekBounds.end > ymdEnd) return false;
    if (fechaIngresoYmd && fechaIngresoYmd > weekBounds.start) return false;
    if (fechaRetiroYmd && fechaRetiroYmd < weekBounds.end) return false;
    return true;
}

function buildRevisionNovedad(row, descansoSemana, markNoDescansoSemana) {
    let resultado = "";
    const contrato = String(row["CONTRATO"] || "").trim();
    const observacion = String(row["OBSERVACION"] || "").trim();
    const marcacionCompleta = String(row["MARCACION COMPLETA"] || "").trim();
    const marcacionPublicada = String(row["MARCACION PUBLICADA"] || "").trim();
    const horas = safeNum(row["HORAS NOMINA"]);
    const dia = String(row["DÍA DE LA SEMANA"] || "").trim().toLowerCase();
    const tipoDia = getTipoDiaRevision(row);
    const fechaYmd = String(row["_FECHA_YMD"] || "").trim();
    const fechaIngresoYmd = String(row["_FECHA_INGRESO_YMD"] || "").trim();
    const fechaRetiroYmd = String(row["_FECHA_RETIRO_YMD"] || "").trim();
    const puntoVentaMarcacion = String(row["PUNTO DE VENTA MARCACION"] || "").trim();
    const centroCostos = normalizeComparable(row["CENTRO DE COSTOS"]);
    const descripcionAsignadaNorm = normalizeComparable(row["DESCRIPCION ASIGNADA"]);
    const programacionNorm = normalizeComparable(row["PROGRAMACION"]);
    const ausentismoNorm = normalizeComparable(row["AUSENTISMO"]);
    const estadoContrato = String(row["ESTADO DEL CONTRATO"] || "").trim().toUpperCase();
    const esFilaDescanso = tipoDia === "descanso" || tipoDia === "descanso y ausentismo" || tipoDia === "marcación y descanso";
    const esContratoCfxpdv = contrato === "CFXPDV - jornada laboral de común acuerdo entre las partes";
    const esViernesSabadoDomingo = ["viernes", "sábado", "sabado", "domingo"].includes(dia);
    const esMantenimientoObras = centroCostos === "MANTENIMIENTO Y OBRAS" || descripcionAsignadaNorm === "MANTENIMIENTO Y OBRAS";
    const cargoNorm = normalizeComparable(row["CARGO"]);
    const cargosMantenimientoAdministrativos = new Set([
        "LIDER DE MANTENIMIENTO PDV",
        "AUXILIAR ADMINISTRATIVO DE MANTENIMIENTO",
        "LIDER DE MANTENIMIENTO DE PPP"
    ]);
    const ausentismosPorHoras = new Set([
        "SALIDA TEMPRANA POR SOLICITUD DEL COLABORADOR",
        "HORAS DE VOTACION",
        "HORA LACTANCIA"
    ]);
    const contratosAdministrativos = [
        "C44ADM - Jornada laboral máxima legal Admon",
        "CMCADM - jornada personal dirección, confianza o manejo (admon)",
        "C46ADM - Jornada laboral máxima legal Admon"
    ];

    if (fechaIngresoYmd && fechaYmd && fechaYmd < fechaIngresoYmd) return "Ok (Personal Nuevo)";
    if (estadoContrato === "R" && fechaRetiroYmd && fechaYmd && fechaYmd > fechaRetiroYmd) return "Ok (Retirado)";

    if (esFilaDescanso && contrato === "CFXPDV - jornada laboral de común acuerdo entre las partes" && descansoSemana > 0) {
        resultado = `Revisar, ${descansoSemana} descansos en la semana`;
    } else if (esFilaDescanso && contrato === "C44PDV - Jornada laboral máxima legal PDV" && descansoSemana > 1) {
        resultado = `Revisar, ${descansoSemana} descansos en la semana`;
    } else if (esFilaDescanso && descansoSemana > 2 && [
        "CMCPDV - jornada personal dirección, confianza o manejo (pdv)",
        "CMCADM - jornada personal dirección, confianza o manejo (admon)",
        "C44ADM - Jornada laboral máxima legal Admon",
        "C46ADM - Jornada laboral máxima legal Admon"
    ].includes(contrato)) {
        resultado = `Revisar, ${descansoSemana} descansos en la semana`;
    } else if (observacion === "Marcación Doble" || observacion === "Marcación Doble (Anular)") {
        resultado = `Revisar, marcación doble de ${formatHoursForMessage(horas)} horas`;
    } else if (marcacionCompleta === "-") {
        resultado = "Revisar, marcación incompleta";
    } else if (observacion === "Trabajo Nocturno") {
        resultado = "Revisar, Trabajo Nocturno";
    } else if (horas > 12.5) {
        resultado = `Revisar, marcación de ${formatHoursForMessage(horas)} horas`;
    } else {
        switch (tipoDia) {
            case "marcación y ausentismo":
                resultado = horas > 3 && ausentismosPorHoras.has(ausentismoNorm)
                    ? "Ok (Marcación Correcta)"
                    : "Revisar, marcación y ausentismo";
                break;
            case "descanso y ausentismo":
            case "marcación y descanso":
                resultado = `Revisar, ${tipoDia}`;
                break;
            case "no identificado":
                if (programacionNorm.includes("CAPACITACION")) {
                    resultado = "Ok (Capacitación Programada)";
                } else if (contrato === "CAP44ADM - Aprendices Admon") {
                    resultado = "Ok (Aprendiz administrativo)";
                } else if (esMantenimientoObras && cargosMantenimientoAdministrativos.has(cargoNorm)) {
                    resultado = "Ok (Administrativo)";
                } else if (esMantenimientoObras) {
                    resultado = "Revisar, día sin marcación";
                } else if (contratosAdministrativos.includes(contrato) && !esMantenimientoObras) {
                    resultado = "Ok (Administrativo)";
                } else if (esContratoCfxpdv && !esViernesSabadoDomingo) {
                    resultado = "Ok (Fines de semana)";
                } else {
                    resultado = "Revisar, día sin marcación";
                }
                break;
            case "marcacion":
                resultado = horas <= 3.9 ? `Revisar, marcación de ${formatHoursForMessage(horas)} horas` : "Ok (Marcación Correcta)";
                break;
            case "ausentismo":
                resultado = puntoVentaMarcacion ? "Ok (Ausentismo Cargado)" : "Revisar, día sin marcación";
                break;
            case "descanso":
                resultado = "Ok (Descanso)";
                break;
            default:
                resultado = "Revisar, día sin marcación";
        }
    }

    if (marcacionPublicada === "-" && marcacionCompleta === "SI" && (!resultado || resultado.startsWith("Ok"))) {
        resultado = "Revisar, marcación sin publicar";
    }

    if (markNoDescansoSemana) {
        resultado = "Revisar, no descanso en la semana";
    }

    return resultado || "Ok (Marcación Correcta)";
}

document.getElementById("btnRun").addEventListener("click", async () => {
    showLoading("Validando selección...");

    setTimeout(async () => {
        try {
            const selected = {
                act: document.getElementById("chkAct").checked,
                rev: document.getElementById("chkRev").checked,
                marc: document.getElementById("chkMarc").checked,
                aus: document.getElementById("chkAus").checked,
                prog: document.getElementById("chkProg").checked
            };
            const includeAllContracts = document.getElementById("includeAllContracts").checked;
            const onlyReviewMain = !!document.getElementById("chkOnlyReviewMain")?.checked;
            const requiresDateRange = selected.rev || selected.marc || selected.aus || selected.prog;

            if (!Object.values(selected).some(Boolean)) {
                notifyWithSpinner("Debe seleccionar al menos un tipo de archivo para procesar.");
                return;
            }

            const dateStartInput = document.getElementById("dateStart").value;
            const dateEndInput = document.getElementById("dateEnd").value;
            let ymdStart = "";
            let ymdEnd = "";

            if (requiresDateRange) {
                if (!dateStartInput || !dateEndInput) {
                    notifyWithSpinner("Debe seleccionar la Fecha de Inicio y la Fecha de Fin de la Revisión.");
                    return;
                }
                ymdStart = dateStartInput.replace(/-/g, "");
                ymdEnd = dateEndInput.replace(/-/g, "");
                if (ymdStart > ymdEnd) {
                    notifyWithSpinner("La Fecha de Inicio no puede ser posterior a la Fecha de Fin.");
                    return;
                }
            }

            const fAct = document.getElementById("fileAct").files[0];
            const fRev = document.getElementById("fileRev").files[0];
            const fMarc = document.getElementById("fileMarc").files[0];
            const fAus = document.getElementById("fileAus").files[0];
            const fProg = document.getElementById("fileProg").files[0];
            const allFiveSelected = !!(selected.act && selected.rev && selected.marc && selected.aus && selected.prog);
            const allFiveLoaded = !!(fAct && fRev && fMarc && fAus && fProg);
            const shouldPersistDashboardSummary = allFiveSelected && allFiveLoaded;

            if (selected.act && !fAct) { notifyWithSpinner("Seleccionó Activos, pero no cargó ese archivo."); return; }
            if (selected.rev && !fRev) { notifyWithSpinner("Debe cargar el archivo de Informe de Seguimiento."); return; }
            if (selected.marc && !fMarc) { notifyWithSpinner("Debe cargar el archivo de Marcaciones."); return; }
            if (selected.aus && !fAus) { notifyWithSpinner("Seleccionó Ausentismos, pero no cargó ese archivo."); return; }
            if (selected.prog && !fProg) { notifyWithSpinner("Seleccionó Programación, pero no cargó ese archivo."); return; }
            if (selected.aus && !fMarc && !selected.marc) {
                hideLoading();
                const continuarSinMarcaciones = confirm("¿Está seguro que quiere procesar los ausentismos sin tener en cuenta las marcaciones?");
                if (!continuarSinMarcaciones) return;
                showLoading("Leyendo archivos seleccionados...");
            }

            setLoadingText("Leyendo archivos seleccionados...");
            const [act, rev, marc, ausRaw, progRaw] = await Promise.all([
                fAct && (selected.act || selected.rev) ? readSheetAOA(fAct) : Promise.resolve(null),
                fRev && selected.rev ? readSheetAOA(fRev) : Promise.resolve(null),
                fMarc && (selected.marc || selected.rev || selected.aus) ? readSheetAOA(fMarc) : Promise.resolve(null),
                fAus && selected.aus ? readSheetAOA(fAus) : Promise.resolve(null),
                fProg && (selected.prog || selected.rev) ? readSheetAOA(fProg) : Promise.resolve(null)
            ]);

            const summary = { activos:0, retirados:0, novedades:0, marcaciones:0, marcacionesCorregidasGrte:0, ausentismos:0, descansos:0, hojas:0, pdvCounts:[], novedadesPorPdv:[], novedadesPorTipo:[] };
            const wb = XLSX.utils.book_new();
            const activosMap = new Map();
            const novedadMap = new Map();
            const marcMap = new Map();
            const ejecMap = new Map();
            const progMap = new Map();
            const ausentismoDetalleMap = new Map();
            const descansosConsol = [];
            const finalRows = [];
            let finalRowsExport = [];
            const marcacionesDetalleFinal = [HEADERS_DETALLE];
            const programacionRows = [];
            let filteredActivosRows = [];
            let ausentismosRows = [];
            let indicesFechasValidas = [];
            let indicesFechasDescansos = [];
            let headerRev = [];
            const expectedReviewDays = diffDaysInclusive(ymdStart, ymdEnd);
            const descansosConsolKeys = new Set();
            const addDescansoConsol = (doc, ymd, nov, actRow, nombre, apellido) => {
                const key = `${doc}|${ymd}`;
                if (!doc || !ymd || descansosConsolKeys.has(key)) return;
                descansosConsolKeys.add(key);
                descansosConsol.push({
                    "LLAVE": doc + ymdToExcelSerial(ymd),
                    "NOVEDAD": nomPropio(nov || ""),
                    "FECHA": formatYMDtoDDMMYYYY(ymd),
                    "DOCUMENTO": doc,
                    "NOMBRE": nombre,
                    "APELLIDO": apellido,
                    "PUNTO DE VENTA ASIGNADO": actRow?.[22] || "",
                    "DESCRIPCION ASIGNADA": actRow?.[23] || ""
                });
            };

            if (act && act.length > ROW_HEADER_ACTIVOS) {
                const limitInferiorYMD = getDateNMonthsAgoYMD(2);
                const limitSuperiorYMD = getDateNextMonthYMD();

                for (let i = ROW_HEADER_ACTIVOS + 1; i < act.length; i++) {
                    const r = act[i];
                    const doc = normalizeDocumentId(r?.[0]);
                    if (!doc) continue;

                    r[1] = nomPropio(r[1]);
                    r[23] = nomPropio(r[23]);
                    r[115] = nomPropio(r[115]);

                    const retiroYMD = toYMDFromAny(r[COL_FECHA_RETIRO]);
                    const include = includeAllContracts || !retiroYMD || (retiroYMD >= limitInferiorYMD && retiroYMD < limitSuperiorYMD);
                    if (!include) continue;

                    const existing = activosMap.get(doc);
                    if (!existing || (!isEstadoActivo(existing[COL_ESTADO_CONTRATO]) && isEstadoActivo(r[COL_ESTADO_CONTRATO]))) {
                        activosMap.set(doc, r);
                    }
                    filteredActivosRows.push(r);
                }

                filteredActivosRows.sort(compareActivosRows);
                summary.activos = filteredActivosRows.filter(row => isEstadoActivo(row[COL_ESTADO_CONTRATO])).length;
                summary.retirados = filteredActivosRows.length - summary.activos;
            }

            if (rev && rev.length > 0) {
                headerRev = rev[0] || [];
                for (let c = REV_DATES_START_IDX; c < headerRev.length; c++) {
                    const ymd = toYMDFromAny(headerRev[c]);
                    if (ymd && ymd >= ymdStart) indicesFechasDescansos.push(c);
                    if (ymd && ymd >= ymdStart && ymd <= ymdEnd) indicesFechasValidas.push(c);
                }

                for (let i = 1; i < rev.length; i++) {
                    const row = rev[i];
                    if (!row || !hasDataInRange(row, 0, REV_STATIC_COLS_END)) continue;
                    const doc = normalizeDocumentId(row[0]);
                    if (!doc) continue;
                    for (const c of indicesFechasValidas) {
                        const ymd = toYMDFromAny(headerRev[c]);
                        if (!ymd) continue;
                        novedadMap.set(doc + ymd, row[c] ? String(row[c]).trim() : "No identificado en Informe");
                    }
                }
            }

            if (marc && marc.length > 1) {
                const resultData = [];
                const dictCedula = new Map();
                const dictFecha = new Map();

                for (let i = 1; i < marc.length; i++) {
                    const r = marc[i];
                    if (!r || r.length < 6) continue;

                    const cedula = normalizeDocumentId(r[4]);
                    const val_eC = getExcelSerialDate(r[5]);
                    const val_sC = getExcelSerialDate(r[6]);
                    const val_eG = getExcelSerialDate(r[8]);
                    const val_sG = getExcelSerialDate(r[9]);
                    if (!cedula && val_eC === null && val_sC === null && val_eG === null && val_sG === null) continue;

                    const preferEntry = val_eG !== null ? val_eG : val_eC;
                    let preferExit = val_sG !== null ? val_sG : val_sC;
                    if (preferEntry === null) continue;
                    if (preferExit === null) preferExit = preferEntry;

                    const ymdCheck = formatExcelSerialToDateOnly(preferEntry);
                    if (requiresDateRange && ymdCheck && (ymdCheck < ymdStart || ymdCheck > ymdEnd)) continue;

                    const startTime = Number(preferEntry);
                    const endTime = Number(preferExit);
                    resultData.push({
                        rawRow: r,
                        cedula,
                        obs: Math.floor(startTime + 0.00001) !== Math.floor(endTime + 0.00001) ? "Trabajo Nocturno" : "0",
                        entry: preferEntry,
                        exit: (val_sG !== null || val_sC !== null) ? preferExit : null,
                        source: (val_eG !== null || val_sG !== null) ? "GRTE" : "COL",
                        startTime,
                        endTime,
                        horas: (endTime - startTime) * 24
                    });
                }

                resultData.sort((a, b) => a.cedula !== b.cedula ? a.cedula.localeCompare(b.cedula) : a.startTime - b.startTime);

                resultData.forEach((currentObj, rowIndex) => {
                    currentObj.index = rowIndex;
                    if (!dictCedula.has(currentObj.cedula)) dictCedula.set(currentObj.cedula, []);
                    const innerList = dictCedula.get(currentObj.cedula);

                    for (const stored of innerList) {
                        const objStored = resultData[stored.index];
                        if (currentObj.startTime <= objStored.endTime && currentObj.endTime >= objStored.startTime) {
                            if (currentObj.horas <= 0.2 && objStored.horas > 0.2) {
                                currentObj.obs = "Marcación Doble (Anular)";
                                if (objStored.obs === "0" || objStored.obs === "Trabajo Nocturno") objStored.obs = "Marcación Doble";
                            } else if (objStored.horas <= 0.2 && currentObj.horas > 0.2) {
                                objStored.obs = "Marcación Doble (Anular)";
                                if (currentObj.obs === "0" || currentObj.obs === "Trabajo Nocturno") currentObj.obs = "Marcación Doble";
                            } else if (objStored.source === "COL" && currentObj.source === "GRTE") {
                                objStored.obs = "Marcación Doble";
                                currentObj.obs = "Marcación Doble (Anular)";
                            } else if (objStored.source === "GRTE" && currentObj.source === "COL") {
                                currentObj.obs = "Marcación Doble";
                                objStored.obs = "Marcación Doble (Anular)";
                            } else {
                                if (objStored.obs === "0" || objStored.obs === "Trabajo Nocturno") objStored.obs = "Marcación Doble";
                                currentObj.obs = "Marcación Doble (Anular)";
                            }
                        }
                    }
                    innerList.push({ startTime: currentObj.startTime, endTime: currentObj.endTime, index: rowIndex });

                    const fechaEntrada = Math.floor(currentObj.startTime + 0.00001);
                    const claveFecha = currentObj.cedula + "|" + fechaEntrada;
                    if (!dictFecha.has(claveFecha)) dictFecha.set(claveFecha, []);
                    const listFechas = dictFecha.get(claveFecha);
                    if (listFechas.length > 0) {
                        for (const idx of listFechas) {
                            const objPrev = resultData[idx];
                            if (objPrev.obs === "0" || objPrev.obs === "Trabajo Nocturno") objPrev.obs = "Marcación Doble";
                        }
                        if (currentObj.obs === "0" || currentObj.obs === "Trabajo Nocturno") currentObj.obs = "Marcación Doble";
                    }
                    listFechas.push(rowIndex);
                });

                for (const item of resultData) {
                    if (item.obs === "Marcación Doble") {
                        const h = item.exit !== null ? (Number(item.exit) - Number(item.entry)) * 24 : 0;
                        if (h <= 0.2) item.obs = "Marcación Doble (Anular)";
                    }

                    const raw = item.rawRow;
                    raw[1] = nomPropio(raw[1]);
                    raw[2] = nomPropio(raw[2]);
                    raw[3] = nomPropio(raw[3]);

                    const ymd = formatExcelSerialToDateOnly(item.entry);
                    const llave = item.cedula + ymd;
                    const llaveVisual = item.cedula + Math.floor(item.entry);
                    if (!marcMap.has(llave)) marcMap.set(llave, []);
                    marcMap.get(llave).push(item);

                    const entrTime = getExcelSerialTime(item.entry);
                    const exitTime = getExcelSerialTime(item.exit);
                    if (entrTime && item.obs !== "Marcación Doble (Anular)") {
                        const descEjec = raw[1] || "";
                        ejecMap.set(llave, item.exit !== null ? `${descEjec} ${entrTime} A ${exitTime}` : `${descEjec} ${entrTime} A`);
                    }

                    marcacionesDetalleFinal.push([
                        llaveVisual, raw[0], raw[1], raw[2], raw[3], raw[4],
                        raw[5], raw[6], raw[7], raw[8], raw[9], raw[10],
                        raw[11], raw[12], raw[13], raw[14], safeNum(raw[15]),
                        item.obs, formatExcelSerialToMacroString(item.entry), formatExcelSerialToMacroString(item.exit), item.source,
                        novedadMap.get(llave) || ""
                    ]);
                }

                summary.marcaciones = marcacionesDetalleFinal.length - 1;
                summary.marcacionesCorregidasGrte = marcacionesDetalleFinal
                    .slice(1)
                    .reduce((acc, row) => acc + (String(row?.[20] || "").trim().toUpperCase() === "GRTE" ? 1 : 0), 0);
                if (selected.marc && appendSheetIfData(wb, "2.Horarios", marcacionesDetalleFinal)) summary.hojas++;
            }

            if (progRaw && progRaw.length > 1) {
                setLoadingText("Leyendo programación...");
                const IDX_PROG = { DOC: 1, FECHA: 2, PDV_DESC: 8, ENTRA_TURNO: 6, SALIDA_TURNO: 7 };
                for (let i = 1; i < progRaw.length; i++) {
                    const r = progRaw[i];
                    const doc = normalizeDocumentId(r?.[IDX_PROG.DOC]);
                    const ymd = toYMDFromAny(r?.[IDX_PROG.FECHA]);
                    if (!doc || !ymd || (requiresDateRange && (ymd < ymdStart || ymd > ymdEnd))) continue;

                    const pdv = nomPropio(r[IDX_PROG.PDV_DESC] || "");
                    const entr = toHHMMFromAny(r[IDX_PROG.ENTRA_TURNO]);
                    const sal = toHHMMFromAny(r[IDX_PROG.SALIDA_TURNO]);
                    const turnoBase = entr && sal ? `${pdv} ${entr} A ${sal}` : `${pdv} ${entr || ""} A ${sal || ""}`;
                    const turno = normalizeProgramacionTurno(turnoBase);
                    progMap.set(doc + ymd, turno);
                    programacionRows.push([doc, formatYMDtoDDMMYYYY(ymd), pdv, entr || "", sal || "", turno]);
                }
            }

            if (ausRaw && ausRaw.length > 1) {
                setLoadingText("Procesando ausentismos...");
                const IDX_AUS = { TIPO_AUS:0, NOMBRE_EMP:1, CEDULA:2, SUCURSAL:3, FECHA_INICIO:4, FECHA_FIN:5, DIAS_SOL:6, HORAS_TOT:7, ESTADO:8, PASO_ACT:9 };
                const terminationDates = new Map();
                const dailyAusentismos = new Map();
                const terminationTypes = ["Renuncia", "Renuncia Bienestar", "Terminación de contrato"];
                const overlapAusTypesByKey = new Map();
                const marcacionPermitidaTiposAus = new Set([
                    "SALIDA TEMPRANA POR SOLICITUD DEL COLABORADOR",
                    "HORAS DE VOTACION",
                    "HORA LACTANCIA",
                    "CITA MEDICA"
                ]);
                const shouldMarkMarcacionObservation = (tipoAus) => !marcacionPermitidaTiposAus.has(normalizeComparable(tipoAus));
                const getMarcacionHorasTexto = (llaveDia) => {
                    const items = marcMap.get(llaveDia) || [];
                    const totalHoras = items.reduce((acc, item) => acc + safeNum(item?.horas), 0);
                    return `Día con marcación registrada (${formatHoursForMessage(totalHoras)} horas)`;
                };

                for (let i = 1; i < ausRaw.length; i++) {
                    const r = ausRaw[i];
                    if (!r || r.length < 10) continue;
                    const estado = String(r[IDX_AUS.ESTADO] || "").trim();
                    if (!["Pendiente", "Aprobada", "Finalizada", "Aprobada y Finalizada"].includes(estado)) continue;

                    const doc = normalizeDocumentId(r[IDX_AUS.CEDULA]);
                    const fechaInicioYMD = toYMDFromAny(r[IDX_AUS.FECHA_INICIO]);
                    const fechaFinYMD = toYMDFromAny(r[IDX_AUS.FECHA_FIN]);
                    if (!doc || !fechaInicioYMD || !fechaFinYMD) continue;

                    const tipo = String(r[IDX_AUS.TIPO_AUS] || "").trim();
                    if (terminationTypes.includes(tipo)) {
                        const existing = terminationDates.get(doc);
                        if (!existing || fechaInicioYMD < existing.ymd) terminationDates.set(doc, { ymd: fechaInicioYMD, rowData: r });
                    }

                    let d = new Date(fechaInicioYMD.slice(0,4), fechaInicioYMD.slice(4,6)-1, fechaInicioYMD.slice(6,8));
                    const end = new Date(fechaFinYMD.slice(0,4), fechaFinYMD.slice(4,6)-1, fechaFinYMD.slice(6,8));
                    while (d <= end) {
                        const ymd = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
                        if (ymd >= ymdStart) {
                            const llaveDia = doc + ymd;
                            if (!terminationTypes.includes(tipo)) {
                                if (!overlapAusTypesByKey.has(llaveDia)) overlapAusTypesByKey.set(llaveDia, new Set());
                                overlapAusTypesByKey.get(llaveDia).add(tipo);
                            }
                            const rowAus = [
                                doc + ymdToExcelSerial(ymd),
                                r[IDX_AUS.TIPO_AUS] || "",
                                nomPropio(r[IDX_AUS.NOMBRE_EMP] || ""),
                                r[IDX_AUS.CEDULA] || "",
                                nomPropio(r[IDX_AUS.SUCURSAL] || ""),
                                formatYMDtoDDMMYYYY(ymd),
                                formatYMDtoDDMMYYYY(ymd),
                                r[IDX_AUS.DIAS_SOL] || "",
                                r[IDX_AUS.HORAS_TOT] || "",
                                r[IDX_AUS.ESTADO] || "",
                                r[IDX_AUS.PASO_ACT] || "",
                                ""
                            ];
                            if (!dailyAusentismos.has(llaveDia)) dailyAusentismos.set(llaveDia, []);
                            dailyAusentismos.get(llaveDia).push(rowAus);
                            const tipoActual = String(r[IDX_AUS.TIPO_AUS] || "").trim();
                            const prevTipo = String(ausentismoDetalleMap.get(llaveDia) || "").trim();
                            if (!prevTipo) {
                                ausentismoDetalleMap.set(llaveDia, tipoActual);
                            } else if (tipoActual && !prevTipo.split(" | ").includes(tipoActual)) {
                                ausentismoDetalleMap.set(llaveDia, `${prevTipo} | ${tipoActual}`);
                            }
                        }
                        d.setDate(d.getDate() + 1);
                    }
                }

                for (const [doc, termInfo] of terminationDates.entries()) {
                    const termYMD = termInfo.ymd;
                    const rawRow = termInfo.rowData;
                    const termDate = new Date(termYMD.slice(0,4), termYMD.slice(4,6)-1, termYMD.slice(6,8));
                    const propagationEndLimitDate = new Date(termDate);
                    propagationEndLimitDate.setDate(termDate.getDate() + 30);
                    const finalPropagationDate = propagationEndLimitDate;
                    let d = new Date(termDate);
                    while (d <= finalPropagationDate) {
                        const ymd = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
                        const llave = doc + ymd;
                        if (ymd >= ymdStart && !dailyAusentismos.has(llave)) {
                            const tipoAus = String(rawRow[IDX_AUS.TIPO_AUS] || "").trim();
                            const propagatedRow = [
                                doc + ymdToExcelSerial(ymd),
                                tipoAus,
                                nomPropio(rawRow[IDX_AUS.NOMBRE_EMP] || ""),
                                rawRow[IDX_AUS.CEDULA] || "",
                                nomPropio(rawRow[IDX_AUS.SUCURSAL] || ""),
                                formatYMDtoDDMMYYYY(ymd),
                                formatYMDtoDDMMYYYY(ymd),
                                "1",
                                "8",
                                "Propagado - " + tipoAus,
                                rawRow[IDX_AUS.PASO_ACT] || "",
                                ""
                            ];
                            dailyAusentismos.set(llave, [propagatedRow]);
                            ausentismoDetalleMap.set(llave, tipoAus);
                        }
                        d.setDate(d.getDate() + 1);
                    }
                }

                for (const [llave, tiposSet] of overlapAusTypesByKey.entries()) {
                    if (!dailyAusentismos.has(llave)) continue;
                    const rowsDia = dailyAusentismos.get(llave) || [];
                    for (const rowAus of rowsDia) {
                        const observaciones = [];
                        if (tiposSet.size > 1) observaciones.push(`Ausentismos sobrepuestos con: ${Array.from(tiposSet).join(" | ")}`);
                        if (marcMap.has(llave) && shouldMarkMarcacionObservation(rowAus[1])) {
                            observaciones.push(getMarcacionHorasTexto(llave));
                        }
                        if (observaciones.length > 0) rowAus[11] = observaciones.join(" | ");
                    }
                }
                for (const [llave, rowsDia] of dailyAusentismos.entries()) {
                    if (!marcMap.has(llave)) continue;
                    const observacionMarcacion = getMarcacionHorasTexto(llave);
                    for (const rowAus of rowsDia) {
                        if (rowAus[11] || !shouldMarkMarcacionObservation(rowAus[1])) continue;
                        rowAus[11] = observacionMarcacion;
                    }
                }

                ausentismosRows = Array.from(dailyAusentismos.values()).flat();
                summary.ausentismos = ausentismosRows.length;
            }

            if (rev && rev.length > 0) {
                setLoadingText("Consolidando revisión de días...");
                const ultimoPdvLookup = buildUltimoPdvLookup(marcMap);
                for (let i = 1; i < rev.length; i++) {
                    const row = rev[i];
                    if (!row || !hasDataInRange(row, 0, REV_STATIC_COLS_END)) continue;
                    const doc = normalizeDocumentId(row[0]);
                    if (!doc) continue;

                    const nombre = nomPropio(row[1] || "");
                    const apellido = nomPropio(row[2] || "");
                    const contrato = row[3] || "";
                    const cargo = nomPropio(row[4] || "");
                    const actRow = activosMap.get(doc) || [];
                    const fCont = getIngresoYmdFromActRow(actRow, row);
                    const fTerm = getTerminacionYmdFromActRow(actRow, row);
                    const estadoContratoActual = isEstadoActivo(actRow[COL_ESTADO_CONTRATO]) ? "A" : "R";
                    for (const c of indicesFechasDescansos) {
                        const ymdDesc = toYMDFromAny(headerRev[c]);
                        if (!ymdDesc) continue;
                        const novDesc = String(row[c] || "").trim();
                        if (novDesc && novDesc.toUpperCase().includes("DESCANSO")) {
                            addDescansoConsol(doc, ymdDesc, novDesc, actRow, nombre, apellido);
                        }
                    }

                    for (const c of indicesFechasValidas) {
                        const ymd = toYMDFromAny(headerRev[c]);
                        if (!ymd) continue;

                        const novRaw = row[c];
                        const nov = novRaw ? String(novRaw).trim() : "No encontrado en Seguimiento";
                        const llaveReporte = doc + ymd;
                        const itemsDia = marcMap.get(llaveReporte) || [];

                        let pv = actRow[22] || "";
                        let desc = nomPropio(actRow[23] || "");
                        let comp = "0";
                        let pub = "0";
                        let hrs = 0;
                        let finalStateForDoble = "0";
                        let progTurno = normalizeProgramacionTurno(progMap.get(llaveReporte));
                        let ejecTurno = ejecMap.get(llaveReporte);
                        const ausentismoDia = ausentismoDetalleMap.get(llaveReporte) || "";

                        if (String(novRaw || "").toUpperCase().includes("DESCANSO")) {
                            progTurno = "Descanso";
                            ejecTurno = "Descanso";
                        } else if (!progTurno) {
                            progTurno = "No programado";
                            if (!ejecTurno) ejecTurno = "No ejecutado";
                        } else if (!ejecTurno) {
                            ejecTurno = "No ejecutado";
                        }

                        if (itemsDia.length > 0) {
                            const first = itemsDia[0];
                            pv = first.rawRow[0] || actRow[22] || "";
                            desc = nomPropio(first.rawRow[1] || actRow[23] || "");
                            comp = first.rawRow[11] || "0";
                            pub = first.rawRow[12] || "0";
                            hrs = itemsDia.reduce((acc, item) => acc + safeNum(item.rawRow[15] || 0), 0);

                            const estados = itemsDia.map(it => it.obs);
                            if (estados.includes("Marcación Doble (Anular)")) finalStateForDoble = "Marcación Doble (Anular)";
                            else if (estados.includes("Marcación Doble")) finalStateForDoble = "Marcación Doble";
                            else if (estados.includes("Trabajo Nocturno")) finalStateForDoble = "Trabajo Nocturno";
                        }

                        const marcacionReferencia = getMarcacionReferencia(doc, ymd, itemsDia, actRow, ultimoPdvLookup);

                        finalRows.push({
                            "LLAVE": doc + ymdToExcelSerial(ymd),
                            "NOVEDAD": nov,
                            "FECHA": formatYMDtoDDMMYYYY(ymd),
                            "DOCUMENTO": doc,
                            "NOMBRE": nombre,
                            "APELLIDO": apellido,
                            "PUNTO DE VENTA ASIGNADO": actRow[22] || "",
                            "DESCRIPCION ASIGNADA": nomPropio(actRow[23] || ""),
                            "CENTRO DE COSTOS": actRow[21] || "",
                            "CONTRATO": contrato,
                            "CARGO": cargo,
                            "FECHA DE CONTRATACION": fCont ? formatYMDtoDDMMYYYY(fCont) : "",
                            "FECHA DE TERMINACION": fTerm ? formatYMDtoDDMMYYYY(fTerm) : "",
                            "PUNTO DE VENTA MARCACION": marcacionReferencia.pdv || getPdvAgrupado(pv),
                            "DESCRIPCION MARCACION": marcacionReferencia.descripcion || desc,
                            "MARCACION COMPLETA": comp,
                            "MARCACION PUBLICADA": pub,
                            "HORAS NOMINA": safeNum(hrs),
                            "OBSERVACION": finalStateForDoble,
                            "ESTADO DEL CONTRATO": estadoContratoActual,
                            "NÚMERO DE SEMANA": getWeekNumber(ymd),
                            "DÍA DE LA SEMANA": getDayOfWeekName(ymd),
                            "PROGRAMACION": progTurno || "",
                            "EJECUCION": ejecTurno || "",
                            "AUSENTISMO": ausentismoDia,
                            "NOVEDADES": "",
                            "MARCACION_PUBLICADA_GRTE": itemsDia.some(item => String(item?.source || "").toUpperCase() === "GRTE") ? "1" : "0",
                            "_FECHA_YMD": ymd,
                            "_FECHA_INGRESO_YMD": fCont || "",
                            "_FECHA_RETIRO_YMD": fTerm || ""
                        });
                    }
                }

                const descansosPorSemana = new Map();
                for (const row of finalRows) {
                    const key = `${row["DOCUMENTO"]}|${row["NÚMERO DE SEMANA"]}`;
                    if (String(row["NOVEDAD"] || "").toUpperCase().includes("DESCANSO")) {
                        descansosPorSemana.set(key, (descansosPorSemana.get(key) || 0) + 1);
                    }
                }
                const firstNoDescansoByWeek = new Map();
                for (const row of finalRows) {
                    const key = `${row["DOCUMENTO"]}|${row["NÚMERO DE SEMANA"]}`;
                    const descansoSemana = descansosPorSemana.get(key) || 0;
                    if (!shouldApplyNoDescansoSemana(row, descansoSemana, ymdStart, ymdEnd)) continue;
                    const fechaYmd = String(row["_FECHA_YMD"] || "");
                    const current = firstNoDescansoByWeek.get(key);
                    if (!current || fechaYmd < current) {
                        firstNoDescansoByWeek.set(key, fechaYmd);
                    }
                }
                for (const row of finalRows) {
                    const key = `${row["DOCUMENTO"]}|${row["NÚMERO DE SEMANA"]}`;
                    const descansoSemana = descansosPorSemana.get(key) || 0;
                    const markNoDescansoSemana = firstNoDescansoByWeek.get(key) === String(row["_FECHA_YMD"] || "");
                    row["NOVEDADES"] = buildRevisionNovedad(row, descansoSemana, markNoDescansoSemana);
                }

                finalRowsExport = finalRows.filter(row => {
                    if (onlyReviewMain && !/^Revisar/i.test(String(row["NOVEDADES"] || "").trim())) return false;
                    return true;
                });
                if (onlyReviewMain && finalRowsExport.length < finalRows.length) {
                    notifyWithSpinner(`Filtro "Solo novedades" activo: ${finalRowsExport.length} de ${finalRows.length} registros.`);
                }

                const pdvCountsMap = new Map();
                const novedadesPorPdvMap = new Map();
                const novedadesPorTipoMap = new Map();
                for (const row of finalRowsExport) {
                    const pdv = String(row["DESCRIPCION MARCACION"] || row["DESCRIPCION ASIGNADA"] || row["PUNTO DE VENTA MARCACION"] || row["PUNTO DE VENTA ASIGNADO"] || "Sin PDV").trim() || "Sin PDV";
                    pdvCountsMap.set(pdv, (pdvCountsMap.get(pdv) || 0) + 1);
                    const novedadTexto = String(row["NOVEDADES"] || "").trim();
                    if (novedadTexto && /^Revisar/i.test(novedadTexto)) {
                        novedadesPorPdvMap.set(pdv, (novedadesPorPdvMap.get(pdv) || 0) + 1);
                        const tipoNovedad = novedadTexto.replace(/^Revisar,\s*/i, "").trim() || "Sin clasificar";
                        novedadesPorTipoMap.set(tipoNovedad, (novedadesPorTipoMap.get(tipoNovedad) || 0) + 1);
                    }
                }
                summary.pdvCounts = Array.from(pdvCountsMap.entries())
                    .map(([pdv, count]) => ({ pdv, count }))
                    .sort((a, b) => b.count - a.count || a.pdv.localeCompare(b.pdv));
                summary.novedadesPorPdv = Array.from(novedadesPorPdvMap.entries())
                    .map(([pdv, count]) => ({ pdv, count }))
                    .sort((a, b) => b.count - a.count || a.pdv.localeCompare(b.pdv));
                summary.novedadesPorTipo = Array.from(novedadesPorTipoMap.entries())
                    .map(([tipo, count]) => ({ tipo, count }))
                    .sort((a, b) => b.count - a.count || a.tipo.localeCompare(b.tipo));
                summary.novedades = finalRowsExport.length;
                summary.descansos = descansosConsol.length;
                if (selected.rev && appendSheetIfData(wb, "3.RevisionDeDias", [HEADERS_CONSOL, ...finalRowsExport.map(row => HEADERS_CONSOL.map(h => normalizeSheetValue(row[h])))])) summary.hojas++;
                if (selected.rev && appendSheetIfData(wb, "5.Descansos", [HEADERS_DESCANSOS, ...descansosConsol.map(row => HEADERS_DESCANSOS.map(h => row[h] || ""))])) summary.hojas++;
            }

            if (selected.aus && appendSheetIfData(wb, "4.Ausentismos", [HEADERS_AUS, ...ausentismosRows])) {
                summary.hojas++;
            }

            if (selected.act && act && filteredActivosRows.length > 0) {
                setLoadingText("Armando hoja de activos...");
                const reviewCountByDoc = new Map();
                for (const row of finalRows) {
                    const doc = normalizeDocumentId(row["DOCUMENTO"]);
                    if (!doc) continue;
                    reviewCountByDoc.set(doc, (reviewCountByDoc.get(doc) || 0) + 1);
                }

                const headersActivos = INDICES_ACTIVOS_EXPORT.map(index => act[ROW_HEADER_ACTIVOS]?.[index] || "");
                const aoaActivos = [headersActivos];
                for (const row of filteredActivosRows) {
                    const visualRow = [...row];
                    visualRow[29] = formatDateFromActivos(row[29]);
                    visualRow[39] = formatDateFromActivos(row[39]);
                    visualRow[41] = formatDateFromActivos(row[41]);
                    aoaActivos.push(INDICES_ACTIVOS_EXPORT.map(index => normalizeSheetValue(visualRow[index])));
                }
                if (appendSheetIfData(wb, "1.Activos", aoaActivos)) summary.hojas++;
            }

            if (summary.hojas === 0) {
                drawSummaryChart(summary);
                notifyWithSpinner("No se generaron hojas con los filtros y archivos seleccionados.");
                return;
            }

            setLoadingText("Generando archivo Excel...");
            const totalNovedadesDetectadas = (summary.novedadesPorTipo || []).reduce((acc, item) => acc + Number(item.count || 0), 0);
            if (shouldPersistDashboardSummary) {
                await saveSummaryData({
                    generatedAt: new Date().toISOString(),
                    rangoRevision: {
                        inicio: ymdStart ? formatYMDtoDDMMYYYY(ymdStart) : "",
                        fin: ymdEnd ? formatYMDtoDDMMYYYY(ymdEnd) : ""
                    },
                    indicadores: {
                        activos: summary.activos || 0,
                        retirados: summary.retirados || 0,
                        novedades: totalNovedadesDetectadas,
                        marcaciones: summary.marcaciones || 0,
                        marcacionesCorregidasGrte: summary.marcacionesCorregidasGrte || 0,
                        ausentismos: summary.ausentismos || 0,
                        descansos: summary.descansos || 0
                    },
                    headers: HEADERS_CONSOL,
                    rows: finalRowsExport.map(row => {
                        const copy = {};
                        for (const header of HEADERS_CONSOL) copy[header] = normalizeSheetValue(row[header]);
                        copy["_FECHA_YMD"] = row["_FECHA_YMD"] || "";
                        copy["_FECHA_INGRESO_YMD"] = row["_FECHA_INGRESO_YMD"] || "";
                        copy["_FECHA_RETIRO_YMD"] = row["_FECHA_RETIRO_YMD"] || "";
                        copy["MARCACION_PUBLICADA_GRTE"] = row["MARCACION_PUBLICADA_GRTE"] || "0";
                        return copy;
                    }),
                    novedadesPorPdv: summary.novedadesPorPdv || [],
                    novedadesPorTipo: summary.novedadesPorTipo || [],
                    pdvTotales: summary.pdvCounts || []
                });
            } else {
                notifyWithSpinner("Ejecución parcial detectada: no se actualizó el resumen guardado del Dashboard.");
            }
            saveUiState();
            reorderWorkbookSheets(wb);
            XLSX.writeFile(wb, "Marcaciones_Revisadas.xlsx");
            drawSummaryChart(summary);
            notifyWithSpinner("Proceso finalizado correctamente.");
        } catch (error) {
            console.error(error);
            notifyWithSpinner("Ocurrió un error inesperado. Revise la consola.");
        }
    }, 10);
});

/* ----------------------------------------------------------
    LÃ“GICA PRINCIPAL DEL PROCESAMIENTO (btnRun)
---------------------------------------------------------- */
const legacyBtnRun = document.getElementById("btnRunLegacy");
if (legacyBtnRun) legacyBtnRun.addEventListener("click", async () => {
    showLoading();
    
    setTimeout(async () => {
        try {
            const dateStartInput = document.getElementById("dateStart").value;
            const dateEndInput = document.getElementById("dateEnd").value;
            
            if (!dateStartInput || !dateEndInput) {
                alert("Debe seleccionar la Fecha de Inicio y la Fecha de Fin de la RevisiÃ³n.");
                hideLoading();
                return;
            }

            const ymdStart = dateStartInput.replace(/-/g, "");
            const ymdEnd = dateEndInput.replace(/-/g, "");
            
            if(ymdStart > ymdEnd) {
                alert("La Fecha de Inicio no puede ser posterior a la Fecha de Fin.");
                hideLoading();
                return;
            }

            const fAct=document.getElementById("fileAct").files[0];
            const fRev=document.getElementById("fileRev").files[0];
            const fMarc=document.getElementById("fileMarc").files[0];
            const fAus=document.getElementById("fileAus").files[0]; 
            const fProg=document.getElementById("fileProg").files[0]; 
            
            if(!fAct||!fRev||!fMarc){
                alert("Debe subir al menos los 3 archivos base (Activos, Seguimiento, Marcaciones).");
                hideLoading();
                return;
            }

            const [act, rev, marc, ausRaw, progRaw]=await Promise.all([
                readSheetAOA(fAct),
                readSheetAOA(fRev),
                readSheetAOA(fMarc),
                fAus ? readSheetAOA(fAus) : Promise.resolve(null),
                fProg ? readSheetAOA(fProg) : Promise.resolve(null)
            ]);
            
            // --- 1. ACTIVOS Y MAPEO ---
            const limitInferiorYMD = getDateNMonthsAgoYMD(2);
            const limitSuperiorYMD = getDateNextMonthYMD();
            
            const activosMap=new Map();
            const filteredActivosRows = [];
            
            const COL_ESTADO_CONTRATO = 38; 
            const COL_FECHA_RETIRO = 39; 
            const ROW_HEADER = 3;

            for(let i=ROW_HEADER+1; i<act.length; i++){
                const r=act[i];
                const doc=r[0]?String(r[0]).trim():"";
                if(!doc) continue;

                // Aplicar NomPropio a los datos que vamos a usar luego
                r[1] = nomPropio(r[1]); // Nombre
                r[115] = nomPropio(r[115]); // Cargo
                r[23] = nomPropio(r[23]); // Desc PDV

                const fRetiroYMD = toYMDFromAny(r[COL_FECHA_RETIRO]);
                let include = !fRetiroYMD || (fRetiroYMD >= limitInferiorYMD && fRetiroYMD < limitSuperiorYMD);
                
                if(include) {
                    const existing = activosMap.get(doc);
                    if(!existing || (existing[COL_ESTADO_CONTRATO] !== 'A' && r[COL_ESTADO_CONTRATO] === 'A')) {
                        activosMap.set(doc, r);
                    }
                    filteredActivosRows.push(r);
                }
            }
            
            // --- 2. NOVEDADES (REV) ---
            const novedadMap = new Map();
            const headerRev = rev[0] || [];
            const startIdx = 7;
            const indicesFechasValidas = []; 
            for (let c = startIdx; c < headerRev.length; c++) {
                const enc = headerRev[c];
                const ymd = toYMDFromAny(enc);
                if (ymd && ymd >= ymdStart && ymd <= ymdEnd) {
                    indicesFechasValidas.push(c); 
                }
            }

            for (let i = 1; i < rev.length; i++) {
                const row = rev[i];
                if (!row) continue;
                const doc = row[0] ? String(row[0]).trim() : "";
                if (!doc) continue;
                for (const c of indicesFechasValidas) {
                    const enc = headerRev[c];
                    const ymd = toYMDFromAny(enc);
                    if (!ymd) continue;
                    const novRaw = row[c];
                    const nov = novRaw ? String(novRaw).trim() : "No identificado en Informe";
                    const llave = doc + ymd;
                    novedadMap.set(llave, nov);
                }
            }
            
            // --- 3. DESCANSOS (Usando NomPropio) ---
            const descansosConsol = []; 
            for (let i = 1; i < rev.length; i++) {
                const row = rev[i];
                if (!row) continue;
                const doc = row[0] ? String(row[0]).trim() : "";
                if (!doc) continue;
                
                const nombre = nomPropio(row[1] || ""); 
                const apellido = nomPropio(row[2] || ""); 
                const actRow = activosMap.get(doc) || [];

                for (let c = startIdx; c < row.length; c++) {
                    const enc = headerRev[c];
                    const ymd = toYMDFromAny(enc);
                    if (!ymd) continue; 
                    
                    const novRaw = row[c];
                    const nov = novRaw ? String(novRaw).trim() : "";
                    
                    if (nov.toUpperCase().includes("DESCANSO")) {
                        // ** CORRECCION LLAVE (Hojas 3,4,5): Usar ymdToExcelSerial **
                        const serialFecha = ymdToExcelSerial(ymd);
                        const llaveVisual = doc + serialFecha;

                        descansosConsol.push({
                            "LLAVE": llaveVisual,
                            "NOVEDAD": nomPropio(nov),
                            "FECHA": formatYMDtoDDMMYYYY(ymd),
                            "DOCUMENTO": doc,
                            "NOMBRE": nombre,
                            "APELLIDO": apellido,
                            "PUNTO DE VENTA ASIGNADO": actRow[22] || "", 
                            "DESCRIPCION ASIGNADA": actRow[23] || "" 
                        });
                    }
                }
            }

                        // --- 4. PROCESAMIENTO MARCACIONES CON LÃ“GICA MACRO (VBA Translation) ---
            const resultData = []; 
            const dictCedula = new Map(); 
            const dictFecha = new Map();  
            
            for (let i = 1; i < marc.length; i++) {
                const r = marc[i];
                if (!r || r.length < 6) continue;
                
                const cedula = r[4] ? String(r[4]).trim() : "";
                
                const val_eC = getExcelSerialDate(r[5]);
                const val_sC = getExcelSerialDate(r[6]);
                const val_eG = getExcelSerialDate(r[8]);
                const val_sG = getExcelSerialDate(r[9]);
                
                if (!cedula && val_eC===null && val_sC===null && val_eG===null && val_sG===null) continue;
                
                const isManagerRow = (val_eG !== null) || (val_sG !== null);
                let preferEntry = null;
                let preferExit = null;
                
                if (val_eG !== null) preferEntry = val_eG;
                else if (val_eC !== null) preferEntry = val_eC;
                
                if (val_sG !== null) preferExit = val_sG;
                else if (val_sC !== null) preferExit = val_sC;
                
                if (preferEntry === null) continue; 
                
                let obs = "0";
                if (preferExit === null) {
                    preferExit = preferEntry; 
                }
                
                const source = isManagerRow ? "GRTE" : "COL";
                const startTime = Number(preferEntry);
                const endTime = Number(preferExit);
                const horas = (endTime - startTime) * 24;
                
                const isNightShift = (Math.floor(startTime + 0.00001) !== Math.floor(endTime + 0.00001));
                if (isNightShift && obs === "0") {
                    obs = "Trabajo Nocturno";
                }
                
                const currentObj = {
                    rawRow: r,
                    cedula: cedula,
                    obs: obs,
                    entry: preferEntry,
                    exit: (val_sG !== null || val_sC !== null) ? preferExit : null, 
                    source: source,
                    startTime: startTime,
                    endTime: endTime,
                    isManager: isManagerRow,
                    horas: horas
                };
                
                const ymdCheck = formatExcelSerialToDateOnly(preferEntry);
                if (ymdCheck && (ymdCheck < ymdStart || ymdCheck > ymdEnd)) {
                    continue; 
                }
                
                resultData.push(currentObj);
            }

            // --- ORDENAR RESULTDATA POR FECHA Y CEDULA ---
            // --- ORDENAMIENTO DE RESULTDATA POR CEDULA Y FECHA ---
            resultData.sort((a, b) => {
                if (a.cedula !== b.cedula) return a.cedula.localeCompare(b.cedula);
                return a.startTime - b.startTime;
            });

            // Re-mapear el Ã­ndice y procesar duplicados tras el ordenamiento
            resultData.forEach((currentObj, rowIndex) => {
                currentObj.index = rowIndex;
                const cedula = currentObj.cedula;
                const startTime = currentObj.startTime;
                const endTime = currentObj.endTime;
                const source = currentObj.source;

                // Solapamientos
                if (!dictCedula.has(cedula)) dictCedula.set(cedula, []);
                const innerList = dictCedula.get(cedula);
                for (const stored of innerList) {
                    const objStored = resultData[stored.index];
                    if (startTime <= objStored.endTime && endTime >= objStored.startTime) {
                        if (currentObj.horas <= 0.2 && objStored.horas > 0.2) {
                             currentObj.obs = "MarcaciÃ³n Doble (Anular)";
                             if (objStored.obs === "0" || objStored.obs === "Trabajo Nocturno") objStored.obs = "MarcaciÃ³n Doble";
                        } else if (objStored.horas <= 0.2 && currentObj.horas > 0.2) {
                             objStored.obs = "MarcaciÃ³n Doble (Anular)";
                             if (currentObj.obs === "0" || currentObj.obs === "Trabajo Nocturno") currentObj.obs = "MarcaciÃ³n Doble";
                        } else {
                            if (objStored.source === "COL" && source === "GRTE") {
                                objStored.obs = "MarcaciÃ³n Doble";
                                currentObj.obs = "MarcaciÃ³n Doble (Anular)";
                            } else if (objStored.source === "GRTE" && source === "COL") {
                                currentObj.obs = "MarcaciÃ³n Doble";
                                objStored.obs = "MarcaciÃ³n Doble (Anular)";
                            } else {
                                if (objStored.obs === "0" || objStored.obs === "Trabajo Nocturno") objStored.obs = "MarcaciÃ³n Doble";
                                currentObj.obs = "MarcaciÃ³n Doble (Anular)";
                            }
                        }
                    }
                }
                innerList.push({startTime, endTime, index: rowIndex});

                // Duplicados Fecha
                const fechaEntrada = Math.floor(startTime + 0.00001); // CorrecciÃ³n para 00:00:00
                const claveFecha = cedula + "|" + fechaEntrada;
                if (!dictFecha.has(claveFecha)) dictFecha.set(claveFecha, []);
                const listFechas = dictFecha.get(claveFecha);
                if (listFechas.length > 0) {
                     for (const idx of listFechas) {
                         const objPrev = resultData[idx];
                         if (objPrev.obs === "0" || objPrev.obs === "Trabajo Nocturno") objPrev.obs = "MarcaciÃ³n Doble";
                     }
                     if (currentObj.obs === "0" || currentObj.obs === "Trabajo Nocturno") currentObj.obs = "MarcaciÃ³n Doble";
                }
                listFechas.push(rowIndex);
            });
            
            // ValidaciÃ³n final
            for (const item of resultData) {
                if (item.obs === "MarcaciÃ³n Doble") {
                    let h = (item.exit !== null) ? (Number(item.exit) - Number(item.entry)) * 24 : 0;
                    if (h <= 0.2) item.obs = "MarcaciÃ³n Doble (Anular)";
                }
            }
            
            // --- 5. PREPARAR DATOS FINALES ---
            const marcMap = new Map(); 
            const ejecMap = new Map(); 
            
            const headersDetalle = [
                "LLAVE", "PDV", "DESCRIPCION", "NOMBRES", "APELLIDOS", "CEDULA", 
                "ENTRA COLAB", "SALIDA COLAB", "HRS COLAB", "ENTRA GRTE", "SALIDA GRTE", "HRS GRTE", 
                "MARCAC.COMP", "MARCAC.PUB", "CONTRATO", "AUSENCIA", "HRS NOMINA", 
                "OBSERVACION", "ENTRADA_REAL", "SALIDA_REAL", "TOMADA_DE", "NOVEDAD"
            ];
            
            const marcacionesDetalleFinal = [headersDetalle];
            
            for (const item of resultData) {
                const raw = item.rawRow;
                raw[0] = raw[0]; raw[1] = nomPropio(raw[1]);
                raw[2] = nomPropio(raw[2]); raw[3] = nomPropio(raw[3]);
                
                const ymd = formatExcelSerialToDateOnly(item.entry);
                const llave = item.cedula + ymd;
                
                // HOJA 2: Marcaciones. AquÃ­ item.entry YA es un serial de excel, asÃ­ que usamos Math.floor directo.
                const llaveVisual = item.cedula + Math.floor(item.entry);

                if (!marcMap.has(llave)) marcMap.set(llave, []);
                marcMap.get(llave).push(item);
                
                const entrTime = getExcelSerialTime(item.entry);
                const exitTime = getExcelSerialTime(item.exit);
                if (entrTime && item.obs !== "MarcaciÃ³n Doble (Anular)") {
                    const descEjec = raw[1] || "";
                    const txt = (item.exit!==null) ? `${descEjec} ${entrTime} A ${exitTime}` : `${descEjec} ${entrTime} A`;
                    ejecMap.set(llave, txt); 
                }
                
                // Obtener novedad para Horarios
                const novHorarios = novedadMap.get(llave) || "";
                 
                const horasNominaNumerico = safeNum(raw[15]);

                marcacionesDetalleFinal.push([
                    llaveVisual, // Usamos llaveVisual
                    raw[0], raw[1], raw[2], raw[3], raw[4], 
                    raw[5], raw[6], raw[7], raw[8], raw[9], raw[10], 
                    raw[11], raw[12], raw[13], raw[14], 
                    horasNominaNumerico,
                    item.obs, 
                    formatExcelSerialToMacroString(item.entry), 
                    formatExcelSerialToMacroString(item.exit),  
                    item.source,
                    novHorarios
                ]);
            }

            // --- 6. PROGRAMACIÃ“N ---
            const progMap = new Map();
            const IDX_PROG = { DOC: 1, FECHA: 2, PDV_DESC: 8, ENTRA_TURNO: 6, SALIDA_TURNO: 7 };
            if (fProg && progRaw && progRaw.length > 1) {
                if (Array.isArray(progRaw[0]) && progRaw[0].length >= 9) {
                    for (let i = 1; i < progRaw.length; i++) {
                        const r = progRaw[i];
                        const doc = r[IDX_PROG.DOC] ? String(r[IDX_PROG.DOC]).trim() : "";
                        const ymd = toYMDFromAny(r[IDX_PROG.FECHA]);
                        if (doc && ymd && ymd >= ymdStart && ymd <= ymdEnd) {
                            const llave = doc + ymd;
                            const pdv = nomPropio(r[IDX_PROG.PDV_DESC] || ""); 
                            const entr = toHHMMFromAny(r[IDX_PROG.ENTRA_TURNO]);
                            const sal = toHHMMFromAny(r[IDX_PROG.SALIDA_TURNO]);
                            if (entr && sal) {
                                progMap.set(llave, `${pdv} ${entr} A ${sal}`);
                            } else if (pdv && (entr || sal)) {
                                 progMap.set(llave, `${pdv} ${entr || ""} A ${sal || ""}`);
                            }
                        }
                    }
                }
            }

            // --- 7. REVISION DE DIAS (Consolidado) ---
            const finalRows=[];
            for (let i = 1; i < rev.length; i++) {
                const row = rev[i];
                if (!row) continue;
                const doc = row[0] ? String(row[0]).trim() : "";
                if (!doc) continue;
                
                const nombre = nomPropio(row[1] || ""); 
                const apellido = nomPropio(row[2] || ""); 
                const contrato = row[3] || "";
                const cargo = nomPropio(row[4] || "");
                const fCont = toYMDFromAny(row[5]);
                const fTerm = toYMDFromAny(row[6]);
                const actRow = activosMap.get(doc) || [];
                
                for (const c of indicesFechasValidas) {
                    const enc = headerRev[c];
                    const ymd = toYMDFromAny(enc);
                    if (!ymd) continue;
                    
                    const novRaw = row[c];
                    const nov = novRaw ? String(novRaw).trim() : "No encontrado en Seguimiento";
                    const llaveReporte = doc + ymd;
                    const itemsDia = marcMap.get(llaveReporte) || [];
                    const tieneMarc = itemsDia.length > 0;

                    // ** CORRECCION LLAVE (HOJA 3) **
                    const serialFecha = ymdToExcelSerial(ymd); // Uso funcion segura
                    const llaveVisual = doc + serialFecha;
                    
                    let pv = actRow[22];
                    let desc = nomPropio(actRow[23] || "");
                    let comp = "0";
                    let pub = "0";
                    let hrs = 0;
                    let finalStateForDoble = "0";
                    let progTurno = progMap.get(llaveReporte); 
                    let ejecTurno = ejecMap.get(llaveReporte); 

                    if (String(novRaw || "").toUpperCase().includes("DESCANSO")) {
                        progTurno = "Descanso";
                        ejecTurno = "Descanso";
                    } else {
                        if (!progTurno) { 
                            progTurno = "No programado";
                            ejecTurno = "No ejecutado";
                        } else if (!ejecTurno) {
                            ejecTurno = "No ejecutado";
                        }
                    }

                    if(tieneMarc){
                        // Tomamos datos del primer registro (o principal)
                        const first = itemsDia[0];
                        pv = first.rawRow[0] || actRow[22];
                        desc = nomPropio(first.rawRow[1] || actRow[23] || "");
                        comp = first.rawRow[11] || "0";
                        pub = first.rawRow[12] || "0";
                        
                        // SUMA DE HORAS
                        hrs = itemsDia.reduce((acc, item) => acc + safeNum(item.rawRow[15] || 0), 0);
                        
                        const estados = itemsDia.map(it => it.obs);
                        
                        // LÃ³gica de estado consolidado
                        if (estados.includes("MarcaciÃ³n Doble (Anular)")) finalStateForDoble = "MarcaciÃ³n Doble (Anular)";
                        else if (estados.includes("MarcaciÃ³n Doble")) finalStateForDoble = "MarcaciÃ³n Doble";
                        else if (estados.includes("Trabajo Nocturno")) finalStateForDoble = "Trabajo Nocturno";
                        else finalStateForDoble = "0";

                    } else {
                         finalStateForDoble = "0";
                    }
                    pv = getPdvAgrupado(pv);
                    
                    finalRows.push({
                        "LLAVE": llaveVisual, // Usamos llaveVisual
                        "NOVEDAD": nov,
                        "FECHA": formatYMDtoDDMMYYYY(ymd),
                        "DOCUMENTO": doc,
                        "NOMBRE": nombre,
                        "APELLIDO": apellido,
                        "PUNTO DE VENTA ASIGNADO": actRow[22],
                        "DESCRIPCION ASIGNADA": nomPropio(actRow[23] || ""),
                        "CENTRO DE COSTOS": actRow[21] || "",
                        "CONTRATO": contrato,
                        "CARGO": cargo,
                        "FECHA DE CONTRATACION": fCont ? formatYMDtoDDMMYYYY(fCont) : "",
                        "FECHA DE TERMINACION": fTerm ? formatYMDtoDDMMYYYY(fTerm) : "",
                        "PUNTO DE VENTA MARCACION": pv,
                        "DESCRIPCION MARCACION": desc,
                        "MARCACION COMPLETA": comp,
                        "MARCACION PUBLICADA": pub,
                        "HORAS NOMINA": safeNum(hrs),
                        "OBSERVACION": finalStateForDoble,
                        "ESTADO DEL CONTRATO": actRow[38] || "",
                        "NÃšMERO DE SEMANA": getWeekNumber(ymd),
                        "DÃA DE LA SEMANA": getDayOfWeekName(ymd),
                        "PROGRAMACION": progTurno, 
                        "EJECUCION": ejecTurno 
                    });
                }
            }
            
            // --- 8. AUSENTISMOS ---
            const TERMINATION_TYPES = ["Renuncia", "Renuncia Bienestar", "TerminaciÃ³n de contrato"];
            const terminationDates = new Map(); 
            const dailyAusentismos = new Map();
            const overlapAusTypesByKey = new Map();
            let ausentismosRows = [];

            const headerAusentismos = [
                "LLAVE", "TIPO DE AUSENTISMO", "NOMBRE EMPLEADO", "CÃ‰DULA", "SUCURSAL", 
                "FECHA INICIO", "FECHA FIN", "DÃAS SOLICITADOS", "HORAS TOTALES", 
                "ESTADO", "PASO ACTUAL", "OBSERVACION" 
            ];
            
            const IDX_AUS = {
                TIPO_AUS: 0, NOMBRE_EMP: 1, CEDULA: 2, SUCURSAL: 3, FECHA_INICIO: 4,
                FECHA_FIN: 5, DIAS_SOL: 6, HORAS_TOT: 7, ESTADO: 8, PASO_ACT: 9
            };
            
            if (fAus && ausRaw && ausRaw.length > 1) {
                if (Array.isArray(ausRaw[0]) && ausRaw[0].length >= 18) {
                     for (let i = 1; i < ausRaw.length; i++) {
                        const r = ausRaw[i];
                        if(r.length < 18) continue; 
                        
                        const estadoRaw = String(r[IDX_AUS.ESTADO] || "").trim();
                        const esEstadoValido = (estadoRaw === "Pendiente" || estadoRaw === "Aprobada" || estadoRaw === "Finalizada" || estadoRaw === "Aprobada y Finalizada");
                        
                        if (!esEstadoValido) continue;

                        const doc = r[IDX_AUS.CEDULA] ? String(r[IDX_AUS.CEDULA]).trim() : "";
                        if (!doc) continue;

                        const fechaInicioYMD = toYMDFromAny(r[IDX_AUS.FECHA_INICIO]);
                        const fechaFinYMD = toYMDFromAny(r[IDX_AUS.FECHA_FIN]);
                        
                        if (fechaInicioYMD && fechaFinYMD) {
                            const tipoAusentismoRaw = String(r[IDX_AUS.TIPO_AUS] || "").trim();
                            if (TERMINATION_TYPES.includes(tipoAusentismoRaw)) { 
                                const existingTermination = terminationDates.get(doc);
                                if (!existingTermination || fechaInicioYMD < existingTermination.ymd) {
                                    terminationDates.set(doc, { ymd: fechaInicioYMD, rowData: r });
                                }
                            }

                            let d = new Date(fechaInicioYMD.slice(0,4), fechaInicioYMD.slice(4,6) - 1, fechaInicioYMD.slice(6,8));
                            const end = new Date(fechaFinYMD.slice(0,4), fechaFinYMD.slice(4,6) - 1, fechaFinYMD.slice(6,8));

                            while (d <= end) {
                                const ymd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
                                if (ymd >= ymdStart) { 
                                     const fechaDesglosadaFmt = formatYMDtoDDMMYYYY(ymd);
                                     const nuevaLlave = doc + ymd;
                                     if (!TERMINATION_TYPES.includes(tipoAusentismoRaw)) {
                                         if (!overlapAusTypesByKey.has(nuevaLlave)) overlapAusTypesByKey.set(nuevaLlave, new Set());
                                         overlapAusTypesByKey.get(nuevaLlave).add(tipoAusentismoRaw);
                                     }

                                     // ** CORRECCION LLAVE (HOJA 4) **
                                     const serialFecha = ymdToExcelSerial(ymd);
                                     const llaveVisual = doc + serialFecha;

                                     const baseRow = [
                                         llaveVisual, // Usamos llaveVisual
                                         r[IDX_AUS.TIPO_AUS] || "", nomPropio(r[IDX_AUS.NOMBRE_EMP] || ""), r[IDX_AUS.CEDULA] || "", nomPropio(r[IDX_AUS.SUCURSAL] || ""),
                                         fechaDesglosadaFmt, fechaDesglosadaFmt, 
                                         r[IDX_AUS.DIAS_SOL] || "", r[IDX_AUS.HORAS_TOT] || "", r[IDX_AUS.ESTADO] || "",
                                         r[IDX_AUS.PASO_ACT] || "", ""
                                     ];
                                     dailyAusentismos.set(nuevaLlave, baseRow);
                                }
                                d.setDate(d.getDate() + 1);
                            }
                        }
                     }
                }
            }
            
            const userEndLimitDate = new Date(ymdEnd.slice(0,4), ymdEnd.slice(4,6) - 1, ymdEnd.slice(6,8));
            for (const [doc, termInfo] of terminationDates.entries()) {
                const termYMD = termInfo.ymd;
                const rawRow = termInfo.rowData;
                const tipoAusentismo = String(rawRow[IDX_AUS.TIPO_AUS] || "").trim();
                const cedula = rawRow[IDX_AUS.CEDULA] || "";
                const nombre = nomPropio(rawRow[IDX_AUS.NOMBRE_EMP] || "");
                const sucursal = nomPropio(rawRow[IDX_AUS.SUCURSAL] || "");
                const termDate = new Date(termYMD.slice(0,4), termYMD.slice(4,6) - 1, termYMD.slice(6,8));
                const propagationEndLimitDate = new Date(termDate);
                propagationEndLimitDate.setDate(termDate.getDate() + 30); 
                const finalPropagationDate = propagationEndLimitDate < userEndLimitDate ? propagationEndLimitDate : userEndLimitDate;
                let d = new Date(termDate); 
                while (d <= finalPropagationDate) {
                    const ymd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
                    const llave = doc + ymd;
                    if (ymd >= ymdStart && !dailyAusentismos.has(llave)) {
                        const fechaDesglosadaFmt = formatYMDtoDDMMYYYY(ymd);

                        // ** CORRECCION LLAVE (HOJA 4 - Propagacion) **
                        const serialFecha = ymdToExcelSerial(ymd);
                        const llaveVisual = doc + serialFecha;

                        const syntheticRow = [
                             llaveVisual, // Usamos llaveVisual
                             tipoAusentismo, nombre, cedula, sucursal,
                             fechaDesglosadaFmt, fechaDesglosadaFmt, "1", "8", 
                             "Propagado - " + tipoAusentismo, rawRow[IDX_AUS.PASO_ACT] || "", ""
                        ];
                        dailyAusentismos.set(llave, syntheticRow);
                    }
                    d.setDate(d.getDate() + 1);
                }
            }
            for (const [llave, tiposSet] of overlapAusTypesByKey.entries()) {
                if (!dailyAusentismos.has(llave) || tiposSet.size <= 1) continue;
                const observacion = `Ausentismos sobrepuestos con: ${Array.from(tiposSet).join(" | ")}`;
                const rowAus = dailyAusentismos.get(llave);
                rowAus[11] = observacion;
            }
            ausentismosRows = Array.from(dailyAusentismos.values());
            
            // --- 9. GENERAR EXCEL ---
            const wb = XLSX.utils.book_new();
            
            // Hoja 5
            const headersDescansos = Object.keys(descansosConsol[0] || {});
            const aoaDescansos = [headersDescansos];
            for (const r of descansosConsol) aoaDescansos.push(headersDescansos.map(h => r[h]));
            let wsDescansos = null;
            if (descansosConsol.length > 0) wsDescansos = XLSX.utils.aoa_to_sheet(aoaDescansos);
            
            // Hoja 3
            const headersConsol = Object.keys(finalRows[0]);
            const aoaConsol = [headersConsol];
            for (const r of finalRows) aoaConsol.push(headersConsol.map(h => r[h]));
            const wsConsol = XLSX.utils.aoa_to_sheet(aoaConsol);
            
            // Hoja 2 (Ya preparada)
            let wsDetalle = null;
            if (marcacionesDetalleFinal.length > 1) {
                wsDetalle = XLSX.utils.aoa_to_sheet(marcacionesDetalleFinal);
            }
            
            // Hoja 1
            const indicesActivos = [0, 1, 21, 22, 23, 29, 38, 39, 41, 115];
            filteredActivosRows.sort((a, b) => {
                const valA = String(a[COL_ESTADO_CONTRATO] || "").toUpperCase();
                const valB = String(b[COL_ESTADO_CONTRATO] || "").toUpperCase();
                return valA < valB ? -1 : (valA > valB ? 1 : 0);
            });
            const aoaActivosFinal = [];
            if (act.length > ROW_HEADER) aoaActivosFinal.push(indicesActivos.map(colIndex => act[ROW_HEADER][colIndex]));
            for(const row of filteredActivosRows) {
                const visualRow = [...row];
                visualRow[29] = formatDateFromActivos(row[29]);
                visualRow[39] = formatDateFromActivos(row[39]);
                visualRow[41] = formatDateFromActivos(row[41]);
                aoaActivosFinal.push(indicesActivos.map(colIndex => visualRow[colIndex] || ""));
            }
            let wsActivos = null;
            if (aoaActivosFinal.length > 0) wsActivos = XLSX.utils.aoa_to_sheet(aoaActivosFinal);

            // Hoja 4
            let wsAus = null;
            if (ausentismosRows.length > 0) {
                 const aoaAus = [headerAusentismos, ...ausentismosRows];
                 wsAus = XLSX.utils.aoa_to_sheet(aoaAus);
            }
            
            const sheetsToAppend = [
                { name: "1.Activos", ws: wsActivos, hasData: aoaActivosFinal.length > 1 },
                { name: "2.Horarios", ws: wsDetalle, hasData: marcacionesDetalleFinal.length > 1 },
                { name: "3.RevisionDeDias", ws: wsConsol, hasData: finalRows.length > 0 },
                { name: "4.Ausentismos", ws: wsAus, hasData: ausentismosRows.length > 0 },
                { name: "5.Descansos", ws: wsDescansos, hasData: descansosConsol.length > 0 }
            ];
            sheetsToAppend.reverse(); 

            for (const sheet of sheetsToAppend) {
                if (sheet.ws && sheet.hasData) {
                    XLSX.utils.book_append_sheet(wb, sheet.ws, sheet.name);
                }
            }

            XLSX.writeFile(wb, "Marcaciones_Revisadas.xlsx");
            hideLoading();
            
        } catch (error) {
            console.error(error);
            hideLoading();
            alert("âŒ Â¡OcurriÃ³ un error inesperado! Revise la consola.");
        }
    }, 10);
});

// Helper de fechas Excel Serial -> Date only format
function formatExcelSerialToDateOnly(serial) {
    if (serial === null || serial === 0) return null;
    let n = Number(serial);
    if (isNaN(n)) return null;
    n += (0.1 / (24 * 60 * 60 * 1000));
    const serialDia = Math.floor(n);
    let ms = (serialDia - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if (isNaN(d)) return null;
    const pad = (num) => String(num).padStart(2, '0');
    return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}`;
}
function getExcelSerialTime(serialDateTime) {
    if (serialDateTime === null || serialDateTime === 0) return null;
    let n = Number(serialDateTime);
    if (isNaN(n)) return null;
    const timeFraction = n - Math.floor(n);
    if (timeFraction === 0) return "00:00"; 
    let totalSeconds = Math.round(timeFraction * 86400);
    let hours = Math.floor(totalSeconds / 3600) % 24;
    let minutes = Math.floor((totalSeconds % 3600) / 60);
    const pad = (num) => String(num).padStart(2, '0');
    return `${pad(hours)}:${pad(minutes)}`;
}
})();

(function(){ if(!document.body.classList.contains('page-dashboard')) return;
if (!AuthGate.requireLocalSession()) return;
AuthGate.verifyRemoteOrLogout();
const APP_DB_NAME = "ProcesadorMarcacionesDB";
const APP_DB_VERSION = 1;
const APP_STORE_NAME = "appState";
const APP_SUMMARY_KEY = "summary";
const APP_IMPORTED_SUMMARY_KEY = "summary_imported";
const APP_DASHBOARD_STATE_KEY = "dashboard_state";

function openAppDb() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(APP_DB_NAME, APP_DB_VERSION);
        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(APP_STORE_NAME)) {
                db.createObjectStore(APP_STORE_NAME);
            }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function loadSummary() {
    try {
        const db = await openAppDb();
        const result = await new Promise((resolve, reject) => {
            const tx = db.transaction(APP_STORE_NAME, "readonly");
            const request = tx.objectStore(APP_STORE_NAME).get(APP_SUMMARY_KEY);
            request.onsuccess = () => resolve(request.result || null);
            request.onerror = () => reject(request.error);
        });
        db.close();
        return result;
    } catch (error) {
        console.warn("No se pudo leer el resumen guardado.", error);
        return null;
    }
}

async function saveSummary(summary) {
    try {
        const db = await openAppDb();
        await new Promise((resolve, reject) => {
            const tx = db.transaction(APP_STORE_NAME, "readwrite");
            tx.objectStore(APP_STORE_NAME).put(summary, APP_SUMMARY_KEY);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error || new Error("No se pudo guardar el resumen."));
            tx.onabort = () => reject(tx.error || new Error("Guardado abortado."));
        });
        db.close();
    } catch (error) {
        console.warn("No se pudo guardar el resumen importado.", error);
    }
}

async function loadImportedSummary() {
    try {
        const db = await openAppDb();
        const result = await new Promise((resolve, reject) => {
            const tx = db.transaction(APP_STORE_NAME, "readonly");
            const request = tx.objectStore(APP_STORE_NAME).get(APP_IMPORTED_SUMMARY_KEY);
            request.onsuccess = () => resolve(request.result || null);
            request.onerror = () => reject(request.error);
        });
        db.close();
        return result;
    } catch (error) {
        console.warn("No fue posible recuperar el importado acumulado.", error);
        return null;
    }
}

async function saveImportedSummary(summary) {
    try {
        const db = await openAppDb();
        await new Promise((resolve, reject) => {
            const tx = db.transaction(APP_STORE_NAME, "readwrite");
            tx.objectStore(APP_STORE_NAME).put(summary, APP_IMPORTED_SUMMARY_KEY);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
        db.close();
    } catch (error) {
        console.warn("No fue posible guardar el importado acumulado.", error);
    }
}

async function saveDashboardUiState(state) {
    try {
        const db = await openAppDb();
        await new Promise((resolve, reject) => {
            const tx = db.transaction(APP_STORE_NAME, "readwrite");
            tx.objectStore(APP_STORE_NAME).put(state, APP_DASHBOARD_STATE_KEY);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error || new Error("No se pudo guardar estado UI."));
            tx.onabort = () => reject(tx.error || new Error("Guardado UI abortado."));
        });
        db.close();
    } catch (error) {
        console.warn("No se pudo guardar estado del dashboard.", error);
    }
}

async function loadDashboardUiState() {
    try {
        const db = await openAppDb();
        const result = await new Promise((resolve, reject) => {
            const tx = db.transaction(APP_STORE_NAME, "readonly");
            const request = tx.objectStore(APP_STORE_NAME).get(APP_DASHBOARD_STATE_KEY);
            request.onsuccess = () => resolve(request.result || null);
            request.onerror = () => reject(request.error);
        });
        db.close();
        return result;
    } catch (error) {
        console.warn("No se pudo leer estado del dashboard.", error);
        return null;
    }
}

function isAdmPdv(row) {
    const values = [
        row["PUNTO DE VENTA MARCACION"],
        row["DESCRIPCION MARCACION"]
    ].map(value => String(value || "").toUpperCase());
    return values.some(value => value === "ADM" || value.includes(" PDV ADM") || value.includes("PDV ADM") || value.startsWith("ADM "));
}

function normalizeComparable(value) {
    return String(value || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

function normalizePdvCode(value) {
    const code = normalizeComparable(value);
    if (code === "PPP1" || code === "PPP2" || code === "PPP") return "PPP";
    if (code === "PPP3") return "PPH";
    if (code === "LOG1" || code === "LOG2") return "LOG";
    return code;
}

function canonicalizePdvDescription(value) {
    const normalized = normalizeComparable(value);
    if ([
        "PLANTA DE PRODUCCION",
        "DIRECCION DE ALIMENTOS PRODUCCION PPP"
    ].includes(normalized)) {
        return "Planta De Producción";
    }
    if ([
        "PLANTA DE HELADOS",
        "PLANTA DE PRODUCCION HELADOS"
    ].includes(normalized)) {
        return "Planta De Helados";
    }
    return toTitleCase(normalized.toLocaleLowerCase("es-CO"));
}

function getAssignedPdv(row) {
    const codigo = normalizePdvCode(row["PUNTO DE VENTA MARCACION"] || row["PUNTO DE VENTA ASIGNADO"] || "");
    const descripcionBase = String(row["DESCRIPCION MARCACION"] || row["DESCRIPCION ASIGNADA"] || "").trim();
    const descripcion = canonicalizePdvDescription(descripcionBase);
    if (!codigo) return "";
    if (codigo === "PPP" || descripcion === "Planta De Producción") return "PPP | Planta De Producción";
    if (codigo === "PPH" || descripcion === "Planta De Helados") return "PPH | Planta De Helados";
    if (codigo && descripcion) return `${codigo} | ${descripcion}`;
    return codigo || "";
}

function getPdvDescriptionFromLabel(label) {
    const parts = String(label || "").split("|");
    return (parts.length > 1 ? parts.slice(1).join("|") : parts[0]).trim();
}

function toTitleCase(value) {
    return String(value || "")
        .toLocaleLowerCase("es-CO")
        .replace(/\s+/g, " ")
        .trim()
        .replace(/(^|[\s(])([a-záéíóúñü])/g, (_, prefix, char) => prefix + char.toLocaleUpperCase("es-CO"));
}

function normalizeNovedadLabel(value) {
    const text = String(value || "").trim().replace(/^Revisar,\s*/i, "");
    if (!text) return "Sin dato";
    if (/^marcación doble de/i.test(text)) return "Marcaciones Dobles";
    if (/^marcación de/i.test(text)) {
        const match = text.match(/de\s+([0-9]+(?:[.,][0-9]+)?)/i);
        const hours = match ? Number(String(match[1]).replace(",", ".")) : NaN;
        if (!Number.isNaN(hours) && hours <= 4) return "Marcaciones Menores A 4 Horas";
        if (!Number.isNaN(hours) && hours > 12.5) return "Marcaciones Mayores A 12.5 Horas";
    }
    return toTitleCase(text);
}

function isOkNovedad(value) {
    return /^Ok\s*\(/i.test(String(value || "").trim());
}

function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, char => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
    }[char]));
}

function filterRows(rows) {
    const includeOk = document.getElementById("chkIncludeOk").checked;
    const includeAdm = document.getElementById("chkIncludeAdm").checked;
    const includeDouble = document.getElementById("chkIncludeDouble").checked;
    const includeNocturno = document.getElementById("chkIncludeNocturno").checked;
    const selectedPdv = document.getElementById("pdvSelector")?.value || "__ALL__";
    return rows.filter(row => {
        if (!includeOk && isOkNovedad(row["NOVEDADES"])) return false;
        if (!includeAdm && isAdmPdv(row)) return false;
        if (!includeDouble && normalizeComparable(row["OBSERVACION"]) === "MARCACION DOBLE") return false;
        if (!includeNocturno && normalizeComparable(row["OBSERVACION"]) === "TRABAJO NOCTURNO") return false;
        if (selectedPdv !== "__ALL__" && getAssignedPdv(row) !== selectedPdv) return false;
        return true;
    });
}

function groupBy(rows, field) {
    const map = new Map();
    for (const row of rows) {
        const rawLabel = String(row[field] || "Sin dato").trim() || "Sin dato";
        const key = normalizeComparable(rawLabel);
        const current = map.get(key) || { label: rawLabel, count: 0 };
        current.count += 1;
        if (rawLabel.length < current.label.length) current.label = rawLabel;
        map.set(key, current);
    }
    return Array.from(map.values())
        .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
}

function renderBars(container, items) {
    if (!container) return;
    if (!items.length) {
        container.innerHTML = '<div class="empty">No hay datos para mostrar con los filtros actuales.</div>';
        return;
    }
    const max = Math.max(1, ...items.map(item => Number(item.count || 0)));
    const rowsHtml = items.map(item => `
        <div class="bar-row">
            <div class="bar-label">${escapeHtml(item.label)}</div>
            <div class="bar-track"><div class="bar-fill" style="width:${(Number(item.count || 0) / max) * 100}%"></div></div>
            <div class="bar-value">${Number(item.count || 0)}</div>
            <div class="bar-rate">${item.ratePct === undefined ? "" : `${Number(item.ratePct).toFixed(1)}%`}</div>
        </div>
    `).join("");
    container.innerHTML = `
        <div class="pdv-grid-head">
            <span>Punto de venta</span>
            <span></span>
            <span>Total</span>
            <span>Tasa</span>
        </div>
        <div class="pdv-grid-rows">${rowsHtml}</div>
    `;
}

function renderVerticalBars(container, items, asPercent = false) {
    if (!container) return;
    container.classList.toggle("dense-chart", items.length > 18);
    if (!items.length) {
        container.innerHTML = '<div class="empty">No hay datos para mostrar con los filtros actuales.</div>';
        return;
    }
    const max = Math.max(1, ...items.map(item => Number(item.count || 0)));
    container.innerHTML = items.map(item => `
        <div class="vbar-col">
            <div class="vbar-plot">
                <div class="vbar-stack">
                    <div class="vbar-value">${Number(item.count || 0) % 1 === 0 ? Number(item.count || 0) : Number(item.count || 0).toFixed(1)}${asPercent ? "%" : ""}</div>
                    <div class="vbar" style="height:${Math.max(10, (Number(item.count || 0) / max) * 205)}px"></div>
                </div>
            </div>
            <div class="vbar-label">${escapeHtml(getPdvDescriptionFromLabel(item.label))}</div>
        </div>
    `).join("");
}

function renderTable(container, items, selectedLabel) {
    if (!container) return;
    if (!items.length) {
        container.innerHTML = '<tr><td colspan="3" class="empty">No hay datos para la tabla.</td></tr>';
        return;
    }
    container.innerHTML = items.slice(0, 15).map(item => `
        <tr data-novedad="${escapeHtml(item.label)}" class="${item.label === selectedLabel ? "active" : ""}" title="Ver puntos de venta de esta novedad">
            <td>${escapeHtml(item.label)}</td>
            <td><strong>${item.count}</strong></td>
            <td><strong>${Number(item.sharePct || 0).toFixed(1)}%</strong></td>
        </tr>
    `).join("");
}

let lastFilteredRows = [];
let lastBaseRows = [];
let lastBaseRowsForDonut = [];
let lastMainData = [];
let lastPdvData = [];
let lastTiposData = [];
let selectedNovedadLabel = "";
let selectedDonutMode = false;
let lastSelectedPdv = "__ALL__";
let currentDashboardData = null;
let lastDashboardUiState = null;

function hasMarcacionRow(row) {
    const completa = String(row["MARCACION COMPLETA"] || "").trim().toUpperCase();
    const horas = Number(row["HORAS NOMINA"] || 0);
    return completa === "SI" || horas > 0;
}

function computeRates(rows, fallbackTotals = {}) {
    const totalSeguimiento = rows.length;
    const totalNovedad = rows.filter(row => !isOkNovedad(row["NOVEDADES"])).length;
    const totalOk = Math.max(0, totalSeguimiento - totalNovedad);
    const efectividadPct = totalSeguimiento > 0 ? (totalOk / totalSeguimiento) * 100 : 0;
    const marcacionesRows = rows.filter(hasMarcacionRow).length;
    const totalMarcaciones = marcacionesRows > 0 ? marcacionesRows : Number(fallbackTotals.marcaciones || 0);
    const corregidasRows = rows.filter(row => String(row["MARCACION_PUBLICADA_GRTE"] || "").trim() === "1").length;
    const totalCorregidasGrte = Math.max(corregidasRows, Number(fallbackTotals.marcacionesCorregidasGrte || 0));
    const correccionPct = totalMarcaciones > 0 ? (totalCorregidasGrte / totalMarcaciones) * 100 : 0;
    return { totalSeguimiento, totalOk, efectividadPct, totalMarcaciones, totalCorregidasGrte, correccionPct };
}

function getSelectedChartId() {
    const checked = document.querySelector(".chart-export-selector:checked");
    return checked?.dataset?.chartId || "";
}

function setupExportSelection() {
    document.querySelectorAll(".chart-export-selector").forEach(chk => {
        chk.addEventListener("change", () => {
            if (!chk.checked) return;
            document.querySelectorAll(".chart-export-selector").forEach(other => {
                if (other !== chk) other.checked = false;
            });
        });
    });
}

function setPrimarySelectionMode(mode) {
    const pdvSelectorEl = document.getElementById("pdvSelector");
    if (mode !== "novedad") selectedNovedadLabel = "";
    if (mode !== "donut") selectedDonutMode = false;
    if (mode !== "pdv" && pdvSelectorEl) {
        pdvSelectorEl.value = "__ALL__";
        lastSelectedPdv = "__ALL__";
    }
}

function setDonut(idBase, value, total, pctOverride = null) {
    const safeTotal = Math.max(0, Number(total || 0));
    const safeValue = Math.max(0, Number(value || 0));
    const pct = pctOverride === null
        ? (safeTotal > 0 ? (safeValue / safeTotal) * 100 : 0)
        : Math.max(0, Number(pctOverride || 0));
    const deg = Math.max(0, Math.min(360, (pct / 100) * 360));
    const donut = document.getElementById(`donut${idBase}`);
    const pctNode = document.getElementById(`donut${idBase}Pct`);
    const metaNode = document.getElementById(`donut${idBase}Meta`);
    if (donut) donut.style.background = `conic-gradient(var(--primary) ${deg}deg, #dce8ff ${deg}deg 360deg)`;
    if (pctNode) pctNode.textContent = `${pct.toFixed(1)}%`;
    if (metaNode) metaNode.textContent = `${safeValue} de ${safeTotal}`;
}

function setDonutByIds(donutId, pctId, metaId, value, total, pctOverride = null) {
    const safeTotal = Math.max(0, Number(total || 0));
    const safeValue = Math.max(0, Number(value || 0));
    const pct = pctOverride === null
        ? (safeTotal > 0 ? (safeValue / safeTotal) * 100 : 0)
        : Math.max(0, Number(pctOverride || 0));
    const deg = Math.max(0, Math.min(360, (pct / 100) * 360));
    const donut = document.getElementById(donutId);
    const pctNode = document.getElementById(pctId);
    const metaNode = document.getElementById(metaId);
    if (donut) donut.style.background = `conic-gradient(var(--primary) ${deg}deg, #dce8ff ${deg}deg 360deg)`;
    if (pctNode) pctNode.textContent = `${pct.toFixed(1)}%`;
    if (metaNode) metaNode.textContent = `${safeValue} de ${safeTotal}`;
}

function exportCurrentExcel(data) {
    const rows = selectedNovedadLabel
        ? lastBaseRows.filter(row => normalizeNovedadLabel(row["NOVEDADES"]) === selectedNovedadLabel)
        : (lastFilteredRows.length ? lastFilteredRows : []);
    const headers = Array.isArray(data.headers) && data.headers.length
        ? data.headers
        : Object.keys(rows[0] || {}).filter(key => !key.startsWith("_"));
    const aoa = [headers, ...rows.map(row => headers.map(header => {
        if (header === "NOVEDADES") return normalizeNovedadLabel(row[header]);
        return row[header] ?? "";
    }))];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "RevisionFiltrada");
    XLSX.writeFile(wb, "Dashboard_Filtrado.xlsx");
}

function findRevisionSheetName(workbook) {
    return (workbook?.SheetNames || []).find(name => {
        const normalized = normalizeComparable(name).replace(/\s+/g, "");
        return normalized === "3.REVISIONDEDIAS" || normalized === "3.REVISIONDEDIAS".replace(/\./g, "") || normalized.endsWith("REVISIONDEDIAS");
    }) || "";
}

async function importProcessedWorkbooks(files) {
    const selectedFiles = Array.from(files || []).filter(Boolean);
    if (!selectedFiles.length) return null;

    let headers = [];
    const rows = [];
    const importedNames = [];

    for (const file of selectedFiles) {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const sheetName = findRevisionSheetName(workbook);
        if (!sheetName) continue;
        const sheet = workbook.Sheets[sheetName];
        const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        if (!aoa.length) continue;

        const localHeaders = aoa[0].map(value => String(value || "").trim());
        if (!headers.length) headers = localHeaders;
        const effectiveHeaders = headers.length ? headers : localHeaders;

        for (let i = 1; i < aoa.length; i++) {
            const rowValues = aoa[i];
            if (!Array.isArray(rowValues)) continue;
            if (!rowValues.some(value => String(value || "").trim() !== "")) continue;
            const row = {};
            for (let c = 0; c < effectiveHeaders.length; c++) {
                row[effectiveHeaders[c]] = rowValues[c] ?? "";
            }
            rows.push(row);
        }
        importedNames.push(file.name);
    }

    if (!headers.length || !rows.length) return null;

    const uniqueDocs = new Set(rows.map(row => String(row["DOCUMENTO"] || "").trim()).filter(Boolean));
    return {
        headers,
        rows,
        generatedAt: new Date().toISOString(),
        rangoRevision: {
            inicio: "Importado",
            fin: importedNames.join(", ")
        },
        indicadores: {
            activos: uniqueDocs.size,
            retirados: 0,
            novedades: rows.length,
            marcaciones: 0,
            ausentismos: 0,
            descansos: 0
        }
    };
}

function drawPdfHeader(pdf, pageWidth, margin, title, subtitle, highlightData) {
    const rightBoxWidth = 72;
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(18);
    pdf.text("Dashboard", margin, 16);
    pdf.setFontSize(13);
    pdf.text(title || "Totales", margin, 24);
    pdf.setFontSize(9);
    pdf.setTextColor(98, 116, 138);
    pdf.text("Total de novedades", pageWidth - margin, 16, { align: "right" });
    pdf.text(highlightData?.rateLabel || "Tasa de efectividad general", pageWidth - margin, 29, { align: "right" });
    pdf.setFont("helvetica", "bold");
    pdf.setTextColor(29, 94, 220);
    pdf.setFontSize(20);
    pdf.text(String(highlightData?.totalNovedades ?? 0), pageWidth - margin, 22, { align: "right" });
    pdf.text(`${Number(highlightData?.efectividadPct ?? 0).toFixed(1)}%`, pageWidth - margin, 35, { align: "right" });
    pdf.setTextColor(32, 48, 71);
    if (subtitle) {
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(10);
        const lines = pdf.splitTextToSize(subtitle, pageWidth - margin * 2 - rightBoxWidth - 8);
        pdf.text(lines, margin, 31);
        return 34 + lines.length * 5;
    }
    return 42;
}

async function exportCurrentPdf() {
    try {
        if (!window.jspdf?.jsPDF) {
            window.print();
            return;
        }

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 10;
        const selectedChart = getSelectedChartId();
        const sourceData = selectedChart === "pdv"
            ? lastPdvData
            : selectedChart === "tipos"
                ? lastTiposData
                : lastMainData;
        const items = (sourceData.length ? sourceData : []).map(item => ({
            label: selectedChart === "tipos" ? String(item.label || "") : getPdvDescriptionFromLabel(item.label),
            count: Number(item.count || 0),
            ratePct: item.ratePct,
            sharePct: item.sharePct
        }));
        if (!items.length) {
            window.print();
            return;
        }

        const max = Math.max(1, ...items.map(item => item.count));
        const chartTitle = document.getElementById("titlePrimary")?.textContent || "Totales";
        let chartSubtitle = selectedChart === "pdv"
            ? (document.getElementById("subtitleSecondary")?.textContent || "")
            : (document.getElementById("subtitlePrimary")?.textContent || "");
        if (selectedChart === "principal" && selectedDonutMode) {
            chartSubtitle = "Ranking de efectividad por punto de venta.";
        }
        if (/^Total de novedades:/i.test(chartSubtitle)) {
            chartSubtitle = "";
        }
        chartSubtitle = chartSubtitle.replace(/\s*\|\s*Promedio general:[^|]+/i, "").trim();
        const metaInfo = document.getElementById("metaInfo")?.textContent || "";
        const generalRowsForRates = Array.isArray(lastBaseRowsForDonut) ? lastBaseRowsForDonut : [];
        const selectedRowsForRates = lastSelectedPdv !== "__ALL__"
            ? generalRowsForRates.filter(row => getAssignedPdv(row) === lastSelectedPdv)
            : generalRowsForRates;
        const contextRates = computeRates(
            selectedRowsForRates,
            currentDashboardData?.indicadores || window.__summaryData?.indicadores || {}
        );
        const totalNovedadesExport = (selectedChart === "principal" && selectedDonutMode)
            ? generalRowsForRates.length
            : items.reduce((acc, item) => acc + Number(item.count || 0), 0);
        const highlightData = {
            totalNovedades: totalNovedadesExport,
            efectividadPct: contextRates.efectividadPct,
            rateLabel: lastSelectedPdv !== "__ALL__" ? `Efectividad ${lastSelectedPdv}` : "Tasa de efectividad general"
        };
        const labelWidth = 56;
        const rowGap = 6;
        const rowHeightBase = 12;
        const barX = margin + labelWidth + 8;
        const barMaxWidth = pageWidth - barX - margin;
        let page = 0;
        let index = 0;

        while (index < items.length) {
            if (page > 0) pdf.addPage();
            let y = drawPdfHeader(
                pdf,
                pageWidth,
                margin,
                chartTitle,
                metaInfo ? `${chartSubtitle}${chartSubtitle ? " | " : ""}${metaInfo}` : chartSubtitle,
                highlightData
            );
            y += 8;

            while (index < items.length) {
                const item = items[index];
                const labelLines = pdf.splitTextToSize(item.label, labelWidth);
                const rowHeight = Math.max(rowHeightBase, labelLines.length * 4.2 + 3);
                if (y + rowHeight > pageHeight - margin) break;

                pdf.setFont("helvetica", "normal");
                pdf.setFontSize(10);
                pdf.text(labelLines, margin, y + 4.5);
                pdf.setDrawColor(214, 223, 239);
                pdf.setFillColor(235, 243, 255);
                pdf.roundedRect(barX, y, barMaxWidth, 8, 2, 2, "FD");
                pdf.setFillColor(43, 124, 255);
                pdf.roundedRect(barX, y, Math.max(8, (item.count / max) * barMaxWidth), 8, 2, 2, "F");
                pdf.setFont("helvetica", "bold");
                pdf.setFontSize(12);
                const valueText = (selectedChart === "principal" && selectedDonutMode)
                    ? `${Number(item.count || 0).toFixed(1)}%`
                    : selectedChart === "pdv" && item.ratePct !== undefined
                        ? `${String(item.count)} | ${Number(item.ratePct).toFixed(1)}%`
                        : selectedChart === "tipos" && item.sharePct !== undefined
                            ? `${String(item.count)} | ${Number(item.sharePct).toFixed(1)}%`
                            : String(item.count);
                pdf.text(valueText, barX + barMaxWidth - 3, y + 5.8, { align: "right" });
                y += rowHeight + rowGap;
                index += 1;
            }
            page += 1;
        }
        pdf.save("Dashboard_Novedades.pdf");
    } catch (error) {
        console.warn("No se pudo exportar el PDF personalizado.", error);
        window.print();
    }
}

function setOptions(select, fields, preferred) {
    if (!select) return;
    select.innerHTML = fields.map(field => `<option value="${field}">${field}</option>`).join("");
    if (fields.includes(preferred)) select.value = preferred;
    else if (fields.length) select.value = fields[0];
}

function renderDashboard(data) {
    currentDashboardData = data || {};
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const pdvLabelByCode = new Map();
    const rowMeta = rows.map(row => {
        const nov = String(row["NOVEDADES"] || "").trim();
        const code = normalizePdvCode(row["PUNTO DE VENTA MARCACION"] || row["PUNTO DE VENTA ASIGNADO"] || "");
        const descRaw = String(row["DESCRIPCION MARCACION"] || row["DESCRIPCION ASIGNADA"] || "").trim();
        const desc = canonicalizePdvDescription(descRaw);
        if (code && !pdvLabelByCode.has(code)) {
            pdvLabelByCode.set(code, desc || code);
        }
        const pdv = code ? `${code} | ${pdvLabelByCode.get(code) || code}` : "";
        return {
            row,
            pdv,
            novLabel: normalizeNovedadLabel(nov),
            isOk: isOkNovedad(nov),
            isNovedad: !isOkNovedad(nov),
            isAdm: isAdmPdv(row),
            isDoble: normalizeComparable(row["OBSERVACION"]) === "MARCACION DOBLE",
            isNocturno: normalizeComparable(row["OBSERVACION"]) === "TRABAJO NOCTURNO"
        };
    });
    const allPdvs = Array.from(new Set(rowMeta.map(meta => meta.pdv))).filter(Boolean).sort((a, b) => a.localeCompare(b));
    const rateByPdv = new Map();
    for (const meta of rowMeta) {
        const current = rateByPdv.get(meta.pdv) || { total: 0, revisar: 0 };
        current.total += 1;
        if (meta.isNovedad) current.revisar += 1;
        rateByPdv.set(meta.pdv, current);
    }
    const pdvSelector = document.getElementById("pdvSelector");

    const apply = () => {
        const includeOk = !!document.getElementById("chkIncludeOk")?.checked;
        const includeAdm = !!document.getElementById("chkIncludeAdm")?.checked;
        const includeDouble = !!document.getElementById("chkIncludeDouble")?.checked;
        const includeNocturno = !!document.getElementById("chkIncludeNocturno")?.checked;
        const baseMetaForDonut = rowMeta.filter(meta => {
            if (!meta.pdv) return false;
            if (!includeAdm && meta.isAdm) return false;
            if (!includeDouble && meta.isDoble) return false;
            if (!includeNocturno && meta.isNocturno) return false;
            return true;
        });
        const baseMeta = rowMeta.filter(meta => {
            if (!meta.pdv) return false;
            if (!includeOk && meta.isOk) return false;
            if (!includeAdm && meta.isAdm) return false;
            if (!includeDouble && meta.isDoble) return false;
            if (!includeNocturno && meta.isNocturno) return false;
            return true;
        });
        const baseRows = baseMeta.map(meta => meta.row);
        const baseRowsForDonut = baseMetaForDonut.map(meta => meta.row);
        const globalRates = computeRates(baseRowsForDonut, data?.indicadores || {});
        const preferredSavedPdv = lastSelectedPdv || "__ALL__";
        const currentSelection = pdvSelector?.value || preferredSavedPdv;
        const pdvOptions = Array.from(new Set(baseMeta.map(meta => meta.pdv)))
            .filter(Boolean)
            .sort((a, b) => a.localeCompare(b));
        if (pdvSelector) {
            pdvSelector.innerHTML = [`<option value="__ALL__">Todos los puntos de venta</option>`, ...pdvOptions.map(pdv => `<option value="${escapeHtml(pdv)}">${escapeHtml(pdv)}</option>`)].join("");
            pdvSelector.value = pdvOptions.includes(currentSelection) ? currentSelection : "__ALL__";
        }
        const selectedPdv = pdvSelector?.value || "__ALL__";
        lastSelectedPdv = selectedPdv;
        const filteredMeta = baseMeta.filter(meta => selectedPdv === "__ALL__" || meta.pdv === selectedPdv);
        const filtered = filteredMeta.map(meta => meta.row);
        const selectedRowsAll = selectedPdv === "__ALL__"
            ? baseMetaForDonut.map(meta => meta.row)
            : baseMetaForDonut.filter(meta => meta.pdv === selectedPdv).map(meta => meta.row);
        const selectedRates = computeRates(selectedRowsAll, data?.indicadores || {});
        const pdvData = groupBy(baseMeta.map(meta => ({ PDV: meta.pdv })), "PDV")
            .slice(0, 50)
            .map(item => {
                const rowsPdv = baseMetaForDonut.filter(meta => meta.pdv === item.label).map(meta => meta.row);
                const ratesPdv = computeRates(rowsPdv, data?.indicadores || {});
                return { ...item, ratePct: ratesPdv.efectividadPct };
            });
        const totalNovedadesBase = Math.max(1, baseMeta.length);
        const totalsByType = groupBy(baseMeta.map(meta => ({ TIPO: meta.novLabel })), "TIPO")
            .slice(0, 50)
            .map(item => ({ ...item, sharePct: (Number(item.count || 0) / totalNovedadesBase) * 100 }));
        if (selectedNovedadLabel && !totalsByType.some(item => item.label === selectedNovedadLabel)) {
            selectedNovedadLabel = "";
        }
        const rowsForSelectedNovedad = selectedNovedadLabel
            ? baseMeta.filter(meta => meta.novLabel === selectedNovedadLabel).map(meta => meta.row)
            : [];
        let mainData = selectedNovedadLabel
            ? groupBy(rowsForSelectedNovedad.map(row => ({ PDV: getAssignedPdv(row) })), "PDV").slice(0, 50)
            : selectedPdv === "__ALL__"
                ? pdvData
                : groupBy(filtered.map(row => ({ TIPO: normalizeNovedadLabel(row["NOVEDADES"]) })), "TIPO").slice(0, 20);
        if (selectedDonutMode) {
            const perPdv = Array.from(new Set(baseMetaForDonut.map(meta => meta.pdv))).filter(Boolean);
            mainData = perPdv.map(pdv => {
                const rowsPdv = baseMetaForDonut.filter(meta => meta.pdv === pdv);
                const agg = { total: rowsPdv.length, revisar: rowsPdv.filter(meta => meta.isNovedad).length };
                const ok = Math.max(0, agg.total - agg.revisar);
                const pct = agg.total > 0 ? (ok / agg.total) * 100 : 0;
                return { label: pdv, count: Number(pct.toFixed(1)) };
            }).sort((a, b) => a.count - b.count || a.label.localeCompare(b.label));
        }
        const pdvAsignadosUnicos = new Set(baseRows.map(getAssignedPdv).filter(Boolean));
        lastFilteredRows = filtered;
        lastBaseRows = baseRows;
        lastBaseRowsForDonut = baseRowsForDonut;
        lastMainData = mainData;
        lastPdvData = pdvData;
        lastTiposData = totalsByType;

        document.getElementById("metricRegistros").textContent = String(baseRowsForDonut.length);
        document.getElementById("metricNovedades").textContent = String(baseMeta.filter(meta => meta.isNovedad).length);
        document.getElementById("metricPdvs").textContent = String(pdvAsignadosUnicos.size);
        document.getElementById("metricActivos").textContent = String((data.indicadores || {}).activos || 0);

        setDonutByIds("donutEfectividadGeneral", "donutEfectividadGeneralPct", "donutEfectividadGeneralMeta", globalRates.totalOk, globalRates.totalSeguimiento, globalRates.efectividadPct);
        setDonutByIds("donutEfectividadPdv", "donutEfectividadPdvPct", "donutEfectividadPdvMeta", selectedRates.totalOk, selectedRates.totalSeguimiento, selectedRates.efectividadPct);
        const donutPdvCard = document.getElementById("cardDonutPdv");
        if (donutPdvCard) donutPdvCard.classList.toggle("active", selectedDonutMode);

        document.getElementById("titlePrimary").textContent = selectedDonutMode
            ? "Tasa de efectividad por punto de venta (%)"
            : selectedNovedadLabel
            ? `Puntos de venta con: ${selectedNovedadLabel}`
            : selectedPdv === "__ALL__"
                ? "Totales de novedades por punto de venta"
                : `Novedades de ${selectedPdv}`;
        const totalMainData = mainData.reduce((acc, item) => acc + Number(item.count || 0), 0);
        document.getElementById("subtitlePrimary").textContent = selectedDonutMode
            ? `Ranking de efectividad (%). Promedio general: ${globalRates.efectividadPct.toFixed(1)}%`
            : `Total de novedades: ${totalMainData}`;
        document.getElementById("subtitleSecondary").textContent = "";

        renderVerticalBars(document.getElementById("mainChart"), mainData, selectedDonutMode);
        renderBars(document.getElementById("pdvList"), pdvData);
        renderTable(document.getElementById("tableBody"), totalsByType, selectedNovedadLabel);

        const uiState = {
            selectedPdv,
            includeOk,
            includeAdm,
            includeDouble,
            includeNocturno,
            selectedNovedadLabel,
            selectedDonutMode,
            selectedChartId: getSelectedChartId()
        };
        lastDashboardUiState = uiState;
        saveDashboardUiState(uiState);
    };

    ["pdvSelector", "chkIncludeOk", "chkIncludeAdm", "chkIncludeDouble", "chkIncludeNocturno"].forEach(id => {
        document.getElementById(id).onchange = () => {
            if (id === "pdvSelector") {
                const pdvValue = document.getElementById("pdvSelector")?.value || "__ALL__";
                if (pdvValue !== "__ALL__") setPrimarySelectionMode("pdv");
                else setPrimarySelectionMode("");
            }
            apply();
        };
    });
    document.getElementById("cardDonutPdv").onclick = () => {
        selectedDonutMode = !selectedDonutMode;
        if (selectedDonutMode) setPrimarySelectionMode("donut");
        apply();
    };
    document.getElementById("tableBody").onclick = event => {
        const row = event.target.closest("tr[data-novedad]");
        if (!row) return;
        const label = row.dataset.novedad || "";
        const enableNovedad = selectedNovedadLabel !== label;
        if (enableNovedad) {
            setPrimarySelectionMode("novedad");
            selectedNovedadLabel = label;
        } else {
            setPrimarySelectionMode("");
        }
        apply();
    };
    apply();
}

document.getElementById("btnBack").addEventListener("click", () => {
    window.location.href = "index.html";
});
document.getElementById("btnExportExcel").addEventListener("click", () => exportCurrentExcel(currentDashboardData || window.__summaryData || {}));
document.getElementById("btnExportPdf").addEventListener("click", exportCurrentPdf);
document.getElementById("btnImportProcessed").addEventListener("click", () => {
    document.getElementById("fileProcessedImport")?.click();
});
document.getElementById("fileProcessedImport").addEventListener("change", async event => {
    const files = event.target?.files;
    const importedData = await importProcessedWorkbooks(files);
    if (!importedData) {
        document.getElementById("metaInfo").textContent = "No se encontraron hojas 3.RevisionDeDias válidas en los archivos seleccionados.";
        event.target.value = "";
        return;
    }
    window.__summaryData = importedData;
    await saveImportedSummary(importedData);
    selectedNovedadLabel = "";
    selectedDonutMode = false;
    const rango = importedData.rangoRevision || {};
    document.getElementById("metaInfo").textContent = `Importación acumulada | Fuente: ${rango.fin || "-"} | Generado: ${new Date(importedData.generatedAt || Date.now()).toLocaleString()}`;
    renderDashboard(importedData);
    event.target.value = "";
});
document.getElementById("btnUseLastRun").addEventListener("click", async () => {
    const latestData = await loadSummary();
    window.__summaryData = latestData || {};
    selectedNovedadLabel = "";
    selectedDonutMode = false;
    if (latestData) {
        const rango = latestData.rangoRevision || {};
        document.getElementById("metaInfo").textContent = `Rango revisado: ${rango.inicio || "-"} a ${rango.fin || "-"} | Generado: ${new Date(latestData.generatedAt || Date.now()).toLocaleString()}`;
    } else {
        document.getElementById("metaInfo").textContent = "No hay un procesamiento guardado todavía. Ejecuta la revisión primero y luego vuelve aquí.";
    }
    renderDashboard(latestData || {});
});
document.getElementById("btnUseImportedRun").addEventListener("click", async () => {
    const importedData = await loadImportedSummary();
    window.__summaryData = importedData || {};
    selectedNovedadLabel = "";
    selectedDonutMode = false;
    if (importedData) {
        const rango = importedData.rangoRevision || {};
        document.getElementById("metaInfo").textContent = `Importación acumulada | Fuente: ${rango.fin || "-"} | Generado: ${new Date(importedData.generatedAt || Date.now()).toLocaleString()}`;
    } else {
        document.getElementById("metaInfo").textContent = "No hay un acumulado importado guardado todavía.";
    }
    renderDashboard(importedData || {});
});

(async () => {
    const [data, importedData, savedUi] = await Promise.all([loadSummary(), loadImportedSummary(), loadDashboardUiState()]);
    const initialData = data || importedData || {};
    window.__summaryData = initialData;
    if (savedUi) {
        if (typeof savedUi.includeOk === "boolean") document.getElementById("chkIncludeOk").checked = savedUi.includeOk;
        if (typeof savedUi.includeAdm === "boolean") document.getElementById("chkIncludeAdm").checked = savedUi.includeAdm;
        if (typeof savedUi.includeDouble === "boolean") document.getElementById("chkIncludeDouble").checked = savedUi.includeDouble;
        if (typeof savedUi.includeNocturno === "boolean") document.getElementById("chkIncludeNocturno").checked = savedUi.includeNocturno;
        selectedNovedadLabel = savedUi.selectedNovedadLabel || "";
        selectedDonutMode = !!savedUi.selectedDonutMode;
        lastSelectedPdv = savedUi.selectedPdv || "__ALL__";
    }
    if (data) {
        const meta = document.getElementById("metaInfo");
        const rango = data.rangoRevision || {};
        meta.textContent = `Rango revisado: ${rango.inicio || "-"} a ${rango.fin || "-"} | Generado: ${new Date(data.generatedAt || Date.now()).toLocaleString()}`;
    } else if (importedData) {
        const meta = document.getElementById("metaInfo");
        const rango = importedData.rangoRevision || {};
        meta.textContent = `Importación acumulada | Fuente: ${rango.fin || "-"} | Generado: ${new Date(importedData.generatedAt || Date.now()).toLocaleString()}`;
    } else {
        document.getElementById("metaInfo").textContent = "No hay un procesamiento guardado todavía. Ejecuta la revisión primero y luego vuelve aquí.";
    }
    renderDashboard(initialData);
    setupExportSelection();
    if (savedUi?.selectedChartId) {
        const target = document.querySelector(`.chart-export-selector[data-chart-id="${savedUi.selectedChartId}"]`);
        if (target) {
            target.checked = true;
            document.querySelectorAll(".chart-export-selector").forEach(other => {
                if (other !== target) other.checked = false;
            });
        }
    }
})();
})();

(function(){
if(!document.body.classList.contains('page-login')) return;
const form = document.getElementById("loginForm");
const emailInput = document.getElementById("loginEmail");
const passInput = document.getElementById("loginPassword");
const msg = document.getElementById("loginMsg");

const setMsg = (text, isError = true) => {
    if (!msg) return;
    msg.textContent = text || "";
    msg.style.color = isError ? "#c62828" : "#1d5edc";
};

(async () => {
    const token = AuthGate.getToken();
    if (!token) return;
    const ok = await AuthGate.verifyRemoteOrLogout();
    if (ok) window.location.href = "index.html";
})();

if (!form) return;
form.addEventListener("submit", async event => {
    event.preventDefault();
    const email = String(emailInput?.value || "").trim();
    const password = String(passInput?.value || "");
    if (!email || !password) {
        setMsg("Ingresa correo y contraseña.");
        return;
    }
    setMsg("Validando acceso...", false);
    try {
        const res = await AuthGate.login(email, password);
        if (!res?.ok || !res?.token) {
            setMsg(res?.error || "Credenciales inválidas.");
            return;
        }
        AuthGate.setSession(res.token, res.email || email);
        setMsg("Ingreso exitoso. Redirigiendo...", false);
        window.location.href = "index.html";
    } catch (_e) {
        setMsg("No se pudo conectar con el servicio de login.");
    }
});
})();
