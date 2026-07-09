import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { Upload, Download, Trash2, Search, Users, ClipboardList, Award, PenLine, RotateCcw, Check, ChevronLeft, ChevronRight, Moon, Sun, Lock, Unlock, Layers, Share2, BarChart3, AlertTriangle, UserX, UserCheck, Undo2, Plus, X, ArrowUpDown } from "lucide-react";

const LIGHT_THEME = {
  paper: "#F3FBF7",
  ink: "#0B4A3A",
  inkSoft: "#5B7A70",
  gold: "#12946B",
  goldSoft: "#DCF2E9",
  goldLine: "#BFE3D3",
  green: "#0E7C5A",
  greenSoft: "#E3F5EE",
  red: "#B23A32",
  redSoft: "#F5E4E1",
  card: "#FFFFFF",
  line: "#D8ECE4",
};

const DARK_THEME = {
  paper: "#0E1B16",
  ink: "#EAF7F1",
  inkSoft: "#9BB5AC",
  gold: "#2FBE8B",
  goldSoft: "#1B3229",
  goldLine: "#294B3E",
  green: "#3ECF98",
  greenSoft: "#153328",
  red: "#E17A72",
  redSoft: "#3A2220",
  card: "#152420",
  line: "#25392F",
};

// Mutated in place so every component reading T.xxx picks up the active
// theme's colors as soon as the app re-renders (triggered by the dark-mode toggle).
const T = { ...LIGHT_THEME };
function applyTheme(mode) {
  Object.assign(T, mode === "dark" ? DARK_THEME : LIGHT_THEME);
}

const FONT_LINK = "https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap";

const CLASSES_KEY = "carnet-notes:classes";
const ACTIVE_KEY = "carnet-notes:activeId";
const LEGACY_KEY = "carnet-notes:current";
const PIN_KEY = "carnet-notes:pin";
const DARK_KEY = "carnet-notes:dark";
const PASSING_AVERAGE = 9; // moyenne générale minimale pour passer en classe supérieure
const SUBJECT_PASS = 10; // seuil de réussite par matière (sur 20)
const EVALUATION_TYPES = ["Évaluation Formative 1", "Évaluation Sommative 1", "Évaluation Sommative 2", "Évaluation Sommative 3"];

function codeFromValues(obtenue, perfectionnement) {
  const oNum = toNum(obtenue);
  const pNum = toNum(perfectionnement);
  if (oNum === null && pNum === null) return "";
  const o = String(oNum ?? 0).replace(".", ",");
  const p = String(Math.min(2, Math.max(0, Math.round(pNum || 0))));
  return `${o} ${p}`;
}

// Format: "<note obtenue, entier ou décimal>  <perfectionnement 0/1/2>"
// Examples: "14,5 2"  "12 0"  "18.5 1"
const CODE_PATTERN = /^(\d{1,2}(?:[.,]\d{1,2})?)\s+([0-2])$/;

function parseCode(code) {
  if (!code) return { obtenue: "", perfectionnement: "" };
  const m = code.trim().match(CODE_PATTERN);
  if (!m) return { obtenue: "", perfectionnement: "" };
  return {
    obtenue: m[1].replace(",", "."),
    perfectionnement: m[2],
  };
}

function isCodeComplete(code) {
  return !!code && CODE_PATTERN.test(code.trim());
}

// Lets the person type freely: digits, one decimal separator (, or .),
// a space, then a single perfectionnement digit (0, 1 or 2).
function sanitizeCode(value) {
  const filtered = value.replace(/[^0-9.,\s]/g, "");
  const spaceIdx = filtered.search(/\s/);

  const cleanObtenue = (raw) => {
    let out = "";
    let seenSep = false;
    let intDigits = 0;
    let decDigits = 0;
    for (const ch of raw) {
      if (/[0-9]/.test(ch)) {
        if (!seenSep) {
          if (intDigits < 2) {
            out += ch;
            intDigits++;
          }
        } else if (decDigits < 2) {
          out += ch;
          decDigits++;
        }
      } else if (/[.,]/.test(ch) && !seenSep && intDigits > 0) {
        out += ch;
        seenSep = true;
      }
    }
    return out;
  };

  if (spaceIdx === -1) {
    return cleanObtenue(filtered);
  }
  const obtenuePart = cleanObtenue(filtered.slice(0, spaceIdx));
  const rest = filtered.slice(spaceIdx).replace(/\s/g, "");
  const perf = rest.length > 0 && ["0", "1", "2"].includes(rest[0]) ? rest[0] : "";
  return perf ? `${obtenuePart} ${perf}` : obtenuePart.length > 0 ? `${obtenuePart} ` : "";
}

function formatNum(n) {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return Number.isInteger(n) ? String(n) : n.toFixed(1).replace(".", ",");
}

function cleanMatricule(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/^'/, "").trim();
}

function toNum(v) {
  if (v === "" || v === null || v === undefined) return null;
  const n = parseFloat(String(v).replace(",", "."));
  return isNaN(n) ? null : n;
}

function computeRanking(roster, subjects, grades, attendance = {}) {
  const list = roster.map((stu) => {
    const absent = attendance[stu.matricule] === false;
    let sum = 0;
    let count = 0;
    subjects.forEach((s) => {
      const g = grades[stu.matricule]?.[s.key];
      if (g && (g.obtenue !== "" && g.obtenue !== undefined || g.perfectionnement !== "" && g.perfectionnement !== undefined)) {
        const o = toNum(g.obtenue) || 0;
        const p = toNum(g.perfectionnement) || 0;
        sum += o + p;
        count += 1;
      }
    });
    const moyenne = !absent && count > 0 ? sum / count : null;
    return { ...stu, moyenne, count, absent };
  });

  const withMoy = list.filter((s) => s.moyenne !== null).sort((a, b) => b.moyenne - a.moyenne);
  const withoutMoy = list.filter((s) => s.moyenne === null);

  let rang = 0;
  let prevMoy = null;
  withMoy.forEach((s, i) => {
    const position = i + 1;
    if (s.moyenne !== prevMoy) {
      rang = position;
      prevMoy = s.moyenne;
    }
    s.rang = rang;
  });
  withoutMoy.forEach((s) => (s.rang = null));

  return [...withMoy, ...withoutMoy];
}

export default function CarnetNotes() {
  const [roster, setRoster] = useState([]);
  const [subjects, setSubjects] = useState([]);
  const [grades, setGrades] = useState({});
  const [attendance, setAttendance] = useState({}); // matricule -> false means absent (absence of key = present)
  const [className, setClassName] = useState("");
  const [evaluationType, setEvaluationType] = useState("");
  const [activeTab, setActiveTab] = useState("parEleve");
  const [studentIndex, setStudentIndex] = useState(0);
  const [showWelcome, setShowWelcome] = useState(true);
  const [search, setSearch] = useState("");
  const [loaded, setLoaded] = useState(false);
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved
  const [newStudent, setNewStudent] = useState({ matricule: "", nom: "", prenoms: "" });

  const [classesList, setClassesList] = useState([]); // catalog of all saved classes (metadata + snapshot)
  const [activeId, setActiveId] = useState(null);
  const [showClassSwitcher, setShowClassSwitcher] = useState(false);
  const [showReorder, setShowReorder] = useState(false);

  const [darkMode, setDarkMode] = useState(() => {
    const dark = typeof window !== "undefined" && localStorage.getItem(DARK_KEY) === "1";
    applyTheme(dark ? "dark" : "light");
    return dark;
  });
  const [, forceThemeRerender] = useState(0);

  const [pinSet, setPinSet] = useState(() => typeof window !== "undefined" && !!localStorage.getItem(PIN_KEY));
  const [locked, setLocked] = useState(() => typeof window !== "undefined" && !!localStorage.getItem(PIN_KEY));

  const fileInputRef = useRef(null);
  const saveTimer = useRef(null);
  const undoStack = useRef([]);
  const [canUndo, setCanUndo] = useState(false);

  // Load persisted data once (stored locally in this browser), migrating any
  // older single-class save into the new multi-class catalog.
  useEffect(() => {
    try {
      const dark = localStorage.getItem(DARK_KEY) === "1";
      setDarkMode(dark);
      applyTheme(dark ? "dark" : "light");

      const pin = localStorage.getItem(PIN_KEY);
      setPinSet(!!pin);
      setLocked(!!pin);

      let list = [];
      const rawClasses = localStorage.getItem(CLASSES_KEY);
      if (rawClasses) list = JSON.parse(rawClasses);

      if (list.length === 0) {
        const legacyRaw = localStorage.getItem(LEGACY_KEY);
        if (legacyRaw) {
          const legacy = JSON.parse(legacyRaw);
          if (legacy.roster && legacy.roster.length > 0) {
            list = [
              {
                id: "legacy",
                className: legacy.className || "",
                roster: legacy.roster || [],
                subjects: legacy.subjects || [],
                grades: legacy.grades || {},
                attendance: {},
                updatedAt: Date.now(),
              },
            ];
          }
        }
      }

      setClassesList(list);
      const savedActiveId = localStorage.getItem(ACTIVE_KEY);
      const active = list.find((c) => c.id === savedActiveId) || list[0];
      if (active) {
        setActiveId(active.id);
        setRoster(active.roster || []);
        setSubjects(active.subjects || []);
        setGrades(active.grades || {});
        setAttendance(active.attendance || {});
        setClassName(active.className || "");
        setEvaluationType(active.evaluationType || "");
        setActiveTab("parEleve");
        setShowWelcome(false);
      }
    } catch (e) {
      // no saved data yet
    }
    setLoaded(true);
  }, []);

  // Autosave (debounced): keeps the active class's data in sync inside the
  // multi-class catalog stored in localStorage.
  useEffect(() => {
    if (!loaded) return;
    if (roster.length === 0 || !activeId) return;
    setSaveState("saving");
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      try {
        setClassesList((prev) => {
          const others = prev.filter((c) => c.id !== activeId);
          const updated = [
            ...others,
            { id: activeId, className, evaluationType, roster, subjects, grades, attendance, updatedAt: Date.now() },
          ];
          localStorage.setItem(CLASSES_KEY, JSON.stringify(updated));
          localStorage.setItem(ACTIVE_KEY, activeId);
          return updated;
        });
        setSaveState("saved");
      } catch (e) {
        setSaveState("idle");
      }
    }, 700);
    return () => clearTimeout(saveTimer.current);
  }, [roster, subjects, grades, attendance, className, evaluationType, activeId, loaded]);

  const toggleDarkMode = () => {
    setDarkMode((prev) => {
      const next = !prev;
      applyTheme(next ? "dark" : "light");
      localStorage.setItem(DARK_KEY, next ? "1" : "0");
      forceThemeRerender((n) => n + 1);
      return next;
    });
  };

  const startNewClass = () => {
    const id = `cls_${Date.now()}`;
    setActiveId(id);
    setRoster([]);
    setSubjects([]);
    setGrades({});
    setAttendance({});
    setClassName("");
    setEvaluationType("");
    setActiveTab("parEleve");
    setStudentIndex(0);
    setShowWelcome(false);
    setShowClassSwitcher(false);
    undoStack.current = [];
    setCanUndo(false);
  };

  const switchToClass = (id) => {
    const rec = classesList.find((c) => c.id === id);
    if (!rec) return;
    setActiveId(id);
    setRoster(rec.roster || []);
    setSubjects(rec.subjects || []);
    setGrades(rec.grades || {});
    setAttendance(rec.attendance || {});
    setClassName(rec.className || "");
    setEvaluationType(rec.evaluationType || "");
    setActiveTab("parEleve");
    setStudentIndex(0);
    setShowClassSwitcher(false);
    undoStack.current = [];
    setCanUndo(false);
  };

  const deleteClass = (id) => {
    if (!window.confirm("Supprimer définitivement cette classe et toutes ses notes ?")) return;
    const remaining = classesList.filter((c) => c.id !== id);
    localStorage.setItem(CLASSES_KEY, JSON.stringify(remaining));
    setClassesList(remaining);
    if (id === activeId) {
      if (remaining.length > 0) {
        switchToClass(remaining[0].id);
      } else {
        setActiveId(null);
        setRoster([]);
        setSubjects([]);
        setGrades({});
        setAttendance({});
        setClassName("");
        localStorage.removeItem(ACTIVE_KEY);
      }
    }
  };

  const pushUndo = (prevGrades) => {
    undoStack.current.push(JSON.parse(JSON.stringify(prevGrades)));
    if (undoStack.current.length > 20) undoStack.current.shift();
    setCanUndo(true);
  };

  const handleUndo = () => {
    const last = undoStack.current.pop();
    if (last) {
      setGrades(last);
      setCanUndo(undoStack.current.length > 0);
      if (navigator.vibrate) navigator.vibrate(20);
    }
  };

  const ranking = useMemo(() => computeRanking(roster, subjects, grades, attendance), [roster, subjects, grades, attendance]);
  const rankByMatricule = useMemo(() => {
    const m = {};
    ranking.forEach((r) => (m[r.matricule] = r));
    return m;
  }, [ranking]);

  const filteredRoster = useMemo(() => {
    if (!search.trim()) return roster;
    const q = search.toLowerCase();
    return roster.filter(
      (s) => s.nom.toLowerCase().includes(q) || s.prenoms.toLowerCase().includes(q) || s.matricule.includes(q)
    );
  }, [roster, search]);

  const handleFile = useCallback((file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });

        const headerRow1 = rows[0] || [];
        const newSubjects = [];
        for (let c = 3; c < headerRow1.length; c++) {
          if (headerRow1[c] && String(headerRow1[c]).trim() !== "") {
            newSubjects.push({ key: `subj_${c}`, label: String(headerRow1[c]).trim(), col: c });
            c += 1; // skip the paired "perfectionnement" column
          }
        }

        const newRoster = [];
        const newGrades = {};
        for (let r = 2; r < rows.length; r++) {
          const row = rows[r] || [];
          const rawMat = row[0];
          if (!rawMat || String(rawMat).trim() === "") continue;
          const matricule = cleanMatricule(rawMat);
          const nom = String(row[1] || "").trim();
          const prenoms = String(row[2] || "").trim();
          newRoster.push({ matricule, rawMatricule: String(rawMat), nom, prenoms });
          newGrades[matricule] = {};
          newSubjects.forEach((s) => {
            const obtenue = row[s.col] !== undefined ? String(row[s.col]) : "";
            const perfectionnement = row[s.col + 1] !== undefined ? String(row[s.col + 1]) : "";
            newGrades[matricule][s.key] = { obtenue, perfectionnement, rawCode: codeFromValues(obtenue, perfectionnement) };
          });
        }

        const guessedClassName = sheetName.replace(/^Notes\s*-\s*/i, "").trim();
        setSubjects(newSubjects);
        setRoster(newRoster);
        setGrades(newGrades);
        setClassName(guessedClassName || "");
        setActiveTab("parEleve");
        setActiveId((prev) => prev || `cls_${Date.now()}`);
      } catch (err) {
        alert("Impossible de lire ce fichier. Vérifie qu'il s'agit bien d'un export EducMaster (.xlsx).");
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const updateGradeCode = (matricule, subjectKey, rawCode) => {
    const { obtenue, perfectionnement } = parseCode(rawCode);
    if (isCodeComplete(rawCode)) {
      pushUndo(grades);
    }
    setGrades((prev) => ({
      ...prev,
      [matricule]: {
        ...prev[matricule],
        [subjectKey]: { rawCode, obtenue, perfectionnement },
      },
    }));
  };

  const toggleAttendance = (matricule) => {
    setAttendance((prev) => ({
      ...prev,
      [matricule]: prev[matricule] === false ? true : false,
    }));
  };

  const addStudent = () => {
    if (!newStudent.nom.trim()) return;
    const matricule = newStudent.matricule.trim() || `TEMP-${Date.now()}`;
    if (roster.some((s) => s.matricule === matricule)) {
      alert("Ce matricule existe déjà dans la classe.");
      return;
    }
    setRoster((prev) => [...prev, { matricule, rawMatricule: matricule, nom: newStudent.nom.trim(), prenoms: newStudent.prenoms.trim() }]);
    setGrades((prev) => ({ ...prev, [matricule]: {} }));
    setNewStudent({ matricule: "", nom: "", prenoms: "" });
  };

  const removeStudent = (matricule) => {
    if (!window.confirm("Retirer cet élève de la classe ?")) return;
    setRoster((prev) => prev.filter((s) => s.matricule !== matricule));
    setGrades((prev) => {
      const next = { ...prev };
      delete next[matricule];
      return next;
    });
  };

  const resetAll = () => {
    if (!activeId) return;
    deleteClass(activeId);
  };

  const handleSetPin = () => {
    const pin = window.prompt("Choisis un code à 4 chiffres pour verrouiller l'application :");
    if (!pin) return;
    if (!/^\d{4}$/.test(pin)) {
      alert("Le code doit contenir exactement 4 chiffres.");
      return;
    }
    localStorage.setItem(PIN_KEY, pin);
    setPinSet(true);
    alert("Code enregistré. Utilise l'icône de cadenas pour verrouiller l'application.");
  };

  const handleRemovePin = () => {
    if (!window.confirm("Supprimer le code PIN ? L'application ne sera plus verrouillable.")) return;
    localStorage.removeItem(PIN_KEY);
    setPinSet(false);
    setLocked(false);
  };

  const handleLockNow = () => {
    if (!pinSet) {
      handleSetPin();
      return;
    }
    setLocked(true);
  };

  const buildWorkbook = () => {
    const header1 = ["Matricule", "Nom", "Prénoms"];
    const header2 = ["", "", ""];
    subjects.forEach((s) => {
      header1.push(s.label, "");
      header2.push("Note obtenue", "Note perfectionnement");
    });
    const wsData = [header1, header2];
    roster.forEach((stu) => {
      const row = [stu.rawMatricule || stu.matricule, stu.nom, stu.prenoms];
      subjects.forEach((s) => {
        const g = grades[stu.matricule]?.[s.key] || {};
        const obtenueDisplay = g.obtenue !== undefined && g.obtenue !== "" ? String(g.obtenue).replace(".", ",") : "";
        row.push(obtenueDisplay, g.perfectionnement ?? "");
      });
      wsData.push(row);
    });
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const merges = subjects.map((s, i) => {
      const c = 3 + i * 2;
      return { s: { r: 0, c }, e: { r: 0, c: c + 1 } };
    });
    ws["!merges"] = merges;
    ws["!cols"] = [{ wch: 16 }, { wch: 18 }, { wch: 22 }, ...subjects.flatMap(() => [{ wch: 14 }, { wch: 16 }])];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Notes - ${className || "Classe"}`.slice(0, 31));

    const recapData = [["Nom", "Prénoms", "Matricule", "Moyenne générale", "Rang", "Décision"]];
    ranking.forEach((r) => {
      const decision = r.absent ? "Absent(e)" : r.moyenne === null ? "-" : r.moyenne >= PASSING_AVERAGE ? "Admis(e)" : "Non admis(e)";
      recapData.push([r.nom, r.prenoms, r.matricule, r.moyenne !== null ? Math.round(r.moyenne * 100) / 100 : "-", r.rang ?? "-", decision]);
    });
    const ws2 = XLSX.utils.aoa_to_sheet(recapData);
    ws2["!cols"] = [{ wch: 18 }, { wch: 22 }, { wch: 16 }, { wch: 16 }, { wch: 8 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws2, "Récapitulatif");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const evalSlug = evaluationType ? `_${evaluationType.replace(/\s+/g, "_")}` : "";
    const filename = `Notes_${(className || "Classe").replace(/\s+/g, "_")}${evalSlug}_EducMaster.xlsx`;
    return { wbout, filename };
  };

  const handleExport = () => {
    if (roster.length === 0) return;
    const { wbout, filename } = buildWorkbook();
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleShare = async () => {
    if (roster.length === 0) return;
    const { wbout, filename } = buildWorkbook();
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const file = new File([blob], filename, { type: "application/octet-stream" });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: filename, text: `Notes de la classe ${className || ""}` });
      } catch (e) {
        // person cancelled the share sheet — nothing to do
      }
    } else {
      alert("Le partage direct n'est pas disponible sur ce navigateur. Utilise le bouton \"Exporter\" puis partage le fichier téléchargé depuis tes fichiers.");
    }
  };

  const activeSubject = subjects.find((s) => s.key === activeTab);

  if (locked) {
    return (
      <PinLockScreen
        onUnlock={() => setLocked(false)}
        onForgot={() => {
          if (window.confirm("Retirer le code PIN oublié ? Tu pourras en redéfinir un nouveau ensuite.")) {
            localStorage.removeItem(PIN_KEY);
            setPinSet(false);
            setLocked(false);
          }
        }}
      />
    );
  }

  return (
    <div style={{ background: T.paper, minHeight: "100vh", fontFamily: "'IBM Plex Sans', sans-serif", color: T.ink }}>
      <style>{`@import url('${FONT_LINK}');
        input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { opacity: 1; }
        .tab-btn { transition: transform .15s ease, box-shadow .15s ease; }
        .tab-btn:hover { transform: translateY(-1px); }
        ::selection { background: ${T.goldSoft}; }
      `}</style>

      {/* Header / ledger cover */}
      <div style={{ background: T.ink, borderBottom: `4px solid ${T.gold}` }} className="px-4 pt-6 pb-5 sm:px-8">
        <div className="max-w-5xl mx-auto flex items-start justify-between gap-4">
          <div className="min-w-0 flex-1">
            <p style={{ color: T.goldSoft, fontFamily: "'IBM Plex Mono', monospace" }} className="text-xs tracking-widest uppercase mb-1">
              NoteExpress · Import EducMaster
            </p>
            <input
              value={className}
              onChange={(e) => setClassName(e.target.value)}
              placeholder="Nom de la classe (ex : CM2 Ganmi A)"
              style={{ fontFamily: "'Fraunces', serif", color: "#FFFFFF", borderBottom: `1px solid ${T.inkSoft}` }}
              className="bg-transparent text-2xl sm:text-3xl font-semibold outline-none w-full max-w-md placeholder-slate-400"
            />
            {roster.length > 0 && evaluationType && (
              <select
                value={evaluationType}
                onChange={(e) => setEvaluationType(e.target.value)}
                className="mt-1.5 text-xs rounded-full px-2.5 py-1 outline-none"
                style={{ background: T.gold, color: "#fff", border: "none" }}
              >
                {EVALUATION_TYPES.map((ev) => (
                  <option key={ev} value={ev} style={{ color: T.ink }}>
                    {ev}
                  </option>
                ))}
              </select>
            )}
          </div>
          <div className="flex items-center gap-1 shrink-0">
            {canUndo && roster.length > 0 && (
              <button onClick={handleUndo} title="Annuler la dernière saisie" className="p-2 rounded-full" style={{ color: T.goldSoft }}>
                <Undo2 size={19} />
              </button>
            )}
            <button
              onClick={() => setShowClassSwitcher((v) => !v)}
              title="Mes classes"
              className="p-2 rounded-full relative"
              style={{ color: T.goldSoft }}
            >
              <Layers size={19} />
              {classesList.length > 1 && (
                <span
                  className="absolute -top-0.5 -right-0.5 w-4 h-4 rounded-full text-[10px] flex items-center justify-center font-semibold"
                  style={{ background: T.gold, color: "#fff" }}
                >
                  {classesList.length}
                </span>
              )}
            </button>
            <button onClick={toggleDarkMode} title="Mode sombre" className="p-2 rounded-full" style={{ color: T.goldSoft }}>
              {darkMode ? <Sun size={19} /> : <Moon size={19} />}
            </button>
            <button
              onClick={handleLockNow}
              title={pinSet ? "Verrouiller" : "Définir un code PIN"}
              className="p-2 rounded-full"
              style={{ color: T.goldSoft }}
            >
              {pinSet ? <Lock size={19} /> : <Unlock size={19} />}
            </button>
            {roster.length > 0 && (
              <button onClick={resetAll} title="Supprimer cette classe" className="p-2 rounded-full" style={{ color: T.goldSoft }}>
                <Trash2 size={19} />
              </button>
            )}
          </div>
        </div>

        {showClassSwitcher && (
          <ClassSwitcher
            classesList={classesList}
            activeId={activeId}
            onSwitch={switchToClass}
            onNew={startNewClass}
            onDelete={deleteClass}
            onClose={() => setShowClassSwitcher(false)}
          />
        )}
      </div>

      <div className="max-w-5xl mx-auto px-4 sm:px-8 py-6">
        {roster.length === 0 && showWelcome ? (
          <WelcomeScreen onStart={() => setShowWelcome(false)} />
        ) : roster.length === 0 ? (
          <ImportScreen onFile={handleFile} fileInputRef={fileInputRef} />
        ) : !evaluationType ? (
          <EvaluationPicker onSelect={setEvaluationType} />
        ) : (
          <>
            {/* Tabs */}
            <div className="flex gap-2 overflow-x-auto pb-1 mb-5 -mx-1 px-1">
              <TabButton active={activeTab === "eleves"} onClick={() => setActiveTab("eleves")} icon={<Users size={15} />}>
                Élèves
              </TabButton>
              <TabButton active={activeTab === "parEleve"} onClick={() => setActiveTab("parEleve")} icon={<ClipboardList size={15} />}>
                Par élève
              </TabButton>
              {subjects.map((s) => (
                <TabButton key={s.key} active={activeTab === s.key} onClick={() => setActiveTab(s.key)} icon={<PenLine size={15} />}>
                  {s.label}
                </TabButton>
              ))}
              <TabButton active={activeTab === "recap"} onClick={() => setActiveTab("recap")} icon={<Award size={15} />}>
                Récapitulatif
              </TabButton>
              <TabButton active={activeTab === "stats"} onClick={() => setActiveTab("stats")} icon={<BarChart3 size={15} />}>
                Statistiques
              </TabButton>
            </div>

            {/* Search + export row */}
            <div className="flex flex-col sm:flex-row gap-3 mb-4">
              <div
                className="flex items-center gap-2 flex-1 rounded-lg px-3 py-2"
                style={{ background: T.card, border: `1px solid ${T.line}` }}
              >
                <Search size={16} style={{ color: T.inkSoft }} />
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Rechercher un élève…"
                  className="bg-transparent outline-none w-full text-sm"
                />
              </div>
              <button
                onClick={() => setShowReorder((v) => !v)}
                title="Réorganiser l'ordre des matières"
                className="flex items-center justify-center gap-2 px-3 py-2 rounded-lg font-medium text-sm shrink-0"
                style={{ background: T.card, color: T.inkSoft, border: `1px solid ${T.line}` }}
              >
                <ArrowUpDown size={16} />
              </button>
              <button
                onClick={handleShare}
                className="flex items-center justify-center gap-2 px-4 py-2 rounded-lg font-medium text-sm shrink-0"
                style={{ background: T.greenSoft, color: T.green, border: `1px solid ${T.goldLine}` }}
              >
                <Share2 size={16} />
                Partager
              </button>
              <button
                onClick={handleExport}
                className="flex items-center justify-center gap-2 px-4 py-2 rounded-lg font-medium text-sm shrink-0"
                style={{ background: T.gold, color: "#FFFFFF" }}
              >
                <Download size={16} />
                Exporter pour EducMaster
              </button>
            </div>

            {showReorder && (
              <ReorderPanel subjects={subjects} setSubjects={setSubjects} onClose={() => setShowReorder(false)} />
            )}

            <p className="text-xs mb-4" style={{ color: T.inkSoft, fontFamily: "'IBM Plex Mono', monospace" }}>
              {saveState === "saving" ? "Enregistrement…" : saveState === "saved" ? "✓ Enregistré sur cet appareil" : ""}
              {"  ·  "} {roster.length} élève{roster.length > 1 ? "s" : ""}
            </p>

            {activeTab === "eleves" && (
              <ElevesTab
                roster={filteredRoster}
                rankByMatricule={rankByMatricule}
                onRemove={removeStudent}
                newStudent={newStudent}
                setNewStudent={setNewStudent}
                onAdd={addStudent}
                attendance={attendance}
                onToggleAttendance={toggleAttendance}
              />
            )}

            {activeTab === "parEleve" && (
              <StudentEntryTab
                roster={roster}
                subjects={subjects}
                grades={grades}
                onChangeCode={updateGradeCode}
                currentIndex={studentIndex}
                setCurrentIndex={setStudentIndex}
                attendance={attendance}
                onToggleAttendance={toggleAttendance}
              />
            )}

            {activeSubject && (
              <SubjectTab
                roster={filteredRoster}
                subject={activeSubject}
                grades={grades}
                onChangeCode={updateGradeCode}
                attendance={attendance}
              />
            )}

            {activeTab === "recap" && <RecapTab ranking={search.trim() ? ranking.filter((r) => filteredRoster.some((f) => f.matricule === r.matricule)) : ranking} />}

            {activeTab === "stats" && <StatsTab roster={roster} subjects={subjects} grades={grades} attendance={attendance} />}
          </>
        )}
      </div>
    </div>
  );

  function TabButton({ active, onClick, icon, children }) {
    return (
      <button
        onClick={onClick}
        className="tab-btn flex items-center gap-1.5 whitespace-nowrap px-3.5 py-2 rounded-t-lg text-sm font-medium"
        style={
          active
            ? { background: T.card, color: T.ink, boxShadow: `0 -2px 0 ${T.gold} inset`, border: `1px solid ${T.line}`, borderBottom: `1px solid ${T.card}` }
            : { background: T.goldSoft, color: T.inkSoft, border: `1px solid ${T.goldLine}` }
        }
      >
        {icon}
        {children}
      </button>
    );
  }
}

function ClassSwitcher({ classesList, activeId, onSwitch, onNew, onDelete, onClose }) {
  return (
    <div className="max-w-5xl mx-auto mt-3">
      <div className="rounded-xl overflow-hidden" style={{ background: T.card, border: `1px solid ${T.goldLine}` }}>
        <div className="flex items-center justify-between px-4 py-2.5" style={{ background: T.greenSoft }}>
          <span className="text-sm font-semibold" style={{ color: T.green }}>Mes classes</span>
          <button onClick={onClose} style={{ color: T.green }}><X size={16} /></button>
        </div>
        <div className="max-h-64 overflow-y-auto">
          {classesList.length === 0 && (
            <p className="px-4 py-3 text-xs" style={{ color: T.inkSoft }}>Aucune classe enregistrée pour l'instant.</p>
          )}
          {classesList.map((c) => (
            <div
              key={c.id}
              className="flex items-center justify-between px-4 py-2.5"
              style={{ borderTop: `1px solid ${T.line}`, background: c.id === activeId ? T.goldSoft : "transparent" }}
            >
              <button onClick={() => onSwitch(c.id)} className="text-left flex-1 min-w-0">
                <div className="text-sm font-medium truncate">{c.className || "Classe sans nom"}</div>
                <div className="text-xs" style={{ color: T.inkSoft }}>{(c.roster || []).length} élève(s)</div>
              </button>
              <button onClick={() => onDelete(c.id)} className="p-1.5 ml-2" style={{ color: T.red }}>
                <Trash2 size={15} />
              </button>
            </div>
          ))}
        </div>
        <button
          onClick={onNew}
          className="w-full flex items-center justify-center gap-1.5 px-4 py-3 text-sm font-medium"
          style={{ borderTop: `1px solid ${T.line}`, color: T.green }}
        >
          <Plus size={16} /> Nouvelle classe
        </button>
      </div>
    </div>
  );
}

function PinLockScreen({ onUnlock, onForgot }) {
  const [pin, setPin] = useState("");
  const [error, setError] = useState(false);

  const handleDigit = (d) => {
    const next = (pin + d).slice(0, 4);
    setPin(next);
    setError(false);
    if (next.length === 4) {
      const saved = localStorage.getItem(PIN_KEY);
      setTimeout(() => {
        if (next === saved) {
          onUnlock();
        } else {
          setError(true);
          setPin("");
          if (navigator.vibrate) navigator.vibrate([30, 40, 30]);
        }
      }, 120);
    }
  };

  return (
    <div style={{ background: T.ink, minHeight: "100vh" }} className="flex items-center justify-center px-6">
      <style>{`@import url('${FONT_LINK}');`}</style>
      <div className="w-full max-w-xs text-center">
        <div
          className="mx-auto mb-6 w-14 h-14 rounded-2xl flex items-center justify-center"
          style={{ background: T.gold }}
        >
          <Lock size={22} color="#fff" />
        </div>
        <p style={{ color: "#FFFFFF", fontFamily: "'Fraunces', serif" }} className="text-lg font-semibold mb-1">
          Application verrouillée
        </p>
        <p className="text-xs mb-6" style={{ color: T.goldSoft }}>
          {error ? "Code incorrect, réessaie." : "Saisis ton code à 4 chiffres"}
        </p>
        <div className="flex items-center justify-center gap-3 mb-8">
          {[0, 1, 2, 3].map((i) => (
            <div
              key={i}
              className="w-3.5 h-3.5 rounded-full"
              style={{ background: i < pin.length ? T.gold : "rgba(255,255,255,0.15)" }}
            />
          ))}
        </div>
        <div className="grid grid-cols-3 gap-3 mb-6">
          {["1", "2", "3", "4", "5", "6", "7", "8", "9"].map((d) => (
            <button
              key={d}
              onClick={() => handleDigit(d)}
              className="py-3.5 rounded-xl text-lg font-semibold"
              style={{ background: "rgba(255,255,255,0.08)", color: "#FFFFFF" }}
            >
              {d}
            </button>
          ))}
          <div />
          <button
            onClick={() => handleDigit("0")}
            className="py-3.5 rounded-xl text-lg font-semibold"
            style={{ background: "rgba(255,255,255,0.08)", color: "#FFFFFF" }}
          >
            0
          </button>
          <button
            onClick={() => setPin((p) => p.slice(0, -1))}
            className="py-3.5 rounded-xl text-sm font-semibold"
            style={{ background: "rgba(255,255,255,0.08)", color: "#FFFFFF" }}
          >
            ⌫
          </button>
        </div>
        <button onClick={onForgot} className="text-xs underline" style={{ color: T.goldSoft }}>
          Code oublié ?
        </button>
      </div>
    </div>
  );
}

function ReorderPanel({ subjects, setSubjects, onClose }) {
  const move = (index, dir) => {
    const target = index + dir;
    if (target < 0 || target >= subjects.length) return;
    const next = [...subjects];
    [next[index], next[target]] = [next[target], next[index]];
    setSubjects(next);
  };

  return (
    <div className="rounded-xl overflow-hidden mb-4" style={{ background: T.card, border: `1px solid ${T.goldLine}` }}>
      <div className="flex items-center justify-between px-4 py-2.5" style={{ background: T.greenSoft }}>
        <span className="text-sm font-semibold" style={{ color: T.green }}>Ordre des matières pour la saisie</span>
        <button onClick={onClose} style={{ color: T.green }}><X size={16} /></button>
      </div>
      <div>
        {subjects.map((s, i) => (
          <div
            key={s.key}
            className="flex items-center justify-between px-4 py-2.5"
            style={{ borderTop: i > 0 ? `1px solid ${T.line}` : "none" }}
          >
            <span className="text-sm">
              <span style={{ color: T.inkSoft, fontFamily: "'IBM Plex Mono', monospace" }} className="mr-2">{i + 1}.</span>
              {s.label}
            </span>
            <div className="flex gap-1">
              <button
                onClick={() => move(i, -1)}
                disabled={i === 0}
                className="p-1.5 rounded disabled:opacity-25"
                style={{ color: T.green }}
              >
                <ChevronLeft size={16} style={{ transform: "rotate(90deg)" }} />
              </button>
              <button
                onClick={() => move(i, 1)}
                disabled={i === subjects.length - 1}
                className="p-1.5 rounded disabled:opacity-25"
                style={{ color: T.green }}
              >
                <ChevronRight size={16} style={{ transform: "rotate(90deg)" }} />
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function EvaluationPicker({ onSelect }) {
  return (
    <div className="rounded-2xl p-8 text-center" style={{ background: T.card, border: `1px solid ${T.line}` }}>
      <div
        className="mx-auto mb-5 w-14 h-14 rounded-full flex items-center justify-center"
        style={{ background: T.goldSoft, color: T.gold }}
      >
        <Layers size={24} />
      </div>
      <h2 style={{ fontFamily: "'Fraunces', serif" }} className="text-xl font-semibold mb-2">
        Choix de l'évaluation
      </h2>
      <p className="text-sm max-w-md mx-auto mb-6" style={{ color: T.inkSoft }}>
        Sélectionne l'évaluation active pour insérer tes notes. Tu pourras en changer à tout moment
        depuis l'en-tête de l'application.
      </p>
      <div className="max-w-sm mx-auto space-y-2.5">
        {EVALUATION_TYPES.map((ev) => (
          <button
            key={ev}
            onClick={() => onSelect(ev)}
            className="w-full text-left px-4 py-3 rounded-lg text-sm font-medium"
            style={{ border: `2px solid ${T.goldLine}`, color: T.ink }}
          >
            {ev}
          </button>
        ))}
      </div>
    </div>
  );
}

function WelcomeScreen({ onStart }) {
  return (
    <div
      className="rounded-2xl overflow-hidden text-center"
      style={{ background: T.card, border: `1px solid ${T.line}` }}
    >
      <div className="px-6 pt-10 pb-8" style={{ background: T.ink }}>
        <div
          className="mx-auto mb-4 w-16 h-16 rounded-2xl flex items-center justify-center"
          style={{ background: T.gold, fontFamily: "'Fraunces', serif" }}
        >
          <span className="text-2xl font-bold" style={{ color: T.ink }}>NE</span>
        </div>
        <p className="text-xs tracking-widest uppercase mb-2" style={{ color: T.goldSoft, fontFamily: "'IBM Plex Mono', monospace" }}>
          Application pour directeurs d'école
        </p>
        <h1 style={{ fontFamily: "'Fraunces', serif", color: "#FFFFFF" }} className="text-2xl font-semibold">
          NoteExpress
        </h1>
      </div>

      <div className="px-6 py-7">
        <p className="text-sm mb-6" style={{ color: T.inkSoft }}>
          Importe la liste EducMaster de ta classe, saisis les notes en quelques secondes par élève,
          puis exporte un fichier prêt à réimporter — sans jongler avec les cellules Excel.
        </p>

        <div className="grid grid-cols-3 gap-3 mb-7 text-left">
          <div className="p-3 rounded-lg" style={{ background: T.greenSoft }}>
            <div className="text-xs font-semibold mb-1" style={{ color: T.green }}>1. Importer</div>
            <div className="text-xs" style={{ color: T.inkSoft }}>Ton fichier EducMaster</div>
          </div>
          <div className="p-3 rounded-lg" style={{ background: T.goldSoft }}>
            <div className="text-xs font-semibold mb-1" style={{ color: T.gold }}>2. Saisir</div>
            <div className="text-xs" style={{ color: T.inkSoft }}>Note + perfectionnement, en un geste</div>
          </div>
          <div className="p-3 rounded-lg" style={{ background: T.greenSoft }}>
            <div className="text-xs font-semibold mb-1" style={{ color: T.green }}>3. Exporter</div>
            <div className="text-xs" style={{ color: T.inkSoft }}>Prêt pour EducMaster</div>
          </div>
        </div>

        <button
          onClick={onStart}
          className="w-full px-5 py-3 rounded-lg font-medium text-sm"
          style={{ background: T.gold, color: "#FFFFFF" }}
        >
          Commencer →
        </button>
        <p className="text-xs mt-4" style={{ color: T.inkSoft }}>
          Conçu par S. Prince LOKO
        </p>
      </div>
    </div>
  );
}

function ImportScreen({ onFile, fileInputRef }) {
  const [dragOver, setDragOver] = useState(false);
  const [showHelp, setShowHelp] = useState(false);
  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        const f = e.dataTransfer.files?.[0];
        if (f) onFile(f);
      }}
      className="rounded-2xl p-10 text-center"
      style={{
        background: T.card,
        border: `2px dashed ${dragOver ? T.gold : T.goldLine}`,
      }}
    >
      <div
        className="mx-auto mb-5 w-14 h-14 rounded-full flex items-center justify-center"
        style={{ background: T.goldSoft, color: T.gold }}
      >
        <Upload size={24} />
      </div>
      <h2 style={{ fontFamily: "'Fraunces', serif" }} className="text-xl font-semibold mb-2">
        Importer la liste de la classe
      </h2>
      <p className="text-sm max-w-md mx-auto mb-4" style={{ color: T.inkSoft }}>
        Dépose ici le fichier Excel EducMaster (vide ou déjà rempli) — matricules, noms et prénoms seront
        récupérés automatiquement, ainsi que la liste des matières.
      </p>

      <button
        onClick={() => setShowHelp((v) => !v)}
        className="text-xs underline mb-6"
        style={{ color: T.green }}
      >
        {showHelp ? "Masquer l'aide" : "Je ne sais pas où trouver ce fichier sur mon téléphone"}
      </button>

      {showHelp && (
        <div
          className="text-left text-xs rounded-lg p-4 mb-6 max-w-md mx-auto"
          style={{ background: T.greenSoft, color: T.green }}
        >
          <ol className="list-decimal ml-4 space-y-1.5">
            <li>Ouvre l'application <strong>Fichiers</strong> (ou <strong>Mes Fichiers / My Files</strong>) sur ton téléphone.</li>
            <li>Va dans le dossier <strong>Téléchargements</strong> (Download) — c'est là que les fichiers reçus par WhatsApp, email ou navigateur atterrissent en général.</li>
            <li>Cherche le fichier envoyé par EducMaster (il se termine par <strong>.xlsx</strong>).</li>
            <li>Reviens ici et appuie sur le bouton ci-dessous — le sélecteur de fichiers de ton téléphone s'ouvrira au même endroit.</li>
          </ol>
        </div>
      )}

      <button
        onClick={() => fileInputRef.current?.click()}
        className="px-5 py-2.5 rounded-lg font-medium text-sm"
        style={{ background: T.ink, color: "#FFFFFF" }}
      >
        Choisir un fichier .xlsx
      </button>
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onFile(f);
        }}
      />
    </div>
  );
}

function StudentEntryTab({ roster, subjects, grades, onChangeCode, currentIndex, setCurrentIndex, attendance, onToggleAttendance }) {
  const safeIndex = Math.min(Math.max(currentIndex, 0), Math.max(roster.length - 1, 0));
  const student = roster[safeIndex];
  const inputRefs = useRef([]);

  useEffect(() => {
    if (currentIndex !== safeIndex) setCurrentIndex(safeIndex);
  }, [safeIndex, currentIndex, setCurrentIndex]);

  useEffect(() => {
    const t = setTimeout(() => inputRefs.current[0]?.focus(), 30);
    return () => clearTimeout(t);
  }, [safeIndex]);

  if (!student) {
    return (
      <div className="rounded-xl p-8 text-center text-sm" style={{ background: T.card, border: `1px solid ${T.line}`, color: T.inkSoft }}>
        Aucun élève à afficher.
      </div>
    );
  }

  const isComplete = (stu) =>
    attendance[stu.matricule] === false ||
    subjects.every((s) => {
      const g = grades[stu.matricule]?.[s.key];
      return g && g.obtenue !== "" && g.obtenue !== undefined;
    });

  const goToNextIncomplete = () => {
    const idx = roster.findIndex((s, i) => i > safeIndex && !isComplete(s));
    const fallback = roster.findIndex((s) => !isComplete(s));
    const target = idx !== -1 ? idx : fallback;
    if (target !== -1) setCurrentIndex(target);
  };

  const incompleteCount = roster.filter((s) => !isComplete(s)).length;
  const absent = attendance[student.matricule] === false;

  const handleCodeChange = (subjIdx, subjectKey, digits) => {
    onChangeCode(student.matricule, subjectKey, digits);
    if (isCodeComplete(digits)) {
      if (navigator.vibrate) navigator.vibrate(12);
      if (subjIdx < subjects.length - 1) {
        setTimeout(() => inputRefs.current[subjIdx + 1]?.focus(), 10);
      } else if (safeIndex < roster.length - 1) {
        setTimeout(() => setCurrentIndex(safeIndex + 1), 250);
      }
    }
  };

  const completedCount = subjects.filter((s) => {
    const g = grades[student.matricule]?.[s.key];
    return g && g.obtenue !== "" && g.obtenue !== undefined;
  }).length;

  return (
    <div className="rounded-xl overflow-hidden" style={{ background: T.card, border: `1px solid ${T.line}` }}>
      {incompleteCount > 0 && (
        <div
          className="px-4 py-2 flex items-center justify-between gap-2 text-xs"
          style={{ background: T.redSoft, color: T.red }}
        >
          <span className="flex items-center gap-1.5">
            <AlertTriangle size={14} /> {incompleteCount} élève(s) avec des notes manquantes
          </span>
          <button onClick={goToNextIncomplete} className="underline font-semibold whitespace-nowrap">
            Aller au prochain
          </button>
        </div>
      )}
      <div
        className="px-4 py-3 flex items-center justify-between gap-2"
        style={{ background: T.ink, color: "#FFFFFF" }}
      >
        <button
          onClick={() => safeIndex > 0 && setCurrentIndex(safeIndex - 1)}
          disabled={safeIndex === 0}
          className="p-1.5 rounded-full disabled:opacity-30"
          style={{ background: "rgba(255,255,255,0.1)" }}
        >
          <ChevronLeft size={18} />
        </button>
        <div className="text-center">
          <div className="text-xs" style={{ color: T.goldSoft, fontFamily: "'IBM Plex Mono', monospace" }}>
            Élève {safeIndex + 1} / {roster.length} · {completedCount}/{subjects.length} matières
          </div>
          <div style={{ fontFamily: "'Fraunces', serif" }} className="font-semibold">
            {student.nom} {student.prenoms}
          </div>
          <button
            onClick={() => onToggleAttendance(student.matricule)}
            className="mt-1 inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold"
            style={{ background: absent ? T.red : T.gold, color: "#fff" }}
          >
            {absent ? <UserX size={12} /> : <UserCheck size={12} />}
            {absent ? "Absent" : "Présent"}
          </button>
        </div>
        <button
          onClick={() => safeIndex < roster.length - 1 && setCurrentIndex(safeIndex + 1)}
          disabled={safeIndex === roster.length - 1}
          className="p-1.5 rounded-full disabled:opacity-30"
          style={{ background: "rgba(255,255,255,0.1)" }}
        >
          <ChevronRight size={18} />
        </button>
      </div>

      <select
        value={safeIndex}
        onChange={(e) => setCurrentIndex(Number(e.target.value))}
        className="w-full px-4 py-2 text-sm border-b outline-none"
        style={{ borderColor: T.line, color: T.inkSoft }}
      >
        {roster.map((s, i) => (
          <option key={s.matricule} value={i}>
            {i + 1}. {s.nom} {s.prenoms}
          </option>
        ))}
      </select>

      <div className="divide-y" style={{ borderColor: T.line, opacity: absent ? 0.5 : 1 }}>
        {subjects.map((s, idx) => {
          const g = grades[student.matricule]?.[s.key] || { obtenue: "", perfectionnement: "", rawCode: "" };
          const displayCode = g.rawCode !== undefined && g.rawCode !== "" ? g.rawCode : codeFromValues(g.obtenue, g.perfectionnement);
          const hasAny = g.obtenue !== "" && g.obtenue !== undefined;
          const total = hasAny ? (toNum(g.obtenue) || 0) + (toNum(g.perfectionnement) || 0) : null;
          const color = total === null ? T.inkSoft : total >= 10 ? T.green : T.red;
          return (
            <div key={s.key} className="p-4" style={{ borderColor: T.line }}>
              <div className="flex items-center justify-between mb-2">
                <span className="text-sm font-medium">{s.label}</span>
                {hasAny && (
                  <span className="inline-flex items-center gap-1 text-xs font-semibold" style={{ color: T.green }}>
                    <Check size={14} /> Fait
                  </span>
                )}
              </div>
              <input
                ref={(el) => (inputRefs.current[idx] = el)}
                type="text"
                inputMode="numeric"
                pattern="[0-9]*"
                maxLength={7}
                disabled={absent}
                placeholder={absent ? "Élève absent" : "Ex: 14,5 2"}
                value={displayCode}
                onChange={(e) => handleCodeChange(idx, s.key, sanitizeCode(e.target.value))}
                className="w-full text-center py-3 rounded-lg text-2xl font-semibold"
                style={{
                  border: `2px solid ${hasAny ? T.green : T.goldLine}`,
                  fontFamily: "'IBM Plex Mono', monospace",
                  letterSpacing: "0.2em",
                }}
              />
              {hasAny && (
                <div className="text-center text-xs mt-1.5" style={{ color, fontFamily: "'IBM Plex Mono', monospace" }}>
                  Note {formatNum(toNum(g.obtenue))} + Perf. {g.perfectionnement} = <strong>{formatNum(total)}/20</strong>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ElevesTab({ roster, rankByMatricule, onRemove, newStudent, setNewStudent, onAdd, attendance, onToggleAttendance }) {
  const absentCount = roster.filter((s) => attendance[s.matricule] === false).length;
  return (
    <div className="rounded-xl overflow-hidden" style={{ background: T.card, border: `1px solid ${T.line}` }}>
      <div className="px-4 py-2 text-xs flex gap-4" style={{ background: T.greenSoft, color: T.green, fontFamily: "'IBM Plex Mono', monospace" }}>
        <span>Inscrits : <strong>{roster.length}</strong></span>
        <span>Présents : <strong>{roster.length - absentCount}</strong></span>
        <span style={{ color: absentCount > 0 ? T.red : T.green }}>Absents : <strong>{absentCount}</strong></span>
      </div>
      <table className="w-full text-sm">
        <thead>
          <tr style={{ background: T.goldSoft }}>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft, fontFamily: "'IBM Plex Mono', monospace" }}>Matricule</th>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Nom</th>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Prénoms</th>
            <th className="text-right px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Moyenne</th>
            <th className="text-center px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Présence</th>
            <th className="px-3 py-2"></th>
          </tr>
        </thead>
        <tbody>
          {roster.map((s, i) => {
            const r = rankByMatricule[s.matricule];
            const absent = attendance[s.matricule] === false;
            return (
              <tr key={s.matricule} style={{ borderTop: `1px solid ${T.line}`, background: absent ? T.redSoft : i % 2 ? T.paper : T.card }}>
                <td className="px-3 py-2" style={{ fontFamily: "'IBM Plex Mono', monospace", color: T.inkSoft }}>{s.matricule}</td>
                <td className="px-3 py-2 font-medium">{s.nom}</td>
                <td className="px-3 py-2">{s.prenoms}</td>
                <td className="px-3 py-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>
                  {absent ? "—" : r?.moyenne !== null && r?.moyenne !== undefined ? r.moyenne.toFixed(2) : "—"}
                </td>
                <td className="px-3 py-2 text-center">
                  <button
                    onClick={() => onToggleAttendance(s.matricule)}
                    className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold"
                    style={{ background: absent ? T.redSoft : T.greenSoft, color: absent ? T.red : T.green }}
                  >
                    {absent ? <UserX size={13} /> : <UserCheck size={13} />}
                    {absent ? "Absent" : "Présent"}
                  </button>
                </td>
                <td className="px-3 py-2 text-right">
                  <button onClick={() => onRemove(s.matricule)} style={{ color: T.red }}>
                    <Trash2 size={15} />
                  </button>
                </td>
              </tr>
            );
          })}
          {roster.length === 0 && (
            <tr>
              <td colSpan={6} className="px-3 py-6 text-center text-sm" style={{ color: T.inkSoft }}>
                Aucun élève ne correspond à la recherche.
              </td>
            </tr>
          )}
        </tbody>
      </table>
      <div className="p-3 flex flex-col sm:flex-row gap-2" style={{ borderTop: `1px solid ${T.line}` }}>
        <input
          placeholder="Matricule (optionnel)"
          value={newStudent.matricule}
          onChange={(e) => setNewStudent({ ...newStudent, matricule: e.target.value })}
          className="px-3 py-2 rounded-lg text-sm flex-1"
          style={{ border: `1px solid ${T.line}` }}
        />
        <input
          placeholder="Nom"
          value={newStudent.nom}
          onChange={(e) => setNewStudent({ ...newStudent, nom: e.target.value })}
          className="px-3 py-2 rounded-lg text-sm flex-1"
          style={{ border: `1px solid ${T.line}` }}
        />
        <input
          placeholder="Prénoms"
          value={newStudent.prenoms}
          onChange={(e) => setNewStudent({ ...newStudent, prenoms: e.target.value })}
          className="px-3 py-2 rounded-lg text-sm flex-1"
          style={{ border: `1px solid ${T.line}` }}
        />
        <button onClick={onAdd} className="px-4 py-2 rounded-lg text-sm font-medium shrink-0" style={{ background: T.green, color: "#FFFFFF" }}>
          + Ajouter
        </button>
      </div>
    </div>
  );
}

function SubjectTab({ roster, subject, grades, onChangeCode, attendance }) {
  const inputRefs = useRef([]);

  const handleChange = (i, matricule, digits) => {
    onChangeCode(matricule, subject.key, digits);
    if (isCodeComplete(digits)) {
      if (navigator.vibrate) navigator.vibrate(12);
      if (i < roster.length - 1) {
        setTimeout(() => inputRefs.current[i + 1]?.focus(), 10);
      }
    }
  };

  const presentRoster = roster.filter((s) => attendance[s.matricule] !== false);
  const doneCount = presentRoster.filter((s) => {
    const g = grades[s.matricule]?.[subject.key];
    return g && g.obtenue !== "" && g.obtenue !== undefined;
  }).length;
  const pct = presentRoster.length > 0 ? Math.round((doneCount / presentRoster.length) * 100) : 0;
  const missingCount = presentRoster.length - doneCount;

  const goToFirstMissing = () => {
    const idx = roster.findIndex((s) => {
      if (attendance[s.matricule] === false) return false;
      const g = grades[s.matricule]?.[subject.key];
      return !(g && g.obtenue !== "" && g.obtenue !== undefined);
    });
    if (idx !== -1) {
      inputRefs.current[idx]?.focus();
      inputRefs.current[idx]?.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  };

  return (
    <div className="rounded-xl overflow-hidden" style={{ background: T.card, border: `1px solid ${T.line}` }}>
      <div className="px-4 py-2.5" style={{ background: T.greenSoft, borderBottom: `1px solid ${T.line}` }}>
        <div className="flex items-center justify-between text-xs mb-1.5" style={{ color: T.green }}>
          <span>
            Saisis la <strong>note obtenue</strong>, un <strong>espace</strong>, puis le <strong>perfectionnement</strong> (0, 1 ou 2).
            Exemple : <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>14,5 2</span> → 14,5 + 2 = <strong>16,5</strong>.
          </span>
        </div>
        <div className="flex items-center gap-2">
          <div className="flex-1 h-1.5 rounded-full overflow-hidden" style={{ background: T.goldLine }}>
            <div className="h-full rounded-full" style={{ width: `${pct}%`, background: T.green }} />
          </div>
          <span className="text-xs font-semibold whitespace-nowrap" style={{ color: T.green, fontFamily: "'IBM Plex Mono', monospace" }}>
            {doneCount}/{presentRoster.length}
          </span>
        </div>
      </div>

      {missingCount > 0 && (
        <div
          className="px-4 py-2 flex items-center justify-between gap-2 text-xs"
          style={{ background: T.redSoft, color: T.red }}
        >
          <span className="flex items-center gap-1.5">
            <AlertTriangle size={14} /> {missingCount} note(s) manquante(s)
          </span>
          <button onClick={goToFirstMissing} className="underline font-semibold whitespace-nowrap">
            Aller à la première
          </button>
        </div>
      )}

      <table className="w-full text-sm">
        <thead>
          <tr style={{ background: T.goldSoft }}>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Élève</th>
            <th className="text-center px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Code (OO-P)</th>
            <th className="text-center px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Détail</th>
            <th className="text-center px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Total</th>
          </tr>
        </thead>
        <tbody>
          {roster.map((s, i) => {
            const absent = attendance[s.matricule] === false;
            const g = grades[s.matricule]?.[subject.key] || { obtenue: "", perfectionnement: "", rawCode: "" };
            const displayCode = g.rawCode !== undefined && g.rawCode !== "" ? g.rawCode : codeFromValues(g.obtenue, g.perfectionnement);
            const hasAny = g.obtenue !== "" && g.obtenue !== undefined;
            const total = hasAny ? (toNum(g.obtenue) || 0) + (toNum(g.perfectionnement) || 0) : null;
            const color = total === null ? T.inkSoft : total >= 10 ? T.green : T.red;
            return (
              <tr key={s.matricule} style={{ borderTop: `1px solid ${T.line}`, background: absent ? T.redSoft : i % 2 ? T.paper : T.card, opacity: absent ? 0.6 : 1 }}>
                <td className="px-3 py-2">
                  <div className="font-medium">{s.nom}</div>
                  <div className="text-xs" style={{ color: T.inkSoft }}>{s.prenoms}{absent ? " · Absent" : ""}</div>
                </td>
                <td className="px-2 py-2">
                  <input
                    ref={(el) => (inputRefs.current[i] = el)}
                    type="text"
                    inputMode="numeric"
                    pattern="[0-9]*"
                    maxLength={7}
                    disabled={absent}
                    placeholder="Ex: 14,5 2"
                    value={displayCode}
                    onChange={(e) => handleChange(i, s.matricule, sanitizeCode(e.target.value))}
                    className="w-24 mx-auto block text-center px-2 py-1.5 rounded-md text-base font-semibold"
                    style={{ border: `1px solid ${T.line}`, fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.15em" }}
                  />
                </td>
                <td className="px-2 py-2 text-center text-xs" style={{ color: T.inkSoft, fontFamily: "'IBM Plex Mono', monospace" }}>
                  {hasAny ? `${g.obtenue} + ${g.perfectionnement}` : "—"}
                </td>
                <td className="px-3 py-2 text-center font-semibold" style={{ color, fontFamily: "'IBM Plex Mono', monospace" }}>
                  {total !== null ? formatNum(total) : "—"}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function RecapTab({ ranking }) {
  const withMoy = ranking.filter((r) => r.moyenne !== null);
  const admis = withMoy.filter((r) => r.moyenne >= PASSING_AVERAGE).length;
  const total = ranking.length;

  return (
    <div className="rounded-xl overflow-hidden" style={{ background: T.card, border: `1px solid ${T.line}` }}>
      <div
        className="px-4 py-3 flex flex-wrap items-center justify-between gap-2"
        style={{ background: T.ink, color: "#FFFFFF" }}
      >
        <span style={{ fontFamily: "'Fraunces', serif" }} className="font-semibold">
          Classement général
        </span>
        <span className="text-xs" style={{ color: T.goldSoft, fontFamily: "'IBM Plex Mono', monospace" }}>
          {admis} admis sur {total} · seuil de passage : {PASSING_AVERAGE}/20
        </span>
      </div>
      <table className="w-full text-sm">
        <thead>
          <tr style={{ background: T.goldSoft }}>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Rang</th>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Nom</th>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Prénoms</th>
            <th className="text-right px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Moyenne générale</th>
            <th className="text-center px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Décision</th>
          </tr>
        </thead>
        <tbody>
          {ranking.map((r, i) => {
            const isPassing = r.moyenne !== null && r.moyenne >= PASSING_AVERAGE;
            return (
              <tr key={r.matricule} style={{ borderTop: `1px solid ${T.line}`, background: i % 2 ? T.paper : T.card }}>
                <td className="px-3 py-2">
                  <span
                    className="inline-flex items-center justify-center w-7 h-7 rounded-full text-xs font-semibold"
                    style={{
                      background: r.rang === 1 ? T.gold : r.rang === 2 || r.rang === 3 ? T.goldSoft : "transparent",
                      color: r.rang === 1 ? "#fff" : T.ink,
                      border: r.rang && r.rang > 3 ? `1px solid ${T.line}` : "none",
                    }}
                  >
                    {r.rang ?? "–"}
                  </span>
                </td>
                <td className="px-3 py-2 font-medium">{r.nom}</td>
                <td className="px-3 py-2">{r.prenoms}</td>
                <td
                  className="px-3 py-2 text-right font-semibold"
                  style={{ fontFamily: "'IBM Plex Mono', monospace", color: r.moyenne === null ? T.inkSoft : isPassing ? T.green : T.red }}
                >
                  {r.moyenne !== null ? r.moyenne.toFixed(2) : r.absent ? "Absent(e)" : "En attente"}
                </td>
                <td className="px-3 py-2 text-center">
                  {r.moyenne === null ? (
                    <span className="text-xs" style={{ color: T.inkSoft }}>{r.absent ? "Absent(e)" : "—"}</span>
                  ) : (
                    <span
                      className="inline-block px-2.5 py-1 rounded-full text-xs font-semibold"
                      style={{
                        background: isPassing ? T.greenSoft : T.redSoft,
                        color: isPassing ? T.green : T.red,
                      }}
                    >
                      {isPassing ? "Admis(e)" : "Non admis(e)"}
                    </span>
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function StatsTab({ roster, subjects, grades, attendance }) {
  const presentRoster = roster.filter((s) => attendance[s.matricule] !== false);

  const stats = subjects.map((s) => {
    let sum = 0;
    let count = 0;
    let passCount = 0;
    presentRoster.forEach((stu) => {
      const g = grades[stu.matricule]?.[s.key];
      if (g && g.obtenue !== "" && g.obtenue !== undefined) {
        const total = (toNum(g.obtenue) || 0) + (toNum(g.perfectionnement) || 0);
        sum += total;
        count += 1;
        if (total >= SUBJECT_PASS) passCount += 1;
      }
    });
    const average = count > 0 ? sum / count : null;
    const passRate = count > 0 ? Math.round((passCount / count) * 100) : null;
    return { key: s.key, label: s.label, average, passRate, count };
  });

  const classAverage = (() => {
    const withAvg = stats.filter((s) => s.average !== null);
    if (withAvg.length === 0) return null;
    return withAvg.reduce((acc, s) => acc + s.average, 0) / withAvg.length;
  })();

  return (
    <div className="space-y-4">
      <div className="rounded-xl overflow-hidden" style={{ background: T.card, border: `1px solid ${T.line}` }}>
        <div className="px-4 py-3 flex items-center justify-between" style={{ background: T.ink, color: "#FFFFFF" }}>
          <span style={{ fontFamily: "'Fraunces', serif" }} className="font-semibold">Statistiques par matière</span>
          <span className="text-xs" style={{ color: T.goldSoft, fontFamily: "'IBM Plex Mono', monospace" }}>
            {presentRoster.length} élève(s) présent(s)
          </span>
        </div>
        {classAverage !== null && (
          <div className="px-4 py-2.5 text-xs" style={{ background: T.greenSoft, color: T.green }}>
            Moyenne générale de la classe (toutes matières confondues) : <strong>{classAverage.toFixed(2)}/20</strong>
          </div>
        )}
        <div className="divide-y" style={{ borderColor: T.line }}>
          {stats.map((s) => (
            <div key={s.key} className="p-4">
              <div className="flex items-center justify-between mb-1.5">
                <span className="text-sm font-medium">{s.label}</span>
                <span className="text-xs" style={{ color: T.inkSoft, fontFamily: "'IBM Plex Mono', monospace" }}>
                  {s.count}/{presentRoster.length} saisies
                </span>
              </div>
              {s.average === null ? (
                <p className="text-xs" style={{ color: T.inkSoft }}>Aucune note saisie pour l'instant.</p>
              ) : (
                <>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="flex-1 h-2 rounded-full overflow-hidden" style={{ background: T.goldLine }}>
                      <div
                        className="h-full rounded-full"
                        style={{ width: `${Math.min(100, (s.average / 20) * 100)}%`, background: s.average >= SUBJECT_PASS ? T.green : T.red }}
                      />
                    </div>
                    <span
                      className="text-xs font-semibold whitespace-nowrap"
                      style={{ color: s.average >= SUBJECT_PASS ? T.green : T.red, fontFamily: "'IBM Plex Mono', monospace" }}
                    >
                      {s.average.toFixed(2)}/20
                    </span>
                  </div>
                  <p className="text-xs" style={{ color: T.inkSoft }}>
                    Taux de réussite (≥ {SUBJECT_PASS}/20) : <strong>{s.passRate}%</strong>
                  </p>
                </>
              )}
            </div>
          ))}
          {stats.length === 0 && (
            <p className="p-4 text-sm text-center" style={{ color: T.inkSoft }}>Aucune matière à afficher.</p>
          )}
        </div>
      </div>
    </div>
  );
}
