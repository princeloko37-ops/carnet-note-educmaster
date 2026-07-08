import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { Upload, Download, Trash2, Search, Users, ClipboardList, Award, PenLine, RotateCcw, Check, ChevronLeft, ChevronRight } from "lucide-react";

const T = {
  paper: "#F8F4EA",
  ink: "#1E2A44",
  inkSoft: "#5B6478",
  gold: "#B4841F",
  goldSoft: "#EADFC0",
  goldLine: "#DDCE9F",
  green: "#2F6B4F",
  greenSoft: "#E4EEE8",
  red: "#AD3A32",
  redSoft: "#F5E4E1",
  card: "#FFFFFF",
  line: "#E7DFCB",
};

const FONT_LINK = "https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap";

const STORAGE_KEY = "carnet-notes:current";
const PASSING_AVERAGE = 9; // moyenne générale minimale pour passer en classe supérieure

function codeFromValues(obtenue, perfectionnement) {
  const oNum = toNum(obtenue);
  const pNum = toNum(perfectionnement);
  if (oNum === null && pNum === null) return "";
  const o = String(Math.max(0, Math.round(oNum || 0))).padStart(2, "0").slice(-2);
  const p = String(Math.max(0, Math.round(pNum || 0))).padStart(2, "0").slice(-2);
  return o + p;
}

function parseCode(code) {
  if (!code || code.length < 4) return { obtenue: "", perfectionnement: "" };
  return {
    obtenue: String(parseInt(code.slice(0, 2), 10)),
    perfectionnement: String(parseInt(code.slice(2, 4), 10)),
  };
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

function computeRanking(roster, subjects, grades) {
  const list = roster.map((stu) => {
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
    const moyenne = count > 0 ? sum / count : null;
    return { ...stu, moyenne, count };
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
  const [className, setClassName] = useState("");
  const [activeTab, setActiveTab] = useState("eleves");
  const [studentIndex, setStudentIndex] = useState(0);
  const [search, setSearch] = useState("");
  const [loaded, setLoaded] = useState(false);
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved
  const [newStudent, setNewStudent] = useState({ matricule: "", nom: "", prenoms: "" });
  const fileInputRef = useRef(null);
  const saveTimer = useRef(null);

  // Load persisted data once (stored locally in this browser)
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        setRoster(parsed.roster || []);
        setSubjects(parsed.subjects || []);
        setGrades(parsed.grades || {});
        setClassName(parsed.className || "");
        if (parsed.roster && parsed.roster.length > 0) setActiveTab("eleves");
      }
    } catch (e) {
      // no saved data yet
    }
    setLoaded(true);
  }, []);

  // Autosave (debounced)
  useEffect(() => {
    if (!loaded) return;
    if (roster.length === 0) return;
    setSaveState("saving");
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      try {
        localStorage.setItem(
          STORAGE_KEY,
          JSON.stringify({ roster, subjects, grades, className })
        );
        setSaveState("saved");
      } catch (e) {
        setSaveState("idle");
      }
    }, 700);
    return () => clearTimeout(saveTimer.current);
  }, [roster, subjects, grades, className, loaded]);

  const ranking = useMemo(() => computeRanking(roster, subjects, grades), [roster, subjects, grades]);
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
        setActiveTab("eleves");
      } catch (err) {
        alert("Impossible de lire ce fichier. Vérifie qu'il s'agit bien d'un export EducMaster (.xlsx).");
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const updateGrade = (matricule, subjectKey, field, value) => {
    setGrades((prev) => ({
      ...prev,
      [matricule]: {
        ...prev[matricule],
        [subjectKey]: {
          ...prev[matricule]?.[subjectKey],
          [field]: value,
        },
      },
    }));
  };

  const updateGradeCode = (matricule, subjectKey, rawCode) => {
    const { obtenue, perfectionnement } = parseCode(rawCode);
    setGrades((prev) => ({
      ...prev,
      [matricule]: {
        ...prev[matricule],
        [subjectKey]: { rawCode, obtenue, perfectionnement },
      },
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

  const resetAll = async () => {
    if (!window.confirm("Effacer toute la classe et repartir de zéro ? Cette action est irréversible.")) return;
    setRoster([]);
    setSubjects([]);
    setGrades({});
    setClassName("");
    setActiveTab("eleves");
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch (e) {}
  };

  const handleExport = () => {
    if (roster.length === 0) return;
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
        row.push(g.obtenue ?? "", g.perfectionnement ?? "");
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
      const decision = r.moyenne === null ? "-" : r.moyenne >= PASSING_AVERAGE ? "Admis(e)" : "Non admis(e)";
      recapData.push([r.nom, r.prenoms, r.matricule, r.moyenne !== null ? Math.round(r.moyenne * 100) / 100 : "-", r.rang ?? "-", decision]);
    });
    const ws2 = XLSX.utils.aoa_to_sheet(recapData);
    ws2["!cols"] = [{ wch: 18 }, { wch: 22 }, { wch: 16 }, { wch: 16 }, { wch: 8 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws2, "Récapitulatif");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Notes_${(className || "Classe").replace(/\s+/g, "_")}_EducMaster.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const activeSubject = subjects.find((s) => s.key === activeTab);

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
          <div>
            <p style={{ color: T.goldSoft, fontFamily: "'IBM Plex Mono', monospace" }} className="text-xs tracking-widest uppercase mb-1">
              Carnet de Notes · Import EducMaster
            </p>
            <input
              value={className}
              onChange={(e) => setClassName(e.target.value)}
              placeholder="Nom de la classe (ex : CM2 Ganmi A)"
              style={{ fontFamily: "'Fraunces', serif", color: "#FFFFFF", borderBottom: `1px solid ${T.inkSoft}` }}
              className="bg-transparent text-2xl sm:text-3xl font-semibold outline-none w-full max-w-md placeholder-slate-400"
            />
          </div>
          {roster.length > 0 && (
            <button onClick={resetAll} title="Repartir de zéro" className="shrink-0 p-2 rounded-full" style={{ color: T.goldSoft }}>
              <RotateCcw size={20} />
            </button>
          )}
        </div>
      </div>

      <div className="max-w-5xl mx-auto px-4 sm:px-8 py-6">
        {roster.length === 0 ? (
          <ImportScreen onFile={handleFile} fileInputRef={fileInputRef} />
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
                onClick={handleExport}
                className="flex items-center justify-center gap-2 px-4 py-2 rounded-lg font-medium text-sm shrink-0"
                style={{ background: T.gold, color: "#FFFFFF" }}
              >
                <Download size={16} />
                Exporter pour EducMaster
              </button>
            </div>

            <p className="text-xs mb-4" style={{ color: T.inkSoft, fontFamily: "'IBM Plex Mono', monospace" }}>
              {saveState === "saving" ? "Enregistrement…" : saveState === "saved" ? "✓ Enregistré sur cet appareil" : ""}
              {"  ·  "} {roster.length} élève{roster.length > 1 ? "s" : ""}
            </p>

            {activeTab === "eleves" && (
              <ElevesTab roster={filteredRoster} rankByMatricule={rankByMatricule} onRemove={removeStudent} newStudent={newStudent} setNewStudent={setNewStudent} onAdd={addStudent} />
            )}

            {activeTab === "parEleve" && (
              <StudentEntryTab
                roster={roster}
                subjects={subjects}
                grades={grades}
                onChangeCode={updateGradeCode}
                currentIndex={studentIndex}
                setCurrentIndex={setStudentIndex}
              />
            )}

            {activeSubject && (
              <SubjectTab
                roster={filteredRoster}
                subject={activeSubject}
                grades={grades}
                onChangeCode={updateGradeCode}
              />
            )}

            {activeTab === "recap" && <RecapTab ranking={search.trim() ? ranking.filter((r) => filteredRoster.some((f) => f.matricule === r.matricule)) : ranking} />}
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

function ImportScreen({ onFile, fileInputRef }) {
  const [dragOver, setDragOver] = useState(false);
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
      <p className="text-sm max-w-md mx-auto mb-6" style={{ color: T.inkSoft }}>
        Dépose ici le fichier Excel EducMaster (vide ou déjà rempli) — matricules, noms et prénoms seront
        récupérés automatiquement, ainsi que la liste des matières.
      </p>
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

function StudentEntryTab({ roster, subjects, grades, onChangeCode, currentIndex, setCurrentIndex }) {
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

  const handleCodeChange = (subjIdx, subjectKey, digits) => {
    onChangeCode(student.matricule, subjectKey, digits);
    if (digits.length === 4) {
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

      <div className="divide-y" style={{ borderColor: T.line }}>
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
                maxLength={4}
                placeholder="Saisir 4 chiffres (ex: 1402)"
                value={displayCode}
                onChange={(e) => handleCodeChange(idx, s.key, e.target.value.replace(/\D/g, "").slice(0, 4))}
                className="w-full text-center py-3 rounded-lg text-2xl font-semibold"
                style={{
                  border: `2px solid ${hasAny ? T.green : T.goldLine}`,
                  fontFamily: "'IBM Plex Mono', monospace",
                  letterSpacing: "0.2em",
                }}
              />
              {hasAny && (
                <div className="text-center text-xs mt-1.5" style={{ color, fontFamily: "'IBM Plex Mono', monospace" }}>
                  Note {g.obtenue} + Perf. {g.perfectionnement} = <strong>{total}/20</strong>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ElevesTab({ roster, rankByMatricule, onRemove, newStudent, setNewStudent, onAdd }) {
  return (
    <div className="rounded-xl overflow-hidden" style={{ background: T.card, border: `1px solid ${T.line}` }}>
      <table className="w-full text-sm">
        <thead>
          <tr style={{ background: T.goldSoft }}>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft, fontFamily: "'IBM Plex Mono', monospace" }}>Matricule</th>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Nom</th>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Prénoms</th>
            <th className="text-right px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Moyenne</th>
            <th className="px-3 py-2"></th>
          </tr>
        </thead>
        <tbody>
          {roster.map((s, i) => {
            const r = rankByMatricule[s.matricule];
            return (
              <tr key={s.matricule} style={{ borderTop: `1px solid ${T.line}`, background: i % 2 ? T.paper : T.card }}>
                <td className="px-3 py-2" style={{ fontFamily: "'IBM Plex Mono', monospace", color: T.inkSoft }}>{s.matricule}</td>
                <td className="px-3 py-2 font-medium">{s.nom}</td>
                <td className="px-3 py-2">{s.prenoms}</td>
                <td className="px-3 py-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>
                  {r?.moyenne !== null && r?.moyenne !== undefined ? r.moyenne.toFixed(2) : "—"}
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
              <td colSpan={5} className="px-3 py-6 text-center text-sm" style={{ color: T.inkSoft }}>
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

function SubjectTab({ roster, subject, grades, onChangeCode }) {
  const inputRefs = useRef([]);

  const handleChange = (i, matricule, digits) => {
    onChangeCode(matricule, subject.key, digits);
    if (digits.length === 4 && i < roster.length - 1) {
      setTimeout(() => inputRefs.current[i + 1]?.focus(), 10);
    }
  };

  return (
    <div className="rounded-xl overflow-hidden" style={{ background: T.card, border: `1px solid ${T.line}` }}>
      <div className="px-4 py-2.5 text-xs" style={{ background: T.greenSoft, color: T.green, borderBottom: `1px solid ${T.line}` }}>
        Saisis <strong>4 chiffres</strong> : les 2 premiers = note obtenue, les 2 derniers = perfectionnement.
        Exemple : <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>1202</span> → 12 + 02 = <strong>14</strong>.
        Le champ suivant se sélectionne automatiquement.
      </div>
      <table className="w-full text-sm">
        <thead>
          <tr style={{ background: T.goldSoft }}>
            <th className="text-left px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Élève</th>
            <th className="text-center px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Code (OOPP)</th>
            <th className="text-center px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Détail</th>
            <th className="text-center px-3 py-2 font-medium" style={{ color: T.inkSoft }}>Total</th>
          </tr>
        </thead>
        <tbody>
          {roster.map((s, i) => {
            const g = grades[s.matricule]?.[subject.key] || { obtenue: "", perfectionnement: "", rawCode: "" };
            const displayCode = g.rawCode !== undefined && g.rawCode !== "" ? g.rawCode : codeFromValues(g.obtenue, g.perfectionnement);
            const hasAny = g.obtenue !== "" && g.obtenue !== undefined;
            const total = hasAny ? (toNum(g.obtenue) || 0) + (toNum(g.perfectionnement) || 0) : null;
            const color = total === null ? T.inkSoft : total >= 10 ? T.green : T.red;
            return (
              <tr key={s.matricule} style={{ borderTop: `1px solid ${T.line}`, background: i % 2 ? T.paper : T.card }}>
                <td className="px-3 py-2">
                  <div className="font-medium">{s.nom}</div>
                  <div className="text-xs" style={{ color: T.inkSoft }}>{s.prenoms}</div>
                </td>
                <td className="px-2 py-2">
                  <input
                    ref={(el) => (inputRefs.current[i] = el)}
                    type="text"
                    inputMode="numeric"
                    maxLength={4}
                    placeholder="0000"
                    value={displayCode}
                    onChange={(e) => handleChange(i, s.matricule, e.target.value.replace(/\D/g, "").slice(0, 4))}
                    className="w-24 mx-auto block text-center px-2 py-1.5 rounded-md text-base font-semibold"
                    style={{ border: `1px solid ${T.line}`, fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.15em" }}
                  />
                </td>
                <td className="px-2 py-2 text-center text-xs" style={{ color: T.inkSoft, fontFamily: "'IBM Plex Mono', monospace" }}>
                  {hasAny ? `${g.obtenue} + ${g.perfectionnement}` : "—"}
                </td>
                <td className="px-3 py-2 text-center font-semibold" style={{ color, fontFamily: "'IBM Plex Mono', monospace" }}>
                  {total !== null ? total : "—"}
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
                  {r.moyenne !== null ? r.moyenne.toFixed(2) : "En attente"}
                </td>
                <td className="px-3 py-2 text-center">
                  {r.moyenne === null ? (
                    <span className="text-xs" style={{ color: T.inkSoft }}>—</span>
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
