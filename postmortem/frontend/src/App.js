//written by Smriti Wadhwa, D365 Software Dev Intern
//August 2025

import React, { useEffect, useState, useRef } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import './App.css';

function App() {
  const [questions, setQuestions] = useState([]);
  const [projectOptions, setProjectOptions] = useState([]);
  const [phaseOptions, setPhaseOptions] = useState([]);
  const [answers, setAnswers] = useState({});
  const [selectedProject, setSelectedProject] = useState("");
  const [selectedPhase, setSelectedPhase] = useState("");
  const fileInputRef = useRef(null);

  const slugify = (str = "") =>
    str.toString().trim().toLowerCase().replace(/[^\w-]+/g, "_");

  // Load Questions and Projects
  useEffect(() => {
    fetch("/questions.xlsx")
      .then(res => res.arrayBuffer())
      .then(buffer => {
        const wb = XLSX.read(buffer, { type: "array" });
        const qs = XLSX.utils.sheet_to_json(wb.Sheets["Questions"], { defval: "" });
        setQuestions(qs);
        const rows = XLSX.utils.sheet_to_json(wb.Sheets["Projects"], { defval: "" });
        setProjectOptions(rows.map(r => r.Projects).filter(v => v));
        setPhaseOptions(rows.map(r => r.Phases).filter(v => v));
      });
  }, []);

  // Load draft from Excel
  const loadDraft = file => {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const data = XLSX.utils.sheet_to_json(wb.Sheets['Responses'], { defval: "" });
      if (!data.length) return;
      // first row has project/phase
      const row0 = data[0];
      setSelectedProject(row0.Project || "");
      setSelectedPhase(row0.Phase || "");
      const newAns = {};
      // map answers for QIDs > 2
      data.forEach(r => {
        if (r.QuestionID > 2) newAns[r.QuestionID] = r.Answer;
      });
      setAnswers(newAns);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleChange = (id, value) => {
    setAnswers(prev => ({ ...prev, [id]: value }));
  };

  const handleSubmit = () => {
    // confirm before final submission
    if (!window.confirm("Are you sure you want to submit? You cannot make edits after this.")) {
      return;
    }
    const missing = questions
      .filter(q => q.Required?.toLowerCase() === "yes")
      .filter(q => {
        if (q['#'] === 1) return !selectedProject;
        if (q['#'] === 2) return !selectedPhase;
        return !answers[q['#']];
      });
    if (missing.length) {
    alert(
      "Please fill out the following required questions:\n\n" +
      missing.map(q => `• ${q.Question}`).join("\n")
    );
    return;
  }

    const rows = [];
    // include project & phase as first two rows
    rows.push({ Project: selectedProject, Phase: selectedPhase, QuestionID: 1, Question: 'Project', Answer: selectedProject });
    rows.push({ Project: selectedProject, Phase: selectedPhase, QuestionID: 2, Question: 'Phase', Answer: selectedPhase });
    Object.entries(answers).forEach(([id, ans]) => {
      const q = questions.find(x => x['#'] === Number(id)) || {};
      rows.push({ Project: selectedProject, Phase: selectedPhase, QuestionID: id, Question: q.Question, Answer: ans });
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Responses');
    const fileName = `${slugify(selectedProject).toUpperCase()}_${slugify(selectedPhase).toUpperCase()}_POSTMORTEM.xlsx`;
    XLSX.writeFile(wb, fileName);

    // reset
    setAnswers({});
    setSelectedProject("");
    setSelectedPhase("");
  };

  const saveDraft = () => {
    const missing = questions
      .filter(q => q.Required?.toLowerCase() === "yes")
      .filter(q => {
        if (q['#'] === 1) return !selectedProject;
        if (q['#'] === 2) return !selectedPhase;
        return !answers[q['#']];
      });
    if (missing.length) {
    alert(
      "Please fill out the following required questions:\n\n" +
      missing.map(q => `• ${q.Question}`).join("\n")
    );
    return;
  }

    const rows = [];
    // include project & phase as first two rows
    rows.push({ Project: selectedProject, Phase: selectedPhase, QuestionID: 1, Question: 'Project', Answer: selectedProject });
    rows.push({ Project: selectedProject, Phase: selectedPhase, QuestionID: 2, Question: 'Phase', Answer: selectedPhase });
    Object.entries(answers).forEach(([id, ans]) => {
      const q = questions.find(x => x['#'] === Number(id)) || {};
      rows.push({ Project: selectedProject, Phase: selectedPhase, QuestionID: id, Question: q.Question, Answer: ans });
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Responses');
   const fileName = `DRAFT_${slugify(selectedProject).toUpperCase()}_${slugify(selectedPhase).toUpperCase()}_POSTMORTEM.xlsx`;
    XLSX.writeFile(wb, fileName);
    setAnswers({});
    setSelectedProject("");
    setSelectedPhase("");
  };

  // filter dependent
  const visible = questions.filter(q => {
    const dep = q['Dependent Question'];
    if(dep === 1 && selectedProject){
      return q.Value.toLowerCase() === selectedProject.toLowerCase()
    }
    if(dep === 2 && selectedPhase){
      return q.Value.toLowerCase() === selectedPhase.toLowerCase()
    }
    if (!dep){
      return true;
    } 
    const answerValue = answers[dep];
    if (!answerValue) return false; // If no answer yet, don't show dependent question
    return answerValue.toString().toLowerCase() === q.Value.toLowerCase();

  });

  // group by category
  const grouped = visible.reduce((acc, q) => {
    (acc[q.Category] ||= []).push(q);
    return acc;
  }, {});

  return (
    <div className="App">
      <header className="App-header" style={{ textAlign: 'center' }}>
        <img src="https://strategiccfo360.com/wp-content/uploads/2021/02/RSM-logo-696x362.png" alt="RSM logo" width="25%" />
        <h1>Postmortem Survey</h1>
        <button
          onClick={() => fileInputRef.current.click()}
          style={{ margin: '10px 0', display: 'block', marginLeft: 'auto', marginRight: 'auto' }}
        >Open Draft</button>
        <input
  type="file"
  accept=".xlsx"
  ref={fileInputRef}
  style={{ display: 'none' }}
  onChange={e => {
    const file = e.target.files[0];
    if (!file) return;
    // only allow filenames that include "DRAFT" (case‐insensitive)
    if (!file.name.toUpperCase().includes('DRAFT')) {
      alert('Please select a draft file (filename must include "DRAFT").');
      e.target.value = null;  // clear selection
      return;
    }
    loadDraft(file);
    e.target.value = null;    // clear for next open
  }}
/>

      </header>
<div className="App-body" style={{ textAlign: 'left', margin: '20px' }}>
  {Object.entries(grouped).map(([category, qs]) => {
    // ─── PRE-FILTER QUESTIONS ────────────────────────────────────────────────
    const filteredQs = qs.filter(q => {
      const num = Number(q['#']);

      // 1 & 2 always show (Project & Phase dropdowns)
      if (num === 1 || num === 2) return true;

      // ── PHASE-GATE ────────────────────────────────────────────────────────
      const phaseVal = (q.Phase ?? '').toString().trim().toLowerCase();
      const selPhase = (selectedPhase ?? '').toString().trim().toLowerCase();
      if (selPhase && phaseVal !== 'all' && phaseVal !== selPhase) {
        return false;
      }

      // ── DEPENDENCY-GATE ───────────────────────────────────────────────────
      // only if the Dependent Question cell is not blank
      const rawDep = q['Dependent Question'];
      const rawDepStr = rawDep != null ? rawDep.toString().trim() : '';
      if (rawDepStr !== '') {
        const depNum = Number(rawDepStr);
        // only proceed if it parsed to a real question number > 0
        if (!isNaN(depNum) && depNum > 0) {
          // if they depend on Phase (#2), compare against selectedPhase
          const actual = depNum === 2
            ? selPhase
            : (answers[depNum] ?? '').toString().trim().toLowerCase();
          const required = (q.Value ?? '').toString().trim().toLowerCase();
          if (actual !== required) {
            return false;
          }
        }
      }

      return true;
    });

    return (
      <section key={category} style={{ marginBottom: 40 }}>
        <h2>{category}</h2>

        {/* ─── RENDER ONLY filteredQs ────────────────────────────────────────── */}
        {filteredQs.map(q => {
          const opts      = (q.Options || '').split(',').map(o => o.trim().toLowerCase());
          const isOption  = opts.length > 0 && q.Type?.toString().trim().toLowerCase() === 'option';
          const isYesNo   = isOption && opts.length === 2 && opts.includes('yes') && opts.includes('no');
          const isTraffic = isOption && opts.length === 3 && opts.includes('green') && opts.includes('yellow') && opts.includes('red');

          // Project (#1)
          if (q['#'] === 1) {
            return (
              <div key="project" style={{ margin: '12px 0' }}>
                <label style={{ fontWeight: 500 }}>
                  Project<span style={{ color:'red' }}>*</span>
                </label>
                <select
                  value={selectedProject}
                  onChange={e => setSelectedProject(e.target.value)}
                  style={{ padding:8, width:'100%' }}
                >
                  <option value="">— select —</option>
                  {projectOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                </select>
              </div>
            );
          }

          // Phase (#2)
          if (q['#'] === 2) {
            return (
              <div key="phase" style={{ margin: '12px 0' }}>
                <label style={{ fontWeight: 500 }}>
                  Phase<span style={{ color:'red' }}>*</span>
                </label>
                <select
                  value={selectedPhase}
                  onChange={e => setSelectedPhase(e.target.value)}
                  style={{ padding:8, width:'100%' }}
                >
                  <option value="">— select —</option>
                  {phaseOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                </select>
              </div>
            );
          }

          // All other questions (exactly as before)
          return (
            <div key={q['#']} style={{ margin: '12px 0' }}>
              <label style={{ display:'block', marginBottom:4, fontWeight:500 }}>
                {q.Question}
                {q.Required?.toString().trim().toLowerCase() === 'yes' && <span style={{ color:'red' }}>*</span>}
              </label>

              {isYesNo ? (
                <div className="radio-group">
                  {opts.map(opt => (
                    <label key={opt} className="radio-label yesno">
                      <input
                        type="radio"
                        name={`q${q['#']}`} value={opt}
                        checked={answers[q['#']] === opt}
                        onChange={() => handleChange(q['#'], opt)}
                        className={`radio-input yesno-${opt}`}
                      />{opt.charAt(0).toUpperCase() + opt.slice(1)}
                    </label>
                  ))}
                </div>
              ) : isTraffic ? (
                <div className="radio-group">
                  {opts.map(opt => (
                    <label key={opt} className={`radio-label ${opt}`}>
                      <input
                        type="radio"
                        name={`q${q['#']}`} value={opt}
                        checked={answers[q['#']] === opt}
                        onChange={() => handleChange(q['#'], opt)}
                        className={`radio-input ${opt}`}
                      />{opt.charAt(0).toUpperCase() + opt.slice(1)}
                    </label>
                  ))}
                </div>
              ) : isOption ? (
                <select
                  value={answers[q['#']] || ''}
                  onChange={e => handleChange(q['#'], e.target.value)}
                  style={{ padding:8, width:'100%' }}
                >
                  <option value="">— select —</option>
                  {opts.map(opt => <option key={opt} value={opt}>{opt.charAt(0).toUpperCase() + opt.slice(1)}</option>)}
                </select>
              ) : (
                <input
                  type="text"
                  value={answers[q['#']] || ''}
                  onChange={e => handleChange(q['#'], e.target.value)}
                  style={{ padding:8, width:'100%' }}
                />
              )}
            </div>
          );
        })}
      </section>
    );
  })}

  <button onClick={saveDraft} style={{ marginRight:10 }}>Save Draft</button>
  <button onClick={handleSubmit}>Submit</button>
</div>
</div>

  );
}

export default App;
