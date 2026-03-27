// =====================================================================
//  FIREBASE CONFIGURATION
//  Legge le classi dal progetto classroomanager (Realtime Database)
// =====================================================================
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyA6O12KFRkZ4gGBi2LEGKZni33c3a2NBcU",
  authDomain: "classroomanager.firebaseapp.com",
  databaseURL: "https://classroomanager-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "classroomanager",
  storageBucket: "classroomanager.firebasestorage.app",
  messagingSenderId: "995055049853",
  appId: "1:995055049853:web:5b3674ae12f7c0b504e514"
};

// =====================================================================
//  STRUTTURA DEL DATABASE classroomanager:
//  /users/{uid}/classes/[array]
//    - id: "1769021271490"         ← ID della classe
//    - name: "1A"                  ← nome della classe
//    - students: [                 ← array di studenti (NO campo id!)
//        { fullName: "Rossi Mario", displayName: "Mario R." },
//        ...
//      ]
//    - seatingByClassroom: { ... }
//
//  IMPORTANTE: gli studenti non hanno un campo "id" proprio.
//  Usiamo "fullName" (es. "Rossi Mario") come identificatore stabile.
// =====================================================================
// Non modificare questo campo — è quello usato da classroomanager
const FB_STUDENT_FULLNAME_FIELD = "fullName";    // identificatore stabile
const FB_STUDENT_DISPLAY_FIELD  = "displayName"; // nome visualizzato (es. "Mario R.")

// =====================================================================
//  Firebase globals
// =====================================================================
let fbApp = null;
let fbDb = null;
let fbAuth = null;
let fbUser = null;
let fbClassesUnsubscribe = null;  // listener classi (read-only da classroomanager)
let fbGradingUnsubscribe = null;  // listener voti/test (read-write, nostro)
let fbSaveTimer = null;           // debounce per non scrivere su Firebase a ogni tasto
let fbIgnoreGrading = false;      // evita loop write→listen→write
let tableRenderTimer = null;      // debounce per non ricaricare la tabella durante navigazione frecce
let commentModalContext = null;   // contesto cella aperta nel dialog commento
let isNavigatingWithArrows = false; // flag per evitare renderTestTable durante navigazione frecce

// =====================================================================
//  STORAGE KEY (invariato)
// =====================================================================
const STORAGE_KEY = "teacher-grading-data-v1";

// Nota: le classi NON sono in defaultData perché arrivano da Firebase.
// Se Firebase non è connesso o l'utente non è autenticato,
// lo stato partirà con classes = [].
const defaultData = {
  classes: [],
  tests: [
    {
      id: "test-1",
      title: "English Midterm Exam",
      sections: [
        {
          id: "sec-grammar",
          name: "Grammar",
          weight: 2,
          max: 10,
          subsections: [
            { id: "sub-g1", name: "Exercise 1", weight: 1, max: 3 },
            { id: "sub-g2", name: "Exercise 2", weight: 1, max: 3 },
            { id: "sub-g3", name: "Exercise 3", weight: 1, max: 4 },
          ],
        },
        {
          id: "sec-vocab",
          name: "Vocabulary",
          weight: 1,
          max: 8,
          subsections: [{ id: "sub-v1", name: "Exercise 4", weight: 1, max: 8 }],
        },
        {
          id: "sec-read",
          name: "Reading",
          weight: 3,
          max: 15,
          subsections: [],
        },
        {
          id: "sec-write",
          name: "Writing",
          weight: 2,
          max: 12,
          subsections: [],
        },
      ],
    },
  ],
  selectedClassId: "class-1",
  selectedTestId: "test-1",
  selectedTestVersionId: null,
  selectedConfigVersionId: null,
  view: "home",
};

const state = loadState();

const navHomeBtn = document.getElementById("navHomeBtn");
const navTestsBtn = document.getElementById("navTestsBtn");
const navEvaluationBtn = document.getElementById("navEvaluationBtn");
const navConfigBtn = document.getElementById("navConfigBtn");
const resetBtn = document.getElementById("resetBtn");

const homeView = document.getElementById("homeView");
const classView = document.getElementById("classView");
const testsView = document.getElementById("testsView");
const testView = document.getElementById("testView");
const configView = document.getElementById("configView");

const classList = document.getElementById("classList");
const newClassNameInput = document.getElementById("newClassNameInput");
const addClassBtn = document.getElementById("addClassBtn");

const backToHomeBtn = document.getElementById("backToHomeBtn");
const classSelect = document.getElementById("classSelect");
const classRenameInput = document.getElementById("classRenameInput");
const renameClassBtn = document.getElementById("renameClassBtn");
const deleteClassBtn = document.getElementById("deleteClassBtn");
const addStudentBtn = document.getElementById("addStudentBtn");
const bulkStudentsInput = document.getElementById("bulkStudentsInput");
const addBulkStudentsBtn = document.getElementById("addBulkStudentsBtn");
const classStudentsTable = document.getElementById("classStudentsTable");

const testsList = document.getElementById("testsList");
const addTestBtn = document.getElementById("addTestBtn");

const testClassSelect = document.getElementById("testClassSelect");
const testSelect = document.getElementById("testSelect");
const testVersionSelect = document.getElementById("testVersionSelect");
const exportBtn = document.getElementById("exportBtn");
const gradeTable = document.getElementById("gradeTable");
const warningArea = document.getElementById("warningArea");

const configTestSelect = document.getElementById("configTestSelect");
const testTitleInput = document.getElementById("testTitle");
const configVersionSelect = document.getElementById("configVersionSelect");
const versionNameInput = document.getElementById("versionNameInput");
const facilitatedVersionSelect = document.getElementById(
  "facilitatedVersionSelect"
);
const addVersionBtn = document.getElementById("addVersionBtn");
const addSectionBtn = document.getElementById("addSectionBtn");
const sectionsContainer = document.getElementById("sectionsContainer");

const sectionTemplate = document.getElementById("sectionTemplate");
const subsectionTemplate = document.getElementById("subsectionTemplate");

init();

function init() {
  // Avvia Firebase (auth + listener classi)
  initFirebase();

  ensureClassState();
  ensureTestState();

  // Auth UI
  const loginBtn = document.getElementById("loginBtn");
  const logoutBtn = document.getElementById("logoutBtn");
  if (loginBtn) {
    loginBtn.addEventListener("click", () => {
      const provider = new firebase.auth.GoogleAuthProvider();
      fbAuth.signInWithPopup(provider).catch((err) => {
        setFirebaseStatus("❌ Login fallito: " + err.message, "error");
      });
    });
  }
  if (logoutBtn) {
    logoutBtn.addEventListener("click", () => {
      fbAuth.signOut();
    });
  }

  navHomeBtn.addEventListener("click", () => setView("home"));
  navTestsBtn.addEventListener("click", () => setView("tests"));
  navEvaluationBtn.addEventListener("click", () => setView("test"));
  if (navConfigBtn) {
    navConfigBtn.addEventListener("click", () => setView("config"));
  }

  backToHomeBtn.addEventListener("click", () => setView("home"));

  addClassBtn.addEventListener("click", () => {
    const name = newClassNameInput.value.trim() || generateClassName();
    const newClass = createClass(name);
    state.classes.push(newClass);
    state.selectedClassId = newClass.id;
    newClassNameInput.value = "";
    saveState();
    render();
    setView("class");
  });

  renameClassBtn.addEventListener("click", () => {
    const selectedClass = getSelectedClass();
    if (!selectedClass) {
      return;
    }
    selectedClass.name = classRenameInput.value.trim() || selectedClass.name;
    saveState();
    render();
  });

  deleteClassBtn.addEventListener("click", () => {
    const selectedClass = getSelectedClass();
    if (!selectedClass) {
      return;
    }
    if (!confirm("Eliminare questa classe?")) {
      return;
    }
    state.classes = state.classes.filter(
      (classItem) => classItem.id !== selectedClass.id
    );
    ensureClassState();
    saveState();
    render();
    setView("home");
  });

  classSelect.addEventListener("change", (event) => {
    state.selectedClassId = event.target.value;
    saveState();
    render();
  });

  testClassSelect.addEventListener("change", (event) => {
    state.selectedClassId = event.target.value;
    saveState();
    renderTestTable();
  });

  testSelect.addEventListener("change", (event) => {
    state.selectedTestId = event.target.value;
    ensureVersionSelections();
    saveState();
    renderTestTable();
  });

  testVersionSelect.addEventListener("change", (event) => {
    state.selectedTestVersionId = event.target.value;
    saveState();
    renderTestTable();
  });

  if (configTestSelect) {
    configTestSelect.addEventListener("change", (event) => {
      state.selectedTestId = event.target.value;
      ensureVersionSelections();
      saveState();
      renderConfig();
    });
  }

  if (configVersionSelect) {
    configVersionSelect.addEventListener("change", (event) => {
      state.selectedConfigVersionId = event.target.value;
      saveState();
      renderConfig();
    });
  }

  if (versionNameInput) {
    versionNameInput.addEventListener("change", (event) => {
      const selectedTest = getSelectedTest();
      if (!selectedTest) {
        return;
      }
      const version = getVersionById(selectedTest, state.selectedConfigVersionId);
      if (!version) {
        return;
      }
      version.name = event.target.value.trim() || version.name;
      saveState();
      renderConfig();
      renderTestsList();
    });
  }

  if (facilitatedVersionSelect) {
    facilitatedVersionSelect.addEventListener("change", (event) => {
      const selectedTest = getSelectedTest();
      if (!selectedTest) {
        return;
      }
      selectedTest.facilitatedVersionId = event.target.value;
      saveState();
      renderTestTable();
    });
  }

  if (addVersionBtn) {
    addVersionBtn.addEventListener("click", () => {
      const selectedTest = getSelectedTest();
      if (!selectedTest) {
        return;
      }
      const baseVersion = getVersionById(
        selectedTest,
        state.selectedConfigVersionId
      );
      const newVersion = createVersionFrom(baseVersion, selectedTest);
      selectedTest.versions.push(newVersion);
      selectedTest.facilitatedVersionId =
        selectedTest.facilitatedVersionId ?? newVersion.id;
      state.selectedConfigVersionId = newVersion.id;
      state.selectedTestVersionId = newVersion.id;
      saveState();
      renderConfig();
      renderTestTable();
    });
  }

  if (testTitleInput) {
    testTitleInput.addEventListener("input", (event) => {
      const selectedTest = getSelectedTest();
      if (!selectedTest) {
        return;
      }
      selectedTest.title = event.target.value;
      saveState();
      renderTestsList();
      renderClassDetail();
    });
  }

  // Dialog elementi
  const newTestDialog = document.getElementById("newTestDialog");
  const newTestForm = document.getElementById("newTestForm");
  const newTestNameInput = document.getElementById("newTestNameInput");
  const newTestSubjectInput = document.getElementById("newTestSubjectInput");
  const cancelNewTestBtn = document.getElementById("cancelNewTestBtn");

  addTestBtn.addEventListener("click", () => {
    newTestNameInput.value = "";
    newTestSubjectInput.value = "";
    newTestDialog.showModal();
  });

  cancelNewTestBtn.addEventListener("click", () => {
    newTestDialog.close();
  });

  // Gestione submit del form per nuova verifica (era erroneamente dentro createTest)
  if (newTestForm) {
    newTestForm.addEventListener("submit", (e) => {
      e.preventDefault();
      const name = newTestNameInput.value.trim() || generateTestName();
      const subject = newTestSubjectInput.value.trim();
      const newTestCategoryInput = document.getElementById("newTestCategoryInput");
      const category = newTestCategoryInput ? newTestCategoryInput.value.trim() : "";
      const newTest = createTest(name, subject, category);
      state.tests.push(newTest);
      state.selectedTestId = newTest.id;
      state.selectedTestVersionId = newTest.versions[0]?.id ?? null;
      state.selectedConfigVersionId = newTest.versions[0]?.id ?? null;
      saveState();
      setView("tests");
      render();
      renderTestsList();
      newTestDialog.close();
    });
  }

  if (addSectionBtn) {
    addSectionBtn.addEventListener("click", () => {
      const selectedTest = getSelectedTest();
      if (!selectedTest) {
        return;
      }
      const version = getVersionById(selectedTest, state.selectedConfigVersionId);
      if (!version) {
        return;
      }
      version.sections.push(createSection());
      saveState();
      renderConfig();
      renderTestTable();
    });
  }

  addStudentBtn.addEventListener("click", () => {
    const selectedClass = getSelectedClass();
    if (!selectedClass) {
      return;
    }
    selectedClass.students.push(createStudent());
    saveState();
    renderClassDetail();
  });

  addBulkStudentsBtn.addEventListener("click", () => {
    const selectedClass = getSelectedClass();
    if (!selectedClass) {
      return;
    }

    const entries = bulkStudentsInput.value
      .replace(/\n/g, ",")
      .split(",")
      .map((entry) => entry.trim())
      .filter(Boolean);

    entries.forEach((fullName) => {
      const student = createStudent();
      student.name = fullName;
      selectedClass.students.push(student);
    });

    bulkStudentsInput.value = "";
    saveState();
    renderClassDetail();
  });

  exportBtn.addEventListener("click", exportCSV);

  // Listener globale per Ctrl+C / Ctrl+V nella tabella voti
  document.addEventListener("keydown", (event) => {
    if (!testView.classList.contains("active")) return;
    const activeEl = document.activeElement;
    if (!activeEl || !gradeTable.contains(activeEl)) return;

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
      if (selectionState.selectedInputs.size > 0) {
        event.preventDefault();
        copySelectedCells();
      }
    } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
      if (selectionState.clipboard !== null) {
        event.preventDefault();
        pasteSelectedCells(activeEl);
      }
    }
  });

  // Listener dialog commento
  const commentDialog   = document.getElementById("commentDialog");
  const commentTextarea = document.getElementById("commentTextarea");
  const commentSaveBtn  = document.getElementById("commentSaveBtn");
  const commentDeleteBtn = document.getElementById("commentDeleteBtn");
  const commentCancelBtn = document.getElementById("commentCancelBtn");

  commentSaveBtn.addEventListener("click", () => {
    if (!commentModalContext) return;
    const { student, testId, sectionId, subsectionId, trigger } = commentModalContext;
    const key = subsectionId ?? "direct";
    ensureScoreStore(student, testId, sectionId);
    if (!student.scores[testId][sectionId].comments) student.scores[testId][sectionId].comments = {};
    const text = commentTextarea.value.trim();
    student.scores[testId][sectionId].comments[key] = text || null;
    trigger.classList.toggle("has-comment", Boolean(text));
    trigger.title = text || "Aggiungi commento";
    saveState();
    commentDialog.close();
  });

  commentDeleteBtn.addEventListener("click", () => {
    if (!commentModalContext) return;
    const { student, testId, sectionId, subsectionId, trigger } = commentModalContext;
    const key = subsectionId ?? "direct";
    if (student.scores?.[testId]?.[sectionId]?.comments) {
      student.scores[testId][sectionId].comments[key] = null;
    }
    trigger.classList.remove("has-comment");
    trigger.title = "Aggiungi commento";
    saveState();
    commentDialog.close();
  });

  commentCancelBtn.addEventListener("click", () => commentDialog.close());

  resetBtn.addEventListener("click", () => {
    if (confirm("Reset all data?")) {
      localStorage.removeItem(STORAGE_KEY);
      window.location.reload();
    }
  });

  classStudentsTable.addEventListener("click", (event) => {
    const target = event.target.closest("button[data-test-id]");
    if (!target) {
      return;
    }
    const testId = target.dataset.testId;
    if (testId) {
      state.selectedTestId = testId;
      saveState();
      renderTestTable();
      setView("test");
    }
  });

  render();
}

function render() {
  ensureClassState();
  ensureTestState();
  renderClassesList();
  renderClassDetail();
  renderTestsList();
  if (configView) {
    renderConfig();
  }
  renderTestTable();
  updateView();
}

function setView(view) {
  state.view = view;
  saveState();
  updateView();
}

function updateView() {
  const views = [homeView, classView, testsView, testView, configView].filter(Boolean);
  views.forEach((view) => view.classList.remove("active"));

  if (state.view === "config" && !configView) {
    state.view = "home";
  }

  switch (state.view) {
    case "class":
      classView.classList.add("active");
      break;
    case "tests":
      testsView.classList.add("active");
      break;
    case "test":
      testView.classList.add("active");
      break;
    case "config":
      if (configView) {
        configView.classList.add("active");
      } else {
        homeView.classList.add("active");
      }
      break;
    default:
      homeView.classList.add("active");
      break;
  }
}

function renderClassesList() {
  classList.innerHTML = "";
  state.classes.forEach((classItem) => {
    const card = document.createElement("div");
    card.classList.add("card");

    const title = document.createElement("h3");
    title.textContent = classItem.name || "Class";
    card.appendChild(title);

    const info = document.createElement("small");
    info.textContent = `${classItem.students.length} studenti`;
    card.appendChild(info);

    const openBtn = document.createElement("button");
    openBtn.classList.add("btn", "btn-secondary");
    openBtn.textContent = "Apri";
    openBtn.addEventListener("click", () => {
      state.selectedClassId = classItem.id;
      saveState();
      renderClassDetail();
      setView("class");
    });
    card.appendChild(openBtn);

    const deleteBtn = document.createElement("button");
    deleteBtn.classList.add("icon-btn", "card-delete");
    deleteBtn.textContent = "×";
    deleteBtn.setAttribute("aria-label", "Elimina classe");
    deleteBtn.addEventListener("click", () => {
      if (!confirm("Eliminare questa classe?")) {
        return;
      }
      state.classes = state.classes.filter(
        (item) => item.id !== classItem.id
      );
      ensureClassState();
      saveState();
      render();
    });
    card.appendChild(deleteBtn);

    classList.appendChild(card);
  });
}

function renderClassDetail() {
  const selectedClass = getSelectedClass();
  classSelect.innerHTML = "";
  state.classes.forEach((classItem) => {
    const option = document.createElement("option");
    option.value = classItem.id;
    option.textContent = classItem.name || "Class";
    if (classItem.id === state.selectedClassId) {
      option.selected = true;
    }
    classSelect.appendChild(option);
  });

  classRenameInput.value = selectedClass?.name || "";

  classStudentsTable.innerHTML = "";
  if (!selectedClass) {
    return;
  }

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  const studentHeader = document.createElement("th");
  studentHeader.textContent = "Studente";
  headerRow.appendChild(studentHeader);

  // Colonna DSA/104
  const dsaHeader = document.createElement("th");
  dsaHeader.textContent = "DSA/104";
  dsaHeader.title = "Studente con DSA o Legge 104 – riceverà automaticamente la versione facilitata";
  dsaHeader.style.cssText = "width:80px;text-align:center;font-size:.85em;";
  headerRow.appendChild(dsaHeader);

  // Filtra i test: mostra solo quelli che hanno almeno un voto nella classe
  const visibleTests = state.tests.filter((test) => testHasGradesInClass(test, selectedClass));
  
  visibleTests.forEach((test) => {
    const th = document.createElement("th");
    let label = test.title || "Verifica";
    // Aggiungi data classe se disponibile
    const dateForClass = test.classDates?.[selectedClass.id];
    if (dateForClass) {
      label += " · " + dateForClass;
    }
    th.innerHTML = `<div>${label}</div>`;
    if (test.subject && test.subject.trim() !== "") {
      const subjectDiv = document.createElement("div");
      subjectDiv.style.cssText = "font-size:11px;color:#666;";
      subjectDiv.textContent = test.subject;
      th.appendChild(subjectDiv);
    }
    if ((test.categories || []).length > 0) {
      const catDiv = document.createElement("div");
      catDiv.style.cssText = "font-size:10px;color:#9f7aea;font-style:italic;";
      catDiv.textContent = test.categories.join(", ");
      th.appendChild(catDiv);
    }
    headerRow.appendChild(th);
  });

  const avgHeader = document.createElement("th");
  avgHeader.textContent = "Media";
  headerRow.appendChild(avgHeader);

  const actionsHeader = document.createElement("th");
  actionsHeader.textContent = "Azioni";
  headerRow.appendChild(actionsHeader);

  thead.appendChild(headerRow);
  classStudentsTable.appendChild(thead);

  const tbody = document.createElement("tbody");

  selectedClass.students.forEach((student) => {
    const isFacilitated = student.facilitated === true;
    const row = document.createElement("tr");
    if (isFacilitated) {
      row.classList.add("facilitated-list-row");
    }

    // Nome studente
    const studentCell = document.createElement("td");
    studentCell.classList.add("student-cell");
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.value = student.name || "";
    nameInput.addEventListener("change", (event) => {
      student.name = event.target.value.trim() || "Student";
      saveState();
      renderClassDetail();
    });
    studentCell.appendChild(nameInput);
    // Badge DSA inline accanto al nome
    if (isFacilitated) {
      const badge = document.createElement("span");
      badge.className = "dsa-badge";
      badge.textContent = "DSA/104";
      studentCell.appendChild(badge);
    }
    row.appendChild(studentCell);

    // Toggle DSA/104
    const dsaCell = document.createElement("td");
    dsaCell.style.textAlign = "center";
    const dsaToggle = document.createElement("input");
    dsaToggle.type = "checkbox";
    dsaToggle.checked = isFacilitated;
    dsaToggle.title = "Segna come DSA / Legge 104";
    dsaToggle.addEventListener("change", (event) => {
      student.facilitated = event.target.checked;
      saveState();
      renderClassDetail();
    });
    dsaCell.appendChild(dsaToggle);
    row.appendChild(dsaCell);

    visibleTests.forEach((test) => {
      const score = getFinalScore(student, test);
      const cell = document.createElement("td");
      const btn = document.createElement("button");
      btn.classList.add("grade-link");
      btn.dataset.testId = test.id;
      btn.textContent = formatScore(score);
      cell.appendChild(btn);
      row.appendChild(cell);
    });

    const avgCell = document.createElement("td");
    const avgScore = getStudentAverage(student);
    avgCell.textContent = formatScore(avgScore);
    if (isLowGrade(avgScore)) {
      avgCell.classList.add("low-grade");
    }
    row.appendChild(avgCell);

    const actionsCell = document.createElement("td");
    const deleteBtn = document.createElement("button");
    deleteBtn.classList.add("btn", "btn-danger", "btn-small");
    deleteBtn.textContent = "Elimina";
    deleteBtn.addEventListener("click", () => {
      selectedClass.students = selectedClass.students.filter(
        (item) => item.id !== student.id
      );
      saveState();
      renderClassDetail();
    });
    actionsCell.appendChild(deleteBtn);
    row.appendChild(actionsCell);

    tbody.appendChild(row);
  });

  classStudentsTable.appendChild(tbody);
}

function renderTestsList() {
  testsList.innerHTML = "";
  state.tests.forEach(ensureTestMeta);

  // ── Valori unici per i filtri ────────────────────────────────────────────
  const allSubjects    = [...new Set(state.tests.map(t => t.subject).filter(Boolean))].sort();
  const allCategories  = [...new Set(state.tests.flatMap(t => t.categories || []).filter(Boolean))].sort();

  const filterClass    = document.getElementById("filterClass")?.value    || "";
  const filterSubject  = document.getElementById("filterSubject")?.value  || "";
  const filterCategory = document.getElementById("filterCategory")?.value || "";

  // ── Barra filtri ─────────────────────────────────────────────────────────
  const existingBar = document.getElementById("testsFilterBar");
  if (existingBar) existingBar.remove();

  const bar = document.createElement("div");
  bar.id = "testsFilterBar";
  bar.className = "filter-bar";
  bar.innerHTML = `
    <span class="filter-bar-label">🔍 Filtra:</span>
    <select id="filterClass">
      <option value="">Tutte le classi</option>
      ${state.classes.map(c => `<option value="${c.id}" ${filterClass===c.id?"selected":""}>${c.name}</option>`).join("")}
    </select>
    <select id="filterSubject">
      <option value="">Tutte le materie</option>
      ${allSubjects.map(s => `<option value="${s}" ${filterSubject===s?"selected":""}>${s}</option>`).join("")}
    </select>
    <select id="filterCategory">
      <option value="">Tutte le categorie</option>
      ${allCategories.map(c => `<option value="${c}" ${filterCategory===c?"selected":""}>${c}</option>`).join("")}
    </select>
    <button class="btn btn-secondary btn-small" id="clearFiltersBtn">✕ Reset</button>
  `;
  testsList.parentElement.insertBefore(bar, testsList);
  bar.querySelectorAll("select").forEach(sel => sel.addEventListener("change", renderTestsList));
  bar.querySelector("#clearFiltersBtn").addEventListener("click", () => {
    bar.querySelectorAll("select").forEach(s => s.value = "");
    renderTestsList();
  });

  // ── Filtro ────────────────────────────────────────────────────────────────
  const filtered = state.tests.filter(test => {
    if (filterSubject  && test.subject !== filterSubject)  return false;
    if (filterCategory && !(test.categories || []).includes(filterCategory)) return false;
    if (filterClass    && !test.classIds.includes(filterClass)) return false;
    return true;
  });

  if (!filtered.length) {
    testsList.innerHTML = `<p style="color:#888;padding:12px;">Nessuna verifica corrisponde ai filtri.</p>`;
    return;
  }

  // ── Render card ───────────────────────────────────────────────────────────
  filtered.forEach((test) => {
    ensureTestVersions(test);
    const card = document.createElement("div");
    card.classList.add("card");

    // Titolo (read-only, modificabile nel pannello)
    const titleEl = document.createElement("h3");
    titleEl.textContent = test.title || "Verifica";
    titleEl.style.cssText = "margin-bottom:6px;";
    card.appendChild(titleEl);

    // Materia + Categorie
    const metaRow = document.createElement("div");
    metaRow.style.cssText = "display:flex;flex-wrap:wrap;gap:6px;margin-bottom:6px;align-items:center;";
    if (test.subject) {
      const sub = document.createElement("span");
      sub.className = "test-meta-chip test-subject-chip";
      sub.textContent = "📚 " + test.subject;
      metaRow.appendChild(sub);
    }
    (test.categories || []).forEach(cat => {
      const chip = document.createElement("span");
      chip.className = "test-meta-chip test-category-chip";
      chip.textContent = "🏷️ " + cat;
      metaRow.appendChild(chip);
    });
    if (metaRow.children.length) card.appendChild(metaRow);

    // Classi + date
    if (test.classIds.length > 0) {
      const classRow = document.createElement("div");
      classRow.style.cssText = "display:flex;flex-wrap:wrap;gap:5px;margin-bottom:8px;";
      test.classIds.forEach(cid => {
        const cls = state.classes.find(c => c.id === cid);
        if (!cls) return;
        const date = test.classDates[cid] || "";
        const chip = document.createElement("span");
        chip.className = "class-date-chip";
        chip.textContent = cls.name + (date ? " · " + date : "");
        classRow.appendChild(chip);
      });
      if (classRow.children.length) card.appendChild(classRow);
    }

    const info = document.createElement("small");
    info.textContent = `${getDefaultVersion(test)?.sections.length ?? 0} sezioni`;
    info.style.color = "#888";
    card.appendChild(info);

    // ── Pannello "Modifica dettagli" ──────────────────────────────────────
    const detailsToggle = document.createElement("button");
    detailsToggle.className = "btn btn-secondary btn-small";
    detailsToggle.style.cssText = "margin-top:8px;font-size:.8em;";
    detailsToggle.textContent = "✏️ Modifica dettagli";

    const detailsPanel = document.createElement("div");
    detailsPanel.className = "test-details-panel";
    detailsPanel.style.display = "none";

    // Costruiamo il pannello in JS (no innerHTML con dati user per sicurezza)
    const makeField = (labelText, child) => {
      const wrap = document.createElement("div");
      wrap.className = "field";
      wrap.style.marginTop = "10px";
      const lbl = document.createElement("label");
      lbl.textContent = labelText;
      wrap.appendChild(lbl);
      wrap.appendChild(child);
      return wrap;
    };

    // — Nome verifica —
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.value = test.title || "";
    nameInput.placeholder = "Nome della verifica";
    detailsPanel.appendChild(makeField("Nome verifica", nameInput));

    // — Materia —
    const subjectInput = document.createElement("input");
    subjectInput.type = "text";
    subjectInput.value = test.subject || "";
    subjectInput.placeholder = "es. English, Storia…";
    detailsPanel.appendChild(makeField("Materia", subjectInput));

    // — Categorie (tag) —
    const catField = document.createElement("div");
    catField.className = "field";
    catField.style.marginTop = "10px";
    const catLabel = document.createElement("label");
    catLabel.textContent = "Categorie";
    catField.appendChild(catLabel);

    const tagBox = document.createElement("div");
    tagBox.className = "tag-box";

    const renderTags = () => {
      tagBox.innerHTML = "";
      (test.categories || []).forEach((cat, idx) => {
        const chip = document.createElement("span");
        chip.className = "tag-chip";
        chip.innerHTML = `${cat} <button class="tag-remove" data-idx="${idx}" title="Rimuovi">×</button>`;
        tagBox.appendChild(chip);
      });
      // Input per aggiungere tag
      const addWrap = document.createElement("div");
      addWrap.style.cssText = "display:flex;gap:6px;margin-top:6px;";
      const tagInput = document.createElement("input");
      tagInput.type = "text";
      tagInput.placeholder = "Aggiungi categoria…";
      tagInput.style.cssText = "flex:1;padding:5px 9px;border-radius:7px;border:1.5px solid #e2b4c8;font-size:.88em;";
      const addBtn = document.createElement("button");
      addBtn.type = "button";
      addBtn.className = "btn btn-secondary btn-small";
      addBtn.textContent = "+ Aggiungi";
      addBtn.style.fontSize = ".82em";
      const doAdd = () => {
        const val = tagInput.value.trim();
        if (val && !test.categories.includes(val)) {
          test.categories.push(val);
          renderTags();
        }
        tagInput.value = "";
      };
      addBtn.addEventListener("click", doAdd);
      tagInput.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); doAdd(); } });
      addWrap.appendChild(tagInput);
      addWrap.appendChild(addBtn);
      tagBox.appendChild(addWrap);

      // Rimozione tag via ×
      tagBox.querySelectorAll(".tag-remove").forEach(btn => {
        btn.addEventListener("click", () => {
          test.categories.splice(parseInt(btn.dataset.idx), 1);
          renderTags();
        });
      });
    };
    renderTags();
    catField.appendChild(tagBox);
    detailsPanel.appendChild(catField);

    // — Classi + date —
    const clsField = document.createElement("div");
    clsField.className = "field";
    clsField.style.marginTop = "10px";
    const clsLabel = document.createElement("label");
    clsLabel.textContent = "Classi che svolgono questa verifica";
    clsField.appendChild(clsLabel);
    const clsList = document.createElement("div");
    clsList.className = "class-check-list";
    state.classes.forEach(c => {
      const row = document.createElement("label");
      row.className = "class-check-row";
      const chk = document.createElement("input");
      chk.type = "checkbox";
      chk.className = "cls-check";
      chk.dataset.classId = c.id;
      chk.checked = test.classIds.includes(c.id);
      const nameSp = document.createElement("span");
      nameSp.textContent = c.name;
      const dateIn = document.createElement("input");
      dateIn.type = "date";
      dateIn.className = "cls-date";
      dateIn.dataset.classId = c.id;
      dateIn.value = test.classDates[c.id] || "";
      dateIn.disabled = !chk.checked;
      dateIn.style.cssText = "margin-left:8px;font-size:.85em;border:1px solid #ddd;border-radius:6px;padding:2px 6px;";
      chk.addEventListener("change", () => { dateIn.disabled = !chk.checked; });
      row.appendChild(chk);
      row.appendChild(nameSp);
      row.appendChild(dateIn);
      clsList.appendChild(row);
    });
    clsField.appendChild(clsList);
    detailsPanel.appendChild(clsField);

    // — Salva —
    const saveBtn = document.createElement("button");
    saveBtn.className = "btn btn-primary btn-small";
    saveBtn.style.marginTop = "10px";
    saveBtn.textContent = "💾 Salva dettagli";
    saveBtn.addEventListener("click", () => {
      // Nome + Materia
      test.title   = nameInput.value.trim()   || test.title;
      test.subject = subjectInput.value.trim();
      // Classi + date (categories già aggiornate live)
      test.classIds = [];
      clsList.querySelectorAll(".cls-check").forEach(chk => {
        if (chk.checked) {
          const cid = chk.dataset.classId;
          test.classIds.push(cid);
          const di = clsList.querySelector(`.cls-date[data-class-id="${cid}"]`);
          if (di?.value) test.classDates[cid] = di.value;
          else delete test.classDates[cid];
        }
      });
      Object.keys(test.classDates).forEach(cid => {
        if (!test.classIds.includes(cid)) delete test.classDates[cid];
      });
      saveState();
      renderTestsList();
    });
    detailsPanel.appendChild(saveBtn);

    detailsToggle.addEventListener("click", () => {
      const open = detailsPanel.style.display === "block";
      detailsPanel.style.display = open ? "none" : "block";
      detailsToggle.textContent = open ? "✏️ Modifica dettagli" : "▲ Chiudi";
    });

    card.appendChild(detailsToggle);
    card.appendChild(detailsPanel);

    // ── Azioni principali ─────────────────────────────────────────────────
    const actions = document.createElement("div");
    actions.classList.add("panel-actions");
    actions.style.marginTop = "10px";

    const evalBtn = document.createElement("button");
    evalBtn.classList.add("btn", "btn-secondary");
    evalBtn.textContent = "Valuta";
    evalBtn.addEventListener("click", () => {
      state.selectedTestId = test.id;
      saveState();
      renderTestTable();
      setView("test");
    });
    actions.appendChild(evalBtn);

    if (configView) {
      const configBtn = document.createElement("button");
      configBtn.classList.add("btn", "btn-secondary");
      configBtn.textContent = "Configura";
      configBtn.addEventListener("click", () => {
        state.selectedTestId = test.id;
        saveState();
        renderConfig();
        setView("config");
      });
      actions.appendChild(configBtn);
    }

    const deleteBtn = document.createElement("button");
    deleteBtn.classList.add("icon-btn", "card-delete");
    deleteBtn.textContent = "×";
    deleteBtn.setAttribute("aria-label", "Elimina verifica");
    deleteBtn.addEventListener("click", () => {
      if (!confirm(`Eliminare la verifica "${test.title}"? Tutti i voti andranno persi.`)) return;
      state.tests = state.tests.filter(t => t.id !== test.id);
      saveState();
      renderTestsList();
      renderTestTable();
    });
    card.appendChild(deleteBtn);

    card.appendChild(actions);
    testsList.appendChild(card);
  });
}

function renderConfig() {
  if (!configView || !configTestSelect || !configVersionSelect || !sectionsContainer) {
    return;
  }
  const selectedTest = getSelectedTest();
  configTestSelect.innerHTML = "";
  state.tests.forEach((test) => {
    const option = document.createElement("option");
    option.value = test.id;
    option.textContent = test.title || "Verifica";
    if (test.id === state.selectedTestId) {
      option.selected = true;
    }
    configTestSelect.appendChild(option);
  });

  testTitleInput.value = selectedTest?.title || "";

  configVersionSelect.innerHTML = "";
  facilitatedVersionSelect.innerHTML = "";
  versionNameInput.value = "";

  if (!selectedTest) {
    sectionsContainer.innerHTML = "";
    return;
  }

  ensureTestVersions(selectedTest);
  ensureVersionSelections();

  selectedTest.versions.forEach((version) => {
    const option = document.createElement("option");
    option.value = version.id;
    option.textContent = version.name || "Versione";
    if (version.id === state.selectedConfigVersionId) {
      option.selected = true;
    }
    configVersionSelect.appendChild(option);

    const facilitatedOption = document.createElement("option");
    facilitatedOption.value = version.id;
    facilitatedOption.textContent = version.name || "Versione";
    if (version.id === selectedTest.facilitatedVersionId) {
      facilitatedOption.selected = true;
    }
    facilitatedVersionSelect.appendChild(facilitatedOption);
  });

  const activeVersion = getVersionById(
    selectedTest,
    state.selectedConfigVersionId
  );
  versionNameInput.value = activeVersion?.name || "";
  renderSections(activeVersion);
}

function renderSections(version) {
  sectionsContainer.innerHTML = "";

  if (!version) {
    return;
  }

  let draggedSection = null;

  version.sections.forEach((section) => {
    const card = sectionTemplate.content.firstElementChild.cloneNode(true);
    const nameInput = card.querySelector(".section-name");
    const subsectionsContainer = card.querySelector(".subsections");

    nameInput.value = section.name;

    nameInput.addEventListener("change", (event) => {
      section.name = event.target.value;
      saveState();
      renderConfig();
      renderTestTable();
    });

    // Drag and drop events
    card.addEventListener("dragstart", (event) => {
      draggedSection = section;
      card.classList.add("dragging");
      event.dataTransfer.effectAllowed = "move";
    });

    card.addEventListener("dragend", () => {
      card.classList.remove("dragging");
      // Remove drag-over styling from all cards
      document.querySelectorAll(".section-card").forEach((c) => {
        c.classList.remove("drag-over");
      });
      draggedSection = null;
    });

    card.addEventListener("dragover", (event) => {
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
      if (draggedSection && draggedSection.id !== section.id) {
        card.classList.add("drag-over");
      }
    });

    card.addEventListener("dragleave", () => {
      card.classList.remove("drag-over");
    });

    card.addEventListener("drop", (event) => {
      event.preventDefault();
      card.classList.remove("drag-over");
      if (draggedSection && draggedSection.id !== section.id) {
        // Find indices
        const draggedIndex = version.sections.findIndex((s) => s.id === draggedSection.id);
        const targetIndex = version.sections.findIndex((s) => s.id === section.id);
        
        if (draggedIndex !== -1 && targetIndex !== -1) {
          // Swap sections
          const temp = version.sections[draggedIndex];
          version.sections[draggedIndex] = version.sections[targetIndex];
          version.sections[targetIndex] = temp;
          
          saveState();
          renderConfig();
          renderTestTable();
        }
      }
    });

    card.querySelector(".remove-section").addEventListener("click", () => {
      version.sections = version.sections.filter((item) => item.id !== section.id);
      removeSectionScores(section.id, getSelectedTest()?.id);
      saveState();
      renderConfig();
      renderTestTable();
    });

    card.querySelector(".add-subsection").addEventListener("click", () => {
      const lastSubsection = section.subsections[section.subsections.length - 1];
      section.subsections.push(createSubsection(lastSubsection));
      saveState();
      renderConfig();
      renderTestTable();
    });

    section.subsections.forEach((subsection) => {
      const subRow = subsectionTemplate.content.firstElementChild.cloneNode(true);
      const subNameInput = subRow.querySelector(".subsection-name");
      const subWeightInput = subRow.querySelector(".subsection-weight");
      const subMaxInput = subRow.querySelector(".subsection-max");
      subNameInput.value = subsection.name;
      subWeightInput.value = subsection.weight ?? "";
      subMaxInput.value = subsection.max ?? "";

      subNameInput.addEventListener("change", (event) => {
        subsection.name = event.target.value;
        saveState();
        renderTestTable();
      });

      subWeightInput.addEventListener("change", (event) => {
        subsection.weight = parseNumber(event.target.value);
        saveState();
        renderConfig();
        renderTestTable();
      });

      subMaxInput.addEventListener("change", (event) => {
        subsection.max = parseNumber(event.target.value);
        saveState();
        renderConfig();
        renderTestTable();
      });

      subRow.querySelector(".remove-subsection").addEventListener("click", () => {
        section.subsections = section.subsections.filter(
          (item) => item.id !== subsection.id
        );
        removeSubsectionScores(section.id, subsection.id, getSelectedTest()?.id);
        saveState();
        renderConfig();
        renderTestTable();
      });

      subsectionsContainer.appendChild(subRow);
    });

    sectionsContainer.appendChild(card);
  });
}

function renderTestTable() {
  // Salva la posizione del focus prima di ricostruire il DOM
  const activeEl = document.activeElement;
  let focusRowIdx = -1, focusCellIdx = -1, focusCursorPos = -1, focusRawValue = null;
  if (activeEl && gradeTable.contains(activeEl) && activeEl.tagName === "INPUT") {
    const row = activeEl.closest("tr");
    const cell = activeEl.closest("td");
    if (row && cell && row.parentElement.tagName === "TBODY") {
      focusRowIdx = Array.from(row.parentElement.children).indexOf(row);
      focusCellIdx = Array.from(row.children).indexOf(cell);
      focusCursorPos = activeEl.selectionStart ?? -1;
      focusRawValue = activeEl.value; // preserva il testo grezzo es. "7." o "1,5"
    }
  }

  gradeTable.innerHTML = "";
  warningArea.innerHTML = "";

  const selectedClass = getSelectedClass();
  const selectedTest = getSelectedTest();
  const students = selectedClass?.students ?? [];

  testClassSelect.innerHTML = "";
  state.classes.forEach((classItem) => {
    const option = document.createElement("option");
    option.value = classItem.id;
    option.textContent = classItem.name || "Class";
    if (classItem.id === state.selectedClassId) {
      option.selected = true;
    }
    testClassSelect.appendChild(option);
  });

  testSelect.innerHTML = "";
  state.tests.forEach((test) => {
    const option = document.createElement("option");
    option.value = test.id;
    option.textContent = test.title || "Verifica";
    if (test.id === state.selectedTestId) {
      option.selected = true;
    }
    testSelect.appendChild(option);
  });

  if (!selectedTest) {
    return;
  }

  ensureTestVersions(selectedTest);
  saveState(); // Salva se versioni sono state aggiunte
  ensureVersionSelections();

  testVersionSelect.innerHTML = "";
  selectedTest.versions.forEach((version) => {
    const option = document.createElement("option");
    option.value = version.id;
    option.textContent = version.name || "Versione";
    if (version.id === state.selectedTestVersionId) {
      option.selected = true;
    }
    testVersionSelect.appendChild(option);
  });

  const defaultVersion = getDefaultVersion(selectedTest);
  const activeVersion =
    getVersionById(selectedTest, state.selectedTestVersionId) ??
    defaultVersion;
  const facilitatedVersionId = getFacilitatedVersionId(selectedTest);

  // Applica la classe CSS per la versione facilitata
  gradeTable.classList.remove("facilitata-version");
  if (state.selectedTestVersionId === facilitatedVersionId) {
    gradeTable.classList.add("facilitata-version");
  }

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const subHeaderRow = document.createElement("tr");
  const weightRow = document.createElement("tr");
  const maxRow = document.createElement("tr");

  const studentHeader = document.createElement("th");
  studentHeader.textContent = "Studente";
  studentHeader.rowSpan = 4;
  headerRow.appendChild(studentHeader);

  const versionHeader = document.createElement("th");
  versionHeader.innerHTML = "DSA<br><small style='font-weight:400;font-size:10px'>/ 104</small>";
  versionHeader.title = "Spunta per assegnare automaticamente la versione facilitata a questo studente in tutte le verifiche";
  versionHeader.rowSpan = 4;
  headerRow.appendChild(versionHeader);

  const sections = activeVersion?.sections ?? [];

  sections.forEach((section) => {
    if (!Array.isArray(section.subsections)) {
      section.subsections = [];
    }
    if (section.subsections.length === 0) {
      section.subsections.push(
        createSubsection({ weight: section.weight, max: section.max })
      );
      saveState();
    }

    const th = document.createElement("th");
    th.colSpan = section.subsections.length;
    const headerWrap = document.createElement("div");
    headerWrap.classList.add("section-header-cell");

    const sectionNameInput = document.createElement("input");
    sectionNameInput.type = "text";
    sectionNameInput.value = section.name || "Section";
    sectionNameInput.addEventListener("change", (event) => {
      section.name = event.target.value;
      saveState();
      renderTestTable();
    });
    headerWrap.appendChild(sectionNameInput);

    const addColumnBtn = document.createElement("button");
    addColumnBtn.type = "button";
    addColumnBtn.classList.add("btn", "btn-secondary", "btn-small");
    addColumnBtn.textContent = "+";
    addColumnBtn.addEventListener("click", () => {
      const lastSubsection = section.subsections[section.subsections.length - 1];
      section.subsections.push(createSubsection(lastSubsection));
      saveState();
      renderTestTable();
    });
    headerWrap.appendChild(addColumnBtn);

    const removeColumnBtn = document.createElement("button");
    removeColumnBtn.type = "button";
    removeColumnBtn.classList.add("btn", "btn-danger", "btn-small");
    removeColumnBtn.textContent = "×";
    removeColumnBtn.title = "Elimina questa section e tutti i voti associati";
    removeColumnBtn.addEventListener("click", () => {
      if (confirm(`Sei sicuro di voler eliminare la section "${section.name}"? Tutti i voti andranno persi.`)) {
        activeVersion.sections = activeVersion.sections.filter((item) => item.id !== section.id);
        removeSectionScores(section.id, selectedTest?.id);
        saveState();
        renderTestTable();
      }
    });
    headerWrap.appendChild(removeColumnBtn);

    th.appendChild(headerWrap);
    headerRow.appendChild(th);

    section.subsections.forEach((subsection, subsectionIndex) => {
      const subTh = document.createElement("th");
      subTh.classList.add("subheader");
      const isLastSubsection = subsectionIndex === section.subsections.length - 1;
      if (isLastSubsection) {
        subTh.classList.add("section-divider");
      }
      const subHeaderWrap = document.createElement("div");
      subHeaderWrap.classList.add("subheader-cell");
      const subNameInput = document.createElement("input");
      subNameInput.type = "text";
      subNameInput.value = subsection.name || "Subsection";
      subNameInput.addEventListener("change", (event) => {
        subsection.name = event.target.value;
        saveState();
        renderTestTable();
      });
      subHeaderWrap.appendChild(subNameInput);

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.classList.add("subsection-remove");
      removeBtn.textContent = "×";
      removeBtn.addEventListener("click", () => {
        section.subsections = section.subsections.filter(
          (item) => item.id !== subsection.id
        );
        removeSubsectionScores(section.id, subsection.id, getSelectedTest()?.id);
        saveState();
        renderTestTable();
      });
      subHeaderWrap.appendChild(removeBtn);

      subTh.appendChild(subHeaderWrap);
      subHeaderRow.appendChild(subTh);

      const weightTh = document.createElement("th");
      weightTh.classList.add("subheader");
      if (isLastSubsection) {
        weightTh.classList.add("section-divider");
      }
      const weightInput = document.createElement("input");
      weightInput.type = "number";
      weightInput.step = "0.1";
      weightInput.min = "0";
      weightInput.value = getSubsectionWeight(subsection);
      weightInput.addEventListener("change", (event) => {
        subsection.weight = parseNumber(event.target.value);
        saveState();
        renderTestTable();
      });
      weightTh.appendChild(weightInput);
      weightRow.appendChild(weightTh);

      const maxTh = document.createElement("th");
      maxTh.classList.add("subheader");
      if (isLastSubsection) {
        maxTh.classList.add("section-divider");
      }
      const maxInput = document.createElement("input");
      maxInput.type = "number";
      maxInput.step = "0.1";
      maxInput.min = "0";
      maxInput.value = getSubsectionMax(
        section,
        subsection,
        getSectionTotals(section).fallbackPerSubMax
      );
      maxInput.addEventListener("change", (event) => {
        subsection.max = parseNumber(event.target.value);
        saveState();
        renderTestTable();
      });
      maxTh.appendChild(maxInput);
      maxRow.appendChild(maxTh);
    });
  });

  const finalHeader = document.createElement("th");
  const finalHeaderWrap = document.createElement("div");
  finalHeaderWrap.classList.add("section-header-cell");
  const addSectionBtn = document.createElement("button");
  addSectionBtn.type = "button";
  addSectionBtn.classList.add("btn", "btn-add-section");
  addSectionBtn.textContent = "+ Section";
  addSectionBtn.addEventListener("click", () => {
    if (!activeVersion) {
      return;
    }
    activeVersion.sections.push(createSection());
    saveState();
    renderTestTable();
  });
  const finalLabel = document.createElement("span");
  finalLabel.textContent = "FINAL";
  finalHeaderWrap.appendChild(addSectionBtn);
  finalHeaderWrap.appendChild(finalLabel);
  finalHeader.appendChild(finalHeaderWrap);
  finalHeader.rowSpan = 4;
  headerRow.appendChild(finalHeader);

  thead.appendChild(headerRow);
  thead.appendChild(subHeaderRow);
  thead.appendChild(weightRow);
  thead.appendChild(maxRow);
  gradeTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  const warnings = [];

  students.forEach((student) => {
    const row = document.createElement("tr");

    // Flag globale: se lo studente è DSA/104 è facilitato su TUTTE le verifiche
    const isFacilitated = student.facilitated === true;

    // Se facilitato, forza automaticamente la versione facilitata per questa verifica
    const effectiveVersionId = isFacilitated && facilitatedVersionId
      ? facilitatedVersionId
      : getStudentVersionId(student, selectedTest.id, defaultVersion?.id);

    const isActiveVersion = effectiveVersionId === activeVersion?.id;

    if (isFacilitated) {
      row.classList.add("facilitated-row");
    }
    if (!isActiveVersion) {
      row.classList.add("version-mismatch");
    }

    const studentCell = document.createElement("td");
    studentCell.classList.add("student-cell");
    const studentInput = document.createElement("input");
    studentInput.type = "text";
    studentInput.value = student.name || "";
    studentInput.addEventListener("change", (event) => {
      student.name = event.target.value;
      saveState();
    });
    studentCell.appendChild(studentInput);
    
    // Aggiungi click handler per cambiare versione
    studentCell.style.cursor = "pointer";
    studentCell.addEventListener("click", (event) => {
      if (event.target !== studentInput) {
        if (isFacilitated) {
          state.selectedTestVersionId = facilitatedVersionId;
        } else {
          state.selectedTestVersionId = activeVersion?.id;
        }
        saveState();
        renderTestTable();
      }
    });
    
    row.appendChild(studentCell);

    const versionCell = document.createElement("td");
    versionCell.style.textAlign = "center";
    const versionToggle = document.createElement("input");
    versionToggle.type = "checkbox";
    versionToggle.checked = isFacilitated;
    versionToggle.title = "DSA / 104 – versione facilitata per tutte le verifiche";
    versionToggle.addEventListener("change", (event) => {
      // Imposta il flag globale sullo studente (vale per TUTTE le verifiche)
      student.facilitated = event.target.checked;
      saveState();
      renderTestTable();
    });
    versionCell.appendChild(versionToggle);
    
    // Aggiungi click handler per cambiare versione (ignora click sulla checkbox)
    versionCell.style.cursor = "pointer";
    versionCell.addEventListener("click", (event) => {
      // Ignora click diretti sulla checkbox stessa
      if (event.target === versionToggle) {
        return;
      }
      if (isFacilitated) {
        state.selectedTestVersionId = facilitatedVersionId;
      } else {
        state.selectedTestVersionId = activeVersion?.id;
      }
      saveState();
      renderTestTable();
    });
    
    row.appendChild(versionCell);

    sections.forEach((section, sectionIndex) => {
      const sectionScore = getSectionScore(student, selectedTest, section);
      const hasSubsections = section.subsections.length > 0;

      if (hasSubsections) {
        section.subsections.forEach((subsection, subsectionIndex) => {
          const cell = document.createElement("td");
          const isLastSubsection = subsectionIndex === section.subsections.length - 1;
          if (isLastSubsection) {
            cell.classList.add("section-divider");
          }
          // Logica di disabilitazione: Standard disabilita chi ha spunta, Facilitata disabilita chi NON l'ha
          const shouldDisable = (state.selectedTestVersionId === facilitatedVersionId) ? 
            (student.facilitated !== true) : 
            (student.facilitated === true);
          const input = createScoreInput(
            student,
            selectedTest.id,
            section.id,
            subsection.id,
            "subsection",
            shouldDisable
          );
          cell.appendChild(input);
          attachCommentTrigger(cell, student, selectedTest.id, section.id, subsection.id);
          
          // Aggiungi click handler per cambiare versione
          cell.style.cursor = "pointer";
          cell.addEventListener("click", (event) => {
            if (event.target !== input && !event.target.classList.contains("comment-trigger")) {
              if (isFacilitated) {
                state.selectedTestVersionId = facilitatedVersionId;
              } else {
                state.selectedTestVersionId = activeVersion?.id;
              }
              saveState();
              renderTestTable();
            }
          });
          
          row.appendChild(cell);
        });
      } else {
        const cell = document.createElement("td");
        // Logica di disabilitazione: Standard disabilita chi ha spunta, Facilitata disabilita chi NON l'ha
        const shouldDisable = (state.selectedTestVersionId === facilitatedVersionId) ? 
          (student.facilitated !== true) : 
          (student.facilitated === true);
        const input = createScoreInput(
          student,
          selectedTest.id,
          section.id,
          null,
          "section",
          shouldDisable
        );
        cell.appendChild(input);
        attachCommentTrigger(cell, student, selectedTest.id, section.id, null);
        
        // Aggiungi click handler per cambiare versione
        cell.style.cursor = "pointer";
        cell.addEventListener("click", (event) => {
          if (event.target !== input && !event.target.classList.contains("comment-trigger")) {
            if (isFacilitated) {
              state.selectedTestVersionId = facilitatedVersionId;
            } else {
              state.selectedTestVersionId = activeVersion?.id;
            }
            saveState();
            renderTestTable();
          }
        });
        
        row.appendChild(cell);
      }

      if (getSectionMax(section) != null && sectionScore > getSectionMax(section)) {
        warnings.push(
          `${student.name || "Student"}: ${section.name || "Section"} exceeds max (${sectionScore} > ${getSectionMax(section)})`
        );
      }
    });

    const finalCell = document.createElement("td");
    finalCell.classList.add("final-cell");
    const finalScore = getFinalScore(student, selectedTest, activeVersion);
    finalCell.textContent = formatScore(finalScore);
    if (isLowGrade(finalScore)) {
      finalCell.classList.add("low-grade");
    }
    row.appendChild(finalCell);

    tbody.appendChild(row);
  });

  gradeTable.appendChild(tbody);

  // Ripristina il focus sulla stessa cella dopo il re-render
  if (focusRowIdx >= 0 && focusCellIdx >= 0) {
    const tbody2 = gradeTable.querySelector("tbody");
    if (tbody2) {
      const targetRow = tbody2.children[focusRowIdx];
      if (targetRow) {
        const targetCell = targetRow.children[focusCellIdx];
        if (targetCell) {
          const targetInput = targetCell.querySelector("input");
          if (targetInput && !targetInput.disabled) {
            // Ripristina il testo grezzo prima del focus, così '7.' non viene troncato
            if (focusRawValue !== null) {
              targetInput.value = focusRawValue;
            }
            targetInput.focus();
            if (focusCursorPos >= 0) {
              try { targetInput.setSelectionRange(focusCursorPos, focusCursorPos); } catch (e) {}
            }
          }
        }
      }
    }
  }

  warnings.forEach((message) => {
    const warning = document.createElement("div");
    warning.classList.add("warning");
    warning.textContent = message;
    warningArea.appendChild(warning);
  });
}

function createScoreInput(
  student,
  testId,
  sectionId,
  subsectionId,
  type,
  isDisabled = false
) {
  ensureScoreStore(student, testId, sectionId);
  const input = document.createElement("input");
  input.type = "number";
  input.step = "0.1";
  input.min = "0";
  input.disabled = Boolean(isDisabled);

  if (type === "subsection") {
    const value = student.scores[testId][sectionId].subsections[subsectionId];
    input.value = value ?? "";
  } else {
    input.value = student.scores[testId][sectionId].direct ?? "";
  }

  input.addEventListener("blur", (event) => {
    const value = parseNumber(event.target.value);
    if (type === "subsection") {
      student.scores[testId][sectionId].subsections[subsectionId] = value;
    } else {
      student.scores[testId][sectionId].direct = value;
    }
    saveState();
    // Re-render completo solo quando si lascia la cella, mai durante la digitazione
    if (!isNavigatingWithArrows) {
      renderTestTable();
    }
  });

  // Aggiorna solo lo stato mentre si digita (nessun re-render DOM)
  input.addEventListener("input", (event) => {
    const value = parseNumber(event.target.value);
    if (type === "subsection") {
      student.scores[testId][sectionId].subsections[subsectionId] = value;
    } else {
      student.scores[testId][sectionId].direct = value;
    }
    // Aggiorna la cella FINAL della riga in-place, senza toccare il DOM dell'input attivo
    updateFinalCellInRow(input, student, getSelectedTest());
  });

  // Navigazione tra celle con frecce della tastiera
  input.addEventListener("keydown", (event) => {
    if (event.key === "ArrowRight") {
      event.preventDefault();
      // Salva il valore PRIMA di navigare
      const value = parseNumber(input.value);
      if (type === "subsection") {
        student.scores[testId][sectionId].subsections[subsectionId] = value;
      } else {
        student.scores[testId][sectionId].direct = value;
      }
      saveState();
      
      isNavigatingWithArrows = true;
      const cell = input.parentElement;
      const nextCell = cell.nextElementSibling;
      if (nextCell) {
        const nextInput = nextCell.querySelector("input");
        if (nextInput && !nextInput.disabled) {
          nextInput.focus();
          nextInput.select();
        }
      }
      isNavigatingWithArrows = false;
    } else if (event.key === "ArrowLeft") {
      event.preventDefault();
      // Salva il valore PRIMA di navigare
      const value = parseNumber(input.value);
      if (type === "subsection") {
        student.scores[testId][sectionId].subsections[subsectionId] = value;
      } else {
        student.scores[testId][sectionId].direct = value;
      }
      saveState();
      
      isNavigatingWithArrows = true;
      const cell = input.parentElement;
      const prevCell = cell.previousElementSibling;
      if (prevCell) {
        const prevInput = prevCell.querySelector("input");
        if (prevInput && !prevInput.disabled) {
          prevInput.focus();
          prevInput.select();
        }
      }
      isNavigatingWithArrows = false;
    } else if (event.key === "ArrowDown") {
      event.preventDefault();
      // Salva il valore PRIMA di navigare
      const value = parseNumber(input.value);
      if (type === "subsection") {
        student.scores[testId][sectionId].subsections[subsectionId] = value;
      } else {
        student.scores[testId][sectionId].direct = value;
      }
      saveState();
      
      isNavigatingWithArrows = true;
      const cell = input.parentElement;
      const row = cell.parentElement;
      const cellIndex = Array.from(row.children).indexOf(cell);
      const nextRow = row.nextElementSibling;
      if (nextRow) {
        const nextCell = nextRow.children[cellIndex];
        if (nextCell) {
          const nextInput = nextCell.querySelector("input");
          if (nextInput && !nextInput.disabled) {
            nextInput.focus();
            nextInput.select();
          }
        }
      }
      isNavigatingWithArrows = false;
    } else if (event.key === "ArrowUp") {
      event.preventDefault();
      // Salva il valore PRIMA di navigare
      const value = parseNumber(input.value);
      if (type === "subsection") {
        student.scores[testId][sectionId].subsections[subsectionId] = value;
      } else {
        student.scores[testId][sectionId].direct = value;
      }
      saveState();
      
      isNavigatingWithArrows = true;
      const cell = input.parentElement;
      const row = cell.parentElement;
      const cellIndex = Array.from(row.children).indexOf(cell);
      const prevRow = row.previousElementSibling;
      if (prevRow) {
        const prevCell = prevRow.children[cellIndex];
        if (prevCell) {
          const prevInput = prevCell.querySelector("input");
          if (prevInput && !prevInput.disabled) {
            prevInput.focus();
            prevInput.select();
          }
        }
      }
      isNavigatingWithArrows = false;
    }
  });

  // Inizializza la multi-selezione per questo input
  initializeInputSelection(input);

  return input;
}

/**
 * Aggiorna solo la cella FINAL della riga senza ricostruire tutto il DOM.
 * Chiamata durante la digitazione per mostrare il voto finale in tempo reale.
 */
/**
 * Aggiunge il trigger (bordo destro cliccabile) per il commento di una cella.
 */
function attachCommentTrigger(cell, student, testId, sectionId, subsectionId) {
  const key = subsectionId ?? "direct";
  const existingComment = student.scores?.[testId]?.[sectionId]?.comments?.[key];

  const trigger = document.createElement("div");
  trigger.className = "comment-trigger";
  if (existingComment) {
    trigger.classList.add("has-comment");
    trigger.title = existingComment;
  } else {
    trigger.title = "Aggiungi commento";
  }

  trigger.addEventListener("click", (e) => {
    e.stopPropagation();
    openCommentModal(student, testId, sectionId, subsectionId, trigger);
  });

  cell.appendChild(trigger);
}

/**
 * Apre la modale per inserire/modificare il commento di una cella.
 */
function openCommentModal(student, testId, sectionId, subsectionId, trigger) {
  commentModalContext = { student, testId, sectionId, subsectionId, trigger };
  const key = subsectionId ?? "direct";
  const existing = student.scores?.[testId]?.[sectionId]?.comments?.[key] ?? "";
  document.getElementById("commentTextarea").value = existing;
  document.getElementById("commentDialog").showModal();
  document.getElementById("commentTextarea").focus();
}

function updateFinalCellInRow(input, student, test) {
  const cell = input.closest("td");
  if (!cell) return;
  const row = cell.closest("tr");
  if (!row) return;
  const finalCell = row.querySelector("td.final-cell");
  if (!finalCell) return;

  const selectedTest = test ?? getSelectedTest();
  if (!selectedTest) return;

  ensureVersionSelections();
  const activeVersion = getVersionById(selectedTest, state.selectedTestVersionId)
    ?? getDefaultVersion(selectedTest);
  const finalScore = getFinalScore(student, selectedTest, activeVersion);
  finalCell.textContent = formatScore(finalScore);
  finalCell.classList.toggle("low-grade", isLowGrade(finalScore));
}

function ensureScoreStore(student, testId, sectionId) {
  if (!student.scores[testId]) {
    student.scores[testId] = {};
  }
  if (!student.scores[testId][sectionId]) {
    student.scores[testId][sectionId] = { subsections: {}, direct: null };
  }
  if (!student.scores[testId][sectionId].subsections) {
    student.scores[testId][sectionId].subsections = {};
  }
}

function getSectionScore(student, test, section) {
  ensureScoreStore(student, test.id, section.id);
  if (!Array.isArray(section.subsections)) section.subsections = [];
  if (section.subsections.length > 0) {
    const totals = getSectionTotals(section);
    if (totals.totalWeight <= 0 || totals.totalMax <= 0) {
      return 0;
    }
    const weightedRatioSum = section.subsections.reduce((sum, subsection) => {
      const value = parseNumber(
        student.scores[test.id][section.id].subsections[subsection.id]
      ) ?? 0;
      const max = getSubsectionMax(section, subsection, totals.fallbackPerSubMax);
      const weight = getSubsectionWeight(subsection);
      if (max <= 0 || weight <= 0) {
        return sum;
      }
      return sum + (value / max) * weight;
    }, 0);
    const averageRatio = weightedRatioSum / totals.totalWeight;
    return averageRatio * totals.totalMax;
  }

  return parseNumber(student.scores[test.id][section.id].direct) ?? 0;
}

function getFinalScore(student, test, version) {
  if (!test) {
    return null;
  }
  const targetVersion =
    version ??
    getVersionById(
      test,
      getStudentVersionId(student, test.id, getDefaultVersion(test)?.id)
    );
  const sections = targetVersion?.sections ?? [];
  let weightedSum = 0;
  let weightedMaxSum = 0;

  sections.forEach((section) => {
    const score = getSectionScore(student, test, section);
    const weight = getSectionWeight(section) ?? 0;
    const max = getSectionMax(section) ?? 0;
    if (weight > 0 && max > 0) {
      weightedSum += score * weight;
      weightedMaxSum += max * weight;
    }
  });

  if (weightedMaxSum === 0) {
    return null;
  }

  return (weightedSum * 10) / weightedMaxSum;
}

// Verifica se un test ha almeno un voto in una classe
function testHasGradesInClass(test, selectedClass) {
  if (!test || !selectedClass) {
    return false;
  }
  // Controlla se almeno uno studente della classe ha un voto per questo test
  return selectedClass.students.some((student) => {
    const finalScore = getFinalScore(student, test);
    return finalScore !== null && finalScore > 0;
  });
}

function removeSectionScores(sectionId, testId) {
  if (!testId) {
    return;
  }
  state.classes.forEach((classItem) => {
    classItem.students.forEach((student) => {
      if (student.scores?.[testId]) {
        delete student.scores[testId][sectionId];
      }
    });
  });
}

function removeSubsectionScores(sectionId, subsectionId, testId) {
  if (!testId) {
    return;
  }
  state.classes.forEach((classItem) => {
    classItem.students.forEach((student) => {
      if (student.scores?.[testId]?.[sectionId]?.subsections) {
        delete student.scores[testId][sectionId].subsections[subsectionId];
      }
    });
  });
}

function formatSectionTitle(section) {
  const name = section.name || "Section";
  const sectionWeight = getSectionWeight(section);
  const sectionMax = getSectionMax(section);
  const weight = sectionWeight != null ? `w:${sectionWeight}` : "w:?";
  const max = sectionMax != null ? `max:${sectionMax}` : "max:?";
  return `${name} (${weight}, ${max})`;
}

function formatScore(score) {
  if (score === null || score === undefined) {
    return "";
  }
  return Math.round(score * 10) / 10;
}

function isLowGrade(value) {
  const numeric = parseNumber(value);
  return numeric != null && numeric < 6;
}

function parseNumber(value) {
  if (value === "" || value === null || value === undefined) {
    return null;
  }
  const number = Number(value);
  return Number.isNaN(number) ? null : number;
}

function ensureTestVersions(test) {
  if (!test) {
    return;
  }
  
  // Inizializza versions se non esiste
  if (!Array.isArray(test.versions)) {
    test.versions = [];
  }
  
  // Se non ha versioni, creane due (Standard e Facilitata)
  if (test.versions.length === 0) {
    const baseSections = Array.isArray(test.sections) ? test.sections : [];
    const standardVersionId = createId("ver");
    const facilitatedVersionId = createId("ver");
    
    test.versions = [
      {
        id: standardVersionId,
        name: "Standard",
        sections: JSON.parse(JSON.stringify(baseSections)), // Deep copy
      },
      {
        id: facilitatedVersionId,
        name: "Facilitata",
        sections: JSON.parse(JSON.stringify(baseSections)), // Deep copy indipendente
      },
    ];
    
    test.facilitatedVersionId = facilitatedVersionId;
  }
  // Se ha solo 1 versione, aggiungi la Facilitata
  else if (test.versions.length === 1) {
    const baseSections = Array.isArray(test.sections) ? test.sections : 
                        (test.versions[0]?.sections ? test.versions[0].sections : []);
    const facilitatedVersionId = createId("ver");
    
    test.versions.push({
      id: facilitatedVersionId,
      name: "Facilitata",
      sections: JSON.parse(JSON.stringify(baseSections)), // Deep copy indipendente
    });
    
    test.facilitatedVersionId = facilitatedVersionId;
  }
  
  // Assicura che tutte le versioni abbiano proprietà valide
  test.versions.forEach((version) => {
    if (!version.id) {
      version.id = createId("ver");
    }
    if (!version.name) {
      version.name = "Versione";
    }
    if (!Array.isArray(version.sections)) {
      version.sections = [];
    }
  });
  
  // Se non c'è una versione facilitata assegnata, usane una
  if (!test.facilitatedVersionId) {
    test.facilitatedVersionId = test.versions[1]?.id ?? test.versions[0]?.id ?? null;
  }
}

function ensureVersionSelections() {
  const selectedTest = getSelectedTest();
  if (!selectedTest) {
    return;
  }
  ensureTestVersions(selectedTest);
  const versions = selectedTest.versions;
  if (
    !state.selectedTestVersionId ||
    !versions.some((version) => version.id === state.selectedTestVersionId)
  ) {
    state.selectedTestVersionId = versions[0]?.id ?? null;
  }
  if (
    !state.selectedConfigVersionId ||
    !versions.some((version) => version.id === state.selectedConfigVersionId)
  ) {
    state.selectedConfigVersionId = versions[0]?.id ?? null;
  }
}

function getVersionById(test, versionId) {
  if (!test || !versionId) {
    return null;
  }
  return test.versions?.find((version) => version.id === versionId) ?? null;
}

function getDefaultVersion(test) {
  if (!test) {
    return null;
  }
  ensureTestVersions(test);
  return test.versions[0] ?? null;
}

function getFacilitatedVersionId(test) {
  if (!test) {
    return null;
  }
  ensureTestVersions(test);
  if (!test.facilitatedVersionId) {
    test.facilitatedVersionId =
      test.versions[1]?.id ?? test.versions[0]?.id ?? null;
  }
  return test.facilitatedVersionId;
}

function getStudentVersionId(student, testId, fallbackVersionId) {
  if (!student) {
    return fallbackVersionId ?? null;
  }
  if (!student.testVersions) {
    student.testVersions = {};
  }
  return student.testVersions[testId] ?? fallbackVersionId ?? null;
}

function setStudentVersionId(student, testId, versionId) {
  if (!student) {
    return;
  }
  if (!student.testVersions) {
    student.testVersions = {};
  }
  student.testVersions[testId] = versionId;
}

function createVersionFrom(baseVersion, test) {
  const existingNames = (test?.versions ?? []).map((version) => version.name);
  const suggestedName = existingNames.includes("Facilitata")
    ? `Versione ${existingNames.length + 1}`
    : "Facilitata";

  return {
    id: createId("ver"),
    name: suggestedName,
    sections: cloneSections(baseVersion?.sections ?? []),
  };
}

function cloneSections(sections) {
  return sections.map((section) => ({
    id: createId("sec"),
    name: section.name,
    weight: section.weight,
    max: section.max,
    subsections: (section.subsections ?? []).map((subsection) => ({
      id: createId("sub"),
      name: subsection.name,
      weight: subsection.weight,
      max: subsection.max,
    })),
  }));
}

function createSection() {
  return {
    id: createId("sec"),
    name: "New Section",
    weight: 1,
    max: 10,
    subsections: [createSubsection()],
  };
}

function createSubsection(base = null) {
  return {
    id: createId("sub"),
    name: "New Subsection",
    weight: parseNumber(base?.weight) ?? 1,
    max: parseNumber(base?.max) ?? 1,
  };
}

function getSubsectionWeight(subsection) {
  return parseNumber(subsection?.weight) ?? 1;
}

function getSectionTotals(section) {
  const subsections = section?.subsections ?? [];
  const fallbackMax = parseNumber(section?.max) ?? 0;
  const fallbackPerSub = subsections.length > 0 ? fallbackMax / subsections.length : 0;
  return subsections.reduce(
    (totals, subsection) => {
      totals.totalWeight += getSubsectionWeight(subsection);
      totals.totalMax += getSubsectionMax(section, subsection, fallbackPerSub);
      return totals;
    },
    { totalWeight: 0, totalMax: 0, fallbackPerSubMax: fallbackPerSub }
  );
}

function getSubsectionMax(section, subsection, fallbackPerSub = null) {
  const explicitMax = parseNumber(subsection?.max);
  if (explicitMax != null) {
    return explicitMax;
  }
  if (fallbackPerSub != null) {
    return fallbackPerSub;
  }
  const fallbackMax = parseNumber(section?.max) ?? 0;
  const count = section?.subsections?.length ?? 0;
  return count > 0 ? fallbackMax / count : 0;
}

function getSectionWeight(section) {
  if (section?.subsections?.length) {
    return getSectionTotals(section).totalWeight;
  }
  return parseNumber(section?.weight) ?? 1;
}

function getSectionMax(section) {
  if (section?.subsections?.length) {
    return getSectionTotals(section).totalMax;
  }
  return parseNumber(section?.max) ?? null;
}

function createStudent() {
  return {
    id: createId("stu"),
    name: "New Student",
    scores: {},
    testVersions: {},
  };
}

function createTest(title, subject, category) {
  const standardVersionId = createId("ver");
  const facilitatedVersionId = createId("ver");
  const baseVersion = {
    id: standardVersionId,
    name: "Standard",
    sections: [],
  };
  const facilitatedVersion = {
    id: facilitatedVersionId,
    name: "Facilitata",
    sections: [],
  };
  return {
    id: createId("test"),
    title: title || "Nuova verifica",
    subject: subject || "",
    categories: category ? [category] : [],
    classIds: [],
    classDates: {},
    versions: [baseVersion, facilitatedVersion],
    facilitatedVersionId: facilitatedVersionId,
  };
}

/** Garantisce che i test vecchi abbiano i nuovi campi */
function ensureTestMeta(test) {
  if (!test) return;
  // Migra vecchio campo "category" (stringa) → "categories" (array)
  if (!Array.isArray(test.categories)) {
    test.categories = test.category ? test.category.split(/[,;]+/).map(s => s.trim()).filter(Boolean) : [];
    delete test.category;
  }
  if (!test.classIds)   test.classIds   = [];
  if (!test.classDates) test.classDates = {};
}

function createClass(name) {
  return {
    id: createId("class"),
    name: name || "New Class",
    students: [],
  };
}

function createId(prefix) {
  return `${prefix}-${Math.random().toString(36).slice(2, 9)}`;
}

function getStudentAverage(student) {
  if (!state.tests.length) {
    return 0;
  }
  // Filtra i voti: esclude null, undefined e voti <= 2 (non svolti)
  // Se lo studente è facilitato, usa la versione facilitata di ogni test
  const scores = state.tests
    .map((test) => {
      if (student.facilitated === true) {
        const facilitatedVersion = getVersionById(test, getFacilitatedVersionId(test));
        return getFinalScore(student, test, facilitatedVersion);
      } else {
        return getFinalScore(student, test);
      }
    })
    .filter((value) => value !== null && value !== undefined && value > 2);

  if (scores.length === 0) {
    return 0;
  }

  const sum = scores.reduce((total, value) => total + value, 0);
  return sum / scores.length;
}

function saveState() {
  // Salva preferenze UI e cache voti in localStorage (fallback offline)
  const dataToSave = {
    tests: state.tests,
    selectedClassId: state.selectedClassId,
    selectedTestId: state.selectedTestId,
    selectedTestVersionId: state.selectedTestVersionId,
    selectedConfigVersionId: state.selectedConfigVersionId,
    view: state.view,
    studentScores: buildStudentScoresMap(),
    studentTestVersions: buildStudentTestVersionsMap(),
    studentFacilitated: buildStudentFacilitatedMap(),
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(dataToSave));

  // Salva voti e test su Firebase con debounce (500ms)
  // così non spammiamo il DB a ogni tasto premuto
  if (fbDb && fbUser) {
    clearTimeout(fbSaveTimer);
    fbSaveTimer = setTimeout(() => saveGradingToFirebase(), 500);
  }
}

/**
 * Scrive voti, test e flag DSA su /users/{uid}/grading/
 * Struttura:
 *   /grading/tests         → array delle verifiche (struttura, sezioni, pesi)
 *   /grading/scores        → { studentId: { testId: { sectionId: ... } } }
 *   /grading/facilitated   → { studentId: true }
 *   /grading/testVersions  → { studentId: { testId: versionId } }
 */
function saveGradingToFirebase() {
  if (!fbDb || !fbUser) return;
  fbIgnoreGrading = true; // evita che il listener ri-scriva quello che abbiamo appena scritto
  const path = `/users/${fbUser.uid}/grading`;
  const data = {
    tests: state.tests,
    scores: buildStudentScoresMap(),
    facilitated: buildStudentFacilitatedMap(),
    testVersions: buildStudentTestVersionsMap(),
    savedAt: Date.now(),
  };
  fbDb.ref(path).set(data)
    .then(() => {
      setFirebaseStatus("🟢 Voti sincronizzati su Firebase");
      setTimeout(() => { fbIgnoreGrading = false; }, 2000);
    })
    .catch((err) => {
      fbIgnoreGrading = false;
      setFirebaseStatus(`❌ Errore salvataggio voti: ${err.message}`, "error");
    });
}

/** Costruisce una mappa { studentId: { testId: { sectionId: ... } } } */
function buildStudentScoresMap() {
  const map = {};
  state.classes.forEach((cls) => {
    cls.students.forEach((student) => {
      if (student.scores && Object.keys(student.scores).length > 0) {
        map[student.id] = student.scores;
      }
    });
  });
  return map;
}

/** Costruisce una mappa { studentId: { testId: versionId } } */
function buildStudentTestVersionsMap() {
  const map = {};
  state.classes.forEach((cls) => {
    cls.students.forEach((student) => {
      if (student.testVersions && Object.keys(student.testVersions).length > 0) {
        map[student.id] = student.testVersions;
      }
    });
  });
  return map;
}

/** Costruisce una mappa { studentId: true } per gli studenti DSA/104 */
function buildStudentFacilitatedMap() {
  const map = {};
  state.classes.forEach((cls) => {
    cls.students.forEach((student) => {
      if (student.facilitated === true) {
        map[student.id] = true;
      }
    });
  });
  return map;
}



function loadState() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) {
    return structuredClone(defaultData);
  }

  try {
    const parsed = JSON.parse(saved);

    // Supporto formato vecchio che aveva le classi dentro
    const classesFromOld = parsed.students
      ? [
          {
            id: createId("class"),
            name: "Class 1",
            students: parsed.students ?? [],
          },
        ]
      : null;

    const testsFromOld = parsed.sections
      ? [
          {
            id: createId("test"),
            title: parsed.title ?? "New Test",
            subject: parsed.subject ?? "",
            sections: parsed.sections ?? [],
          },
        ]
      : null;

    // Nel formato nuovo le classi sono vuote (arrivano da Firebase);
    // nel formato vecchio potrebbero essere presenti temporaneamente.
    const classes = parsed.classes ?? classesFromOld ?? [];
    const tests = (parsed.tests ?? testsFromOld ?? []).map(test => ({
      ...test,
      subject: typeof test.subject === 'string' ? test.subject : ""
    }));
    
    // Assicura che ogni test abbia due versioni (Standard e Facilitata)
    tests.forEach(test => ensureTestVersions(test));

    const selectedClassId = parsed.selectedClassId ?? classes[0]?.id ?? null;
    const selectedTestId = parsed.selectedTestId ?? tests[0]?.id ?? null;
    const selectedTestVersionId = parsed.selectedTestVersionId ?? null;
    const selectedConfigVersionId = parsed.selectedConfigVersionId ?? null;
    const view = parsed.view ?? "home";

    // Ripristina le mappe dei voti, versioni e flag DSA per studente
    const studentScores = parsed.studentScores ?? {};
    const studentTestVersions = parsed.studentTestVersions ?? {};
    const studentFacilitated = parsed.studentFacilitated ?? {};

    // Applica i dati salvati agli studenti già presenti (formato vecchio)
    classes.forEach((cls) => {
      cls.students.forEach((student) => {
        if (!student.scores) student.scores = {};
        if (!student.testVersions) student.testVersions = {};
        if (studentScores[student.id]) {
          Object.assign(student.scores, studentScores[student.id]);
        }
        if (studentTestVersions[student.id]) {
          Object.assign(student.testVersions, studentTestVersions[student.id]);
        }
        if (studentFacilitated[student.id]) {
          student.facilitated = true;
        }
      });
    });

    const data = {
      classes,
      tests,
      selectedClassId,
      selectedTestId,
      selectedTestVersionId,
      selectedConfigVersionId,
      view,
      // Teniamo le mappe in state così mergeFirebaseClasses può usarle
      _studentScores: studentScores,
      _studentTestVersions: studentTestVersions,
      _studentFacilitated: studentFacilitated,
    };

    tests.forEach((testItem) => {
      ensureTestVersions(testItem);
      ensureTestMeta(testItem);
    });

    if (tests[0]) {
      normalizeStudentScores(classes, tests[0]);
    }

    return data;
  } catch (error) {
    return structuredClone(defaultData);
  }
}

function exportCSV() {
  const selectedClass = getSelectedClass();
  const selectedTest = getSelectedTest();
  const students = selectedClass?.students ?? [];
  if (!selectedTest) {
    return;
  }
  ensureTestVersions(selectedTest);
  ensureVersionSelections();
  const activeVersion = getVersionById(
    selectedTest,
    state.selectedTestVersionId
  );
  const headers = ["Student", "Versione"];

  (activeVersion?.sections ?? []).forEach((section) => {
    if (section.subsections.length > 0) {
      section.subsections.forEach((subsection) => {
        headers.push(`${section.name} - ${subsection.name}`);
      });
      headers.push(`${section.name} - Total`);
    } else {
      headers.push(section.name);
    }
  });

  headers.push("Final Grade");

  const rows = students.map((student) => {
    const versionId = getStudentVersionId(
      student,
      selectedTest.id,
      getDefaultVersion(selectedTest)?.id
    );
    const versionLabel = getVersionById(selectedTest, versionId)?.name || "";
    const row = [student.name, versionLabel];

    (activeVersion?.sections ?? []).forEach((section) => {
      if (section.subsections.length > 0) {
        section.subsections.forEach((subsection) => {
          const value =
            student.scores[selectedTest.id]?.[section.id]?.subsections?.[
              subsection.id
            ] ?? "";
          row.push(value);
        });
        row.push(getSectionScore(student, selectedTest, section));
      } else {
        const value =
          student.scores[selectedTest.id]?.[section.id]?.direct ?? "";
        row.push(value);
      }
    });

    row.push(formatScore(getFinalScore(student, selectedTest, activeVersion)));
    return row;
  });

  const csvContent = [headers, ...rows]
    .map((row) => row.map(escapeCsv).join(","))
    .join("\n");

  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  const classLabel = selectedClass?.name ? `-${selectedClass.name}` : "";
  link.download = `${(selectedTest.title || "grades").replace(/\s+/g, "-")}${classLabel.replace(/\s+/g, "-")}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function getSelectedClass() {
  if (!state.classes || state.classes.length === 0) {
    return null;
  }
  const selected = state.classes.find(
    (classItem) => classItem.id === state.selectedClassId
  );
  return selected ?? state.classes[0];
}

function getSelectedTest() {
  if (!state.tests || state.tests.length === 0) {
    return null;
  }
  const selected = state.tests.find(
    (testItem) => testItem.id === state.selectedTestId
  );
  return selected ?? state.tests[0];
}

function ensureClassState() {
  // Se Firebase è connesso e sta caricando, non creare classi vuote
  if (fbUser && (!state.classes || state.classes.length === 0)) {
    // Le classi arriveranno dal listener Firebase – non fare nulla
    return;
  }
  if (!state.classes || state.classes.length === 0) {
    state.classes = [createClass("Class 1")];
  }
  const hasSelected = state.classes.some(
    (classItem) => classItem.id === state.selectedClassId
  );
  if (!state.selectedClassId || !hasSelected) {
    state.selectedClassId = state.classes[0]?.id ?? null;
  }
}

function ensureTestState() {
  if (!state.tests || state.tests.length === 0) {
    state.tests = [createTest("Test 1")];
  }
  state.tests.forEach((test) => { ensureTestVersions(test); ensureTestMeta(test); });
  const hasSelected = state.tests.some(
    (testItem) => testItem.id === state.selectedTestId
  );
  if (!state.selectedTestId || !hasSelected) {
    state.selectedTestId = state.tests[0]?.id ?? null;
  }
  ensureVersionSelections();
}

function generateClassName() {
  const count = state.classes?.length ?? 0;
  return `Class ${count + 1}`;
}

function generateTestName() {
  const count = state.tests?.length ?? 0;
  return `Test ${count + 1}`;
}

function normalizeStudentScores(classes, test) {
  const sectionIds = (getDefaultVersion(test)?.sections ?? []).map(
    (section) => section.id
  );
  classes.forEach((classItem) => {
    classItem.students.forEach((student) => {
      if (!student.scores) {
        student.scores = {};
      }
      const hasTestScores = Boolean(student.scores[test.id]);
      const looksOld = Object.keys(student.scores).some((key) =>
        sectionIds.includes(key)
      );
      if (!hasTestScores && looksOld) {
        const oldScores = student.scores;
        student.scores = {};
        student.scores[test.id] = oldScores;
      }
    });
  });
}

function escapeCsv(value) {
  if (value === null || value === undefined) {
    return "";
  }
  const stringValue = String(value);
  if (stringValue.includes(",") || stringValue.includes("\n")) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }
  return stringValue;
}


// =====================================================================
//  FIREBASE INTEGRATION (Realtime Database)
//  Struttura DB: /users/{uid}/classes/[{id, name, students: {...}}]
// =====================================================================

function initFirebase() {
  try {
    fbApp = firebase.initializeApp(FIREBASE_CONFIG);
  } catch (e) {
    fbApp = firebase.app();
  }
  fbAuth = firebase.auth();
  fbDb = firebase.database();

  const loginBtn = document.getElementById("loginBtn");
  const logoutBtn = document.getElementById("logoutBtn");
  const userInfo = document.getElementById("userInfo");
  const firebaseBadge = document.getElementById("firebaseBadge");

  fbAuth.onAuthStateChanged((user) => {
    fbUser = user;

    if (user) {
      if (loginBtn) loginBtn.style.display = "none";
      if (logoutBtn) logoutBtn.style.display = "";
      if (userInfo) {
        userInfo.style.display = "";
        userInfo.textContent = `👤 ${user.displayName || user.email}`;
      }
      if (firebaseBadge) firebaseBadge.style.display = "";
      setFirebaseStatus("🟢 Connesso — caricamento classi e voti…");
      startClassesListener();
    } else {
      if (loginBtn) loginBtn.style.display = "";
      if (logoutBtn) logoutBtn.style.display = "none";
      if (userInfo) userInfo.style.display = "none";
      if (firebaseBadge) firebaseBadge.style.display = "none";
      setFirebaseStatus("🔒 Accedi con Google per caricare le classi da Firebase");

      if (fbClassesUnsubscribe && fbUser) {
        fbDb.ref(`/users/${fbUser.uid}/classes`).off("value", fbClassesUnsubscribe);
        fbClassesUnsubscribe = null;
      }
      if (fbGradingUnsubscribe && fbUser) {
        fbDb.ref(`/users/${fbUser.uid}/grading`).off("value", fbGradingUnsubscribe);
        fbGradingUnsubscribe = null;
      }
      clearTimeout(fbSaveTimer);
      state.classes = [];
      state._fbGrading = null;
      render();
    }
  });
}

/**
 * Listener real-time su /users/{uid}/classes
 * La struttura in Firebase è un array (indici 0,1,2…) di oggetti:
 *   { id: "...", name: "1A", students: { key: { name: "...", ... } } }
 */
function startClassesListener() {
  if (!fbDb || !fbUser) return;

  // Listener 1: classi (read-only da classroomanager)
  const classesRef = fbDb.ref(`/users/${fbUser.uid}/classes`);
  fbClassesUnsubscribe = classesRef.on(
    "value",
    (snapshot) => {
      const data = snapshot.val();
      if (!data) {
        setFirebaseStatus("⚠️ Nessuna classe trovata nel database");
        state.classes = [];
        render();
        return;
      }
      mergeFirebaseClasses(data);
      render();
      setFirebaseStatus(`🟢 Connesso — ${state.classes.length} classi, voti sincronizzati`);
    },
    (error) => {
      console.error("Firebase classes error:", error);
      setFirebaseStatus(`❌ ${error.message} — controlla le regole del Realtime Database`, "error");
    }
  );

  // Listener 2: voti/test (read-write, nostro grading app)
  const gradingRef = fbDb.ref(`/users/${fbUser.uid}/grading`);
  fbGradingUnsubscribe = gradingRef.on(
    "value",
    (snapshot) => {
      // Ignora l'aggiornamento se siamo stati noi a scrivere (evita loop)
      if (fbIgnoreGrading) return;
      const data = snapshot.val();
      if (!data) return; // nessun dato grading ancora: va bene, partiamo vuoti
      applyFirebaseGrading(data);
    },
    (error) => {
      console.error("Firebase grading error:", error);
    }
  );
}

/**
 * Applica i dati di grading arrivati da Firebase allo state.
 * Chiamata dal listener real-time OPPURE all'avvio per caricare i dati da un altro dispositivo.
 */
function applyFirebaseGrading(data) {
  // Aggiorna i test (struttura verifiche, sezioni, pesi)
  if (Array.isArray(data.tests) && data.tests.length > 0) {
    state.tests = data.tests;
    state.tests.forEach((t) => { ensureTestVersions(t); ensureTestMeta(t); });
  }

  // Memorizza le mappe in state così mergeFirebaseClasses le userà
  state._studentScores = data.scores || {};
  state._studentTestVersions = data.testVersions || {};
  state._studentFacilitated = data.facilitated || {};

  // Applica subito agli studenti già caricati
  state.classes.forEach((cls) => {
    cls.students.forEach((student) => {
      if (state._studentScores[student.id]) {
        student.scores = state._studentScores[student.id];
      }
      if (state._studentTestVersions[student.id]) {
        student.testVersions = state._studentTestVersions[student.id];
      }
      student.facilitated = state._studentFacilitated[student.id] === true;
    });
  });

  render();
}

/**
 * Converte i dati da Firebase al formato interno, preservando i voti locali.
 *
 * Firebase restituisce gli array come oggetti { "0": {...}, "1": {...} }
 * Ogni classe ha: { id, name, students: { key: { name, ... } } }
 * Ogni studente ha i campi "fullName" (usato come ID stabile) e "displayName" (visualizzato)
 */
function mergeFirebaseClasses(fbData) {
  // Usa i dati di grading da Firebase o dalla cache locale
  const studentScores = state._studentScores || buildStudentScoresMap();
  const studentTestVersions = state._studentTestVersions || buildStudentTestVersionsMap();
  const studentFacilitated = state._studentFacilitated || buildStudentFacilitatedMap();

  // Firebase converte gli array JS in oggetti { "0": {...}, "1": {...} }
  const rawClasses = Array.isArray(fbData) ? fbData : Object.values(fbData);

  const newClasses = rawClasses
    .filter(Boolean)
    .map((classData) => {
      // L'id della CLASSE è numerico (timestamp), usalo come stringa
      const classId = String(classData.id || createId("class"));
      const className = classData.name || classId;

      // Gli studenti in classroomanager sono un array di { fullName, displayName }
      // senza campo "id" — Firebase li converte in oggetto con chiavi "0","1","2"...
      const studentsRaw = classData.students || {};
      const studentsArray = Array.isArray(studentsRaw)
        ? studentsRaw
        : Object.values(studentsRaw);

      const students = studentsArray
        .filter(Boolean)
        .map((studentData) => {
          // ⚠️ Gli studenti NON hanno un campo "id" in classroomanager.
          // Usiamo fullName (es. "Rossi Mario") come ID stabile — non cambia mai.
          const studentId = studentData[FB_STUDENT_FULLNAME_FIELD]
            || studentData.fullName
            || studentData.name
            || String(Math.random()); // ultimo fallback (non dovrebbe mai capitare)

          // Nome visualizzato nella tabella voti
          const studentName = studentData[FB_STUDENT_DISPLAY_FIELD]
            || studentData.displayName
            || studentId;

          return {
            id: studentId,           // "Rossi Mario" — stabile tra sessioni
            name: studentName,       // "Mario R."    — visualizzato
            scores: studentScores[studentId] || {},
            testVersions: studentTestVersions[studentId] || {},
            facilitated: studentFacilitated[studentId] === true,
          };
        });

      // Ordina per cognome (il fullName inizia con il cognome)
      students.sort((a, b) => a.id.localeCompare(b.id, "it"));

      return { id: classId, name: className, students };
    });

  newClasses.sort((a, b) => a.name.localeCompare(b.name, "it"));
  state.classes = newClasses;

  const hasSelected = state.classes.some((c) => c.id === state.selectedClassId);
  if (!hasSelected && state.classes.length > 0) {
    state.selectedClassId = state.classes[0].id;
  }
}

function setFirebaseStatus(message, type = "info") {
  const el = document.getElementById("firebaseStatus");
  if (!el) return;
  el.textContent = message;
  el.style.color = type === "error" ? "#ffb3b3" : "rgba(255,255,255,0.85)";
}

// Le classi sono read-only da classroomanager — non scriviamo su Firebase
function saveClassesToFirebase() {}

// =====================================================================
//  MULTI-SELECTION & COPY/PASTE FUNCTIONALITY
//  Permette di selezionare più celle e copiarle/incollarle come Excel
// =====================================================================

// Stato di selezione
let selectionState = {
  selectedInputs: new Set(),
  clipboard: null,
  clipboardLayout: null, // { rows, cols } per paste smarter
  isSelecting: false,
  startInput: null,
};

/**
 * Inizializza la multi-selezione per un input della tabella.
 * Logica:
 * - Click semplice       → focus normale, nessuna interferenza
 * - Shift+Click          → range selection
 * - Ctrl+Click           → toggle selezione singola
 * - Mousedown + trascinamento → selezione rettangolare
 */
function initializeInputSelection(input) {
  input.addEventListener("mousedown", (event) => {
    if (event.button !== 0) return;

    if (event.shiftKey && selectionState.startInput) {
      // Shift+Click: seleziona il range, ma lascia il focus all'input
      selectRangeBetween(selectionState.startInput, input);
      event.preventDefault();
      return;
    }

    if (event.ctrlKey || event.metaKey) {
      // Ctrl+Click: toggle cella singola senza perdere il focus
      toggleInputSelection(input);
      event.preventDefault();
      return;
    }

    // Click semplice: NON fare nulla qui — lascia che il browser gestisca
    // focus e cursor normalmente. Pulisci la selezione visiva solo su focus,
    // così non interferisce con la digitazione.
    selectionState.startInput = input;
    selectionState.isSelecting = false;

    // Inizia il drag solo se il mouse si muove mentre il tasto è premuto
    const onMouseMove = () => {
      if (!selectionState.isSelecting) {
        selectionState.isSelecting = true;
        clearSelection();
        addToSelection(input);
      }
    };

    const onMouseUp = () => {
      selectionState.isSelecting = false;
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
    };

    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
  });

  input.addEventListener("focus", () => {
    // Se non siamo in drag, in navigazione frecce, e nessuna selezione attiva su questo input, pulisci
    if (!selectionState.isSelecting && !isNavigatingWithArrows && !selectionState.selectedInputs.has(input)) {
      clearSelection();
    }
    selectionState.startInput = input;
  });

  // Estendi la selezione durante il trascinamento
  input.addEventListener("mouseover", () => {
    if (!selectionState.isSelecting) return;
    if (selectionState.startInput && input !== selectionState.startInput) {
      selectRangeBetween(selectionState.startInput, input);
    }
  });
}

/**
 * Aggiunge un input alla selezione e lo evidenzia
 */
function addToSelection(input) {
  selectionState.selectedInputs.add(input);
  input.classList.add("selected");
}

/**
 * Rimuove un input dalla selezione
 */
function removeFromSelection(input) {
  selectionState.selectedInputs.delete(input);
  input.classList.remove("selected");
}

/**
 * Alterna lo stato di selezione di un input (Ctrl+Click)
 */
function toggleInputSelection(input) {
  if (selectionState.selectedInputs.has(input)) {
    removeFromSelection(input);
  } else {
    addToSelection(input);
  }
  selectionState.startInput = input;
}

/**
 * Seleziona tutti gli input tra due celle (Shift+Click o range selection)
 */
function selectRangeBetween(input1, input2) {
  const table = gradeTable;
  const rows = Array.from(table.querySelectorAll("tbody tr"));
  
  // Trova i due input nella tabella
  let start = -1, end = -1;
  let startCol = -1, endCol = -1;
  
  rows.forEach((row, rowIdx) => {
    const cells = Array.from(row.querySelectorAll("td"));
    cells.forEach((cell, colIdx) => {
      const cellInput = cell.querySelector("input");
      if (cellInput === input1) {
        start = rowIdx;
        startCol = colIdx;
      }
      if (cellInput === input2) {
        end = rowIdx;
        endCol = colIdx;
      }
    });
  });
  
  if (start === -1 || end === -1 || startCol === -1 || endCol === -1) return;
  
  // Normalizza start/end (il range può andare in qualsiasi direzione)
  const minRow = Math.min(start, end);
  const maxRow = Math.max(start, end);
  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);
  
  // Seleziona tutti gli input nel rettangolo
  clearSelection();
  for (let r = minRow; r <= maxRow; r++) {
    const cells = Array.from(rows[r].querySelectorAll("td"));
    for (let c = minCol; c <= maxCol; c++) {
      if (cells[c]) {
        const cellInput = cells[c].querySelector("input");
        if (cellInput) {
          addToSelection(cellInput);
        }
      }
    }
  }
}

/**
 * Cancella tutta la selezione
 */
function clearSelection() {
  selectionState.selectedInputs.forEach((input) => {
    input.classList.remove("selected");
  });
  selectionState.selectedInputs.clear();
}

/**
 * Copia i valori delle celle selezionate negli appunti
 * Formato: tab-separated per righe, newline-separated per colonne
 */
function copySelectedCells() {
  if (selectionState.selectedInputs.size === 0) {
    return;
  }
  
  // Estrai gli input selezionati dalla tabella mantenendo la struttura
  const table = gradeTable;
  const rows = Array.from(table.querySelectorAll("tbody tr"));
  const clipboard = [];
  const layout = { rows: new Set(), cols: new Set() };
  
  rows.forEach((row, rowIdx) => {
    const cells = Array.from(row.querySelectorAll("td"));
    cells.forEach((cell, colIdx) => {
      const cellInput = cell.querySelector("input");
      if (cellInput && selectionState.selectedInputs.has(cellInput)) {
        layout.rows.add(rowIdx);
        layout.cols.add(colIdx);
        clipboard.push({
          row: rowIdx,
          col: colIdx,
          value: cellInput.value || "",
        });
      }
    });
  });
  
  // Converti in formato tab-separated per Excel
  if (clipboard.length > 0) {
    const minRow = Math.min(...layout.rows);
    const minCol = Math.min(...layout.cols);
    const maxRow = Math.max(...layout.rows);
    const maxCol = Math.max(...layout.cols);
    
    const lines = [];
    for (let r = minRow; r <= maxRow; r++) {
      const line = [];
      for (let c = minCol; c <= maxCol; c++) {
        const cell = clipboard.find(item => item.row === r && item.col === c);
        line.push(cell ? cell.value : "");
      }
      lines.push(line.join("\t"));
    }
    
    selectionState.clipboard = lines.join("\n");
    selectionState.clipboardLayout = {
      rows: Array.from(layout.rows),
      cols: Array.from(layout.cols),
    };

    // Copia negli appunti del browser con textarea trick (funziona sempre)
    const ta = document.createElement("textarea");
    ta.value = selectionState.clipboard;
    ta.style.cssText = "position:fixed;opacity:0;pointer-events:none;top:0;left:0;";
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    try { document.execCommand("copy"); } catch (e) {}
    document.body.removeChild(ta);
  }
}

/**
 * Incolla i valori degli appunti a partire dalla cella selezionata
 */
function pasteSelectedCells(startInput) {
  if (!selectionState.clipboard) {
    return;
  }
  
  const table = gradeTable;
  const rows = Array.from(table.querySelectorAll("tbody tr"));
  
  // Trova il punto di partenza del paste
  let startRow = -1, startCol = -1;
  rows.forEach((row, rowIdx) => {
    const cells = Array.from(row.querySelectorAll("td"));
    cells.forEach((cell, colIdx) => {
      const cellInput = cell.querySelector("input");
      if (cellInput === startInput) {
        startRow = rowIdx;
        startCol = colIdx;
      }
    });
  });
  
  if (startRow === -1 || startCol === -1) return;
  
  // Parsa il clipboard
  const lines = selectionState.clipboard.split("\n");
  lines.forEach((line, lineIdx) => {
    const values = line.split("\t");
    const targetRow = startRow + lineIdx;
    
    if (targetRow >= rows.length) return;
    
    const cells = Array.from(rows[targetRow].querySelectorAll("td"));
    values.forEach((value, valueIdx) => {
      const targetCol = startCol + valueIdx;
      if (targetCol >= cells.length) return;
      
      const targetCell = cells[targetCol];
      if (!targetCell) return;
      
      const targetInput = targetCell.querySelector("input");
      if (!targetInput || targetInput.disabled) return;
      
      // Incolla il valore
      targetInput.value = value;
      targetInput.dispatchEvent(new Event("input", { bubbles: true }));
      targetInput.dispatchEvent(new Event("change", { bubbles: true }));
    });
  });
}