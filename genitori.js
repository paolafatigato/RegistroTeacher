// =====================================================================
//  AREA GENITORI — pagina standalone, sola lettura
//  Usa lo stesso progetto Firebase dell'app insegnante (classroomanager),
//  ma un flusso di login separato (email + password) pensato per le famiglie.
//
//  Percorsi letti:
//    /parentAccounts/{uid}                      → { teacherUid, studentId, studentName, className }
//    /publishedGrades/{teacherUid}/{studentId}   → { [testId]: { ...snapshot pubblicato } }
//
//  Queste pagine funzionano solo se:
//   1) nella Console Firebase è attivo il provider "Email/Password";
//   2) le regole del Realtime Database sono state aggiornate come
//      descritto in FIREBASE_SETUP.md.
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

const PARENT_DISPLAY_MODES = {
  grade:  { label: "Voto" },
  good:   { label: "🟢 Ottimo" },
  mid:    { label: "🟡 OK" },
  bad:    { label: "🔴 Da migliorare" },
  hidden: { label: null }, // non mostrato
};

let fbApp = null;
let fbAuth = null;
let fbDb = null;
let fbUser = null;
let publishedUnsub = null;
let publishedRefPath = null;
let currentParentTeacherUid = null;
let currentParentStudents = []; // { studentId, studentName, className } — può contenere più figli (fratelli/sorelle)

// ── Elementi DOM ──────────────────────────────────────────────────────
const loginScreen = document.getElementById("parentLoginScreen");
const dataScreen = document.getElementById("parentDataScreen");
const linkErrorScreen = document.getElementById("parentLinkErrorScreen");
const loginForm = document.getElementById("parentLoginForm");
const loginError = document.getElementById("parentLoginError");
const userBar = document.getElementById("parentUserBar");
const userInfo = document.getElementById("parentUserInfo");
const logoutBtn = document.getElementById("parentLogoutBtn");
const studentHeader = document.getElementById("parentStudentHeader");
const studentSwitcher = document.getElementById("parentStudentSwitcher");
const testsList = document.getElementById("parentTestsList");
const emptyMessage = document.getElementById("parentEmptyMessage");

function init() {
  try {
    fbApp = firebase.initializeApp(FIREBASE_CONFIG);
  } catch (e) {
    fbApp = firebase.app();
  }
  fbAuth = firebase.auth();
  fbDb = firebase.database();

  loginForm.addEventListener("submit", (e) => {
    e.preventDefault();
    loginError.style.display = "none";
    const email = document.getElementById("parentEmailInput").value.trim();
    const password = document.getElementById("parentPasswordInput").value;
    const submitBtn = loginForm.querySelector("button[type='submit']");
    submitBtn.disabled = true;
    fbAuth.signInWithEmailAndPassword(email, password)
      .catch((err) => {
        loginError.textContent = describeAuthError(err);
        loginError.style.display = "";
      })
      .finally(() => {
        submitBtn.disabled = false;
      });
  });

  logoutBtn.addEventListener("click", () => {
    detachPublishedListener();
    fbAuth.signOut();
  });

  fbAuth.onAuthStateChanged((user) => {
    fbUser = user;
    if (user) {
      showScreen("loading");
      userBar.style.display = "flex";
      userInfo.textContent = `👤 ${user.email}`;
      loadParentLinkAndData(user.uid);
    } else {
      userBar.style.display = "none";
      detachPublishedListener();
      currentParentTeacherUid = null;
      currentParentStudents = [];
      if (studentSwitcher) { studentSwitcher.style.display = "none"; studentSwitcher.innerHTML = ""; }
      showScreen("login");
    }
  });
}

function showScreen(which) {
  loginScreen.style.display = which === "login" ? "" : "none";
  dataScreen.style.display = which === "data" ? "" : "none";
  linkErrorScreen.style.display = which === "error" ? "" : "none";
}

function describeAuthError(err) {
  switch (err.code) {
    case "auth/invalid-email":
      return "Email non valida.";
    case "auth/user-not-found":
    case "auth/wrong-password":
    case "auth/invalid-credential":
      return "Email o password non corrette.";
    case "auth/too-many-requests":
      return "Troppi tentativi. Riprova tra qualche minuto.";
    default:
      return "Accesso non riuscito: " + err.message;
  }
}

/** Recupera il collegamento genitore→studenti (uno o più, es. fratelli) e
 *  mostra il primo, con uno switcher se ce n'è più di uno. */
function loadParentLinkAndData(parentUid) {
  fbDb.ref(`/parentAccounts/${parentUid}`).once("value")
    .then((snapshot) => {
      const link = snapshot.val();
      const studentsMap = normalizeParentStudents(link);
      const studentIds = Object.keys(studentsMap);
      if (!link || !link.teacherUid || studentIds.length === 0) {
        showScreen("error");
        return;
      }
      currentParentTeacherUid = link.teacherUid;
      currentParentStudents = studentIds
        .map((id) => studentsMap[id])
        .sort((a, b) => (a.studentName || "").localeCompare(b.studentName || "", "it"));

      renderStudentSwitcher(currentParentStudents);
      selectParentStudent(currentParentStudents[0].studentId);
      showScreen("data");
    })
    .catch(() => {
      showScreen("error");
    });
}

/** Converte /parentAccounts/{uid} nella forma { studentId: {...} },
 *  gestendo sia il formato attuale (students: {...}, più figli possibili)
 *  sia quello "legacy" con un solo studentId in cima al record. */
function normalizeParentStudents(link) {
  if (!link) return {};
  if (link.students && typeof link.students === "object") return link.students;
  if (link.studentId) {
    return {
      [link.studentId]: {
        studentId: link.studentId,
        studentName: link.studentName || "",
        className: link.className || "",
      },
    };
  }
  return {};
}

/** Mostra un piccolo selettore per passare da un figlio all'altro. Se il
 *  genitore ha un solo figlio collegato, resta nascosto (nessun cambiamento
 *  visibile rispetto a prima). */
function renderStudentSwitcher(students) {
  if (!studentSwitcher) return;
  if (students.length <= 1) {
    studentSwitcher.style.display = "none";
    studentSwitcher.innerHTML = "";
    return;
  }
  studentSwitcher.style.display = "";
  studentSwitcher.innerHTML = "";

  const label = document.createElement("span");
  label.className = "parent-switcher-label";
  label.textContent = "👨‍👩‍👧 Vedi i voti di:";
  studentSwitcher.appendChild(label);

  students.forEach((s) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "parent-switcher-btn";
    btn.textContent = s.studentName || "Studente";
    btn.dataset.studentId = s.studentId;
    btn.addEventListener("click", () => selectParentStudent(s.studentId));
    studentSwitcher.appendChild(btn);
  });
}

/** Mostra i dati dello studente scelto (o dell'unico disponibile). */
function selectParentStudent(studentId) {
  const student = currentParentStudents.find((s) => s.studentId === studentId);
  if (!student) return;

  if (studentSwitcher) {
    Array.from(studentSwitcher.querySelectorAll(".parent-switcher-btn")).forEach((btn) => {
      btn.classList.toggle("is-active", btn.dataset.studentId === studentId);
    });
  }

  renderStudentHeader({ studentName: student.studentName, className: student.className });
  attachPublishedListener({ teacherUid: currentParentTeacherUid, studentId: student.studentId });
}

function renderStudentHeader(link) {
  studentHeader.innerHTML = "";
  const name = document.createElement("h2");
  name.textContent = link.studentName || "Studente";
  studentHeader.appendChild(name);
  if (link.className) {
    const cls = document.createElement("span");
    cls.className = "parent-student-class";
    cls.textContent = `Classe ${link.className}`;
    studentHeader.appendChild(cls);
  }
}

function attachPublishedListener(link) {
  detachPublishedListener();
  publishedRefPath = `/publishedGrades/${link.teacherUid}/${link.studentId}`;
  const ref = fbDb.ref(publishedRefPath);
  publishedUnsub = ref.on("value", (snapshot) => {
    const data = snapshot.val();
    renderTests(data);
  }, () => {
    testsList.innerHTML = "";
    emptyMessage.style.display = "";
    emptyMessage.textContent = "Non è stato possibile caricare i dati. Riprova più tardi.";
  });
}

function detachPublishedListener() {
  if (publishedUnsub && publishedRefPath) {
    fbDb.ref(publishedRefPath).off("value", publishedUnsub);
  }
  publishedUnsub = null;
  publishedRefPath = null;
}

function renderTests(data) {
  testsList.innerHTML = "";

  const tests = data ? Object.values(data) : [];
  if (tests.length === 0) {
    emptyMessage.style.display = "";
    emptyMessage.textContent = "Non ci sono ancora verifiche pubblicate.";
    return;
  }
  emptyMessage.style.display = "none";

  tests
    .sort((a, b) => (b.publishedAt || 0) - (a.publishedAt || 0))
    .forEach((test) => {
      testsList.appendChild(buildTestCard(test));
    });
}

function buildTestCard(test) {
  const card = document.createElement("article");
  card.className = "parent-test-card";

  const head = document.createElement("div");
  head.className = "parent-test-head";
  const title = document.createElement("h3");
  title.textContent = test.testTitle || "Verifica";
  head.appendChild(title);
  if (test.subject) {
    const subject = document.createElement("span");
    subject.className = "parent-test-subject";
    subject.textContent = test.subject;
    head.appendChild(subject);
  }
  card.appendChild(head);

  if (test.finalScore !== null && test.finalScore !== undefined && test.finalScore !== "") {
    const finalWrap = document.createElement("div");
    finalWrap.className = "parent-final-score";
    finalWrap.innerHTML = `<span class="parent-final-score-label">Voto finale</span><span class="parent-final-score-value">${escapeHtml(String(test.finalScore))}</span>`;
    card.appendChild(finalWrap);
  }

  const sections = Array.isArray(test.sections) ? test.sections : Object.values(test.sections || {});
  const visibleSections = sections.filter((s) => s && s.mode !== "hidden");
  if (visibleSections.length > 0) {
    const sectionsWrap = document.createElement("div");
    sectionsWrap.className = "parent-test-sections";
    visibleSections.forEach((section) => {
      const row = document.createElement("div");
      row.className = "parent-test-section-row mode-" + section.mode;
      const label = document.createElement("span");
      label.className = "parent-test-section-label";
      label.textContent = section.name || "Sezione";
      row.appendChild(label);

      const value = document.createElement("span");
      value.className = "parent-test-section-value";
      if (section.mode === "grade") {
        value.textContent = `${formatMaybeNumber(section.score)} / ${formatMaybeNumber(section.max)}`;
      } else {
        value.textContent = PARENT_DISPLAY_MODES[section.mode]?.label || "";
      }
      row.appendChild(value);

      sectionsWrap.appendChild(row);
    });
    card.appendChild(sectionsWrap);
  }

  const rubricEntries = Array.isArray(test.rubric) ? test.rubric : Object.values(test.rubric || {});
  if (rubricEntries.length > 0) {
    const rubricWrap = document.createElement("div");
    rubricWrap.className = "parent-test-rubric";
    const rubricTitle = document.createElement("h4");
    rubricTitle.className = "parent-test-rubric-title";
    rubricTitle.textContent = "📐 Griglia di valutazione";
    rubricWrap.appendChild(rubricTitle);

    rubricEntries.forEach((entry) => {
      const row = document.createElement("div");
      row.className = "parent-rubric-row";

      const head = document.createElement("div");
      head.className = "parent-rubric-row-head";
      const label = document.createElement("span");
      label.className = "parent-rubric-row-label";
      label.textContent = entry.section ? `${entry.section} — ${entry.name || "Voce"}` : (entry.name || "Voce");
      head.appendChild(label);
      const score = document.createElement("span");
      score.className = "parent-rubric-row-score";
      score.textContent = `${formatMaybeNumber(entry.score)} / ${formatMaybeNumber(entry.max)}`;
      head.appendChild(score);
      row.appendChild(head);

      if (entry.judgment) {
        const judgment = document.createElement("p");
        judgment.className = "parent-rubric-row-judgment";
        judgment.textContent = entry.judgment;
        row.appendChild(judgment);
      }

      rubricWrap.appendChild(row);
    });

    card.appendChild(rubricWrap);
  }

  if (test.generalComment) {
    const commentBox = document.createElement("div");
    commentBox.className = "parent-test-comment";
    commentBox.textContent = test.generalComment;
    card.appendChild(commentBox);
  }

  if (test.publishedAt) {
    const dateEl = document.createElement("div");
    dateEl.className = "parent-test-date";
    dateEl.textContent = "Pubblicato il " + formatDate(test.publishedAt);
    card.appendChild(dateEl);
  }

  return card;
}

function formatMaybeNumber(value) {
  if (value === null || value === undefined || value === "") return "–";
  return value;
}

function formatDate(timestamp) {
  try {
    return new Date(timestamp).toLocaleString("it-IT", {
      day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit",
    });
  } catch (e) {
    return "";
  }
}

function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

document.addEventListener("DOMContentLoaded", init);