let allQuestions = Array.isArray(window.PREGUNTES) ? window.PREGUNTES : [];
let quiz = [];
let index = 0;
let score = 0;
let mode = "practica";
let answers = [];
let timerInterval = null;
let startTime = null;
let starred = JSON.parse(localStorage.getItem("starredQuestions") || "[]");

const $ = (id) => document.getElementById(id);

window.addEventListener("DOMContentLoaded", init);
$("startBtn").addEventListener("click", start);
$("nextBtn").addEventListener("click", next);
$("starBtn").addEventListener("click", toggleStar);

function init() {
  if (!allQuestions.length) {
    $("loadStatus").textContent = "No s’han carregat preguntes. Comprova que preguntes.js és a la mateixa carpeta.";
    $("startBtn").disabled = true;
    return;
  }

  const temaSelect = $("tema");
  temaSelect.innerHTML = "";
  const temes = ["Tots", ...new Set(allQuestions.map(q => q.tema))];

  temes.forEach(t => {
    const opt = document.createElement("option");
    opt.value = t;
    opt.textContent = t;
    temaSelect.appendChild(opt);
  });

  $("loadStatus").textContent = `${allQuestions.length} preguntes carregades.`;
}

function start() {
  mode = $("mode").value;
  const tema = $("tema").value;
  const num = Math.max(1, parseInt($("num").value || "10", 10));
  const shuffleQ = $("shuffleQ").checked;
  const shuffleO = $("shuffleO").checked;
  const oldQ = $("oldQ").checked;
  const starredOnly = $("starredOnly").checked;

  let filtered = allQuestions.filter(q =>
    (tema === "Tots" || q.tema === tema) &&
    (oldQ || !q.convocatoria_anterior) &&
    (!starredOnly || starred.includes(q.id))
  );

  if (!filtered.length) {
    alert("No hi ha preguntes amb aquests filtres.");
    return;
  }

  if (shuffleQ) filtered = shuffleArray(filtered);
  quiz = filtered.slice(0, Math.min(num, filtered.length)).map(q => prepareQuestion(q, shuffleO));

  index = 0;
  score = 0;
  answers = [];
  startTime = Date.now();

  $("config").classList.add("hidden");
  $("result").classList.add("hidden");
  $("quiz").classList.remove("hidden");

  if (mode === "test") startTimer();
  else stopTimer();

  showQuestion();
}

function prepareQuestion(q, shuffleOptions) {
  const newQ = JSON.parse(JSON.stringify(q));

  if (shuffleOptions) {
    const paired = newQ.opcions.map((text, i) => ({ text, correct: i === newQ.correcta }));
    const shuffled = shuffleArray(paired);
    newQ.opcions = shuffled.map(p => p.text);
    newQ.correcta = shuffled.findIndex(p => p.correct);
  }

  return newQ;
}

function showQuestion() {
  const q = quiz[index];
  $("progress").textContent = `Pregunta ${index + 1} / ${quiz.length}`;
  $("question").textContent = q.pregunta;
  $("nextBtn").classList.add("hidden");

  const isStarred = starred.includes(q.id);
  $("starBtn").textContent = isStarred ? "★" : "☆";
  $("starBtn").classList.toggle("active", isStarred);

  const optionsDiv = $("options");
  optionsDiv.innerHTML = "";

  q.opcions.forEach((opt, i) => {
    const div = document.createElement("div");
    div.className = "option";
    div.textContent = `${String.fromCharCode(97 + i)}) ${opt}`;
    div.addEventListener("click", () => answer(div, i));
    optionsDiv.appendChild(div);
  });
}

function answer(selectedDiv, selectedIndex) {
  const q = quiz[index];
  const options = [...document.querySelectorAll(".option")];
  options.forEach(o => {
    o.classList.add("disabled");
    o.replaceWith(o.cloneNode(true));
  });

  const correct = selectedIndex === q.correcta;
  if (correct) score++;

  answers.push({
    id: q.id,
    pregunta: q.pregunta,
    opcions: q.opcions,
    correcta: q.correcta,
    resposta: selectedIndex,
    correct,
    tema: q.tema,
    convocatoria_anterior: q.convocatoria_anterior
  });

  const freshOptions = [...document.querySelectorAll(".option")];

  if (mode === "practica") {
    if (correct) freshOptions[selectedIndex].classList.add("correct");
    else {
      freshOptions[selectedIndex].classList.add("wrong");
      freshOptions[q.correcta].classList.add("correct");
    }
    $("nextBtn").classList.remove("hidden");
  } else {
    // Mode examen: només indiquem visualment què has triat, sense dir si és correcte.
    freshOptions[selectedIndex].classList.add("selected");
    setTimeout(next, 450);
  }
}

function next() {
  index++;
  if (index >= quiz.length) return showResult();
  showQuestion();
}

function showResult() {
  stopTimer();
  $("quiz").classList.add("hidden");
  $("result").classList.remove("hidden");

  const percent = Math.round((score / quiz.length) * 100);
  const wrong = answers.filter(a => !a.correct);
  const elapsed = startTime ? formatTime(Math.floor((Date.now() - startTime) / 1000)) : "";

  $("result").innerHTML = `
    <h2>Resultat</h2>
    <p><strong>${score} / ${quiz.length}</strong> encerts (${percent}%)</p>
    ${mode === "test" ? `<p>Temps total: ${elapsed}</p>` : ""}
    <p>Errors: ${wrong.length}</p>
    <div class="result-actions">
      <button class="primary" onclick="restart()">Nou test</button>
      ${wrong.length ? `<button onclick="reviewErrors()">Repassar errors</button>` : ""}
      ${starred.length ? `<button onclick="startStarredPractice()">Practicar marcades ⭐</button>` : ""}
    </div>
  `;
}

function restart() {
  $("result").classList.add("hidden");
  $("config").classList.remove("hidden");
}

function reviewErrors() {
  quiz = answers.filter(a => !a.correct).map(a => ({
    id: a.id,
    pregunta: a.pregunta,
    opcions: a.opcions,
    correcta: a.correcta,
    tema: a.tema,
    convocatoria_anterior: a.convocatoria_anterior,
    explicacio: ""
  }));
  index = 0;
  score = 0;
  answers = [];
  mode = "practica";
  $("result").classList.add("hidden");
  $("quiz").classList.remove("hidden");
  showQuestion();
}

function startStarredPractice() {
  quiz = allQuestions
    .filter(q => starred.includes(q.id))
    .map(q => prepareQuestion(q, true));

  if (!quiz.length) {
    alert("Encara no tens preguntes marcades amb estrella.");
    return;
  }

  quiz = shuffleArray(quiz);
  index = 0;
  score = 0;
  answers = [];
  mode = "practica";
  stopTimer();

  $("result").classList.add("hidden");
  $("config").classList.add("hidden");
  $("quiz").classList.remove("hidden");
  showQuestion();
}

function toggleStar() {
  if (!quiz.length) return;
  const id = quiz[index].id;
  if (starred.includes(id)) starred = starred.filter(x => x !== id);
  else starred.push(id);
  localStorage.setItem("starredQuestions", JSON.stringify(starred));
  showQuestion();
}

function startTimer() {
  $("timer").classList.remove("hidden");
  timerInterval = setInterval(() => {
    const seconds = Math.floor((Date.now() - startTime) / 1000);
    $("timer").textContent = `Temps: ${formatTime(seconds)}`;
  }, 1000);
}

function stopTimer() {
  if (timerInterval) clearInterval(timerInterval);
  timerInterval = null;
  $("timer").classList.add("hidden");
}

function formatTime(totalSeconds) {
  const m = Math.floor(totalSeconds / 60).toString().padStart(2, "0");
  const s = (totalSeconds % 60).toString().padStart(2, "0");
  return `${m}:${s}`;
}

function shuffleArray(array) {
  const arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}
