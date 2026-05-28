// ==========================================================================
// Antigravity Premium Adaptive Learning Controller (Frontend Logic Engine)
// ==========================================================================

// Global state variables
let currentJourney = null;
let currentLearningBank = null;
let currentExcelInfo = null;
let activeComplexityQuestion = null;
let activeScenarios = {}; // key: topic, value: active challenge object

// Initialize Application on Page Load
document.addEventListener('DOMContentLoaded', async () => {
  // Update timestamp
  document.getElementById('local-time-stamp').innerText = new Date().toLocaleTimeString();
  setInterval(() => {
    document.getElementById('local-time-stamp').innerText = new Date().toLocaleTimeString();
  }, 1000);

  // Load API Key from localStorage & update status
  loadSavedApiKey();
  
  // Fetch initial topics and user progress journey
  await fetchTopicsAndJourney();
});

// ----------------------------------------------------
// SETTINGS & LOCAL STORAGE
// ----------------------------------------------------
function getSavedApiKey() {
  return localStorage.getItem('GEMINI_API_KEY') || '';
}

function loadSavedApiKey() {
  const key = getSavedApiKey();
  const badge = document.getElementById('api-status-badge');
  const input = document.getElementById('gemini-key-input');
  
  if (input) {
    input.value = key;
  }
  
  if (key) {
    badge.className = 'badge-api live';
    badge.innerHTML = '<i class="fa-solid fa-bolt animate-pulse"></i> <span>Live Gemini Active</span>';
  } else {
    badge.className = 'badge-api demo';
    badge.innerHTML = '<i class="fa-solid fa-circle-nodes"></i> <span>Demo Mode (Active)</span>';
  }
}

function saveApiKey() {
  const key = document.getElementById('gemini-key-input').value.trim();
  localStorage.setItem('GEMINI_API_KEY', key);
  loadSavedApiKey();
  
  // Show notification
  alert('Configurations saved successfully! Connected in ' + (key ? 'Live Mode' : 'Demo Mode') + '.');
  switchTab('dashboard');
}

// ----------------------------------------------------
// ROUTING & NAVIGATION
// ----------------------------------------------------
function switchTab(viewId) {
  // Update active navigation item in sidebar
  const navItems = document.querySelectorAll('.nav-item');
  navItems.forEach(item => {
    item.classList.remove('active');
    // Check if link matches viewId
    if (item.getAttribute('onclick').includes(viewId)) {
      item.classList.add('active');
    }
  });

  // Switch display sections
  const sections = document.querySelectorAll('.view-section');
  sections.forEach(sec => {
    sec.classList.remove('active');
  });

  const activeSection = document.getElementById(`${viewId}-view`);
  if (activeSection) {
    activeSection.classList.add('active');
  }

  // Update Title and Subtitle dynamically
  const viewTitle = document.getElementById('view-title');
  const viewSubtitle = document.getElementById('view-subtitle');

  switch(viewId) {
    case 'dashboard':
      viewTitle.innerText = 'Dashboard';
      viewSubtitle.innerText = 'Track your adaptive learning journey and practice live engineering scenarios.';
      renderDashboard();
      break;
    case 'ai-ml-nlp':
      viewTitle.innerText = 'AI, ML & NLP Fundamentals';
      viewSubtitle.innerText = 'Study fundamental AI, Machine Learning algorithms, and Language processing principles.';
      renderAiMlNlp();
      break;
    case 'dsa':
      viewTitle.innerText = 'DSA Concept Bank & Complexity Chat';
      viewSubtitle.innerText = 'Master Big O equations, sorting algorithms, and guess dynamic code runtimes.';
      renderDsa();
      break;
    case 'system-design':
      viewTitle.innerText = 'System Design Practice Sandbox';
      viewSubtitle.innerText = 'Design resilient architectures or diagnose & repair high-traffic failure points.';
      resetScenarioState('System Design');
      break;
    case 'advanced-genai':
      viewTitle.innerText = 'Advanced GenAI Infrastructure';
      viewSubtitle.innerText = 'Build complex multi-source RAG syncs and handle vector database scale bottlenecks.';
      resetScenarioState('RAG, MCP, Vector Database, LLM, Tokens');
      break;
    case 'multi-agent':
      viewTitle.innerText = 'Multi-Agent Cycles & Coordination';
      viewSubtitle.innerText = 'Design stateful multi-agent orchestrators or resolve infinite loop tooling lock-ups.';
      resetScenarioState('Langraph, MultiAgent, Crew AI');
      break;
    case 'settings':
      viewTitle.innerText = 'System Settings';
      viewSubtitle.innerText = 'Add your Google Gemini API key to activate real-time generated quizzes and live evaluations.';
      break;
  }
}

// Quick launch buttons from Dashboard
function quickLaunch(topic) {
  if (topic === 'DSA') {
    switchTab('dsa');
    loadComplexityQuestion();
  } else if (topic === 'System Design') {
    switchTab('system-design');
    loadScenario('System Design');
  } else if (topic === 'RAG, MCP, Vector Database, LLM, Tokens') {
    switchTab('advanced-genai');
    loadScenario('RAG, MCP, Vector Database, LLM, Tokens');
  } else if (topic === 'Langraph, MultiAgent, Crew AI') {
    switchTab('multi-agent');
    loadScenario('Langraph, MultiAgent, Crew AI');
  }
}

// ----------------------------------------------------
// API REQUESTS & STATE SYNC
// ----------------------------------------------------
async function fetchTopicsAndJourney() {
  try {
    const topicsRes = await fetch('/api/topics');
    const topicsData = await topicsRes.json();
    currentLearningBank = topicsData.learningBank;
    currentExcelInfo = topicsData.excel;

    const journeyRes = await fetch('/api/journey');
    currentJourney = await journeyRes.json();

    renderDashboard();
  } catch (err) {
    console.error('Failed fetching core learning data:', err);
  }
}

// ----------------------------------------------------
// VIEW RENDERING: DASHBOARD
// ----------------------------------------------------
function renderDashboard() {
  if (!currentJourney) return;

  // Calculate percentages
  // 1. AI ML NLP Progress
  const aimlTopics = ['AI', 'ML', 'NLP'];
  let aimlLearnt = 0;
  let aimlTotal = 0;
  aimlTopics.forEach(t => {
    aimlLearnt += (currentJourney.progress[t]?.completedCards || []).length;
    aimlTotal += (currentLearningBank?.[t] || []).length;
  });
  const aimlPerc = aimlTotal > 0 ? Math.round((aimlLearnt / aimlTotal) * 100) : 0;
  updateProgressRing('aiml', aimlPerc);

  // 2. DSA Progress
  const dsaLearnt = (currentJourney.progress['DSA']?.completedCards || []).length;
  const dsaTotalCards = (currentLearningBank?.['DSA'] || []).length;
  const dsaComplexityTotal = currentJourney.progress['DSA']?.complexityTotal || 0;
  const dsaComplexityCorrect = currentJourney.progress['DSA']?.complexityCorrect || 0;
  // Combine learning cards and correct complexity guess scores to calculate total mastery
  const dsaScoreVal = dsaComplexityTotal > 0 ? (dsaComplexityCorrect / dsaComplexityTotal) * 100 : 0;
  const dsaPerc = Math.round((dsaLearnt + (dsaComplexityCorrect > 0 ? 1 : 0)) / (dsaTotalCards + 1) * 100);
  updateProgressRing('dsa', dsaPerc);
  document.getElementById('progress-text-dsa').innerText = `${dsaPerc}% complete (${dsaComplexityCorrect}/${dsaComplexityTotal} Correct Complexity)`;

  // 3. System Design (Sections 3, 4, 5)
  const designTopics = ['System Design', 'RAG, MCP, Vector Database, LLM, Tokens', 'Langraph, MultiAgent, Crew AI'];
  let designLearnt = 0;
  let designTotal = 0;
  designTopics.forEach(t => {
    designLearnt += (currentJourney.progress[t]?.completedCards || []).length;
    designTotal += (currentLearningBank?.[t] || []).length;
    
    // Add attempted scenarios to weight
    const attempted = currentJourney.progress[t]?.scenariosAttempted || 0;
    designLearnt += (attempted > 0 ? 1 : 0);
    designTotal += 1;
  });
  const designPerc = designTotal > 0 ? Math.round((designLearnt / designTotal) * 100) : 0;
  updateProgressRing('sys', designPerc);

  // Render recent activity ledger log
  const ledgerContainer = document.getElementById('journey-ledger-list');
  if (!ledgerContainer) return;

  const logs = [];
  // Gather complexity scores
  currentJourney.complexityHistory.forEach(h => {
    logs.push({
      timestamp: new Date(h.timestamp),
      title: `DSA Complexity Guess: ${h.topic}`,
      desc: h.isCorrect ? 'Correct Answer' : `Incorrect Guess (Expected ${h.correctOption})`,
      badgeClass: h.isCorrect ? 'color-success' : 'color-danger',
      icon: h.isCorrect ? 'fa-circle-check' : 'fa-circle-xmark'
    });
  });

  // Gather system design evaluations
  currentJourney.scenariosHistory.forEach(h => {
    logs.push({
      timestamp: new Date(h.timestamp),
      title: `${h.subtype === 'happy_path' ? 'Happy Path' : 'Failure Path'}: ${h.title}`,
      desc: `Grade: ${h.feedback.grade} - ${h.feedback.summary.substring(0, 70)}...`,
      badgeClass: h.feedback.grade === 'A' ? 'accent-cyan' : h.feedback.grade === 'B' ? 'accent-blue' : 'color-warning',
      icon: 'fa-server'
    });
  });

  // Sort logs by date descending
  logs.sort((a, b) => b.timestamp - a.timestamp);

  if (logs.length === 0) {
    ledgerContainer.innerHTML = '<p style="color: var(--text-muted); text-align: center; padding: 2rem;">No logged activity yet. Go practice to track your stats!</p>';
  } else {
    ledgerContainer.innerHTML = logs.slice(0, 5).map(log => `
      <div class="ledger-item">
        <div class="ledger-item-info">
          <h4>${log.title}</h4>
          <span>${log.desc}</span>
        </div>
        <div style="font-family: var(--font-mono); font-size: 0.75rem; color: var(--text-muted); display: flex; align-items: center; gap: 0.5rem;">
          <i class="fa-solid ${log.icon}" style="color: var(--${log.badgeClass})"></i>
          ${log.timestamp.toLocaleDateString()}
        </div>
      </div>
    `).join('');
  }
}

// Progress Ring helper
function updateProgressRing(id, percent) {
  const textEl = document.getElementById(`progress-text-${id}`);
  const circle = document.getElementById(`progress-circle-${id}`);
  if (textEl) {
    textEl.innerText = `${percent}% complete`;
  }
  if (circle) {
    const radius = circle.r.baseVal.value;
    const circumference = radius * 2 * Math.PI;
    const offset = circumference - (percent / 100) * circumference;
    circle.style.strokeDasharray = `${circumference} ${circumference}`;
    circle.style.strokeDashoffset = offset;
  }
}

// ----------------------------------------------------
// VIEW RENDERING: AI-ML-NLP CARDS
// ----------------------------------------------------
function renderAiMlNlp() {
  const grid = document.getElementById('ai-study-grid');
  if (!grid || !currentLearningBank) return;

  const topics = ['AI', 'ML', 'NLP'];
  let cardsHtml = '';

  topics.forEach(topicName => {
    const cards = currentLearningBank[topicName] || [];
    cards.forEach(card => {
      const isCompleted = currentJourney.progress[topicName]?.completedCards?.includes(card.id);
      
      cardsHtml += `
        <div class="flip-card-container" id="container-${card.id}">
          <div class="flip-card" onclick="toggleCardFlip('${card.id}')">
            <!-- Front Face -->
            <div class="card-face front">
              <div class="card-header">
                <span>${topicName}</span>
                <i class="fa-solid fa-brain" style="color: var(--accent-cyan);"></i>
              </div>
              <div class="card-body">
                <h3>${card.question}</h3>
              </div>
              <div class="card-footer">
                <span>Click to Flip</span>
                <button class="btn-complete ${isCompleted ? 'done' : ''}" onclick="markCardComplete(event, '${topicName}', '${card.id}')">
                  ${isCompleted ? '<i class="fa-solid fa-check"></i> Learnt' : 'Mark as Learnt'}
                </button>
              </div>
            </div>
            <!-- Back Face -->
            <div class="card-face back">
              <div class="card-header">
                <span>${topicName} Detail</span>
                <i class="fa-solid fa-graduation-cap"></i>
              </div>
              <div class="card-body">
                <p>${card.answer.replace(/\n/g, '<br>')}</p>
              </div>
              <div class="card-footer">
                <span>Click to Flip</span>
              </div>
            </div>
          </div>
        </div>
      `;
    });
  });

  grid.innerHTML = cardsHtml;
}

function toggleCardFlip(cardId) {
  const cardContainer = document.getElementById(`container-${cardId}`);
  if (cardContainer) {
    cardContainer.classList.toggle('flipped');
  }
}

async function markCardComplete(event, topic, cardId) {
  event.stopPropagation(); // Prevent flip card toggle trigger
  
  try {
    const res = await fetch('/api/journey/complete-card', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ topic, cardId })
    });
    const result = await res.json();
    currentJourney = result.db;
    
    // Play particle burst or flash
    const btn = event.target.closest('.btn-complete');
    btn.className = 'btn-complete done';
    btn.innerHTML = '<i class="fa-solid fa-check"></i> Learnt';
    
    // Refresh list and dashboard progress ratios
    setTimeout(() => {
      renderAiMlNlp();
    }, 200);
  } catch (err) {
    console.error('Failed saving card completion:', err);
  }
}

// ----------------------------------------------------
// VIEW RENDERING: DSA & COMPLEXITY QUIZ
// ----------------------------------------------------
function renderDsa() {
  const foundationsAccordion = document.getElementById('dsa-foundations-accordion');
  const patternsAccordion = document.getElementById('dsa-patterns-accordion');
  
  if (!currentLearningBank) return;

  const dsaCards = currentLearningBank['DSA'] || [];
  
  let foundationsHtml = '';
  let patternsHtml = '';

  dsaCards.forEach((card) => {
    const isCompleted = currentJourney.progress['DSA']?.completedCards?.includes(card.id);
    
    const cardHtml = `
      <div class="accordion-item" id="accordion-${card.id}">
        <div class="accordion-header" onclick="toggleAccordion('${card.id}')">
          <span style="font-family: var(--font-display); font-weight: 600; color: var(--text-bright);">
            <i class="fa-solid fa-code" style="color: var(--accent-blue); margin-right: 0.5rem;"></i>
            ${card.question}
          </span>
          <div style="display: flex; align-items: center; gap: 1rem;">
            <button class="btn-complete ${isCompleted ? 'done' : ''}" onclick="markCardComplete(event, 'DSA', ${JSON.stringify(card.id)})" style="padding: 0.25rem 0.5rem; font-size: 0.7rem;">
              ${isCompleted ? '<i class="fa-solid fa-check"></i> Studied' : 'Mark as Studied'}
            </button>
            <i class="fa-solid fa-chevron-down"></i>
          </div>
        </div>
        <div class="accordion-content">
          <div>${card.answer}</div>
        </div>
      </div>
    `;
    
    if (card.id.startsWith('dsa_found_')) {
      foundationsHtml += cardHtml;
    } else if (card.id.startsWith('dsa_pattern_')) {
      patternsHtml += cardHtml;
    }
  });

  if (foundationsAccordion) foundationsAccordion.innerHTML = foundationsHtml;
  if (patternsAccordion) patternsAccordion.innerHTML = patternsHtml;
}

function toggleAccordion(cardId) {
  const el = document.getElementById(`accordion-${cardId}`);
  if (el) {
    el.classList.toggle('active');
  }
}

// Dynamic Complexity Guessing Challenge Loader
async function loadComplexityQuestion() {
  const consoleTopic = document.getElementById('console-topic-tag');
  const codePane = document.getElementById('console-code-pane');
  const optionsContainer = document.getElementById('quiz-options-container');
  const explanationBox = document.getElementById('dsa-explanation-box');
  
  consoleTopic.innerText = 'DSA: Loading Dynamic Snippet...';
  codePane.innerText = '// Fetching algorithm complexity challenge...';
  optionsContainer.style.display = 'none';
  explanationBox.style.display = 'none';

  try {
    const res = await fetch('/api/generate-challenge', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        topic: 'DSA',
        type: 'complexity',
        apiKey: getSavedApiKey()
      })
    });
    
    const result = await res.json();
    activeComplexityQuestion = result.data;
    
    // Display code and parameters
    consoleTopic.innerText = activeComplexityQuestion.topic;
    codePane.innerText = activeComplexityQuestion.code;
    
    // Load floating buttons
    for (let i = 0; i < 4; i++) {
      const btn = document.querySelectorAll('.option-btn')[i];
      const optText = document.getElementById(`opt-text-${i}`);
      btn.className = 'option-btn';
      btn.disabled = false;
      optText.innerText = activeComplexityQuestion.options[i];
    }
    
    optionsContainer.style.display = 'grid';
  } catch (err) {
    console.error('Failed generating complexity snippet:', err);
    codePane.innerText = '// Error fetching challenge from backend server: ' + err.message;
  }
}

// Select Complexity Guess Option
async function selectComplexityOption(selectedIndex) {
  if (!activeComplexityQuestion) return;

  const correctIndex = activeComplexityQuestion.answer;
  const isCorrect = selectedIndex === correctIndex;
  
  // Highlight options
  const buttons = document.querySelectorAll('.option-btn');
  buttons.forEach((btn, index) => {
    btn.disabled = true;
    if (index === correctIndex) {
      btn.className = 'option-btn correct';
      btn.querySelector('i').className = 'fa-solid fa-circle-check';
    } else if (index === selectedIndex) {
      btn.className = 'option-btn wrong';
      btn.querySelector('i').className = 'fa-solid fa-circle-xmark';
    } else {
      btn.className = 'option-btn';
      btn.querySelector('i').className = 'fa-regular fa-circle';
    }
  });

  // Display explanation
  const explanationBox = document.getElementById('dsa-explanation-box');
  const explanationText = document.getElementById('dsa-explanation-text');
  
  explanationText.innerText = activeComplexityQuestion.explanation;
  explanationBox.style.display = 'block';
  
  // Send metrics back to server journey log
  try {
    const res = await fetch('/api/journey/update-complexity', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        isCorrect,
        selectedOption: activeComplexityQuestion.options[selectedIndex],
        correctOption: activeComplexityQuestion.options[correctIndex],
        questionTopic: activeComplexityQuestion.topic
      })
    });
    
    const result = await res.json();
    currentJourney = result.db;
  } catch (err) {
    console.error('Failed recording score in database:', err);
  }
}

// ----------------------------------------------------
// VIEW RENDERING: SCENARIO PLAYGROUNDS (SYSTEM DESIGN, GENAI, MULTIAGENT)
// ----------------------------------------------------
function getTopicPrefix(topic) {
  if (topic === 'System Design') return 'sys';
  if (topic === 'RAG, MCP, Vector Database, LLM, Tokens') return 'genai';
  return 'agent';
}

function resetScenarioState(topic) {
  const prefix = getTopicPrefix(topic);
  
  // Populate existing study flashcards for this specific topic as a concepts list
  const container = document.getElementById(`${prefix}-sandbox`);
  if (!container) return;
  
  // If activeScenario exists, keep it, otherwise show instructions
  const challenge = activeScenarios[topic];
  if (!challenge) {
    document.getElementById(`${prefix}-challenge-title`).innerText = 'Dynamic Architecture Sandbox';
    document.getElementById(`${prefix}-challenge-description`).innerText = 'Press the "Generate Scenario" button above to request a specific design task.';
    document.getElementById(`${prefix}-challenge-prompt`).style.display = 'none';
    document.getElementById(`${prefix}-type-pill`).style.display = 'none';
    document.getElementById(`${prefix}-solution-input`).value = '';
    document.getElementById(`${prefix}-solution-input`).disabled = true;
    document.getElementById(`${prefix}-submit-btn`).disabled = true;
    document.getElementById(`${prefix}-feedback-box`).style.display = 'none';
  }
}

// Generate Scenarios
async function loadScenario(topic) {
  const prefix = getTopicPrefix(topic);
  
  const titleEl = document.getElementById(`${prefix}-challenge-title`);
  const descEl = document.getElementById(`${prefix}-challenge-description`);
  const promptEl = document.getElementById(`${prefix}-challenge-prompt`);
  const pillEl = document.getElementById(`${prefix}-type-pill`);
  const inputEl = document.getElementById(`${prefix}-solution-input`);
  const submitBtn = document.getElementById(`${prefix}-submit-btn`);
  const feedbackEl = document.getElementById(`${prefix}-feedback-box`);

  titleEl.innerText = 'Requesting Architecture Scenario...';
  descEl.innerHTML = `
    <div class="spinner-container">
      <div class="spinner"></div>
      <p style="color: var(--text-muted); font-size: 0.85rem;">Communicating with Google Gemini LLM...</p>
    </div>
  `;
  promptEl.style.display = 'none';
  pillEl.style.display = 'none';
  feedbackEl.style.display = 'none';
  inputEl.disabled = true;
  submitBtn.disabled = true;

  try {
    const res = await fetch('/api/generate-challenge', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        topic,
        type: 'scenario',
        apiKey: getSavedApiKey()
      })
    });
    
    const result = await res.json();
    const challenge = result.data;
    
    // Attach topic identifier
    challenge.topic = topic;
    activeScenarios[topic] = challenge;

    // Display
    titleEl.innerText = challenge.title;
    descEl.innerText = challenge.description;
    promptEl.innerText = challenge.prompt;
    promptEl.style.display = 'block';

    // Show Path Tag
    pillEl.style.display = 'inline-block';
    if (challenge.subtype === 'happy_path') {
      pillEl.className = 'challenge-type-pill happy';
      pillEl.innerText = 'Happy Path Challenge';
    } else {
      pillEl.className = 'challenge-type-pill failure';
      pillEl.innerText = 'Failure Path Challenge';
    }

    // Enable Editor
    inputEl.value = '';
    inputEl.disabled = false;
    inputEl.placeholder = challenge.subtype === 'happy_path'
      ? "Describe your happy path layout. Include: Load Balancers, API gateways, database partitioning, query channels, cache updates..."
      : "Describe your failure resolution flow. Include: Circuit breakers, message queues, retry strategies, fallbacks, recovery rates...";
    
    submitBtn.disabled = false;
  } catch (err) {
    console.error('Failed generating scenario:', err);
    titleEl.innerText = 'Connection Error';
    descEl.innerText = 'Could not establish connection to the backend generative service: ' + err.message;
  }
}

// Submit Written Solutions
async function submitScenario(topic) {
  const prefix = getTopicPrefix(topic);
  const challenge = activeScenarios[topic];
  const answer = document.getElementById(`${prefix}-solution-input`).value.trim();
  const submitBtn = document.getElementById(`${prefix}-submit-btn`);
  const feedbackBox = document.getElementById(`${prefix}-feedback-box`);

  if (!challenge) return;
  if (!answer) {
    alert('Please write out your proposed architectural design before submitting!');
    return;
  }

  // Spinner on submit button
  submitBtn.disabled = true;
  submitBtn.innerHTML = '<i class="fa-solid fa-circle-notch animate-spin"></i> Grading Design...';
  feedbackBox.style.display = 'none';

  try {
    const res = await fetch('/api/evaluate-scenario', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        challenge,
        answer,
        apiKey: getSavedApiKey()
      })
    });
    
    const result = await res.json();
    const evaluation = result.evaluation;
    
    // Sync current journey state
    currentJourney = result.db;

    // Display grade visual
    const gradeBadge = document.getElementById(`${prefix}-feedback-grade`);
    const summaryEl = document.getElementById(`${prefix}-feedback-summary`);
    const criticalEl = document.getElementById(`${prefix}-feedback-critical`);

    gradeBadge.innerText = evaluation.grade;
    gradeBadge.className = `grade-badge ${evaluation.grade.toLowerCase()}`;
    
    summaryEl.innerText = evaluation.summary;
    criticalEl.innerHTML = `<strong>Architectural Recommendations:</strong><br>${evaluation.criticalReview.replace(/\n/g, '<br>')}`;
    
    feedbackBox.style.display = 'block';
  } catch (err) {
    console.error('Failed evaluating design solution:', err);
    alert('Server failed grading response: ' + err.message);
  } finally {
    submitBtn.disabled = false;
    submitBtn.innerHTML = '<i class="fa-solid fa-paper-plane"></i> Submit Solution';
  }
}

// Switch between Core Concepts (accordion) and Visual Cheatsheets
function switchDsaSubTab(subTab) {
  const foundationsBtn = document.getElementById('dsa-btn-foundations');
  const patternsBtn = document.getElementById('dsa-btn-patterns');
  const cheatsheetsBtn = document.getElementById('dsa-btn-cheatsheets');
  
  const foundationsContainer = document.getElementById('dsa-foundations-container');
  const patternsContainer = document.getElementById('dsa-patterns-container');
  const cheatsheetsContainer = document.getElementById('dsa-cheatsheets-container');
  
  // Reset active buttons and hide containers
  if (foundationsBtn) foundationsBtn.classList.remove('done');
  if (patternsBtn) patternsBtn.classList.remove('done');
  if (cheatsheetsBtn) cheatsheetsBtn.classList.remove('done');
  
  if (foundationsContainer) foundationsContainer.style.display = 'none';
  if (patternsContainer) patternsContainer.style.display = 'none';
  if (cheatsheetsContainer) cheatsheetsContainer.style.display = 'none';
  
  if (subTab === 'foundations') {
    if (foundationsBtn) foundationsBtn.classList.add('done');
    if (foundationsContainer) foundationsContainer.style.display = 'block';
  } else if (subTab === 'patterns') {
    if (patternsBtn) patternsBtn.classList.add('done');
    if (patternsContainer) patternsContainer.style.display = 'block';
  } else {
    if (cheatsheetsBtn) cheatsheetsBtn.classList.add('done');
    if (cheatsheetsContainer) cheatsheetsContainer.style.display = 'grid';
  }
}

// Global Fullscreen Lightbox Modal Control
function openLightbox(imgSrc, title) {
  const modal = document.getElementById('lightbox-modal');
  const img = document.getElementById('lightbox-img');
  const titleEl = document.getElementById('lightbox-title');
  const downloadEl = document.getElementById('lightbox-download');
  
  if (modal && img && titleEl && downloadEl) {
    img.src = imgSrc;
    titleEl.innerText = title;
    downloadEl.href = imgSrc;
    modal.style.display = 'flex';
    document.body.style.overflow = 'hidden'; // Lock background scrolling
  }
}

function closeLightbox() {
  const modal = document.getElementById('lightbox-modal');
  if (modal) {
    modal.style.display = 'none';
    document.body.style.overflow = 'auto'; // Unlock background scrolling
  }
}

// Global Sliding Journey Ledger Drawer Control
function openLedgerDrawer() {
  const drawer = document.getElementById('ledger-drawer');
  if (!drawer) return;
  const content = drawer.querySelector('.drawer-content');
  drawer.style.display = 'block';
  // Force a reflow to trigger slide-in transition
  void drawer.offsetWidth;
  if (content) {
    content.style.transform = 'translateX(0)';
  }
  document.body.style.overflow = 'hidden'; // Lock background scrolling
}

function closeLedgerDrawer() {
  const drawer = document.getElementById('ledger-drawer');
  if (!drawer) return;
  const content = drawer.querySelector('.drawer-content');
  if (content) {
    content.style.transform = 'translateX(100%)';
  }
  // Wait for sliding animation to finish before hiding display
  setTimeout(() => {
    drawer.style.display = 'none';
    document.body.style.overflow = 'auto'; // Unlock background scrolling
  }, 300);
}
