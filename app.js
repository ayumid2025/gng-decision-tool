// app.js – Full Go/No‑Go Decision System with Excel Export

// ==================== DATA STRUCTURES ====================

// All questions, grouped by team, with weight and factor group
const QUESTIONS = {
    finance: [
        { id: 'costAlloc', text: 'Can we charge support/administrative costs directly or through cost allocation?', weight: 'Medium', group: 'Financial' },
        { id: 'match', text: 'Can we meet donor match requirements?', weight: 'Medium', group: 'Financial' },
        { id: 'cashflow', text: 'What are the pre‑financing needs and how will they be covered?', weight: 'Medium', group: 'Financial' },
        { id: 'subrecipient', text: 'Are there subrecipients/partners? Estimated allocation?', weight: 'Medium', group: 'Financial' },
        { id: 'taxImplications', text: 'Are there tax implications for this contract?', weight: 'Medium', group: 'Financial', onlyContracts: true },
        { id: 'indemnity', text: 'Is Professional Indemnity Insurance required?', weight: 'Medium', group: 'Financial', onlyContracts: true },
        { id: 'riskPool', text: 'Has the FO contributed to the Central Risk Pool?', weight: 'Medium', group: 'Financial', onlyContracts: true },
        { id: 'riskPlan', text: 'Is there a Risk Management Plan?', weight: 'Medium', group: 'Financial', onlyLRF: true }
    ],
    supplychain: [
        { id: 'procurePlan', text: 'Is procurement planning embedded in budgeting?', weight: 'Medium', group: 'Implementation' },
        { id: 'staffCapacity', text: 'Does SCM have sufficient staff for this grant/contract?', weight: 'Medium', group: 'Implementation' },
        { id: 'regularItems', text: 'Are required items regularly purchased by the FO?', weight: 'Low', group: 'Implementation' },
        { id: 'donorExperience', text: 'Does SCM have experience with this donor’s procurement rules?', weight: 'High', group: 'Implementation' },
        { id: 'manageAll', text: 'Will WV manage all procurement?', weight: 'Low', group: 'Implementation' }
    ],
    hr: [
        { id: 'hireCapability', text: 'Can we hire/assign qualified people (including GESI, legal, compliance)?', weight: 'High', group: 'Implementation' },
        { id: 'salariesAcceptable', text: 'Are proposed salaries acceptable within FO structure?', weight: 'Medium', group: 'Implementation' }
    ],
    security: [
        { id: 'securityRisk', text: 'Are there current/potential security issues in the target area?', weight: 'High', group: 'Risk' },
        { id: 'riskManagement', text: 'Does the FO have a proven risk management approach?', weight: 'Medium', group: 'Risk' }
    ],
    technical: [
        { id: 'donorRel', text: 'Is the FO well positioned with the donor?', weight: 'High', group: 'Donor' },
        { id: 'gesi', text: 'Is team familiar with donor’s GESI requirements?', weight: 'Medium', group: 'Donor' },
        { id: 'partners', text: 'Do we have or can we quickly establish competitive partners?', weight: 'Medium', group: 'Partners' },
        { id: 'competitors', text: 'Are there strong competitors? How do we rate?', weight: 'Medium', group: 'Partners' },
        { id: 'partnerRisk', text: 'How do we assess and mitigate partnership risks?', weight: 'Medium', group: 'Partners' },
        { id: 'strategyFit', text: 'Does it fit with FO/SO strategy and CWB priorities?', weight: 'High', group: 'Alignment' },
        { id: 'oios', text: 'Does it align with OIOS initiative (impact reporting)?', weight: 'Medium', group: 'Alignment' },
        { id: 'timeFunding', text: 'Adequate time and funds to write a winning proposal?', weight: 'High', group: 'ProposalWriting' },
        { id: 'proposalTeam', text: 'Do we have a high‑quality, experienced proposal team?', weight: 'High', group: 'ProposalWriting' },
        { id: 'proposalContent', text: 'Do we have strong, relevant project models and evidence?', weight: 'High', group: 'ProposalWriting' },
        { id: 'pastPerformance', text: 'Strong past performance in similar projects? Cite evidence.', weight: 'High', group: 'Implementation' },
        { id: 'auditFindings', text: 'Any outstanding audit findings?', weight: 'High', group: 'Implementation' },
        { id: 'legalObligations', text: 'Have we reviewed the agreement template?', weight: 'Medium', group: 'Implementation' }
    ],
    risk: [
        { id: 'politicalRisk', text: 'Are there political risks?', weight: 'High', group: 'Risk' },
        { id: 'impactMeasure', text: 'Can we measure and disaggregate data (sex, age, disability)?', weight: 'Medium', group: 'Risk' },
        { id: 'paymentArrangement', text: 'What is the payment arrangement (for contracts)?', weight: 'Medium', group: 'Risk', onlyContracts: true }
    ]
};

// Helper to filter questions based on opportunity type
function getQuestionsForTeam(team, oppType) {
    let qs = QUESTIONS[team] || [];
    if (oppType === 'Grants') {
        qs = qs.filter(q => !q.onlyContracts);
    } else if (oppType === 'Contracts') {
        qs = qs.filter(q => !q.onlyLRF);
    }
    return qs;
}

// ==================== STATE MANAGEMENT ====================

let currentOpportunityId = null;
let opportunities = {};

// Load from localStorage on startup
function loadOpportunities() {
    const stored = localStorage.getItem('gngOpportunities');
    if (stored) {
        try {
            opportunities = JSON.parse(stored);
        } catch (e) {
            opportunities = {};
        }
    }
    renderSidebar();
}

// Save to localStorage
function saveOpportunities() {
    localStorage.setItem('gngOpportunities', JSON.stringify(opportunities));
}

// Render sidebar list
function renderSidebar() {
    const list = document.getElementById('oppList');
    list.innerHTML = '';
    Object.keys(opportunities).forEach(id => {
        const li = document.createElement('li');
        li.textContent = opportunities[id].name || 'Unnamed';
        li.dataset.id = id;
        li.addEventListener('click', () => selectOpportunity(id));
        if (id === currentOpportunityId) {
            li.classList.add('active');
        }
        list.appendChild(li);
    });
}

// ==================== OPPORTUNITY CREATION ====================

// Show modal to enter details
function showNewOppModal() {
    const modal = document.getElementById('oppModal');
    modal.style.display = 'flex';
    document.getElementById('oppDetailsForm').reset();
}

// Close modal
function closeModal() {
    document.getElementById('oppModal').style.display = 'none';
}

// Create new opportunity from modal data
function createOpportunity(e) {
    e.preventDefault();
    const name = document.getElementById('oppName').value.trim();
    if (!name) {
        alert('Opportunity name is required');
        return;
    }
    const id = Date.now().toString();
    opportunities[id] = {
        id: id,
        name: name,
        leadOffice: document.getElementById('leadOffice').value,
        donor: document.getElementById('donorName').value,
        oppType: document.getElementById('oppType').value,
        solicitation: document.getElementById('solicitationType').value,
        totalBudget: parseFloat(document.getElementById('totalBudget').value) || 0,
        dueDate: document.getElementById('dueDate').value,
        teams: {} // each team's responses
    };
    currentOpportunityId = id;
    saveOpportunities();
    renderSidebar();
    closeModal();
    selectOpportunity(id);
}

// ==================== NAVIGATION & UI ====================

function selectOpportunity(id) {
    currentOpportunityId = id;
    renderSidebar();
    showTeamSelector();
}

function showTeamSelector() {
    const opp = opportunities[currentOpportunityId];
    if (!opp) return;

    const container = document.getElementById('formContainer');
    container.innerHTML = `
        <h2>${opp.name}</h2>
        <p><strong>Donor:</strong> ${opp.donor || 'Not specified'} | <strong>Type:</strong> ${opp.oppType} | <strong>Budget:</strong> $${opp.totalBudget.toLocaleString()}</p>
        <p>Select a team to enter or edit ratings:</p>
        <div class="team-buttons">
            <button class="team-btn" data-team="finance">Finance</button>
            <button class="team-btn" data-team="supplychain">Supply Chain</button>
            <button class="team-btn" data-team="hr">HR</button>
            <button class="team-btn" data-team="security">Security</button>
            <button class="team-btn" data-team="technical">Technical / Programme</button>
            <button class="team-btn" data-team="risk">Risk / MEAL</button>
        </div>
        <hr>
        <button id="viewResultsBtn" class="next">View Results & Decision</button>
        <button id="editDetailsBtn" class="prev">Edit Details</button>
    `;

    document.querySelectorAll('.team-btn').forEach(btn => {
        btn.addEventListener('click', () => showTeamForm(btn.dataset.team));
    });
    document.getElementById('viewResultsBtn').addEventListener('click', showResults);
    document.getElementById('editDetailsBtn').addEventListener('click', editOpportunityDetails);
}

function editOpportunityDetails() {
    const opp = opportunities[currentOpportunityId];
    if (!opp) return;
    // Pre‑fill modal and show
    document.getElementById('oppName').value = opp.name || '';
    document.getElementById('leadOffice').value = opp.leadOffice || '';
    document.getElementById('donorName').value = opp.donor || '';
    document.getElementById('oppType').value = opp.oppType || 'Grants';
    document.getElementById('solicitationType').value = opp.solicitation || 'INT. - Call for Proposal';
    document.getElementById('totalBudget').value = opp.totalBudget || 0;
    document.getElementById('dueDate').value = opp.dueDate || '';
    const modal = document.getElementById('oppModal');
    modal.style.display = 'flex';
    // Change form submit to update instead of create
    const form = document.getElementById('oppDetailsForm');
    form.onsubmit = (e) => {
        e.preventDefault();
        opp.name = document.getElementById('oppName').value.trim();
        opp.leadOffice = document.getElementById('leadOffice').value;
        opp.donor = document.getElementById('donorName').value;
        opp.oppType = document.getElementById('oppType').value;
        opp.solicitation = document.getElementById('solicitationType').value;
        opp.totalBudget = parseFloat(document.getElementById('totalBudget').value) || 0;
        opp.dueDate = document.getElementById('dueDate').value;
        saveOpportunities();
        closeModal();
        showTeamSelector(); // refresh
    };
}

// ==================== TEAM FORM ====================

function showTeamForm(team) {
    const opp = opportunities[currentOpportunityId];
    const teamData = opp.teams[team] || {};
    const questions = getQuestionsForTeam(team, opp.oppType);

    let html = `<h3>${team.charAt(0).toUpperCase() + team.slice(1)} Team</h3>`;
    html += `<form id="teamForm">`;

    questions.forEach(q => {
        const saved = teamData[q.id] || {};
        html += `
            <div class="question">
                <label>${q.text}</label>
                <select name="${q.id}_rating" required>
                    <option value="">-- Select rating --</option>
                    <option value="Strength" ${saved.rating === 'Strength' ? 'selected' : ''}>Strength</option>
                    <option value="Neutral" ${saved.rating === 'Neutral' ? 'selected' : ''}>Neutral</option>
                    <option value="Weakness" ${saved.rating === 'Weakness' ? 'selected' : ''}>Weakness</option>
                    <option value="Must address" ${saved.rating === 'Must address' ? 'selected' : ''}>Must address</option>
                    <option value="Success not possible" ${saved.rating === 'Success not possible' ? 'selected' : ''}>Success not possible</option>
                    <option value="Don't know" ${saved.rating === "Don't know" ? 'selected' : ''}>Don't know</option>
                    <option value="Not applicable" ${saved.rating === 'Not applicable' ? 'selected' : ''}>Not applicable</option>
                </select>
                <textarea name="${q.id}_explanation" placeholder="Brief rationale, answers, actions...">${saved.explanation || ''}</textarea>
            </div>
        `;
    });

    html += `
        <button type="submit" class="save">Save Ratings</button>
        <button type="button" class="prev" onclick="showTeamSelector()">Back to Teams</button>
    </form>`;

    document.getElementById('formContainer').innerHTML = html;

    document.getElementById('teamForm').addEventListener('submit', (e) => {
        e.preventDefault();
        saveTeamForm(team);
    });
}

function saveTeamForm(team) {
    const form = document.getElementById('teamForm');
    const formData = new FormData(form);
    const teamData = {};
    const questions = getQuestionsForTeam(team, opportunities[currentOpportunityId].oppType);

    questions.forEach(q => {
        const rating = formData.get(`${q.id}_rating`);
        const explanation = formData.get(`${q.id}_explanation`) || '';
        if (rating) {
            teamData[q.id] = { rating, explanation };
        }
    });

    opportunities[currentOpportunityId].teams[team] = teamData;
    saveOpportunities();
    alert('Ratings saved!');
    showTeamSelector(); // return to team selector
}

// ==================== SCORING LOGIC ====================

function ratingToValue(rating) {
    switch (rating) {
        case 'Strength': return 1;
        case 'Neutral': return 0;
        case 'Weakness': return -1;
        case 'Must address': return -1;
        case 'Success not possible': return -1000;
        case 'Don\'t know': return -1;
        case 'Not applicable': return 0;
        default: return null;
    }
}

function weightToNumber(weight) {
    switch (weight) {
        case 'Low': return 1;
        case 'Medium': return 2;
        case 'High': return 4;
        default: return 0;
    }
}

function calculateScores(opportunity) {
    const groups = {
        Donor: { sumWeighted: 0, sumWeights: 0, count: 0, stop: false },
        Partners: { sumWeighted: 0, sumWeights: 0, count: 0, stop: false },
        Alignment: { sumWeighted: 0, sumWeights: 0, count: 0, stop: false },
        ProposalWriting: { sumWeighted: 0, sumWeights: 0, count: 0, stop: false },
        Implementation: { sumWeighted: 0, sumWeights: 0, count: 0, stop: false },
        Financial: { sumWeighted: 0, sumWeights: 0, count: 0, stop: false },
        Risk: { sumWeighted: 0, sumWeights: 0, count: 0, stop: false }
    };

    // Iterate over all teams
    Object.keys(QUESTIONS).forEach(team => {
        const teamQuestions = QUESTIONS[team];
        const teamData = opportunity.teams[team] || {};
        teamQuestions.forEach(q => {
            // Skip if not applicable to opportunity type
            if (opportunity.oppType === 'Grants' && q.onlyContracts) return;
            if (opportunity.oppType === 'Contracts' && q.onlyLRF) return;

            const response = teamData[q.id];
            if (!response || !response.rating) return; // not rated

            const value = ratingToValue(response.rating);
            const weight = weightToNumber(q.weight);
            const group = q.group;

            if (value === -1000) {
                groups[group].stop = true;
            }

            groups[group].sumWeighted += value * weight;
            groups[group].sumWeights += weight;
            groups[group].count++;
        });
    });

    // Calculate average per group
    const groupScores = {};
    let totalWeightedScore = 0;
    let totalWeights = 0;
    let anyStop = false;

    Object.keys(groups).forEach(group => {
        const g = groups[group];
        if (g.stop) {
            groupScores[group] = 'STOP';
            anyStop = true;
        } else if (g.sumWeights > 0) {
            groupScores[group] = g.sumWeighted / g.sumWeights;
            totalWeightedScore += groupScores[group] * g.sumWeights;
            totalWeights += g.sumWeights;
        } else {
            groupScores[group] = 'No Rating Yet';
        }
    });

    const overall = anyStop ? 'STOP' : (totalWeights > 0 ? totalWeightedScore / totalWeights : 'No Rating Yet');

    return { groupScores, overall, anyStop };
}

// ==================== RESULTS DISPLAY ====================

function showResults() {
    const opp = opportunities[currentOpportunityId];
    const scores = calculateScores(opp);
    const container = document.getElementById('formContainer');

    let html = `<h2>${opp.name} – Results</h2>`;

    if (scores.anyStop) {
        html += `<p style="color:red; font-weight:bold;">⚠️ STOP – One or more "Success not possible" ratings.</p>`;
    }

    html += `<table class="results-table">
        <tr><th>Factor Group</th><th>Score</th><th>Rating</th></tr>`;

    Object.keys(scores.groupScores).forEach(group => {
        const score = scores.groupScores[group];
        let ratingText = '';
        if (score === 'STOP') ratingText = 'STOP';
        else if (score === 'No Rating Yet') ratingText = 'No Rating Yet';
        else {
            if (score > 0.6) ratingText = 'Very Strong';
            else if (score > 0.2) ratingText = 'Strong';
            else if (score > -0.2) ratingText = 'Neutral';
            else if (score > -0.6) ratingText = 'Weak';
            else ratingText = 'Very Weak';
        }
        html += `<tr><td>${group}</td><td>${score === 'No Rating Yet' ? '-' : score.toFixed(2)}</td><td>${ratingText}</td></tr>`;
    });

    let overallRating = '';
    if (scores.overall === 'STOP') overallRating = 'STOP';
    else if (scores.overall === 'No Rating Yet') overallRating = 'Incomplete';
    else {
        if (scores.overall > 0.6) overallRating = 'Very Strong';
        else if (scores.overall > 0.2) overallRating = 'Strong';
        else if (scores.overall > -0.2) overallRating = 'Neutral';
        else if (scores.overall > -0.6) overallRating = 'Weak';
        else overallRating = 'Very Weak';
    }

    html += `<tr><th>Overall</th><th>${typeof scores.overall === 'number' ? scores.overall.toFixed(2) : scores.overall}</th><th>${overallRating}</th></tr>`;
    html += `</table>`;

    html += `<button onclick="showTeamSelector()" class="prev">Back to Teams</button>`;
    html += `<button onclick="exportToExcel()" class="next" style="margin-left:10px;">📥 Export to Excel</button>`;

    container.innerHTML = html;
}

// ==================== EXPORT TO EXCEL ====================

function exportToExcel() {
    const opp = opportunities[currentOpportunityId];
    if (!opp) return;

    // 1. Prepare data for each sheet
    const sheets = {};

    // --- Sheet 1: Opportunity Details ---
    sheets['Details'] = XLSX.utils.json_to_sheet([
        { Field: 'Opportunity Name', Value: opp.name },
        { Field: 'Lead Implementing Office', Value: opp.leadOffice || '' },
        { Field: 'Donor', Value: opp.donor || '' },
        { Field: 'Opportunity Type', Value: opp.oppType },
        { Field: 'Solicitation Type', Value: opp.solicitation },
        { Field: 'Total Budget (USD)', Value: opp.totalBudget },
        { Field: 'Proposal Due Date', Value: opp.dueDate || '' }
    ]);

    // --- Sheet 2..n: Team Ratings ---
    const teams = ['finance', 'supplychain', 'hr', 'security', 'technical', 'risk'];
    teams.forEach(team => {
        const teamData = opp.teams[team] || {};
        const questions = getQuestionsForTeam(team, opp.oppType);
        if (questions.length === 0) return;

        const rows = [];
        questions.forEach(q => {
            const resp = teamData[q.id] || {};
            rows.push({
                Question: q.text,
                Rating: resp.rating || '',
                Explanation: resp.explanation || '',
                Weight: q.weight,
                FactorGroup: q.group
            });
        });
        if (rows.length > 0) {
            sheets[`Team ${team}`] = XLSX.utils.json_to_sheet(rows);
        }
    });

    // --- Sheet: Scores & Results ---
    const scores = calculateScores(opp);
    const scoreRows = [];
    Object.keys(scores.groupScores).forEach(group => {
        const score = scores.groupScores[group];
        let ratingText = '';
        if (score === 'STOP') ratingText = 'STOP';
        else if (score === 'No Rating Yet') ratingText = 'No Rating Yet';
        else {
            if (score > 0.6) ratingText = 'Very Strong';
            else if (score > 0.2) ratingText = 'Strong';
            else if (score > -0.2) ratingText = 'Neutral';
            else if (score > -0.6) ratingText = 'Weak';
            else ratingText = 'Very Weak';
        }
        scoreRows.push({
            FactorGroup: group,
            Score: (typeof score === 'number') ? score.toFixed(2) : score,
            Rating: ratingText
        });
    });
    // Overall
    let overallRating = '';
    if (scores.overall === 'STOP') overallRating = 'STOP';
    else if (scores.overall === 'No Rating Yet') overallRating = 'Incomplete';
    else {
        if (scores.overall > 0.6) overallRating = 'Very Strong';
        else if (scores.overall > 0.2) overallRating = 'Strong';
        else if (scores.overall > -0.2) overallRating = 'Neutral';
        else if (scores.overall > -0.6) overallRating = 'Weak';
        else overallRating = 'Very Weak';
    }
    scoreRows.push({
        FactorGroup: 'OVERALL',
        Score: (typeof scores.overall === 'number') ? scores.overall.toFixed(2) : scores.overall,
        Rating: overallRating
    });
    sheets['Scores'] = XLSX.utils.json_to_sheet(scoreRows);

    // 2. Create a new workbook and add all sheets
    const wb = XLSX.utils.book_new();
    Object.keys(sheets).forEach(sheetName => {
        XLSX.utils.book_append_sheet(wb, sheets[sheetName], sheetName);
    });

    // 3. Generate Excel file and trigger download
    const fileName = opp.name.replace(/\s+/g, '_') + '_GNG.xlsx';
    XLSX.writeFile(wb, fileName);
}

// ==================== INITIALISATION ====================

document.addEventListener('DOMContentLoaded', () => {
    loadOpportunities();

    // Modal close events
    document.querySelector('.close').addEventListener('click', closeModal);
    window.addEventListener('click', (e) => {
        if (e.target.classList.contains('modal')) closeModal();
    });

    document.getElementById('newOppBtn').addEventListener('click', () => {
        // Reset form for new opportunity
        document.getElementById('oppDetailsForm').reset();
        document.getElementById('oppDetailsForm').onsubmit = createOpportunity;
        showNewOppModal();
    });

    // If no opportunities, optionally create a sample
    if (Object.keys(opportunities).length === 0) {
        // Optionally create a demo opportunity
        // For now, just show empty sidebar
    } else {
        selectOpportunity(Object.keys(opportunities)[0]);
    }
});

// Expose functions to global for inline onclick handlers
window.showTeamSelector = showTeamSelector;
window.showResults = showResults;
window.exportToExcel = exportToExcel;
window.editOpportunityDetails = editOpportunityDetails; // if needed
