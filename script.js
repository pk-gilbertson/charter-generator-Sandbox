const form = document.getElementById('charter-form');
const statusEl = document.getElementById('form-status');
const fillButton = document.getElementById('fill-starter-content');
const downloadNameInput = document.getElementById('download-name');
const filenamePreview = document.getElementById('filename-preview');
const completionText = document.getElementById('completion-text');
const completionBar = document.getElementById('completion-bar');
const submitButton = form?.querySelector('button[type="submit"]');
const submitLabel = submitButton?.textContent?.trim() || 'Generate Charter (.docx)';
const membersTbody = document.getElementById('members-tbody');
const addMemberBtn = document.getElementById('add-member-row');
const sortRoleDefinitionsBtn = document.getElementById('sort-role-definitions');
const jumpTopButton = document.getElementById('jump-to-top');
const resetButton = document.getElementById('reset-form-button');

const REQUIRED_FIELD_IDS = ['agency-name', 'charter-name'];
const PROGRESS_FIELD_IDS = [
  'agency-name',
  'charter-name',
  'committee-type',
  'agency-scope',
  'purpose',
  'vision',
  'mission',
  'in-scope',
  'guiding-principles-input',
  'responsibilities-input',
  'meeting-frequency',
  'decision-making'
];

const DEFAULT_ROLE_DEFINITION_LINES = [
  'Executive Sponsor: Provides executive support, alignment, and escalation authority.',
  'Chair: Leads meetings, sets priorities, and guides committee decisions.',
  'Members: Participate in deliberation, decision-making, and follow-through.',
  'Advisors: Provide subject matter expertise in support of the committee.'
];

const FUNCTION_TO_ROLE_DEFINITIONS = {
  'Business / Program Leadership': 'Business / Program Leadership: Senior leaders representing business or program priorities.',
  'Data Ownership': 'Data Ownership: Staff that own the purpose, use, access, and sharing expectations for the data asset.',
  'Data Stewardship': 'Data Stewardship: Staff responsible for operational quality, definitions, and day-to-day governance practices.',
  'Data Custodian (IT / Data Platform)': 'Data Custodian (IT / Data Platform): Technical staff supporting systems, integration, or data platforms.',
  'Information Security': 'Information Security: Staff representing security requirements and risk controls.',
  'Privacy / Confidentiality': 'Privacy / Confidentiality: Staff representing privacy, confidentiality, or disclosure requirements.',
  'Legal / Compliance': 'Legal / Compliance: Staff representing statutory, regulatory, contractual, or policy obligations.',
  'Records / Information Management': 'Records / Information Management: Staff representing retention, records management, or information lifecycle practices.',
  'Analytics / BI / Reporting': 'Analytics / BI / Reporting: Staff representing reporting, dashboards, analytics, or downstream use.',
  'Operations / Service Delivery': 'Operations / Service Delivery: Staff representing operational processes affected by the data.',
  'External Partner / Interagency Liaison': 'External Partner / Interagency Liaison: Staff representing partner coordination or external participation.'
};

const STARTER_CONTENT = {
  'agency-name': 'Department of Health',
  'charter-name': 'Data Governance Steering Committee Charter',
  'committee-type': 'Data Governance Steering Committee',
  'agency-scope': 'Agency / Department',
  'executive-sponsor': 'Deputy Secretary of Health',
  'chair-lead': 'Director of Enterprise Data Strategy',
  'term-review-cycle': 'This charter remains in effect until retired and should be reviewed annually.',
  'purpose': 'Provide a clear forum for governance decisions related to shared departmental data, standards, access, and quality improvement.',
  'vision': 'Trusted, well-governed data that supports consistent operations, informed decisions, and responsible sharing.',
  'mission': 'Guide governance decisions, clarify ownership and stewardship expectations, and support practical coordination across business and technical teams.',
  'objectives': 'Define governance priorities and decisions\nStandardize key data definitions and practices\nSupport issue escalation and resolution\nImprove coordination across programs and technology teams',
  'success-metrics': 'Priority decisions are documented and tracked\nMeeting participation includes required perspectives\nData standards and definitions are approved and reused\nEscalated issues are resolved through a defined path',
  'in-scope': 'Shared data standards\nCross-program governance issues\nData access and sharing practices\nPriority data quality issues',
  'out-of-scope': 'Day-to-day system administration\nProject management for unrelated initiatives\nIndependent decisions that remain fully within a single program',
  'decision-authority': 'Recommends to executive sponsor',
  'escalation-path': 'Issues the committee cannot resolve are escalated to the executive sponsor with a summary of the options considered and the recommended path forward.',
  'guiding-principles-input': 'Use data responsibly and consistently\nClarify decision authority before escalation\nDocument decisions and ownership\nInclude the right business and technical perspectives',
  'role-definitions': DEFAULT_ROLE_DEFINITION_LINES.join('\n'),
  'responsibilities-input': 'Review governance issues and recommendations\nMaintain shared governance priorities\nSupport standards, definitions, and ownership decisions\nTrack actions and follow-up items',
  'priorities': 'Launch the committee and cadence\nApprove initial governance priorities\nCreate a shared decision log\nClarify data ownership and stewardship expectations',
  'deliverables': 'Approved charter\nDecision log\nShared standards or guidance\nPeriodic status updates',
  'meeting-frequency': 'Monthly',
  'quorum': 'Simple majority of voting members',
  'decision-making': 'Consensus',
  'meeting-administration': 'The chair and support staff prepare agendas, document meeting notes, track action items, and maintain governance records.',
  'policy-alignment': 'Relevant agency policies, statewide data standards, privacy requirements, and records management obligations should be reviewed as part of formal governance decisions.',
  'privacy-security': 'Privacy, security, and legal review should be involved when decisions affect restricted data, public release, or external sharing.',
  'data-sharing': 'The committee should support consistent review of internal sharing, external sharing, access requests, and data use considerations.',
  'working-groups': 'Data Standards Working Group\nAccess and Sharing Review Group',
  'version-history': '1.0, 2026-04-16, Enterprise Data Strategy Lead, Initial draft generated from charter tool'
};

let initialFormSnapshot = '';


function slugifyFileName(value) {
  const cleaned = value
    .trim()
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_');

  return cleaned || 'Data_Governance_Steering_Committee_Charter';
}

function updateDownloadNameFromCharterName() {
  const charterName = document.getElementById('charter-name')?.value || '';
  if (!downloadNameInput) return;

  const nextName = `${slugifyFileName(charterName)}.docx`;
  if (!charterName.trim()) {
    downloadNameInput.value = 'Data_Governance_Steering_Committee_Charter.docx';
  } else {
    downloadNameInput.value = nextName;
  }
  updateFilenamePreview();
}

function updateFilenamePreview() {
  if (!downloadNameInput || !filenamePreview) return;
  const value = downloadNameInput.value.trim() || 'Data_Governance_Steering_Committee_Charter.docx';
  filenamePreview.textContent = value.endsWith('.docx') ? value : `${value}.docx`;
}

function linesFromValue(value) {
  return String(value || '')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);
}

function getSelectedRequiredFunctions() {
  const container = document.getElementById('required-functions');
  if (!container) return [];
  const values = Array.from(container.querySelectorAll('input[type="checkbox"]:checked'))
    .map((input) => input.value)
    .filter((value) => value && value !== 'Other');

  const otherValue = document.getElementById('required-functions-other')?.value?.trim();
  if (document.getElementById('required-functions-other-check')?.checked && otherValue) {
    values.push(otherValue);
  }

  return values;
}

function toggleOtherInput(selectId, otherInputId) {
  const selectEl = document.getElementById(selectId);
  const otherInput = document.getElementById(otherInputId);
  if (!selectEl || !otherInput) return;

  const isOther = selectEl.value === 'Other';
  otherInput.hidden = !isOther;
  if (!isOther) {
    otherInput.value = '';
  }
}

function handleRequiredFunctionsOther() {
  const checkbox = document.getElementById('required-functions-other-check');
  const otherInput = document.getElementById('required-functions-other');
  if (!checkbox || !otherInput) return;
  otherInput.hidden = !checkbox.checked;
  if (!checkbox.checked) {
    otherInput.value = '';
  }
}

function buildRoleDefinitionsFromFunctions() {
  const textarea = document.getElementById('role-definitions');
  if (!textarea) return;

  const selectedFunctions = getSelectedRequiredFunctions();
  const currentLines = new Set(linesFromValue(textarea.value));
  if (!currentLines.size) {
    DEFAULT_ROLE_DEFINITION_LINES.forEach((line) => currentLines.add(line));
  }

  selectedFunctions.forEach((value) => {
    if (FUNCTION_TO_ROLE_DEFINITIONS[value]) {
      currentLines.add(FUNCTION_TO_ROLE_DEFINITIONS[value]);
    }
  });

  textarea.value = Array.from(currentLines).join('\n');
}

function createMemberRow(member = {}) {
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" name="member-name" value="${member.name || ''}" /></td>
    <td><input type="text" name="member-title" value="${member.title || ''}" /></td>
    <td><input type="text" name="member-role" value="${member.role || ''}" /></td>
    <td>
      <select name="member-voting">
        <option value="">Select</option>
        <option value="Voting" ${member.voting === 'Voting' ? 'selected' : ''}>Voting</option>
        <option value="Non-Voting" ${member.voting === 'Non-Voting' ? 'selected' : ''}>Non-Voting</option>
      </select>
    </td>
    <td><button type="button" class="remove-row" aria-label="Remove member row">Remove</button></td>
  `;

  tr.querySelector('.remove-row')?.addEventListener('click', () => {
    tr.remove();
    updateProgress();
  });

  tr.querySelectorAll('input, select').forEach((element) => {
    element.addEventListener('input', updateProgress);
    element.addEventListener('change', updateProgress);
  });

  return tr;
}

function seedDefaultMembers() {
  if (!membersTbody) return;
  membersTbody.innerHTML = '';
  [
    { name: 'Deputy Secretary of Health', title: 'Deputy Secretary', role: 'Executive Sponsor', voting: 'Non-Voting' },
    { name: 'Director of Enterprise Data Strategy', title: 'Director', role: 'Chair', voting: 'Voting' },
    { name: 'Program Representative', title: 'Program Manager', role: 'Member', voting: 'Voting' }
  ].forEach((member) => membersTbody.appendChild(createMemberRow(member)));
}

function serializeFormState() {
  const data = new FormData(form);
  return JSON.stringify(Array.from(data.entries()));
}

function requiredFieldsCompleteCount() {
  return REQUIRED_FIELD_IDS.reduce((count, id) => {
    const el = document.getElementById(id);
    return count + (el && String(el.value).trim() ? 1 : 0);
  }, 0);
}

function progressFieldCompleteCount() {
  return PROGRESS_FIELD_IDS.reduce((count, id) => {
    const el = document.getElementById(id);
    return count + (el && String(el.value).trim() ? 1 : 0);
  }, 0);
}

function updateProgress() {
  if (!statusEl || !completionBar || !completionText) return;

  const requiredComplete = requiredFieldsCompleteCount();
  const progressComplete = progressFieldCompleteCount();
  const progressPercent = Math.round((progressComplete / PROGRESS_FIELD_IDS.length) * 100);

  completionBar.style.width = `${progressPercent}%`;
  completionText.textContent = `${progressPercent}% complete`;

  if (requiredComplete < REQUIRED_FIELD_IDS.length) {
    statusEl.textContent = 'Start with the charter basics and purpose sections.';
  } else if (progressPercent < 40) {
    statusEl.textContent = 'Good start. Continue with scope, authority, and guiding principles.';
  } else if (progressPercent < 75) {
    statusEl.textContent = 'Your draft is taking shape. Review membership, responsibilities, and operating model details.';
  } else {
    statusEl.textContent = 'Your draft is nearly ready. Review the sections, then generate the Word document.';
  }
}

function setValue(id, value) {
  const element = document.getElementById(id);
  if (!element) return;
  element.value = value;
}

function loadStarterContent() {
  Object.entries(STARTER_CONTENT).forEach(([id, value]) => setValue(id, value));

  seedDefaultMembers();

  toggleOtherInput('committee-type', 'committee-type-other');
  toggleOtherInput('agency-scope', 'agency-scope-other');
  toggleOtherInput('decision-authority', 'decision-authority-other');
  toggleOtherInput('meeting-frequency', 'meeting-frequency-other');
  toggleOtherInput('quorum', 'quorum-other');
  toggleOtherInput('decision-making', 'decision-making-other');

  const requiredFunctionsContainer = document.getElementById('required-functions');
  if (requiredFunctionsContainer) {
    const desired = [
      'Business / Program Leadership',
      'Data Ownership',
      'Data Stewardship',
      'Data Custodian (IT / Data Platform)',
      'Information Security',
      'Privacy / Confidentiality',
      'Legal / Compliance'
    ];
    requiredFunctionsContainer.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
      checkbox.checked = desired.includes(checkbox.value);
    });
  }

  handleRequiredFunctionsOther();
  buildRoleDefinitionsFromFunctions();
  updateDownloadNameFromCharterName();
  updateProgress();
}

function resetFormToInitial() {
  form.reset();
  if (membersTbody) {
    membersTbody.innerHTML = '';
    membersTbody.appendChild(createMemberRow());
  }

  [
    ['committee-type', 'committee-type-other'],
    ['agency-scope', 'agency-scope-other'],
    ['decision-authority', 'decision-authority-other'],
    ['meeting-frequency', 'meeting-frequency-other'],
    ['quorum', 'quorum-other'],
    ['decision-making', 'decision-making-other']
  ].forEach(([selectId, otherId]) => toggleOtherInput(selectId, otherId));

  handleRequiredFunctionsOther();
  document.getElementById('role-definitions').value = DEFAULT_ROLE_DEFINITION_LINES.join('\n');
  updateDownloadNameFromCharterName();
  updateProgress();
}

function sanitizeDownloadName() {
  if (!downloadNameInput) return 'Data_Governance_Steering_Committee_Charter.docx';
  let value = downloadNameInput.value.trim() || 'Data_Governance_Steering_Committee_Charter.docx';
  if (!value.toLowerCase().endsWith('.docx')) {
    value = `${value}.docx`;
  }
  return value;
}

function getTextValue(id) {
  return document.getElementById(id)?.value?.trim() || '';
}

function getSelectValue(id, otherId) {
  const baseValue = getTextValue(id);
  if (baseValue !== 'Other') return baseValue;
  return getTextValue(otherId);
}

function memberRowsData() {
  if (!membersTbody) return [];
  return Array.from(membersTbody.querySelectorAll('tr'))
    .map((row) => {
      const inputs = row.querySelectorAll('input, select');
      return {
        name: inputs[0]?.value?.trim() || '',
        title: inputs[1]?.value?.trim() || '',
        role: inputs[2]?.value?.trim() || '',
        voting: inputs[3]?.value?.trim() || ''
      };
    })
    .filter((row) => Object.values(row).some(Boolean));
}

function appendSectionHeading(docxApi, children, text) {
  if (!text) return;
  children.push(
    new docxApi.Paragraph({
      text,
      heading: docxApi.HeadingLevel.HEADING_1,
      spacing: { before: 260, after: 120 }
    })
  );
}

function appendSubHeading(docxApi, children, text) {
  if (!text) return;
  children.push(
    new docxApi.Paragraph({
      text,
      heading: docxApi.HeadingLevel.HEADING_2,
      spacing: { before: 180, after: 60 }
    })
  );
}

function appendParagraph(docxApi, children, text) {
  if (!text) return;
  children.push(
    new docxApi.Paragraph({
      text,
      spacing: { after: 120 }
    })
  );
}

function appendBullets(docxApi, children, lines) {
  lines.forEach((line) => {
    children.push(
      new docxApi.Paragraph({
        text: line,
        bullet: { level: 0 },
        spacing: { after: 40 }
      })
    );
  });
}

async function generateDocx(event) {
  event.preventDefault();

  const docxApi = window.docx;
  if (!docxApi?.Document || !docxApi?.Packer) {
    alert('The Word export library did not load. Refresh the page and try again.');
    return;
  }

  submitButton.disabled = true;
  submitButton.textContent = 'Generating document…';

  try {
    const children = [];
    const charterName = getTextValue('charter-name') || 'Data Governance Charter';
    const committeeType = getSelectValue('committee-type', 'committee-type-other');
    const organizationalScope = getSelectValue('agency-scope', 'agency-scope-other');

    children.push(
      new docxApi.Paragraph({
        text: charterName,
        heading: docxApi.HeadingLevel.TITLE,
        spacing: { after: 220 }
      })
    );

    appendParagraph(docxApi, children, `Agency / Department: ${getTextValue('agency-name')}`);
    appendParagraph(docxApi, children, committeeType ? `Committee Type: ${committeeType}` : '');
    appendParagraph(docxApi, children, organizationalScope ? `Organizational Scope: ${organizationalScope}` : '');
    appendParagraph(docxApi, children, getTextValue('executive-sponsor') ? `Executive Sponsor: ${getTextValue('executive-sponsor')}` : '');
    appendParagraph(docxApi, children, getTextValue('chair-lead') ? `Chair / Lead: ${getTextValue('chair-lead')}` : '');
    appendParagraph(docxApi, children, getTextValue('effective-date') ? `Effective Date: ${getTextValue('effective-date')}` : '');
    appendParagraph(docxApi, children, getTextValue('term-review-cycle') ? `Term & Review Cycle: ${getTextValue('term-review-cycle')}` : '');

    appendSectionHeading(docxApi, children, 'Purpose, Vision, Mission & Outcomes');
    appendSubHeading(docxApi, children, 'Purpose');
    appendParagraph(docxApi, children, getTextValue('purpose'));
    appendSubHeading(docxApi, children, 'Vision');
    appendParagraph(docxApi, children, getTextValue('vision'));
    appendSubHeading(docxApi, children, 'Mission');
    appendParagraph(docxApi, children, getTextValue('mission'));
    appendSubHeading(docxApi, children, 'Objectives');
    appendBullets(docxApi, children, linesFromValue(getTextValue('objectives')));
    appendSubHeading(docxApi, children, 'Success Metrics');
    appendBullets(docxApi, children, linesFromValue(getTextValue('success-metrics')));

    appendSectionHeading(docxApi, children, 'Scope & Authority');
    appendSubHeading(docxApi, children, 'In-Scope Activities');
    appendBullets(docxApi, children, linesFromValue(getTextValue('in-scope')));
    appendSubHeading(docxApi, children, 'Out-of-Scope Activities');
    appendBullets(docxApi, children, linesFromValue(getTextValue('out-of-scope')));
    appendSubHeading(docxApi, children, 'Decision Authority');
    appendParagraph(docxApi, children, getSelectValue('decision-authority', 'decision-authority-other'));
    appendSubHeading(docxApi, children, 'Escalation Path');
    appendParagraph(docxApi, children, getTextValue('escalation-path'));

    appendSectionHeading(docxApi, children, 'Guiding Principles');
    appendBullets(docxApi, children, linesFromValue(getTextValue('guiding-principles-input')));

    appendSectionHeading(docxApi, children, 'Membership & Roles');
    const members = memberRowsData();
    if (members.length) {
      members.forEach((member) => {
        appendParagraph(docxApi, children, `${member.name || 'Member'} | ${member.title || ''} | ${member.role || ''} | ${member.voting || ''}`.replace(/\s+\|\s+\|\s+/g, ' | '));
      });
    } else {
      appendParagraph(docxApi, children, 'No members listed.');
    }

    appendSubHeading(docxApi, children, 'Required Functions / Perspectives');
    appendBullets(docxApi, children, getSelectedRequiredFunctions());

    appendSubHeading(docxApi, children, 'Role Definitions');
    appendBullets(docxApi, children, linesFromValue(getTextValue('role-definitions')));

    appendSectionHeading(docxApi, children, 'Responsibilities & Deliverables');
    appendSubHeading(docxApi, children, 'Committee Responsibilities');
    appendBullets(docxApi, children, linesFromValue(getTextValue('responsibilities-input')));
    appendSubHeading(docxApi, children, 'Annual or Initial Priorities');
    appendBullets(docxApi, children, linesFromValue(getTextValue('priorities')));
    appendSubHeading(docxApi, children, 'Key Deliverables');
    appendBullets(docxApi, children, linesFromValue(getTextValue('deliverables')));

    appendSectionHeading(docxApi, children, 'Operating Model');
    appendParagraph(docxApi, children, getSelectValue('meeting-frequency', 'meeting-frequency-other') ? `Meeting Frequency: ${getSelectValue('meeting-frequency', 'meeting-frequency-other')}` : '');
    appendParagraph(docxApi, children, getSelectValue('quorum', 'quorum-other') ? `Quorum: ${getSelectValue('quorum', 'quorum-other')}` : '');
    appendParagraph(docxApi, children, getSelectValue('decision-making', 'decision-making-other') ? `Decision-Making Process: ${getSelectValue('decision-making', 'decision-making-other')}` : '');
    appendParagraph(docxApi, children, getTextValue('meeting-administration') ? `Meeting Administration: ${getTextValue('meeting-administration')}` : '');

    appendSectionHeading(docxApi, children, 'Advanced Sections');
    appendSubHeading(docxApi, children, 'Policy / Legal / Regulatory Alignment');
    appendParagraph(docxApi, children, getTextValue('policy-alignment'));
    appendSubHeading(docxApi, children, 'Privacy, Security & Data Release Considerations');
    appendParagraph(docxApi, children, getTextValue('privacy-security'));
    appendSubHeading(docxApi, children, 'Data Sharing & Access Considerations');
    appendParagraph(docxApi, children, getTextValue('data-sharing'));
    appendSubHeading(docxApi, children, 'Standing Working Groups / Subcommittees');
    appendBullets(docxApi, children, linesFromValue(getTextValue('working-groups')));
    appendSubHeading(docxApi, children, 'Version History');
    appendBullets(docxApi, children, linesFromValue(getTextValue('version-history')));

    const document = new docxApi.Document({
      creator: 'Data Governance Charter Generator',
      description: 'Generated charter draft',
      sections: [
        {
          properties: {},
          children
        }
      ]
    });

    const blob = await docxApi.Packer.toBlob(document);
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = sanitizeDownloadName();
    downloadLink.click();
    setTimeout(() => URL.revokeObjectURL(downloadLink.href), 1000);
  } catch (error) {
    console.error(error);
    alert('Something went wrong while generating the document.');
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = submitLabel;
  }
}

function handleJumpTopVisibility() {
  if (!jumpTopButton) return;
  jumpTopButton.classList.toggle('is-visible', window.scrollY > 520);
}

function init() {
  if (!form) return;

  membersTbody?.appendChild(createMemberRow());
  document.getElementById('role-definitions').value = DEFAULT_ROLE_DEFINITION_LINES.join('\n');

  [
    ['committee-type', 'committee-type-other'],
    ['agency-scope', 'agency-scope-other'],
    ['decision-authority', 'decision-authority-other'],
    ['meeting-frequency', 'meeting-frequency-other'],
    ['quorum', 'quorum-other'],
    ['decision-making', 'decision-making-other']
  ].forEach(([selectId, otherId]) => {
    const select = document.getElementById(selectId);
    select?.addEventListener('change', () => {
      toggleOtherInput(selectId, otherId);
      updateProgress();
    });
  });

  document.getElementById('required-functions-other-check')?.addEventListener('change', () => {
    handleRequiredFunctionsOther();
    buildRoleDefinitionsFromFunctions();
    updateProgress();
  });

  document.querySelectorAll('#required-functions input[type="checkbox"]').forEach((checkbox) => {
    checkbox.addEventListener('change', () => {
      if (checkbox.value !== 'Other') buildRoleDefinitionsFromFunctions();
      updateProgress();
    });
  });

  addMemberBtn?.addEventListener('click', () => {
    membersTbody?.appendChild(createMemberRow());
  });

  fillButton?.addEventListener('click', loadStarterContent);
  resetButton?.addEventListener('click', resetFormToInitial);
  sortRoleDefinitionsBtn?.addEventListener('click', () => {
    const textarea = document.getElementById('role-definitions');
    if (!textarea) return;
    const sorted = linesFromValue(textarea.value).sort((a, b) => a.localeCompare(b));
    textarea.value = sorted.join('\n');
  });

  document.getElementById('charter-name')?.addEventListener('input', updateDownloadNameFromCharterName);
  downloadNameInput?.addEventListener('input', updateFilenamePreview);

  form.addEventListener('input', updateProgress);
  form.addEventListener('change', updateProgress);
  form.addEventListener('submit', generateDocx);

  jumpTopButton?.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
  window.addEventListener('scroll', handleJumpTopVisibility, { passive: true });
  handleJumpTopVisibility();

  updateDownloadNameFromCharterName();
  updateProgress();
  initialFormSnapshot = serializeFormState();
}

document.addEventListener('DOMContentLoaded', init);
