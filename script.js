document.addEventListener('DOMContentLoaded', () => {
  'use strict';

  // Cache frequently accessed DOM nodes and shared UI state.
  const form = document.getElementById('charter-form');
  const statusEl = document.getElementById('form-status');
  const fillButton = document.getElementById('fill-starter-content');
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
  const jumpNavCard = document.getElementById('jump-nav-card');
  const jumpNavToggle = document.getElementById('jump-nav-toggle');
  const resetConfirm = document.getElementById('reset-confirm');
  const resetCancelBtn = document.getElementById('reset-cancel');
  const resetConfirmBtn = document.getElementById('reset-confirm-btn');
  const completionProgress = document.getElementById('completion-progress');
  const JUMP_NAV_COLLAPSE_THRESHOLD = 120;

  let docx = null;
  let initialFormSnapshot = '';
  let saveTimer = null;

  if (!form) {
    console.error('Charter form not found.');
    return;
  }

  const MEMBER_COLUMNS = ['Name', 'Title', 'Role', 'Voting Status'];
  const DEFAULT_ROLE_DEFINITION_LINES = [
    'Executive Sponsor: Provides executive support, alignment, and escalation authority.',
    'Chair: Leads meetings, sets priorities, and guides committee decisions.',
    'Members: Participate in deliberation, decision-making, and follow-through.',
    'Advisors: Provide subject matter expertise in support of the committee.'
  ];
  const FUNCTION_TO_ROLE_DEFINITIONS = {
    'Business / Program Leadership': 'Business / Program Leadership: Senior leaders representing business or program priorities.',
    'Data Ownership':
      'Data Ownership: Staff that own the purpose, use, access, and sharing expectations for the data asset.',
    'Data Stewardship':
      'Data Stewardship: Staff responsible for operational quality, definitions, and day-to-day governance practices.',
    'Data Custodian (IT / Data Platform)':
      'Data Custodian (IT / Data Platform): Technical staff supporting systems, integration, or data platforms.',
    'Information Security': 'Information Security: Staff representing security requirements and risk controls.',
    'Privacy / Confidentiality':
      'Privacy / Confidentiality: Staff representing privacy, confidentiality, or disclosure requirements.',
    'Legal / Compliance':
      'Legal / Compliance: Staff representing statutory, regulatory, contractual, or policy obligations.',
    'Records / Information Management':
      'Records / Information Management: Staff representing retention, records management, or information lifecycle practices.',
    'Analytics / BI / Reporting':
      'Analytics / BI / Reporting: Staff representing reporting, dashboards, analytics, or downstream use.',
    'Operations / Service Delivery':
      'Operations / Service Delivery: Staff representing operational processes affected by the data.',
    'External Partner / Interagency Liaison':
      'External Partner / Interagency Liaison: Staff representing partner coordination or external participation.'
  };
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
    'guiding-principles',
    'required-functions',
    'responsibilities',
    'meeting-frequency',
    'decision-making'
  ];

  // Configuration metadata used to render structured controls and helper text.
  const STRUCTURED_FIELDS = {
    'committee-type': {
      type: 'select',
      placeholder: 'Select committee type',
      noteId: 'committee-type-note',
      otherFieldId: 'committee-type-other',
      otherWrapId: 'committee-type-other-wrap',
      defaultNote: 'Choose the governance structure that best fits this body.',
      options: [
        {
          value: 'Data Governance Steering Committee',
          definition: 'Formal body with oversight and decision-making authority for governance.'
        },
        {
          value: 'Data Governance Advisory Group',
          definition: 'Provides guidance and recommendations, but does not usually hold final authority.'
        },
        {
          value: 'Data Stewardship Council',
          definition: 'Focuses on operational governance, stewardship, data quality, and standards.'
        },
        {
          value: 'Working Group',
          definition: 'Temporary or task-focused group addressing a specific governance need.'
        },
        {
          value: 'Cross-Agency Governance Group',
          definition: 'Coordinates governance across multiple agencies or departments.'
        },
        {
          value: 'Other',
          definition: 'Use when the group follows a different or custom governance model.'
        }
      ]
    },
    'agency-scope': {
      type: 'select',
      placeholder: 'Select organizational scope',
      noteId: 'agency-scope-note',
      otherFieldId: 'agency-scope-other',
      otherWrapId: 'agency-scope-other-wrap',
      defaultNote: 'Choose the organizational level this charter covers.',
      options: [
        {
          value: 'Enterprise / Statewide',
          definition: 'Applies across the full enterprise or state government.'
        },
        {
          value: 'Agency / Department',
          definition: 'Applies across one agency or department.'
        },
        {
          value: 'Division / Program / Bureau',
          definition: 'Applies to one major internal unit or program.'
        },
        {
          value: 'Cross-Agency / Interagency',
          definition: 'Covers multiple agencies or departments working together.'
        },
        {
          value: 'Project / Initiative Specific',
          definition: 'Limited to one named initiative, program, or temporary effort.'
        },
        {
          value: 'Other',
          definition: 'Use for a scope that does not fit the common models above.'
        }
      ]
    },
    'decision-authority': {
      type: 'select',
      placeholder: 'Select decision authority',
      noteId: 'decision-authority-note',
      otherFieldId: 'decision-authority-other',
      otherWrapId: 'decision-authority-other-wrap',
      defaultNote: 'Choose the authority model that best matches how this group makes or recommends decisions.',
      options: [
        {
          value: 'Advisory Only',
          definition: 'The group provides input and recommendations but does not make binding decisions.'
        },
        {
          value: 'Recommends to Executive Sponsor',
          definition: 'The group develops recommendations that require sponsor approval.'
        },
        {
          value: 'Delegated Authority Within Scope',
          definition: 'The group can make binding decisions within the authority defined in the charter.'
        },
        {
          value: 'Approves Standards and Practices',
          definition: 'The group has authority to approve governance standards, definitions, or operating practices.'
        },
        {
          value: 'Escalates Major Decisions',
          definition: 'The group can decide routine matters but escalates major or enterprise-impact decisions.'
        },
        {
          value: 'Other',
          definition: 'Use when the decision model is different or more complex.'
        }
      ]
    },
    'required-functions': {
      type: 'checkbox',
      containerId: 'required-functions-options',
      otherFieldId: 'required-functions-other',
      otherWrapId: 'required-functions-other-wrap',
      options: [
        {
          value: 'Business / Program Leadership',
          definition: 'Senior leaders representing business or program priorities.'
        },
        {
          value: 'Data Ownership',
          definition: 'Persons accountable for a data domain and its use.'
        },
        {
          value: 'Data Stewardship',
          definition: 'Persons responsible for operational quality, definitions, and day-to-day governance practices.'
        },
        {
          value: 'Data Custodian (IT / Data Platform)',
          definition: 'Technical staff supporting systems, integration, or data platforms.'
        },
        {
          value: 'Information Security',
          definition: 'Staff representing security requirements and risk controls.'
        },
        {
          value: 'Privacy / Confidentiality',
          definition: 'Staff representing privacy, confidentiality, or disclosure requirements.'
        },
        {
          value: 'Legal / Compliance',
          definition: 'Staff representing statutory, regulatory, contractual, or policy obligations.'
        },
        {
          value: 'Records / Information Management',
          definition: 'Staff representing retention, records management, or information lifecycle practices.'
        },
        {
          value: 'Analytics / BI / Reporting',
          definition: 'Staff representing reporting, dashboards, analytics, or downstream use.'
        },
        {
          value: 'Operations / Service Delivery',
          definition: 'Staff representing operational processes affected by the data.'
        },
        {
          value: 'External Partner / Interagency Liaison',
          definition: 'Staff representing partner coordination or external participation.'
        },
        {
          value: 'Other',
          definition: 'Use for a required perspective not captured above.'
        }
      ]
    },
    'meeting-frequency': {
      type: 'select',
      placeholder: 'Select meeting frequency',
      noteId: 'meeting-frequency-note',
      otherFieldId: 'meeting-frequency-other',
      otherWrapId: 'meeting-frequency-other-wrap',
      defaultNote: 'Choose how often the committee meets.',
      options: [
        { value: 'Weekly', definition: 'Meets every week.' },
        { value: 'Biweekly', definition: 'Meets every two weeks.' },
        { value: 'Monthly', definition: 'Meets once each month.' },
        { value: 'Bimonthly', definition: 'Meets every two months.' },
        { value: 'Quarterly', definition: 'Meets once each quarter.' },
        { value: 'Semiannual', definition: 'Meets twice per year.' },
        { value: 'Annual', definition: 'Meets once per year.' },
        { value: 'As Needed', definition: 'Meets based on demand rather than a fixed schedule.' },
        { value: 'Other', definition: 'Use for a different schedule.' }
      ]
    },
    quorum: {
      type: 'select',
      placeholder: 'Select quorum rule',
      noteId: 'quorum-note',
      otherFieldId: 'quorum-other',
      otherWrapId: 'quorum-other-wrap',
      defaultNote: 'Choose how quorum is established for the group.',
      options: [
        {
          value: 'Majority of voting members',
          definition: 'More than half of designated voting members must be present.'
        },
        {
          value: 'Majority of total members',
          definition: 'More than half of all committee members must be present, whether voting or not.'
        },
        {
          value: 'Fixed number',
          definition: 'A specific minimum number of members is required.'
        },
        {
          value: 'Percentage threshold',
          definition: 'A defined percentage of membership is required.'
        },
        {
          value: 'Functional representation required',
          definition: 'Quorum requires representation from specific roles, divisions, or functions.'
        },
        {
          value: 'Chair required',
          definition: 'The Chair or delegated lead must be present for quorum.'
        },
        {
          value: 'No formal quorum',
          definition: 'The group may meet and proceed without a minimum attendance requirement.'
        },
        {
          value: 'Other',
          definition: 'Use a custom quorum rule.'
        }
      ]
    },
    'decision-making': {
      type: 'select',
      placeholder: 'Select decision-making process',
      noteId: 'decision-making-note',
      otherFieldId: 'decision-making-other',
      otherWrapId: 'decision-making-other-wrap',
      defaultNote: 'Choose the method the committee uses to make decisions.',
      options: [
        {
          value: 'Consensus',
          definition: 'The group works toward agreement without a formal vote whenever possible.'
        },
        {
          value: 'Simple Majority Vote',
          definition: 'A decision passes with more than half of votes cast.'
        },
        {
          value: 'Supermajority Vote',
          definition: 'A decision passes only when a higher threshold is met, such as two-thirds.'
        },
        {
          value: 'Chair Determines After Input',
          definition: 'The chair makes the decision after hearing group input.'
        },
        {
          value: 'Sponsor Approval Required',
          definition: 'The group discusses and recommends, but final approval rests with the sponsor.'
        },
        {
          value: 'Advisory Recommendation Only',
          definition: 'The group documents recommendations for another authority to decide.'
        },
        {
          value: 'Other',
          definition: 'Use for another decision approach.'
        }
      ]
    }
  };

  const MEMBER_ROLE_OPTIONS = ['Chair', 'Member', 'Advisor'];
  const MEMBER_VOTING_OPTIONS = ['Voting', 'Non-Voting'];

  const DEFAULT_MEMBERS = [
    { name: 'Jane Doe', title: 'Chief Data Officer', role: 'Chair', voting: 'Voting' },
    { name: 'John Smith', title: 'Program Director', role: 'Member', voting: 'Voting' },
    { name: 'Mary Jones', title: 'Privacy Officer', role: 'Advisor', voting: 'Non-Voting' },
    { name: 'Alex Brown', title: 'IT Director', role: 'Member', voting: 'Voting' }
  ];

  const DEFAULTS = {
    'agency-name': 'Department of Example',
    'charter-name': 'Draft Data Governance Steering Committee',
    'committee-type': 'Data Governance Steering Committee',
    'agency-scope': 'Agency / Department',
    'executive-sponsor': 'Secretary Example Sponsor',
    'chair-lead': 'Committee Chair',
    // Getters so the date is always "today" even in long-lived sessions.
    get 'effective-date'() { return new Date().toISOString().slice(0, 10); },
    'term-review': 'Effective until revised or rescinded; reviewed annually.',
    purpose:
      'The purpose of this committee is to establish direction, accountability, and oversight for the management and use of data as a strategic asset in support of agency operations, policy, reporting, and responsible innovation.',
    vision:
      'Trusted, timely, secure, and well-understood data supports better services, decision-making, and public stewardship.',
    mission:
      'To guide agency-wide data governance through clear roles, practical standards, coordinated decision-making, and responsible access and use.',
    objectives: [
      'Promote consistent accountability for priority data assets.',
      'Improve data quality, documentation, and standardization.',
      'Support lawful, secure, and efficient data sharing and access.',
      'Resolve cross-functional data issues and decision points.',
      'Advance a practical, sustainable culture of data governance.'
    ].join('\n'),
    'success-metrics': [
      'Priority data domains have assigned owners and stewards.',
      'Core definitions and standards are documented and approved.',
      'Data issues are tracked and resolved through a defined process.',
      'Requests for data access or sharing are reviewed consistently.',
      'Governance deliverables are completed according to committee priorities.'
    ].join('\n'),
    'in-scope': [
      'Data standards, definitions, and business rules.',
      'Data quality priorities and issue resolution.',
      'Metadata, documentation, and stewardship practices.',
      'Data access, sharing, and escalation workflows.',
      'Governance priorities related to reporting, analytics, and responsible data use.'
    ].join('\n'),
    'out-of-scope': [
      'Routine system administration and platform maintenance.',
      'Project management activities outside approved governance responsibilities.',
      'Operational decisions that remain within program management authority unless escalated.'
    ].join('\n'),
    'decision-authority': 'Approves Standards and Practices',
    'escalation-path':
      'Issues that cannot be resolved by the committee, or that carry enterprise, legal, privacy, security, or significant operational impact, will be escalated through the executive sponsor and appropriate leadership channels.',
    'guiding-principles': [
      'Treat data as a strategic asset.',
      'Protect sensitive and regulated information.',
      'Promote responsible access and appropriate sharing.',
      'Standardize where practical while respecting business context.',
      'Assign clear accountability for data decisions.',
      'Use governance to enable operations, not create unnecessary burden.'
    ].join('\n'),
    'required-functions': [
      'Business / Program Leadership',
      'Data Ownership',
      'Data Stewardship',
      'Data Custodian (IT / Data Platform)',
      'Privacy / Confidentiality',
      'Legal / Compliance',
      'Analytics / BI / Reporting'
    ],
    'role-definitions': DEFAULT_ROLE_DEFINITION_LINES.join('\n'),
    responsibilities: [
      'Review and approve governance priorities, standards, and supporting guidance.',
      'Clarify ownership, stewardship, and accountability for priority data assets.',
      'Monitor governance issues, risks, and implementation progress.',
      'Resolve or escalate conflicts related to data definitions, quality, access, and use.',
      'Support practical coordination across business, technical, privacy, and legal stakeholders.'
    ].join('\n'),
    'annual-priorities': [
      'Establish a governance issue intake and tracking process.',
      'Document core data elements and definitions.',
      'Assign accountable roles for priority datasets.',
      'Create or refine standard templates for governance and sharing.'
    ].join('\n'),
    'key-deliverables': [
      'Committee charter',
      'Governance issue log',
      'Priority data glossary or data dictionary',
      'Standards, guidelines, or decision records',
      'Periodic status or progress summary'
    ].join('\n'),
    'meeting-frequency': 'Monthly',
    quorum: 'Majority of voting members',
    'decision-making': 'Consensus',
    'meeting-administration':
      'The chair or designee will prepare agendas, document decisions, maintain meeting records, and track action items and escalations.',
    'policy-alignment':
      'The committee will operate in alignment with applicable laws, regulations, statewide policy, agency policy, privacy requirements, security expectations, and records management obligations.',
    'privacy-security-considerations':
      'Privacy, security, legal, and other control functions will be engaged when governance issues involve confidential data, regulated data, release decisions, new uses of data, or elevated risk.',
    'data-sharing':
      'The committee may review or support processes related to internal sharing, external sharing, access requests, and associated agreements or approvals, consistent with agency and enterprise requirements.',
    subcommittees: 'Data Quality Working Group\nMetadata and Standards Working Group',
    get 'version-history'() { return `1.0, ${new Date().toISOString().slice(0, 10)}, System, Initial charter generated`; }
  };

  // Show/hide the "Back to top" control based on scroll depth.
  function getJumpThreshold() {
    return Math.max(360, Math.round(window.innerHeight * 0.45));
  }

  function setJumpNavCollapsed(collapsed) {
    if (!jumpNavCard || !jumpNavToggle) return;
    jumpNavCard.classList.toggle('is-collapsed', collapsed);
    jumpNavToggle.setAttribute('aria-expanded', String(!collapsed));
  }

  function updateJumpTopVisibility() {
    if (!jumpTopButton) return;
    const shouldShow = window.scrollY > getJumpThreshold();
    jumpTopButton.classList.toggle('is-visible', shouldShow);

    // Auto-collapse Jump to Section when scrolling down, re-open near top
    setJumpNavCollapsed(window.scrollY > JUMP_NAV_COLLAPSE_THRESHOLD);
  }

  function scrollToTop() {
    const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    window.scrollTo({ top: 0, behavior: reducedMotion ? 'auto' : 'smooth' });
  }

  function wireTooltipAccessibility() {
    document.querySelectorAll('.field-help').forEach((container, index) => {
      const button = container.querySelector('.field-help__button');
      const tooltip = container.querySelector('.field-help__tooltip');
      if (!button || !tooltip) return;

      const tooltipId = `field-tooltip-${index}`;
      tooltip.id = tooltipId;
      button.setAttribute('aria-describedby', tooltipId);

      // Remove trailing inline hints like "One per line …" from the aria-label
      const label = button.getAttribute('aria-label') || '';
      const cleaned = label.replace(/\s+One per line(\s+in the format[^"]*)?$/i, '').trim();
      if (cleaned !== label) button.setAttribute('aria-label', cleaned);
    });
  }

  function resolveDocx() {
    const library = window.docx;
    if (!library) return null;

    const requiredKeys = [
      'Document',
      'Paragraph',
      'TextRun',
      'Table',
      'TableRow',
      'TableCell',
      'HeadingLevel',
      'AlignmentType',
      'WidthType',
      'TableLayoutType',
      'VerticalAlign',
      'BorderStyle',
      'Packer'
    ];

    return requiredKeys.every((key) => key in library) ? library : null;
  }

  // Deferred scripts execute before DOMContentLoaded, so the library is available here.
  docx = resolveDocx();

  // Generic value helpers keep null/empty handling consistent across fields.
  function getField(id) {
    return document.getElementById(id);
  }

  function getStructuredConfig(id) {
    return STRUCTURED_FIELDS[id] || null;
  }

  function getOptionalValue(id) {
    const field = getField(id);
    if (!field || typeof field.value !== 'string') return '';
    return field.value.trim();
  }

  function getTextValue(id, fallback = '') {
    const value = getOptionalValue(id);
    return value || fallback;
  }

  function setTextValue(id, value) {
    const field = getField(id);
    if (!field) return;
    field.value = value;
  }

  function mergeUniqueLines(existingLines, linesToAdd) {
    const unique = [];
    const seen = new Set();

    [...existingLines, ...linesToAdd].forEach((line) => {
      const normalized = String(line || '').trim();
      if (!normalized) return;
      const key = normalized.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      unique.push(normalized);
    });

    return unique;
  }

  function getSuggestedRoleDefinitionLines() {
    const values = getStructuredValues('required-functions');

    return values
      .map((value) => FUNCTION_TO_ROLE_DEFINITIONS[String(value || '').trim()])
      .filter(Boolean);
  }

  function syncRoleDefinitionsFromRequiredFunctions() {
    const roleDefinitionsField = getField('role-definitions');
    if (!roleDefinitionsField) return;

    const existing = toLines(roleDefinitionsField.value);
    const suggested = getSuggestedRoleDefinitionLines();
    const merged = mergeUniqueLines(existing, suggested);

    if (merged.join('\n') !== existing.join('\n')) {
      roleDefinitionsField.value = merged.join('\n');
    }
  }

  function toLines(text) {
    return String(text || '')
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean);
  }

  function sortRoleDefinitions() {
    const roleDefinitionsField = getField('role-definitions');
    if (!roleDefinitionsField) return;

    const sortedLines = toLines(roleDefinitionsField.value).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: 'base' })
    );

    roleDefinitionsField.value = sortedLines.join('\n');
    updateHelpers();
    scheduleSave();
  }

  // Splits a comma-delimited line into exactly `expectedParts` columns.
  // Any surplus commas are folded back into the last column, so a Summary
  // field like "Initial draft, scope added" stays intact in column 4 of 4.
  function splitWithLimit(line, expectedParts) {
    const rawParts = String(line || '')
      .split(',')
      .map((part) => part.trim());

    if (rawParts.length <= expectedParts) {
      while (rawParts.length < expectedParts) rawParts.push('');
      return rawParts;
    }

    const fixed = rawParts.slice(0, expectedParts - 1);
    fixed.push(rawParts.slice(expectedParts - 1).join(', ').trim());
    return fixed;
  }

  function formatDate(dateString) {
    if (!dateString) return '';
    const date = new Date(`${dateString}T00:00:00`);
    if (Number.isNaN(date.getTime())) return dateString;

    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
  }

  function safeFileName(text) {
    const cleaned = String(text || 'Data_Governance_Charter')
      .trim()
      .replace(/[^a-z0-9]+/gi, '_')
      .replace(/^_+|_+$/g, '');
    return cleaned || 'Data_Governance_Charter';
  }

  function setStatus(message = '', state = '') {
    if (!statusEl) return;
    statusEl.textContent = message;
    statusEl.classList.remove('success', 'error');
    if (state) statusEl.classList.add(state);
  }

  function setAriaInvalid(id, invalid) {
    const field = getField(id);
    if (!field) return;
    if (invalid) {
      field.setAttribute('aria-invalid', 'true');
    } else {
      field.removeAttribute('aria-invalid');
    }
  }

  function validateForm() {
    let firstInvalid = null;

    REQUIRED_FIELD_IDS.forEach((id) => {
      const invalid = !getOptionalValue(id);
      setAriaInvalid(id, invalid);
      if (invalid && !firstInvalid) firstInvalid = getField(id);
    });

    return { isValid: !firstInvalid, firstInvalid };
  }

  function renderStructuredSelect(id, config) {
    const select = getField(id);
    if (!select) return;

    select.innerHTML = '';
    const placeholderOption = new Option(config.placeholder, '');
    placeholderOption.title = config.defaultNote || config.placeholder || '';
    select.appendChild(placeholderOption);
    config.options.forEach((option) => {
      const selectOption = new Option(option.value, option.value);
      selectOption.title = option.definition || option.value;
      selectOption.dataset.definition = option.definition || '';
      select.appendChild(selectOption);
    });
  }

  function renderStructuredCheckboxes(id, config) {
    const container = getField(config.containerId);
    if (!container) return;

    container.innerHTML = '';
    config.options.forEach((option, index) => {
      const itemId = `${id}-${index}`;
      const label = document.createElement('label');
      label.className = 'checkbox-option checkbox-option--compact';
      label.setAttribute('for', itemId);

      const input = document.createElement('input');
      input.type = 'checkbox';
      input.id = itemId;
      input.name = `${id}[]`;
      input.value = option.value;
      input.dataset.structuredField = id;

      const copy = document.createElement('span');
      copy.className = 'checkbox-option__copy checkbox-option__copy--centered';

      const title = document.createElement('span');
      title.className = 'checkbox-option__label';
      title.textContent = option.value;

      const definition = document.createElement('span');
      definition.className = 'checkbox-option__definition';
      definition.textContent = option.definition;

      copy.append(title, definition);
      label.append(input, copy);
      container.appendChild(label);
    });
  }

  function renderStructuredFields() {
    Object.entries(STRUCTURED_FIELDS).forEach(([id, config]) => {
      if (config.type === 'select') {
        renderStructuredSelect(id, config);
      }
      if (config.type === 'checkbox') {
        renderStructuredCheckboxes(id, config);
      }
    });
  }

  function getSelectDefinition(id) {
    const config = getStructuredConfig(id);
    const value = getOptionalValue(id);
    if (!config || !value) return config?.defaultNote || '';
    const option = config.options.find((item) => item.value === value);
    return option?.definition || config.defaultNote || '';
  }

  function setConditionalVisibility(id, visible) {
    const config = getStructuredConfig(id);
    const wrap = config ? getField(config.otherWrapId) : null;
    if (!wrap) return;
    wrap.hidden = !visible;
  }

  function updateStructuredFieldState(id) {
    const config = getStructuredConfig(id);
    if (!config) return;

    if (config.type === 'select') {
      const value = getOptionalValue(id);
      const note = config.noteId ? getField(config.noteId) : null;
      const field = getField(id);
      const showingOther = value === 'Other';
      setConditionalVisibility(id, showingOther);
      if (!showingOther) {
        const otherField = getField(config.otherFieldId);
        if (otherField) otherField.value = '';
      }
      if (note) {
        note.textContent = value ? getSelectDefinition(id) : config.defaultNote || '';
      }
      if (field) {
        field.title = value ? getSelectDefinition(id) : config.defaultNote || '';
      }
    }

    if (config.type === 'checkbox') {
      const values = getStructuredValues(id);
      const showingOther = values.some((value) => value === 'Other');
      setConditionalVisibility(id, showingOther);
      if (!showingOther) {
        const otherField = getField(config.otherFieldId);
        if (otherField) otherField.value = '';
      }
    }
  }

  function updateAllStructuredFieldStates() {
    Object.keys(STRUCTURED_FIELDS).forEach(updateStructuredFieldState);
  }

  function getStructuredValues(id) {
    const config = getStructuredConfig(id);
    if (!config) return [];

    if (config.type === 'select') {
      const rawValue = getOptionalValue(id);
      if (!rawValue) return [];
      return [rawValue];
    }

    if (config.type === 'checkbox') {
      return Array.from(form.querySelectorAll(`input[name="${id}[]"]:checked`)).map((input) => input.value);
    }

    return [];
  }

  function getFieldValue(id, fallback = '') {
    const config = getStructuredConfig(id);

    if (!config) {
      return getTextValue(id, fallback);
    }

    if (config.type === 'select') {
      const selected = getOptionalValue(id);
      if (!selected) return fallback;
      if (selected !== 'Other') return selected;
      const otherText = getOptionalValue(config.otherFieldId);
      return otherText || 'Other';
    }

    if (config.type === 'checkbox') {
      const selected = getStructuredValues(id);
      const hasOther = selected.includes('Other');
      const otherText = getOptionalValue(config.otherFieldId);
      const values = selected
        .filter((value) => value !== 'Other')
        .concat(hasOther && otherText ? [otherText] : hasOther ? ['Other'] : []);

      if (values.length > 0) return values;
      return Array.isArray(fallback) ? fallback : toLines(fallback);
    }

    return fallback;
  }

  function isFieldComplete(id) {
    const config = getStructuredConfig(id);

    if (!config) {
      return Boolean(getOptionalValue(id));
    }

    if (config.type === 'select') {
      const selected = getOptionalValue(id);
      if (!selected) return false;
      if (selected !== 'Other') return true;
      return Boolean(getOptionalValue(config.otherFieldId));
    }

    if (config.type === 'checkbox') {
      const selected = getStructuredValues(id);
      if (selected.length === 0) return false;
      if (!selected.includes('Other')) return true;
      return Boolean(getOptionalValue(config.otherFieldId)) || selected.some((value) => value !== 'Other');
    }

    return false;
  }

  function setFieldValue(id, value) {
    const config = getStructuredConfig(id);

    if (!config) {
      setTextValue(id, value);
      return;
    }

    if (config.type === 'select') {
      const select = getField(id);
      const otherField = getField(config.otherFieldId);
      const optionValues = new Set(config.options.map((option) => option.value));
      const normalizedValue = String(value || '').trim();

      if (!select) return;
      if (!normalizedValue) {
        select.value = '';
        if (otherField) otherField.value = '';
        updateStructuredFieldState(id);
        return;
      }

      if (optionValues.has(normalizedValue)) {
        select.value = normalizedValue;
        if (otherField) otherField.value = '';
      } else if (optionValues.has('Other')) {
        select.value = 'Other';
        if (otherField) otherField.value = normalizedValue;
      }

      updateStructuredFieldState(id);
      return;
    }

    if (config.type === 'checkbox') {
      const values = Array.isArray(value) ? value : toLines(value);
      const optionValues = new Set(config.options.map((option) => option.value));
      const otherField = getField(config.otherFieldId);
      const checkboxes = Array.from(form.querySelectorAll(`input[name="${id}[]"]`));
      checkboxes.forEach((checkbox) => {
        checkbox.checked = false;
      });

      const extras = [];
      values.forEach((item) => {
        const normalized = String(item || '').trim();
        if (!normalized) return;
        const match = checkboxes.find((checkbox) => checkbox.value === normalized);
        if (match) {
          match.checked = true;
        } else if (optionValues.has('Other')) {
          const otherCheckbox = checkboxes.find((checkbox) => checkbox.value === 'Other');
          if (otherCheckbox) otherCheckbox.checked = true;
          extras.push(normalized);
        }
      });

      if (otherField) {
        otherField.value = extras.join(', ');
      }

      updateStructuredFieldState(id);
    }
  }

  const STORAGE_KEY = 'charter-generator-v1';

  function saveToStorage() {
    try {
      localStorage.setItem(STORAGE_KEY, serializeFormState());
    } catch {
      // localStorage unavailable (private browsing, quota exceeded, etc.)
    }
  }

  function clearStorage() {
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch {}
  }

  function scheduleSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveToStorage, 800);
  }

  function applyFormState(state) {
    if (!state || !Array.isArray(state.fields)) return;

    // Group repeated checkbox keys (key ending in "[]") into arrays
    const fieldMap = new Map();
    state.fields.forEach(([key, value]) => {
      if (key.endsWith('[]')) {
        const base = key.slice(0, -2);
        if (!fieldMap.has(base)) fieldMap.set(base, []);
        fieldMap.get(base).push(value);
      } else {
        fieldMap.set(key, value);
      }
    });

    fieldMap.forEach((value, key) => setFieldValue(key, value));

    if (Array.isArray(state.members) && state.members.some(isMeaningfulMemberRow)) {
      populateMembersTable(state.members.filter(isMeaningfulMemberRow));
    }
  }

  function restoreFromStorage() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return false;
      applyFormState(JSON.parse(raw));
      return true;
    } catch {
      return false;
    }
  }

  function serializeFormState() {
    const formData = new FormData(form);
    const fields = Array.from(formData.entries())
      .map(([key, value]) => [key, String(value)])
      .sort(([a], [b]) => a.localeCompare(b));

    return JSON.stringify({
      fields,
      members: getMemberRows()
    });
  }

  function hasFormChanges() {
    return serializeFormState() !== initialFormSnapshot;
  }

  function createMemberField(config, value = '') {
    if (config.type === 'select') {
      const select = document.createElement('select');
      select.className = 'members-input members-select';
      select.setAttribute('aria-label', config.label);
      select.dataset.memberKey = config.key;
      select.appendChild(new Option(config.placeholder, ''));
      config.options.forEach((option) => {
        select.appendChild(new Option(option, option));
      });
      select.value = value || '';
      select.addEventListener('change', updateHelpers);
      return select;
    }

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'members-input';
    input.placeholder = config.placeholder;
    input.value = value || '';
    input.setAttribute('aria-label', config.label);
    input.dataset.memberKey = config.key;
    input.addEventListener('input', updateHelpers);
    return input;
  }

  function createMemberRow(data = {}) {
    const tr = document.createElement('tr');
    tr.className = 'members-row';

    const fields = [
      { key: 'name', label: 'Name', placeholder: 'e.g., Jane Doe', type: 'text' },
      { key: 'title', label: 'Title', placeholder: 'e.g., Chief Data Officer', type: 'text' },
      {
        key: 'role',
        label: 'Role',
        placeholder: 'Select role',
        type: 'select',
        options: MEMBER_ROLE_OPTIONS
      },
      {
        key: 'voting',
        label: 'Voting Status',
        placeholder: 'Select voting status',
        type: 'select',
        options: MEMBER_VOTING_OPTIONS
      }
    ];

    fields.forEach((config) => {
      const td = document.createElement('td');
      td.appendChild(createMemberField(config, data[config.key] || ''));
      tr.appendChild(td);
    });

    const tdDelete = document.createElement('td');
    tdDelete.className = 'members-cell--delete';

    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'members-delete-btn';
    deleteButton.setAttribute('aria-label', 'Remove this member');
    deleteButton.title = 'Remove row';
    deleteButton.innerHTML = '&times;';
    deleteButton.addEventListener('click', () => {
      tr.remove();
      if (!membersTbody?.querySelector('tr.members-row')) {
        initMembersTable();
      }
      updateHelpers();
      scheduleSave();
    });

    tdDelete.appendChild(deleteButton);
    tr.appendChild(tdDelete);

    return tr;
  }

  function initMembersTable() {
    if (!membersTbody) return;
    membersTbody.innerHTML = '';
    membersTbody.appendChild(createMemberRow());
  }

  function populateMembersTable(members) {
    if (!membersTbody) return;
    membersTbody.innerHTML = '';
    members.forEach((member) => {
      membersTbody.appendChild(
        createMemberRow({
          ...member,
          role: MEMBER_ROLE_OPTIONS.includes(member?.role) ? member.role : '',
          voting: MEMBER_VOTING_OPTIONS.includes(member?.voting) ? member.voting : ''
        })
      );
    });
  }

  function getMemberRows() {
    if (!membersTbody) return [];

    return Array.from(membersTbody.querySelectorAll('tr.members-row')).map((row) => {
      const fields = Array.from(row.querySelectorAll('[data-member-key]'));
      const values = Object.fromEntries(
        fields.map((field) => [field.dataset.memberKey, typeof field.value === 'string' ? field.value.trim() : ''])
      );

      return {
        name: values.name || '',
        title: values.title || '',
        role: values.role || '',
        voting: values.voting || ''
      };
    });
  }

  function isMeaningfulMemberRow(row) {
    return Boolean(row?.name || row?.title || row?.role || row?.voting);
  }

  function hasMemberData() {
    return getMemberRows().some((row) => isMeaningfulMemberRow(row));
  }

  function updateFilenamePreview() {
    if (!filenamePreview) return;
    const charterName = getTextValue('charter-name', DEFAULTS['charter-name']);
    filenamePreview.textContent = `${safeFileName(charterName)}_Charter.docx`;
  }

  function updateCompletion() {
    if (!completionText || !completionBar) return;

    const coreCompleted = PROGRESS_FIELD_IDS.filter((id) => isFieldComplete(id)).length;
    const membersCompleted = hasMemberData() ? 1 : 0;
    const total = PROGRESS_FIELD_IDS.length + 1;
    const completed = coreCompleted + membersCompleted;
    const percent = Math.round((completed / total) * 100);

    completionBar.style.width = `${percent}%`;
    if (completionProgress) completionProgress.setAttribute('aria-valuenow', String(percent));

    if (percent < 35) {
      completionText.textContent = 'Start with the charter basics and purpose sections.';
    } else if (percent < 70) {
      completionText.textContent = 'The draft is taking shape. Add membership, responsibilities, and operating details next.';
    } else if (percent < 100) {
      completionText.textContent = 'Nearly complete. Review advanced sections and export when ready.';
    } else {
      completionText.textContent = 'Core sections are complete. You are ready to generate the charter.';
    }
  }

  function updateHelpers() {
    updateFilenamePreview();
    updateCompletion();
  }

  function fillStarterContent() {
    Object.entries(DEFAULTS).forEach(([id, value]) => {
      if (!isFieldComplete(id)) {
        setFieldValue(id, value);
      }
    });

    if (!hasMemberData()) {
      populateMembersTable(DEFAULT_MEMBERS);
    }

    syncRoleDefinitionsFromRequiredFunctions();
    updateAllStructuredFieldStates();
    updateHelpers();
    setStatus('Starter content loaded. Review and customize before export.', 'success');
    scheduleSave();
  }

  // DOCX helper factories keep document construction readable and reusable.
  const DOCX_THEME = {
    font: 'Aptos',
    colors: {
      ink: '1F2933',
      subtitle: '5A6470',
      heading1: '2F5D50',
      heading2: 'A54A2A',
      white: 'FFFFFF',
      border: 'D8CEC2',
      borderLight: 'EEE5DA',
      rowAlt: 'FBF8F2',
      neutralFill: 'F7F3EC'
    },
    sizes: {
      title: 48,
      subtitle: 24,
      heading1: 30,
      heading2: 24,
      body: 21,
      label: 19,
      metadataHeader: 30,
      metadataLabel: 20,
      metadataValue: 21
    },
    spacing: {
      titleAfter: 80,
      subtitleAfter: 220,
      sectionDividerBefore: 240,
      sectionDividerAfter: 80,
      heading1Before: 0,
      heading1After: 100,
      heading2Before: 200,
      heading2After: 40,
      bodyAfter: 80,
      bodyLine: 276,
      tableCellVertical: 80,
      blankAfter: 120
    },
    sections: {
      owner: { color: 'A54A2A', tint: 'EDDFD5' },
      steward: { color: '2F5D50', tint: 'DFE1D9' },
      custodian: { color: '5A2D5C', tint: 'E4DBDB' },
      support: { color: '7A8691', tint: 'EEF1F4' }
    }
  };

  function createTextRun(text, options = {}) {
    return new docx.TextRun({
      text: String(text ?? ''),
      font: options.font || DOCX_THEME.font,
      size: options.size || DOCX_THEME.sizes.body,
      color: options.color || DOCX_THEME.colors.ink,
      bold: Boolean(options.bold),
      italics: Boolean(options.italics),
      allCaps: Boolean(options.allCaps)
    });
  }

  function getSectionTheme(sectionKey = 'support') {
    return DOCX_THEME.sections[sectionKey] || DOCX_THEME.sections.support;
  }

  function blankParagraph(after = DOCX_THEME.spacing.blankAfter) {
    return new docx.Paragraph({
      children: [createTextRun('')],
      spacing: { after }
    });
  }

  function createSectionDivider(sectionKey) {
    const sectionTheme = getSectionTheme(sectionKey);
    return new docx.Paragraph({
      border: {
        top: {
          style: docx.BorderStyle.SINGLE,
          size: 12,
          color: sectionTheme.color,
          space: 1
        }
      },
      spacing: {
        before: DOCX_THEME.spacing.sectionDividerBefore,
        after: DOCX_THEME.spacing.sectionDividerAfter
      },
      children: [createTextRun('')]
    });
  }

  function heading(text, level) {
    const isPrimaryHeading = level === docx.HeadingLevel.HEADING_1;

    return new docx.Paragraph({
      heading: level,
      keepNext: true,
      spacing: {
        before: isPrimaryHeading ? DOCX_THEME.spacing.heading1Before : DOCX_THEME.spacing.heading2Before,
        after: isPrimaryHeading ? DOCX_THEME.spacing.heading1After : DOCX_THEME.spacing.heading2After
      },
      children: [
        createTextRun(text, {
          bold: true,
          color: isPrimaryHeading ? DOCX_THEME.colors.heading1 : DOCX_THEME.colors.heading2,
          size: isPrimaryHeading ? DOCX_THEME.sizes.heading1 : DOCX_THEME.sizes.heading2
        })
      ]
    });
  }

  function bodyParagraph(text, options = {}) {
    return new docx.Paragraph({
      alignment: options.alignment,
      spacing: {
        after: options.after ?? DOCX_THEME.spacing.bodyAfter,
        line: DOCX_THEME.spacing.bodyLine
      },
      children: [
        createTextRun(text, {
          bold: Boolean(options.bold),
          italics: Boolean(options.italics),
          color: options.color || DOCX_THEME.colors.ink,
          size: options.size || DOCX_THEME.sizes.body
        })
      ]
    });
  }

  function multiParagraphs(text) {
    return toLines(text).map((line) => bodyParagraph(line));
  }

  function bulletList(items) {
    const lines = Array.isArray(items) ? items.filter(Boolean) : toLines(items);
    return lines.map(
      (item) =>
        new docx.Paragraph({
          bullet: { level: 0 },
          spacing: { after: DOCX_THEME.spacing.bodyAfter, line: DOCX_THEME.spacing.bodyLine },
          children: [createTextRun(item)]
        })
    );
  }

  function getTableBorders(accentColor) {
    return {
      top: {
        style: docx.BorderStyle.SINGLE,
        size: accentColor ? 6 : 1,
        color: accentColor || DOCX_THEME.colors.border
      },
      bottom: { style: docx.BorderStyle.SINGLE, size: 1, color: DOCX_THEME.colors.border },
      left: { style: docx.BorderStyle.SINGLE, size: 1, color: DOCX_THEME.colors.border },
      right: { style: docx.BorderStyle.SINGLE, size: 1, color: DOCX_THEME.colors.border },
      insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1, color: DOCX_THEME.colors.borderLight },
      insideVertical: { style: docx.BorderStyle.SINGLE, size: 1, color: DOCX_THEME.colors.borderLight }
    };
  }

  function tableCell(text, options = {}) {
    return new docx.TableCell({
      shading: options.shading ? { fill: options.shading } : undefined,
      verticalAlign: docx.VerticalAlign.CENTER,
      width: options.width,
      children: [
        new docx.Paragraph({
          spacing: {
            before: DOCX_THEME.spacing.tableCellVertical,
            after: DOCX_THEME.spacing.tableCellVertical,
            line: DOCX_THEME.spacing.bodyLine
          },
          children: [
            createTextRun(text, {
              bold: Boolean(options.bold),
              allCaps: Boolean(options.allCaps),
              color: options.color || DOCX_THEME.colors.ink,
              size: options.size || DOCX_THEME.sizes.body
            })
          ]
        })
      ]
    });
  }

  function createKeyValueTable(rows, sectionKey = 'support') {
    const sectionTheme = getSectionTheme(sectionKey);

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      layout: docx.TableLayoutType.FIXED,
      borders: getTableBorders(sectionTheme.color),
      rows: rows.map(
        ([label, value], index) =>
          new docx.TableRow({
            children: [
              tableCell(label, {
                width: { size: 28, type: docx.WidthType.PERCENTAGE },
                shading: sectionTheme.tint,
                bold: true,
                color: sectionTheme.color
              }),
              tableCell(value, {
                width: { size: 72, type: docx.WidthType.PERCENTAGE },
                shading: index % 2 === 0 ? 'FFFFFF' : DOCX_THEME.colors.rowAlt
              })
            ]
          })
      )
    });
  }

  function createMetadataTable(rows, sectionKey = 'owner') {
    const sectionTheme = getSectionTheme(sectionKey);

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      layout: docx.TableLayoutType.FIXED,
      borders: getTableBorders(sectionTheme.color),
      rows: [
        new docx.TableRow({
          tableHeader: true,
          children: [
            new docx.TableCell({
              columnSpan: 2,
              shading: { fill: sectionTheme.tint },
              children: [
                new docx.Paragraph({
                  alignment: docx.AlignmentType.LEFT,
                  spacing: {
                    before: DOCX_THEME.spacing.tableCellVertical,
                    after: DOCX_THEME.spacing.tableCellVertical
                  },
                  children: [
                    createTextRun('Charter Overview', {
                      bold: true,
                      size: DOCX_THEME.sizes.metadataHeader,
                      color: sectionTheme.color
                    })
                  ]
                })
              ]
            })
          ]
        }),
        ...rows.map(
          ([label, value], index) =>
            new docx.TableRow({
              children: [
                tableCell(label, {
                  width: { size: 28, type: docx.WidthType.PERCENTAGE },
                  shading: DOCX_THEME.colors.neutralFill,
                  bold: true,
                  allCaps: true,
                  color: DOCX_THEME.colors.subtitle,
                  size: DOCX_THEME.sizes.metadataLabel
                }),
                tableCell(value, {
                  width: { size: 72, type: docx.WidthType.PERCENTAGE },
                  shading: DOCX_THEME.colors.white,
                  bold: true,
                  color: DOCX_THEME.colors.ink,
                  size: DOCX_THEME.sizes.metadataValue
                })
              ]
            })
        )
      ]
    });
  }

  function createDataTable(headers, linesText, expectedParts, fallbackText, sectionKey = 'support') {
    const lines = toLines(linesText || fallbackText);
    const rows = lines.map((line) => splitWithLimit(line, expectedParts));
    const sectionTheme = getSectionTheme(sectionKey);

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      layout: docx.TableLayoutType.FIXED,
      borders: getTableBorders(sectionTheme.color),
      rows: [
        new docx.TableRow({
          tableHeader: true,
          children: headers.map((header) =>
            tableCell(header, {
              bold: true,
              shading: sectionTheme.color,
              color: DOCX_THEME.colors.white,
              size: DOCX_THEME.sizes.label
            })
          )
        }),
        ...rows.map(
          (row, index) =>
            new docx.TableRow({
              children: row.map((value) =>
                tableCell(value, { shading: index % 2 === 0 ? 'FFFFFF' : DOCX_THEME.colors.rowAlt })
              )
            })
        )
      ]
    });
  }

  function createMembersDocTable(members, sectionKey = 'steward') {
    const rows = members.length > 0 ? members : [{ name: '', title: '', role: '', voting: '' }];
    const sectionTheme = getSectionTheme(sectionKey);

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      layout: docx.TableLayoutType.FIXED,
      borders: getTableBorders(sectionTheme.color),
      rows: [
        new docx.TableRow({
          tableHeader: true,
          children: MEMBER_COLUMNS.map((column) =>
            tableCell(column, {
              bold: true,
              shading: sectionTheme.color,
              color: DOCX_THEME.colors.white,
              size: DOCX_THEME.sizes.label
            })
          )
        }),
        ...rows.map((member, index) =>
          new docx.TableRow({
            children: [member.name, member.title, member.role, member.voting].map((value) =>
              tableCell(value, { shading: index % 2 === 0 ? 'FFFFFF' : DOCX_THEME.colors.rowAlt })
            )
          })
        )
      ]
    });
  }

  function createSection(title, sectionKey, buildContent) {
    const content = buildContent();
    if (!Array.isArray(content) || content.length === 0) return [];
    return [createSectionDivider(sectionKey), heading(title, docx.HeadingLevel.HEADING_1), ...content, blankParagraph()];
  }

  function buildPurposeSection() {
    return createSection('1. Purpose', 'owner', () =>
      multiParagraphs(getTextValue('purpose', DEFAULTS.purpose))
    );
  }

  function buildVisionMissionSection() {
    return createSection('2. Vision & Mission', 'owner', () => [
      heading('Vision', docx.HeadingLevel.HEADING_2),
      ...multiParagraphs(getTextValue('vision', DEFAULTS.vision)),
      heading('Mission', docx.HeadingLevel.HEADING_2),
      ...multiParagraphs(getTextValue('mission', DEFAULTS.mission))
    ]);
  }

  function buildObjectivesSection() {
    return createSection('3. Objectives', 'owner', () =>
      bulletList(getTextValue('objectives', DEFAULTS.objectives))
    );
  }

  function buildSuccessMetricsSection() {
    return createSection('4. Success Metrics', 'owner', () =>
      bulletList(getTextValue('success-metrics', DEFAULTS['success-metrics']))
    );
  }

  function buildScopeSection() {
    return createSection('5. Scope & Authority', 'owner', () => [
      heading('In Scope', docx.HeadingLevel.HEADING_2),
      ...bulletList(getTextValue('in-scope', DEFAULTS['in-scope'])),
      heading('Out of Scope', docx.HeadingLevel.HEADING_2),
      ...bulletList(getTextValue('out-of-scope', DEFAULTS['out-of-scope'])),
      heading('Decision Authority', docx.HeadingLevel.HEADING_2),
      bodyParagraph(getFieldValue('decision-authority', DEFAULTS['decision-authority'])),
      heading('Escalation Path', docx.HeadingLevel.HEADING_2),
      ...multiParagraphs(getTextValue('escalation-path', DEFAULTS['escalation-path']))
    ]);
  }

  function buildGuidingPrinciplesSection() {
    return createSection('6. Guiding Principles', 'steward', () =>
      bulletList(getTextValue('guiding-principles', DEFAULTS['guiding-principles']))
    );
  }

  function buildMembershipSection() {
    const memberRows = getMemberRows().filter(isMeaningfulMemberRow);
    const votingMembers = memberRows.filter((row) => row.voting === 'Voting');
    const nonVotingMembers = memberRows.filter((row) => row.voting === 'Non-Voting');
    const requiredFunctions = getFieldValue('required-functions', DEFAULTS['required-functions']);

    return createSection('7. Membership & Representation', 'steward', () => [
      heading('Committee Members - Voting', docx.HeadingLevel.HEADING_2),
      createMembersDocTable(votingMembers, 'steward'),
      heading('Committee Members - Non-Voting', docx.HeadingLevel.HEADING_2),
      createMembersDocTable(nonVotingMembers, 'steward'),
      heading('Required Functions / Perspectives', docx.HeadingLevel.HEADING_2),
      ...bulletList(requiredFunctions),
      heading('Role Definitions', docx.HeadingLevel.HEADING_2),
      ...bulletList(getTextValue('role-definitions', DEFAULTS['role-definitions']))
    ]);
  }

  function buildResponsibilitiesSection() {
    return createSection('8. Responsibilities & Deliverables', 'steward', () => [
      heading('Committee Responsibilities', docx.HeadingLevel.HEADING_2),
      ...bulletList(getTextValue('responsibilities', DEFAULTS.responsibilities)),
      heading('Annual or Initial Priorities', docx.HeadingLevel.HEADING_2),
      ...bulletList(getTextValue('annual-priorities', DEFAULTS['annual-priorities'])),
      heading('Key Deliverables', docx.HeadingLevel.HEADING_2),
      ...bulletList(getTextValue('key-deliverables', DEFAULTS['key-deliverables']))
    ]);
  }

  function buildOperatingSection() {
    return createSection('9. Operating Model', 'custodian', () => {
      const operatingModelTable = createKeyValueTable(
        [
          ['Meeting Frequency', getFieldValue('meeting-frequency', DEFAULTS['meeting-frequency'])],
          ['Quorum', getFieldValue('quorum', DEFAULTS.quorum)],
          ['Decision-Making Process', getFieldValue('decision-making', DEFAULTS['decision-making'])]
        ],
        'custodian'
      );

      return [
        operatingModelTable,
        heading('Meeting Administration', docx.HeadingLevel.HEADING_2),
        ...multiParagraphs(getTextValue('meeting-administration', DEFAULTS['meeting-administration']))
      ];
    });
  }

  function buildPolicySection() {
    return createSection('10. Policy, Privacy, Security & Sharing', 'custodian', () => {
      const content = [];
      const policyAlignment = getOptionalValue('policy-alignment');
      const privacySecurity = getOptionalValue('privacy-security-considerations');
      const dataSharing = getOptionalValue('data-sharing');

      if (policyAlignment) {
        content.push(heading('Policy / Legal / Regulatory Alignment', docx.HeadingLevel.HEADING_2));
        content.push(...multiParagraphs(policyAlignment));
      }
      if (privacySecurity) {
        content.push(heading('Privacy, Security & Data Release Considerations', docx.HeadingLevel.HEADING_2));
        content.push(...multiParagraphs(privacySecurity));
      }
      if (dataSharing) {
        content.push(heading('Data Sharing & Access Considerations', docx.HeadingLevel.HEADING_2));
        content.push(...multiParagraphs(dataSharing));
      }

      return content;
    });
  }

  function buildSubcommitteesSection() {
    return createSection('11. Working Groups & Subcommittees', 'custodian', () => {
      const subcommittees = getOptionalValue('subcommittees');
      return subcommittees ? bulletList(subcommittees) : [];
    });
  }

  function buildVersionHistorySection() {
    const historyText = getOptionalValue('version-history') || DEFAULTS['version-history'];
    const versionHistoryTable = createDataTable(
      ['Version', 'Date', 'Author', 'Summary of Changes'],
      historyText,
      4,
      historyText,
      'custodian'
    );
    return createSection('12. Version History', 'custodian', () => [versionHistoryTable]);
  }

  function buildDocument() {
    const charterName = getTextValue('charter-name', DEFAULTS['charter-name']);
    const agencyName = getTextValue('agency-name', DEFAULTS['agency-name']);
    const committeeType = getFieldValue('committee-type', DEFAULTS['committee-type']);
    const agencyScope = getFieldValue('agency-scope', DEFAULTS['agency-scope']);
    const executiveSponsor = getTextValue('executive-sponsor', DEFAULTS['executive-sponsor']);
    const chairLead = getTextValue('chair-lead', DEFAULTS['chair-lead']);
    const effectiveDate = formatDate(getTextValue('effective-date', DEFAULTS['effective-date']));
    const termReview = getTextValue('term-review', DEFAULTS['term-review']);

    const metadataTable = createMetadataTable(
      [
        ['Agency / Department', agencyName],
        ['Committee Type', committeeType],
        ['Organizational Scope', agencyScope],
        ['Executive Sponsor', executiveSponsor],
        ['Chair / Lead', chairLead],
        ['Effective Date', effectiveDate],
        ['Term & Review Cycle', termReview]
      ],
      'owner'
    );

    const children = [
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: DOCX_THEME.spacing.titleAfter },
        children: [createTextRun(charterName, { bold: true, size: DOCX_THEME.sizes.title, color: DOCX_THEME.colors.ink })]
      }),
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: DOCX_THEME.spacing.subtitleAfter },
        children: [createTextRun(agencyName, { size: DOCX_THEME.sizes.subtitle, color: DOCX_THEME.colors.subtitle })]
      }),
      metadataTable,
      blankParagraph(60),
      ...buildPurposeSection(),
      ...buildVisionMissionSection(),
      ...buildObjectivesSection(),
      ...buildSuccessMetricsSection(),
      ...buildScopeSection(),
      ...buildGuidingPrinciplesSection(),
      ...buildMembershipSection(),
      ...buildResponsibilitiesSection(),
      ...buildOperatingSection(),
      ...buildPolicySection(),
      ...buildSubcommitteesSection(),
      ...buildVersionHistorySection()
    ];

    return new docx.Document({
      creator: 'Data Governance Charter Generator',
      title: charterName,
      description: 'Generated charter document for a data governance committee.',
      sections: [{ children }]
    });
  }

  form.addEventListener('input', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    if (target.id && REQUIRED_FIELD_IDS.includes(target.id)) {
      setAriaInvalid(target.id, !getOptionalValue(target.id));
    }

    if (target.id) {
      Object.entries(STRUCTURED_FIELDS).forEach(([id, config]) => {
        if (target.id === id || target.id === config.otherFieldId) {
          updateStructuredFieldState(id);
        }
      });
    }

    updateHelpers();
    scheduleSave();
  });

  form.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    if (target.id && getStructuredConfig(target.id)) {
      updateStructuredFieldState(target.id);
    }

    const structuredFieldId = target.dataset?.structuredField;
    if (structuredFieldId) {
      updateStructuredFieldState(structuredFieldId);
      if (structuredFieldId === 'required-functions') {
        syncRoleDefinitionsFromRequiredFunctions();
      }
    }

    updateHelpers();
    scheduleSave();
  });

  form.addEventListener('reset', () => {
    setTimeout(() => {
      clearStorage();
      REQUIRED_FIELD_IDS.forEach((id) => setAriaInvalid(id, false));
      setStatus('');
      initMembersTable();
      setTextValue('role-definitions', DEFAULT_ROLE_DEFINITION_LINES.join('\n'));
      updateAllStructuredFieldStates();
      updateHelpers();
    }, 0);
  });

  if (fillButton) {
    fillButton.addEventListener('click', fillStarterContent);
  }

  if (resetButton) {
    resetButton.addEventListener('click', () => {
      if (!hasFormChanges()) {
        form.reset();
        return;
      }
      if (resetConfirm) {
        resetConfirm.hidden = false;
        resetCancelBtn?.focus();
      } else {
        form.reset();
      }
    });
  }

  if (resetConfirm) {
    resetConfirm.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        resetConfirm.hidden = true;
        resetButton?.focus();
        return;
      }
      if (e.key !== 'Tab') return;
      const focusable = [resetCancelBtn, resetConfirmBtn].filter(Boolean);
      if (focusable.length < 2) return;
      if (e.shiftKey) {
        if (document.activeElement === focusable[0]) {
          e.preventDefault();
          focusable[focusable.length - 1].focus();
        }
      } else {
        if (document.activeElement === focusable[focusable.length - 1]) {
          e.preventDefault();
          focusable[0].focus();
        }
      }
    });
    resetConfirm.addEventListener('click', (e) => {
      if (e.target === resetConfirm) {
        resetConfirm.hidden = true;
        resetButton?.focus();
      }
    });
  }

  if (resetCancelBtn) {
    resetCancelBtn.addEventListener('click', () => {
      if (resetConfirm) resetConfirm.hidden = true;
      resetButton?.focus();
    });
  }

  if (resetConfirmBtn) {
    resetConfirmBtn.addEventListener('click', () => {
      if (resetConfirm) resetConfirm.hidden = true;
      form.reset();
      resetButton?.focus();
    });
  }

  if (addMemberBtn) {
    addMemberBtn.addEventListener('click', () => {
      if (!membersTbody) return;
      const row = createMemberRow();
      membersTbody.appendChild(row);
      row.querySelector('[data-member-key="name"]')?.focus();
      updateHelpers();
      scheduleSave();
    });
  }

  if (sortRoleDefinitionsBtn) {
    sortRoleDefinitionsBtn.addEventListener('click', sortRoleDefinitions);
  }

  if (jumpTopButton) {
    jumpTopButton.addEventListener('click', scrollToTop);
  }

  if (jumpNavToggle && jumpNavCard) {
    jumpNavToggle.addEventListener('click', () => {
      const isCollapsed = jumpNavCard.classList.contains('is-collapsed');
      setJumpNavCollapsed(!isCollapsed);
    });
  }

  window.addEventListener('scroll', updateJumpTopVisibility, { passive: true });
  window.addEventListener('resize', updateJumpTopVisibility, { passive: true });

  form.addEventListener('submit', async (event) => {
    event.preventDefault();

    const { isValid, firstInvalid } = validateForm();
    if (!isValid) {
      setStatus('Please complete the required charter basics before generating the document.', 'error');
      firstInvalid?.focus();
      return;
    }

    if (!docx || typeof window.saveAs !== 'function') {
      setStatus('Document libraries did not load. Refresh the page and try again.', 'error');
      return;
    }

    try {
      if (submitButton) {
        submitButton.disabled = true;
        submitButton.textContent = 'Generating...';
      }

      setStatus('Generating your charter document...', 'success');
      const documentDefinition = buildDocument();
      const blob = await docx.Packer.toBlob(documentDefinition);
      const fileName = `${safeFileName(getTextValue('charter-name', DEFAULTS['charter-name']))}_Charter.docx`;
      window.saveAs(blob, fileName);
      setStatus(`Done. Download started for ${fileName}.`, 'success');
    } catch (error) {
      console.error('Error generating charter:', error);
      setStatus('There was an error generating the charter. Check the browser console for details.', 'error');
    } finally {
      if (submitButton) {
        submitButton.disabled = false;
        submitButton.textContent = submitLabel;
      }
    }
  });

  renderStructuredFields();
  if (!getOptionalValue('role-definitions')) {
    setTextValue('role-definitions', DEFAULT_ROLE_DEFINITION_LINES.join('\n'));
  }
  initMembersTable();
  updateAllStructuredFieldStates();

  // Capture the empty-form baseline before any restoration, so hasFormChanges()
  // correctly treats restored content as "changed" relative to a blank form.
  initialFormSnapshot = serializeFormState();

  const restored = restoreFromStorage();
  if (restored) {
    syncRoleDefinitionsFromRequiredFunctions();
    updateAllStructuredFieldStates();
    setStatus('Previous session content has been restored.', 'success');
  }

  updateHelpers();
  updateJumpTopVisibility();
  wireTooltipAccessibility();
});
