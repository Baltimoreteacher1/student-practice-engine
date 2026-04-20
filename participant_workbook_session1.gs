/**
 * AI for Teachers: Participant Workbook (Session 1)
 * Access Without Lowering Demand
 *
 * SETUP:
 * 1. Optional: paste an existing Google Doc URL or ID into
 *    CONFIG.templateDocIdOrUrl.
 * 2. Leave CONFIG.templateDocIdOrUrl blank if you want the script to create
 *    a fresh workbook document for you.
 * 3. The script can create the workbook document and missing tabs
 *    automatically. If your Apps Script project uses a custom Google Cloud
 *    project, make sure the Google Docs API is enabled there.
 * 4. Run run().
 *
 * The script writes directly into the target document and ensures these tabs
 * exist:
 * Cover, Part 1, Part 2, Part 3, Part 4, Part 5
 *
 * SCOPES:
 * https://www.googleapis.com/auth/documents
 * https://www.googleapis.com/auth/script.external_request
 */

const CONFIG = (function () {
  const ML = 54;
  const MR = 54;
  const MT = 54;
  const MB = 54;
  const PAGE_WIDTH = 612;
  const W = PAGE_WIDTH - ML - MR;

  return {
    templateDocIdOrUrl: '',
    outputTitle: 'AI for Teachers - Participant Workbook (Session 1)',
    organizationLine: 'Baltimore City Schools | EdTech 2026',

    TABS: {
      cover: 'Cover',
      part1: 'Part 1',
      part2: 'Part 2',
      part3: 'Part 3',
      part4: 'Part 4',
      part5: 'Part 5'
    },

    PAGE: { ML, MR, MT, MB, W },

    F: {
      body: 'Arial',
      heading: 'Georgia'
    },

    C: {
      ink: '#1F2933',
      dim: '#52606D',
      muted: '#7B8794',
      navy: '#17324D',
      navySoft: '#ECF3F9',
      blue: '#2667A6',
      blueDark: '#1B4F72',
      blueSoft: '#EFF6FF',
      gold: '#A06A00',
      goldSoft: '#FFF7E6',
      green: '#2D6A4F',
      greenSoft: '#ECF8F1',
      red: '#A61E1E',
      redSoft: '#FDEEEE',
      soft: '#F7F9FB',
      softAlt: '#FBFCFD',
      border: '#D9E2EC',
      borderMid: '#BCCCDC',
      white: '#FFFFFF'
    },

    SZ: {
      coverKicker: 9,
      coverTitle: 21,
      coverSubtitle: 13,
      title: 17,
      h2: 13,
      h3: 11,
      body: 10,
      table: 10,
      label: 9,
      micro: 8
    }
  };
})();

/* -------------------------------------------------------------------------- */
/* ENTRY                                                                      */
/* -------------------------------------------------------------------------- */

function run() {
  try {
    const docId = getOrCreateTargetDocumentId_();
    ensureRequiredTabs_(docId);

    const doc = DocumentApp.openById(docId);
    const tabs = getRequiredTabBodies_(doc);
    const orderedBodies = [
      tabs.cover,
      tabs.part1,
      tabs.part2,
      tabs.part3,
      tabs.part4,
      tabs.part5
    ];

    for (let i = 0; i < orderedBodies.length; i++) {
      setupTabBody_(orderedBodies[i]);
    }

    buildCover_(tabs.cover);
    buildPart1_(tabs.part1);
    buildPart2_(tabs.part2);
    buildPart3_(tabs.part3);
    buildPart4_(tabs.part4);
    buildPart5_(tabs.part5);

    doc.saveAndClose();
    Logger.log('Done: ' + doc.getUrl());
    return doc.getUrl();
  } catch (error) {
    const message = error && error.stack ? error.stack : String(error);
    Logger.log('ERROR: ' + message);
    return 'ERROR: ' + String(error);
  }
}

/* -------------------------------------------------------------------------- */
/* DOCUMENT + TAB RESOLUTION                                                  */
/* -------------------------------------------------------------------------- */

function getOrCreateTargetDocumentId_() {
  const existingId = extractGoogleId_(CONFIG.templateDocIdOrUrl);
  if (existingId) return existingId;

  const created = docsApiCreateDocument_(CONFIG.outputTitle);
  Utilities.sleep(250);
  return created.documentId;
}

function ensureRequiredTabs_(docId) {
  const doc = DocumentApp.openById(docId);
  const existingTitles = getTabTitles_(doc);
  const requiredTitles = getRequiredTabTitles_();
  const missingTitles = getMissingTitles_(requiredTitles, existingTitles);

  if (!missingTitles.length) return;

  const apiDoc = docsApiGetDocument_(docId, true);
  const requests = buildTabCreationRequests_(apiDoc.tabs || [], requiredTitles);
  if (!requests.length) return;

  docsApiBatchUpdate_(docId, requests);
  Utilities.sleep(250);
}

function buildTabCreationRequests_(rootTabs, requiredTitles) {
  const requests = [];
  const flatTabs = flattenApiTabs_(rootTabs);
  const existingByTitle = {};

  for (let i = 0; i < flatTabs.length; i++) {
    const props = flatTabs[i].tabProperties || {};
    if (props.title && !existingByTitle[props.title]) {
      existingByTitle[props.title] = props;
    }
  }

  const hasAnyRequiredTitle = getMissingTitles_(
    requiredTitles,
    Object.keys(existingByTitle)
  ).length !== requiredTitles.length;

  if (!hasAnyRequiredTitle && canReuseFirstRootTab_(rootTabs)) {
    const firstProps = rootTabs[0].tabProperties || {};
    if (firstProps.tabId) {
      requests.push({
        updateDocumentTabProperties: {
          tabProperties: {
            tabId: firstProps.tabId,
            title: requiredTitles[0],
            index: 0
          },
          fields: 'title,index'
        }
      });
      existingByTitle[requiredTitles[0]] = {
        tabId: firstProps.tabId,
        title: requiredTitles[0],
        index: 0
      };
    }
  }

  for (let i = 0; i < requiredTitles.length; i++) {
    const title = requiredTitles[i];
    if (existingByTitle[title]) continue;

    requests.push({
      addDocumentTab: {
        tabProperties: {
          title: title,
          index: i
        }
      }
    });
  }

  return requests;
}

function canReuseFirstRootTab_(rootTabs) {
  if (!rootTabs || rootTabs.length !== 1) return false;
  const first = rootTabs[0];
  return !(first.childTabs && first.childTabs.length);
}

function flattenApiTabs_(tabs) {
  const out = [];
  for (let i = 0; i < tabs.length; i++) {
    out.push(tabs[i]);
    const children = tabs[i].childTabs || [];
    const nested = flattenApiTabs_(children);
    for (let j = 0; j < nested.length; j++) out.push(nested[j]);
  }
  return out;
}

function getRequiredTabTitles_() {
  return [
    CONFIG.TABS.cover,
    CONFIG.TABS.part1,
    CONFIG.TABS.part2,
    CONFIG.TABS.part3,
    CONFIG.TABS.part4,
    CONFIG.TABS.part5
  ];
}

function getMissingTitles_(requiredTitles, existingTitles) {
  const seen = {};
  for (let i = 0; i < existingTitles.length; i++) {
    seen[existingTitles[i]] = true;
  }

  const missing = [];
  for (let i = 0; i < requiredTitles.length; i++) {
    if (!seen[requiredTitles[i]]) missing.push(requiredTitles[i]);
  }
  return missing;
}

function getRequiredTabBodies_(doc) {
  const flatTabs = flattenTabs_(doc.getTabs());
  return {
    cover: getTabBodyByTitle_(flatTabs, CONFIG.TABS.cover),
    part1: getTabBodyByTitle_(flatTabs, CONFIG.TABS.part1),
    part2: getTabBodyByTitle_(flatTabs, CONFIG.TABS.part2),
    part3: getTabBodyByTitle_(flatTabs, CONFIG.TABS.part3),
    part4: getTabBodyByTitle_(flatTabs, CONFIG.TABS.part4),
    part5: getTabBodyByTitle_(flatTabs, CONFIG.TABS.part5)
  };
}

function getTabTitles_(doc) {
  const flatTabs = flattenTabs_(doc.getTabs());
  const titles = [];
  for (let i = 0; i < flatTabs.length; i++) {
    titles.push(flatTabs[i].getTitle());
  }
  return titles;
}

function flattenTabs_(tabs) {
  const out = [];
  for (let i = 0; i < tabs.length; i++) {
    out.push(tabs[i]);
    const children = tabs[i].getChildTabs();
    if (children && children.length) {
      const nested = flattenTabs_(children);
      for (let j = 0; j < nested.length; j++) out.push(nested[j]);
    }
  }
  return out;
}

function getTabBodyByTitle_(flatTabs, title) {
  for (let i = 0; i < flatTabs.length; i++) {
    if (flatTabs[i].getTitle() === title) {
      return flatTabs[i].asDocumentTab().getBody();
    }
  }

  throw new Error(
    'Missing required tab: "' +
      title +
      '". Check your document tabs or enable the Google Docs API service.'
  );
}

function extractGoogleId_(value) {
  if (!value) return '';
  const trimmed = String(value).trim();
  const match = trimmed.match(/[-\w]{25,}/);
  return match ? match[0] : '';
}

function docsApiCreateDocument_(title) {
  if (typeof Docs !== 'undefined' && Docs.Documents) {
    return Docs.Documents.create({ title: title });
  }

  return docsApiRequest_('post', 'https://docs.googleapis.com/v1/documents', {
    title: title
  });
}

function docsApiGetDocument_(docId, includeTabsContent) {
  if (typeof Docs !== 'undefined' && Docs.Documents) {
    return Docs.Documents.get(docId, {
      includeTabsContent: !!includeTabsContent
    });
  }

  const url =
    'https://docs.googleapis.com/v1/documents/' +
    encodeURIComponent(docId) +
    '?includeTabsContent=' +
    (!!includeTabsContent);
  return docsApiRequest_('get', url);
}

function docsApiBatchUpdate_(docId, requests) {
  if (typeof Docs !== 'undefined' && Docs.Documents) {
    return Docs.Documents.batchUpdate({ requests: requests }, docId);
  }

  const url =
    'https://docs.googleapis.com/v1/documents/' +
    encodeURIComponent(docId) +
    ':batchUpdate';
  return docsApiRequest_('post', url, { requests: requests });
}

function docsApiRequest_(method, url, payload) {
  const options = {
    method: method,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };

  if (payload !== undefined) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();
  const data = text ? JSON.parse(text) : {};

  if (code >= 200 && code < 300) return data;

  const apiMessage =
    data &&
    data.error &&
    data.error.message
      ? data.error.message
      : 'Google Docs API request failed.';

  if (
    code === 403 &&
    /SERVICE_DISABLED|has not been used|is disabled|not enabled/i.test(apiMessage)
  ) {
    throw new Error(
      'The Google Docs API is disabled for this Apps Script project. Enable ' +
        'the Google Docs API in the linked Google Cloud project, then run ' +
        'the script again.'
    );
  }

  throw new Error('Google Docs API error (' + code + '): ' + apiMessage);
}

function setupTabBody_(body) {
  clearBody_(body);
  body
    .setMarginLeft(CONFIG.PAGE.ML)
    .setMarginRight(CONFIG.PAGE.MR)
    .setMarginTop(CONFIG.PAGE.MT)
    .setMarginBottom(CONFIG.PAGE.MB);
}

function clearBody_(body) {
  while (body.getNumChildren() > 0) {
    body.removeChild(body.getChild(0));
  }
}

/* -------------------------------------------------------------------------- */
/* COVER                                                                      */
/* -------------------------------------------------------------------------- */

function buildCover_(body) {
  heroPanel_(
    body,
    'AI FOR ACCESSIBLE INSTRUCTION',
    'Participant Workbook',
    'Session 1 | Access Without Lowering Demand',
    CONFIG.organizationLine
  );

  callout_(
    body,
    'Session outcome',
    'By the end of this session, you will leave with one ready-to-use support that removes an access barrier while preserving the thinking students are expected to do.',
    'blue'
  );

  bodyText_(
    body,
    'Use this workbook alongside the session slides. Each section includes short written responses, analysis tasks, and planning spaces designed to help you leave with a practical, classroom-ready support.'
  );

  sectionHeader_(body, 'Session roadmap');
  addInfoTable_(
    body,
    ['Part', 'Focus', 'What you will produce'],
    [
      [
        'Part 1',
        'Name the barrier',
        'A clear access barrier and the thinking students still need to do'
      ],
      [
        'Part 2',
        'Use OSAMR',
        'A stronger understanding of access versus overreliance'
      ],
      [
        'Part 3',
        'Redesign a weak example',
        'An improved support that keeps rigor intact'
      ],
      [
        'Part 4',
        'Build your own support',
        'A ready-to-use classroom support and a reusable workflow'
      ],
      [
        'Part 5',
        'Commit and share',
        'An implementation plan and accountability step'
      ]
    ],
    [90, 150, CONFIG.PAGE.W - 240]
  );

  callout_(
    body,
    'Working principle',
    'Increase access without reducing rigor. The goal is not to make work easier. The goal is to make the thinking more reachable.',
    'soft'
  );
}

/* -------------------------------------------------------------------------- */
/* PART 1                                                                     */
/* -------------------------------------------------------------------------- */

function buildPart1_(body) {
  partBanner_(
    body,
    'Part 1',
    'Start with a real barrier',
    'Name the access barrier clearly, then protect the thinking students still need to do.'
  );

  bodyText_(
    body,
    'Think of one student or one group of students who struggled with a recent task. The challenge was not simply effort or compliance. Something about the text, language, structure, or format created an unnecessary obstacle.'
  );

  promptLabel_(body, 'Describe one barrier you have seen in your classroom.');
  writeBox_(body, 4);

  divider_(body);

  sectionHeader_(body, 'Common access barriers');
  bodyText_(
    body,
    'Select the barrier that best matches the challenge you want to address today.'
  );

  addInfoTable_(
    body,
    ['Barrier', 'What it often looks like in class'],
    [
      [
        'Reading level is too high',
        'Students cannot access the text well enough to begin the task independently.'
      ],
      [
        'Directions are unclear',
        'Students are unsure what to do, where to start, or what a successful response should include.'
      ],
      [
        'Vocabulary is unfamiliar',
        'Academic or content-specific terms block comprehension of the task.'
      ],
      [
        'Language demands are too heavy',
        'Students can think about the content, but they need support expressing their thinking.'
      ],
      [
        'Task format is unfamiliar',
        'Students have not yet learned how to navigate the response structure or format.'
      ]
    ],
    [Math.round(CONFIG.PAGE.W * 0.34), Math.round(CONFIG.PAGE.W * 0.66)]
  );

  divider_(body);

  sectionHeader_(body, '1-2-4 protocol');
  bodyText_(
    body,
    'Use the sequence below to sharpen your thinking before you move into the framework.'
  );

  addInfoTable_(
    body,
    ['Time', 'Round', 'Participant action'],
    [
      [
        '1 min',
        'Individual',
        'Silently identify one specific barrier in an upcoming lesson or assignment.'
      ],
      [
        '2 min',
        'Pair',
        'Share the barrier with a partner and explain what students still need to think about or do.'
      ],
      [
        '2 min',
        'Whole group',
        'Listen for patterns as the facilitator captures common barriers across the room.'
      ]
    ],
    [56, 120, CONFIG.PAGE.W - 176]
  );

  promptLabel_(body, 'The one barrier I am addressing today');
  writeBox_(body, 3);

  promptLabel_(body, 'The thinking I still want students to do');
  writeBox_(body, 3);

  navHint_(body, 'Continue to Part 2 when the facilitator advances.');
}

/* -------------------------------------------------------------------------- */
/* PART 2                                                                     */
/* -------------------------------------------------------------------------- */

function buildPart2_(body) {
  partBanner_(
    body,
    'Part 2',
    'Use OSAMR as a decision tool',
    'Classify supports carefully so access improves while rigor stays intact.'
  );

  callout_(
    body,
    'Guiding question',
    'Does this use of AI increase access while preserving the thinking students are expected to do?',
    'blue'
  );

  bodyText_(
    body,
    'OSAMR is not a ladder to climb. It is a practical lens for checking whether a support improves access, protects rigor, and avoids doing the work for students.'
  );

  addInfoTable_(
    body,
    ['Level', 'What is happening', 'Barrier removed?', 'Thinking preserved?', 'What to watch for'],
    [
      [
        'S - Substitution',
        'Technology replaces a paper-based move with little or no functional change.',
        'No',
        'Yes',
        'The support may not meaningfully improve access.'
      ],
      [
        'A - Augmentation',
        'Technology adds a helpful feature such as text-to-speech or clearer organization.',
        'Yes',
        'Yes',
        'Often a strong fit for access support.'
      ],
      [
        'M - Modification',
        'The task is redesigned in a way that changes how students engage with it.',
        'Yes',
        'Yes',
        'Check that the redesign still aligns to the learning target.'
      ],
      [
        'R - Redefinition',
        'Technology enables a task that would not be possible without it.',
        'Yes',
        'Yes',
        'Use when it truly expands what students can do.'
      ],
      [
        'O - Overreliance',
        'AI completes the thinking the student was supposed to do.',
        'No',
        'No',
        'Avoid this. It lowers demand instead of increasing access.'
      ]
    ],
    [90, 180, 70, 88, CONFIG.PAGE.W - 428]
  );

  callout_(
    body,
    'Key principle',
    'A well-aligned support at S, A, M, or R is stronger than an impressive-looking tool that drifts into overreliance.',
    'soft'
  );

  divider_(body);

  sectionHeader_(body, 'Case study analysis');
  bodyText_(
    body,
    'Read each scenario with your group. Then identify the barrier, the preserved thinking, and the OSAMR level that best fits the support.'
  );

  subHeader_(body, 'Case study 1: Sentence frames that preserve reasoning');
  scenarioBox_(
    body,
    'Ms. Rodriguez teaches Grade 6 science. Her students struggle to write claims with evidence. She uses AI to generate sentence frames such as "I claim that ____ because ____" and "The data shows ____, which supports ____." Students still select their own evidence from the lab data and complete the reasoning themselves.',
    'green'
  );

  addPromptResponseTable_(
    body,
    ['Discussion prompt', 'Group notes'],
    [
      ['What barrier was removed?', 2],
      ['What thinking did students still do?', 2],
      ['Which OSAMR level best fits this example?', 2],
      ['What evidence supports your classification?', 3]
    ],
    [Math.round(CONFIG.PAGE.W * 0.46), Math.round(CONFIG.PAGE.W * 0.54)]
  );

  subHeader_(body, 'Case study 2: Summary that replaces the text');
  scenarioBox_(
    body,
    'Mr. Johnson teaches English. The class is reading a complex chapter from "To Kill a Mockingbird." He asks AI to summarize the chapter in simpler language and gives students the AI summary instead of the original text. Students answer comprehension questions using the summary only.',
    'red'
  );

  addPromptResponseTable_(
    body,
    ['Discussion prompt', 'Group notes'],
    [
      ['What barrier was the teacher trying to remove?', 2],
      ['What student thinking was lost?', 2],
      ['Which OSAMR level best fits this example?', 2],
      ['How could the support be redesigned to preserve rigor?', 3]
    ],
    [Math.round(CONFIG.PAGE.W * 0.46), Math.round(CONFIG.PAGE.W * 0.54)]
  );

  callout_(
    body,
    'Whole-group share-out',
    'Be prepared to share one classification decision and one red flag your group noticed. Listen for patterns that separate true access supports from overreliance.',
    'soft'
  );

  navHint_(body, 'Continue to Part 3 when the facilitator advances.');
}

/* -------------------------------------------------------------------------- */
/* PART 3                                                                     */
/* -------------------------------------------------------------------------- */

function buildPart3_(body) {
  partBanner_(
    body,
    'Part 3',
    'Redesign an overreliance example',
    'Keep the barrier in view, but move the cognitive work back to students.'
  );

  bodyText_(
    body,
    'Your group will receive one scenario in which AI reduced the work students were supposed to do. Redesign the support so that the original barrier is addressed while the learning remains intact.'
  );

  callout_(
    body,
    'Design constraints',
    'Keep the same barrier, preserve the original standard and success criteria, and make sure students still do the target thinking.',
    'soft'
  );

  addGuidedWriteTable_(
    body,
    [
      ['1. Original barrier', 'What access issue was the teacher trying to solve?', 3],
      ['2. Overreliance problem', 'What thinking did the original support remove or replace?', 3],
      ['3. Redesigned support', 'Describe the improved support. Explain how it removes the barrier and keeps the cognitive work with students.', 5],
      ['4. OSAMR classification', 'Identify the new level and justify the choice.', 3]
    ],
    [Math.round(CONFIG.PAGE.W * 0.38), Math.round(CONFIG.PAGE.W * 0.62)]
  );

  callout_(
    body,
    'Stress test before sharing',
    'Ask this question: If a student used only this support, would the student still need to demonstrate the target skill? If the answer is no, revise the support.',
    'soft'
  );

  divider_(body);

  sectionHeader_(body, 'Gallery walk notes');
  bodyText_(
    body,
    'As each redesigned support is shared, capture one move you may want to borrow.'
  );

  addPromptResponseTable_(
    body,
    ['Prompt', 'Notes'],
    [
      ['What barrier was this support solving?', 2],
      ['What made the redesign effective?', 2],
      ['What is one move I could use in my own classroom?', 2]
    ],
    [Math.round(CONFIG.PAGE.W * 0.42), Math.round(CONFIG.PAGE.W * 0.58)]
  );

  promptLabel_(body, 'Top takeaways from the gallery walk');
  writeBox_(body, 5);

  navHint_(body, 'Continue to Part 4 when the facilitator advances.');
}

/* -------------------------------------------------------------------------- */
/* PART 4                                                                     */
/* -------------------------------------------------------------------------- */

function buildPart4_(body) {
  partBanner_(
    body,
    'Part 4',
    'Build your own support',
    'Use a clear workflow so the support removes a barrier without removing the thinking.'
  );

  sectionHeader_(body, 'Build readiness check');
  bodyText_(
    body,
    'Complete this planning step before opening any AI tool. The stronger your inputs and guardrails, the better your support will be.'
  );

  addChecklistPlanningTable_(
    body,
    [
      ['What barrier will I address?', 'Name the specific student need from Part 1.'],
      ['What instructional material will I input?', 'List the task, text, directions, or assignment prompt you will use.'],
      ['What do I want the AI to produce?', 'For example: chunked directions, sentence frames, a graphic organizer, or translated directions.'],
      ['How will I verify quality?', 'Plan how you will check that access improved and rigor stayed intact.']
    ]
  );

  callout_(
    body,
    'Privacy requirement',
    'Do not paste student names, ID numbers, IEP details, or any identifying information into an AI tool. Use neutral placeholders such as "a Grade 6 student" or "an emerging bilingual learner."',
    'gold'
  );

  divider_(body);

  sectionHeader_(body, 'Prompt templates');
  bodyText_(
    body,
    'Select the template that best matches your barrier. Customize the bracketed sections before you run it.'
  );

  promptTemplate_(
    body,
    'Template 1 | Clarify directions',
    'Best for unclear directions',
    'Rewrite these directions for [grade level] students. Break them into numbered steps. Use concise, student-friendly language. Keep the task and required thinking exactly the same.\n\n[Paste directions here]'
  );

  promptTemplate_(
    body,
    'Template 2 | Generate sentence starters',
    'Best for language production demands',
    'Create three sentence starters that help students [analyze, compare, justify, explain]. Guide the response structure, but do not provide the content or reasoning.\n\n[Paste the learning target or task here]'
  );

  promptTemplate_(
    body,
    'Template 3 | Chunk a complex text',
    'Best for dense reading without simplifying the source',
    'Break this passage into three manageable chunks. Add one focus question before each chunk to direct attention. Keep the original text intact. Do not simplify the wording.\n\n[Paste passage here]'
  );

  promptTemplate_(
    body,
    'Template 4 | Translate directions while protecting academic language',
    'Best for multilingual access',
    'Translate these directions into [language]. Keep academic vocabulary in English and include the translation in parentheses the first time each term appears.\n\n[Paste directions here]'
  );

  divider_(body);

  sectionHeader_(body, 'Build sprint');
  bodyText_(
    body,
    'As you create your support, keep naming the OSAMR level you are aiming for. If the support starts doing the thinking for students, revise immediately.'
  );

  callout_(
    body,
    'Coaching questions',
    'Does this remove the barrier? Are students still doing the hard thinking? Is the support aligned to the original learning target? Would you feel comfortable showing it to a coach, co-teacher, or family member?',
    'soft'
  );

  divider_(body);

  sectionHeader_(body, 'Verification checkpoint');
  bodyText_(
    body,
    'Before you document your work, confirm that all three statements below are true.'
  );

  addVerificationTable_(
    body,
    [
      ['1. The barrier is genuinely reduced.', 'A student who was previously blocked can now access the task more effectively.'],
      ['2. The student still does the thinking.', 'The support guides the work but does not complete the reasoning, reading, or analysis.'],
      ['3. The support still matches the standard.', 'The cognitive demand and success criteria remain aligned to the original target.']
    ]
  );

  divider_(body);

  sectionHeader_(body, 'Build recipe');
  bodyText_(
    body,
    'Use the space below to capture your workflow so you can repeat it later with less effort.'
  );

  addGuidedWriteTable_(
    body,
    [
      ['1. Input', 'What material did I provide to the AI tool?', 3],
      ['2. Output', 'What did the tool generate?', 3],
      ['3. Verification', 'How did I check that the barrier was reduced and rigor stayed intact?', 3],
      ['4. OSAMR lens', 'Which level best describes this support, and how did I avoid overreliance?', 3]
    ],
    [Math.round(CONFIG.PAGE.W * 0.36), Math.round(CONFIG.PAGE.W * 0.64)]
  );

  callout_(
    body,
    'Why document this?',
    'A clear build recipe makes your process reusable. The goal is not just one strong support today. The goal is a repeatable workflow you can use again.',
    'soft'
  );

  navHint_(body, 'Continue to Part 5 when the facilitator advances.');
}

/* -------------------------------------------------------------------------- */
/* PART 5                                                                     */
/* -------------------------------------------------------------------------- */

function buildPart5_(body) {
  partBanner_(
    body,
    'Part 5',
    'Commit and share',
    'Lock in an implementation plan and a simple accountability move before you leave.'
  );

  sectionHeader_(body, 'Peer showcase notes');
  bodyText_(
    body,
    'As participant examples are shared, capture evidence of quality and one idea you may want to adapt.'
  );

  addPromptResponseTable_(
    body,
    ['Prompt', 'Notes'],
    [
      ['What barrier was this support designed to solve?', 2],
      ['Why does this support protect rigor?', 2],
      ['What is one move I could adapt for my classroom?', 2]
    ],
    [Math.round(CONFIG.PAGE.W * 0.42), Math.round(CONFIG.PAGE.W * 0.58)]
  );

  divider_(body);

  sectionHeader_(body, 'OSAMR self-reflection');
  bodyText_(
    body,
    'Reflect on your own support before finalizing your implementation plan.'
  );

  addGuidedWriteTable_(
    body,
    [
      ['1. OSAMR level', 'Identify the level that best fits your support.', 2],
      ['2. Evidence for your decision', 'Explain why that classification is appropriate.', 3],
      ['3. Overreliance safeguard', 'Name one specific step you used to keep the thinking with students.', 3]
    ],
    [Math.round(CONFIG.PAGE.W * 0.42), Math.round(CONFIG.PAGE.W * 0.58)]
  );

  divider_(body);

  sectionHeader_(body, 'Implementation commitment');
  bodyText_(
    body,
    'Complete the three planning boxes below so that your support moves from workshop idea to classroom use.'
  );

  promptLabel_(body, 'Lesson or task in which I will use this support within the next 10 days');
  writeBox_(body, 2);

  promptLabel_(body, 'Evidence I will look for in student work to know the support helped');
  writeBox_(body, 2);

  promptLabel_(body, 'Accountability partner and how we will check in');
  writeBox_(body, 2);

  divider_(body);

  sectionHeader_(body, 'Accountability partner exchange');
  addInfoTable_(
    body,
    ['Exchange', 'Agreements for the check-in'],
    [
      [
        'Name and contact information',
        'Discuss whether the support worked, what student work revealed, and what you would revise before using it again.'
      ]
    ],
    [Math.round(CONFIG.PAGE.W * 0.38), Math.round(CONFIG.PAGE.W * 0.62)]
  );

  divider_(body);

  sectionHeader_(body, 'Final pre-use check');
  addBulletChecklist_(
    body,
    [
      'No student names, ID numbers, or protected information were used in any AI prompt.',
      'I reviewed the support for accuracy before planning to share it with students.',
      'I checked that the support aligns to the standard and protects the intended rigor.',
      'I considered whether the language, examples, and translations are appropriate for my students.'
    ]
  );

  callout_(
    body,
    'Closing reminder',
    'You are the instructional decision-maker. AI is a tool that can help remove a barrier, but it does not replace professional judgment. The goal is to improve access while keeping the learning work with students.',
    'gold'
  );
}

/* -------------------------------------------------------------------------- */
/* STYLE HELPERS                                                              */
/* -------------------------------------------------------------------------- */

function p_(para, opts) {
  opts = opts || {};
  para.setFontFamily(opts.font || CONFIG.F.body);
  if (opts.sz !== undefined) para.setFontSize(opts.sz);
  para.setBold(!!opts.bold);
  para.setItalic(!!opts.italic);
  if (opts.color) para.setForegroundColor(opts.color);
  if (opts.after !== undefined) para.setSpacingAfter(opts.after);
  if (opts.before !== undefined) para.setSpacingBefore(opts.before);
  if (opts.align) para.setAlignment(opts.align);
  return para;
}

function clearElementChildren_(element) {
  while (element.getNumChildren && element.getNumChildren() > 0) {
    element.removeChild(element.getChild(0));
  }
}

function setCellText_(cell, text, opts) {
  clearElementChildren_(cell);
  const para = cell.appendParagraph(text || '');
  p_(para, opts || {});
  return para;
}

function setCellParagraphs_(cell, items) {
  clearElementChildren_(cell);
  for (let i = 0; i < items.length; i++) {
    const item = items[i] || {};
    const para = cell.appendParagraph(item.text || '');
    p_(para, item.opts || {});
  }
}

function setTableWidths_(table, widths) {
  for (let i = 0; i < widths.length; i++) {
    table.setColumnWidth(i, widths[i]);
  }
}

function blankMatrix_(rows, cols) {
  const out = [];
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) row.push('');
    out.push(row);
  }
  return out;
}

function appendSpacer_(body, pts) {
  const para = body.appendParagraph('');
  p_(para, { font: CONFIG.F.body, sz: 2, after: pts || 8 });
}

function heroPanel_(body, kicker, title, subtitle, orgLine) {
  const table = body.appendTable([['']]);
  table.setBorderWidth(1).setBorderColor(CONFIG.C.blue);
  table.setColumnWidth(0, CONFIG.PAGE.W);

  const cell = table.getCell(0, 0);
  cell
    .setBackgroundColor(CONFIG.C.navySoft)
    .setPaddingTop(18)
    .setPaddingBottom(18)
    .setPaddingLeft(18)
    .setPaddingRight(18);

  setCellParagraphs_(cell, [
    {
      text: kicker,
      opts: {
        font: CONFIG.F.body,
        sz: CONFIG.SZ.coverKicker,
        bold: true,
        color: CONFIG.C.blueDark,
        after: 6
      }
    },
    {
      text: title,
      opts: {
        font: CONFIG.F.heading,
        sz: CONFIG.SZ.coverTitle,
        bold: true,
        color: CONFIG.C.navy,
        after: 4
      }
    },
    {
      text: subtitle,
      opts: {
        font: CONFIG.F.heading,
        sz: CONFIG.SZ.coverSubtitle,
        color: CONFIG.C.ink,
        after: 6
      }
    },
    {
      text: orgLine,
      opts: {
        font: CONFIG.F.body,
        sz: CONFIG.SZ.label,
        color: CONFIG.C.dim,
        after: 0
      }
    }
  ]);

  appendSpacer_(body, 10);
}

function partBanner_(body, partLabel, title, summary) {
  const table = body.appendTable([['']]);
  table.setBorderWidth(1).setBorderColor(CONFIG.C.borderMid);
  table.setColumnWidth(0, CONFIG.PAGE.W);

  const cell = table.getCell(0, 0);
  cell
    .setBackgroundColor(CONFIG.C.soft)
    .setPaddingTop(14)
    .setPaddingBottom(14)
    .setPaddingLeft(14)
    .setPaddingRight(14);

  setCellParagraphs_(cell, [
    {
      text: partLabel,
      opts: {
        font: CONFIG.F.body,
        sz: CONFIG.SZ.label,
        bold: true,
        color: CONFIG.C.blueDark,
        after: 4
      }
    },
    {
      text: title,
      opts: {
        font: CONFIG.F.heading,
        sz: CONFIG.SZ.title,
        bold: true,
        color: CONFIG.C.navy,
        after: 3
      }
    },
    {
      text: summary,
      opts: {
        font: CONFIG.F.body,
        sz: CONFIG.SZ.body,
        color: CONFIG.C.dim,
        after: 0
      }
    }
  ]);

  appendSpacer_(body, 10);
}

function sectionHeader_(body, label) {
  const para = body.appendParagraph(label);
  p_(para, {
    font: CONFIG.F.heading,
    sz: CONFIG.SZ.h2,
    bold: true,
    color: CONFIG.C.navy,
    before: 2,
    after: 6
  });
}

function subHeader_(body, label) {
  const para = body.appendParagraph(label);
  p_(para, {
    font: CONFIG.F.heading,
    sz: CONFIG.SZ.h3,
    bold: true,
    color: CONFIG.C.ink,
    before: 2,
    after: 4
  });
}

function bodyText_(body, text) {
  const para = body.appendParagraph(text);
  p_(para, {
    font: CONFIG.F.body,
    sz: CONFIG.SZ.body,
    color: CONFIG.C.ink,
    after: 8
  });
}

function promptLabel_(body, label) {
  const para = body.appendParagraph(label);
  p_(para, {
    font: CONFIG.F.body,
    sz: CONFIG.SZ.label,
    bold: true,
    color: CONFIG.C.blueDark,
    after: 4
  });
}

function writeBox_(body, lines) {
  const table = body.appendTable([['']]);
  table.setBorderWidth(1).setBorderColor(CONFIG.C.borderMid);
  table.setColumnWidth(0, CONFIG.PAGE.W);

  const row = table.getRow(0);
  row.setMinimumHeight(Math.max(44, (lines || 2) * 26));

  const cell = table.getCell(0, 0);
  cell
    .setBackgroundColor(CONFIG.C.white)
    .setPaddingTop(10)
    .setPaddingBottom(10)
    .setPaddingLeft(10)
    .setPaddingRight(10);

  setCellText_(cell, '', {
    font: CONFIG.F.body,
    sz: CONFIG.SZ.body,
    color: CONFIG.C.ink,
    after: 0
  });

  appendSpacer_(body, 8);
}

function divider_(body) {
  body.appendHorizontalRule();
  appendSpacer_(body, 6);
}

function navHint_(body, text) {
  const para = body.appendParagraph(text);
  p_(para, {
    font: CONFIG.F.body,
    sz: CONFIG.SZ.micro,
    italic: true,
    color: CONFIG.C.dim,
    after: 0
  });
}

function callout_(body, label, text, type) {
  const palette = {
    blue: { bg: CONFIG.C.blueSoft, border: CONFIG.C.blue, label: CONFIG.C.blueDark },
    soft: { bg: CONFIG.C.soft, border: CONFIG.C.borderMid, label: CONFIG.C.navy },
    gold: { bg: CONFIG.C.goldSoft, border: CONFIG.C.gold, label: CONFIG.C.gold }
  };
  const colors = palette[type] || palette.soft;

  const table = body.appendTable([['']]);
  table.setBorderWidth(1).setBorderColor(colors.border);
  table.setColumnWidth(0, CONFIG.PAGE.W);

  const cell = table.getCell(0, 0);
  cell
    .setBackgroundColor(colors.bg)
    .setPaddingTop(10)
    .setPaddingBottom(10)
    .setPaddingLeft(12)
    .setPaddingRight(12);

  setCellParagraphs_(cell, [
    {
      text: label,
      opts: {
        font: CONFIG.F.body,
        sz: CONFIG.SZ.label,
        bold: true,
        color: colors.label,
        after: 3
      }
    },
    {
      text: text,
      opts: {
        font: CONFIG.F.body,
        sz: CONFIG.SZ.table,
        color: CONFIG.C.ink,
        after: 0
      }
    }
  ]);

  appendSpacer_(body, 8);
}

function scenarioBox_(body, text, tone) {
  const palette = {
    green: { bg: CONFIG.C.greenSoft, border: CONFIG.C.green },
    red: { bg: CONFIG.C.redSoft, border: CONFIG.C.red }
  };
  const colors = palette[tone] || palette.green;

  const table = body.appendTable([['']]);
  table.setBorderWidth(1).setBorderColor(colors.border);
  table.setColumnWidth(0, CONFIG.PAGE.W);

  const cell = table.getCell(0, 0);
  cell
    .setBackgroundColor(colors.bg)
    .setPaddingTop(10)
    .setPaddingBottom(10)
    .setPaddingLeft(12)
    .setPaddingRight(12);

  setCellText_(cell, text, {
    font: CONFIG.F.body,
    sz: CONFIG.SZ.table,
    color: CONFIG.C.ink,
    after: 0
  });

  appendSpacer_(body, 8);
}

function promptTemplate_(body, title, subtitle, promptText) {
  const table = body.appendTable([['']]);
  table.setBorderWidth(1).setBorderColor(CONFIG.C.border);
  table.setColumnWidth(0, CONFIG.PAGE.W);

  const cell = table.getCell(0, 0);
  cell
    .setBackgroundColor(CONFIG.C.white)
    .setPaddingTop(10)
    .setPaddingBottom(10)
    .setPaddingLeft(12)
    .setPaddingRight(12);

  setCellParagraphs_(cell, [
    {
      text: title,
      opts: {
        font: CONFIG.F.heading,
        sz: CONFIG.SZ.body,
        bold: true,
        color: CONFIG.C.navy,
        after: 2
      }
    },
    {
      text: subtitle,
      opts: {
        font: CONFIG.F.body,
        sz: CONFIG.SZ.label,
        italic: true,
        color: CONFIG.C.dim,
        after: 6
      }
    },
    {
      text: promptText,
      opts: {
        font: CONFIG.F.body,
        sz: CONFIG.SZ.label,
        color: CONFIG.C.ink,
        after: 0
      }
    }
  ]);

  appendSpacer_(body, 8);
}

/* -------------------------------------------------------------------------- */
/* TABLE HELPERS                                                              */
/* -------------------------------------------------------------------------- */

function addInfoTable_(body, headers, rows, colWidths) {
  const allRows = [headers].concat(rows);
  const table = body.appendTable(blankMatrix_(allRows.length, headers.length));
  table.setBorderWidth(1).setBorderColor(CONFIG.C.border);
  setTableWidths_(table, colWidths);

  for (let r = 0; r < allRows.length; r++) {
    for (let c = 0; c < headers.length; c++) {
      const cell = table.getCell(r, c);
      cell.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(8).setPaddingRight(8);

      if (r === 0) {
        cell.setBackgroundColor(CONFIG.C.navySoft);
        setCellText_(cell, allRows[r][c], {
          font: CONFIG.F.body,
          sz: CONFIG.SZ.table,
          bold: true,
          color: CONFIG.C.navy,
          after: 0
        });
      } else {
        cell.setBackgroundColor(r % 2 === 0 ? CONFIG.C.softAlt : CONFIG.C.white);
        setCellText_(cell, allRows[r][c], {
          font: CONFIG.F.body,
          sz: CONFIG.SZ.table,
          color: CONFIG.C.ink,
          after: 0
        });
      }
    }
  }

  appendSpacer_(body, 8);
}

function addPromptResponseTable_(body, headers, prompts, colWidths) {
  const table = body.appendTable(blankMatrix_(prompts.length + 1, 2));
  table.setBorderWidth(1).setBorderColor(CONFIG.C.border);
  setTableWidths_(table, colWidths);

  for (let c = 0; c < 2; c++) {
    const headerCell = table.getCell(0, c);
    headerCell
      .setBackgroundColor(CONFIG.C.navySoft)
      .setPaddingTop(8)
      .setPaddingBottom(8)
      .setPaddingLeft(8)
      .setPaddingRight(8);

    setCellText_(headerCell, headers[c], {
      font: CONFIG.F.body,
      sz: CONFIG.SZ.table,
      bold: true,
      color: CONFIG.C.navy,
      after: 0
    });
  }

  for (let i = 0; i < prompts.length; i++) {
    const rowIndex = i + 1;
    table.getRow(rowIndex).setMinimumHeight(Math.max(36, (prompts[i][1] || 2) * 24));

    const promptCell = table.getCell(rowIndex, 0);
    promptCell
      .setBackgroundColor(CONFIG.C.soft)
      .setPaddingTop(8)
      .setPaddingBottom(8)
      .setPaddingLeft(8)
      .setPaddingRight(8);

    setCellText_(promptCell, prompts[i][0], {
      font: CONFIG.F.body,
      sz: CONFIG.SZ.table,
      color: CONFIG.C.ink,
      after: 0
    });

    const responseCell = table.getCell(rowIndex, 1);
    responseCell
      .setBackgroundColor(CONFIG.C.white)
      .setPaddingTop(8)
      .setPaddingBottom(8)
      .setPaddingLeft(8)
      .setPaddingRight(8);

    setCellText_(responseCell, '', {
      font: CONFIG.F.body,
      sz: CONFIG.SZ.table,
      color: CONFIG.C.ink,
      after: 0
    });
  }

  appendSpacer_(body, 8);
}

function addGuidedWriteTable_(body, rows, colWidths) {
  const table = body.appendTable(blankMatrix_(rows.length, 2));
  table.setBorderWidth(1).setBorderColor(CONFIG.C.border);
  setTableWidths_(table, colWidths);

  for (let i = 0; i < rows.length; i++) {
    const label = rows[i][0];
    const guidance = rows[i][1];
    const lines = rows[i][2] || 3;

    table.getRow(i).setMinimumHeight(Math.max(48, lines * 24));

    const left = table.getCell(i, 0);
    left
      .setBackgroundColor(CONFIG.C.navySoft)
      .setPaddingTop(8)
      .setPaddingBottom(8)
      .setPaddingLeft(8)
      .setPaddingRight(8);

    setCellParagraphs_(left, [
      {
        text: label,
        opts: {
          font: CONFIG.F.body,
          sz: CONFIG.SZ.table,
          bold: true,
          color: CONFIG.C.navy,
          after: 3
        }
      },
      {
        text: guidance,
        opts: {
          font: CONFIG.F.body,
          sz: CONFIG.SZ.label,
          italic: true,
          color: CONFIG.C.dim,
          after: 0
        }
      }
    ]);

    const right = table.getCell(i, 1);
    right
      .setBackgroundColor(CONFIG.C.white)
      .setPaddingTop(8)
      .setPaddingBottom(8)
      .setPaddingLeft(8)
      .setPaddingRight(8);

    setCellText_(right, '', {
      font: CONFIG.F.body,
      sz: CONFIG.SZ.table,
      color: CONFIG.C.ink,
      after: 0
    });
  }

  appendSpacer_(body, 8);
}

function addChecklistPlanningTable_(body, items) {
  const colWidths = [36, CONFIG.PAGE.W - 36];
  const table = body.appendTable(blankMatrix_(items.length, 2));
  table.setBorderWidth(1).setBorderColor(CONFIG.C.border);
  setTableWidths_(table, colWidths);

  for (let i = 0; i < items.length; i++) {
    table.getRow(i).setMinimumHeight(42);

    const boxCell = table.getCell(i, 0);
    boxCell
      .setBackgroundColor(CONFIG.C.soft)
      .setPaddingTop(8)
      .setPaddingBottom(8)
      .setPaddingLeft(8)
      .setPaddingRight(4);

    setCellText_(boxCell, '[ ]', {
      font: CONFIG.F.body,
      sz: 11,
      color: CONFIG.C.ink,
      after: 0
    });

    const promptCell = table.getCell(i, 1);
    promptCell
      .setBackgroundColor(CONFIG.C.white)
      .setPaddingTop(8)
      .setPaddingBottom(8)
      .setPaddingLeft(8)
      .setPaddingRight(8);

    setCellParagraphs_(promptCell, [
      {
        text: items[i][0],
        opts: {
          font: CONFIG.F.body,
          sz: CONFIG.SZ.table,
          bold: true,
          color: CONFIG.C.ink,
          after: 3
        }
      },
      {
        text: items[i][1],
        opts: {
          font: CONFIG.F.body,
          sz: CONFIG.SZ.label,
          italic: true,
          color: CONFIG.C.dim,
          after: 0
        }
      }
    ]);
  }

  appendSpacer_(body, 8);
}

function addVerificationTable_(body, rows) {
  const colWidths = [Math.round(CONFIG.PAGE.W * 0.42), Math.round(CONFIG.PAGE.W * 0.58)];
  const table = body.appendTable(blankMatrix_(rows.length, 2));
  table.setBorderWidth(1).setBorderColor(CONFIG.C.border);
  setTableWidths_(table, colWidths);

  for (let i = 0; i < rows.length; i++) {
    table.getRow(i).setMinimumHeight(38);

    const left = table.getCell(i, 0);
    left
      .setBackgroundColor(CONFIG.C.greenSoft)
      .setPaddingTop(8)
      .setPaddingBottom(8)
      .setPaddingLeft(8)
      .setPaddingRight(8);

    setCellText_(left, rows[i][0], {
      font: CONFIG.F.body,
      sz: CONFIG.SZ.table,
      bold: true,
      color: CONFIG.C.green,
      after: 0
    });

    const right = table.getCell(i, 1);
    right
      .setBackgroundColor(CONFIG.C.white)
      .setPaddingTop(8)
      .setPaddingBottom(8)
      .setPaddingLeft(8)
      .setPaddingRight(8);

    setCellText_(right, rows[i][1], {
      font: CONFIG.F.body,
      sz: CONFIG.SZ.table,
      italic: true,
      color: CONFIG.C.dim,
      after: 0
    });
  }

  appendSpacer_(body, 8);
}

function addBulletChecklist_(body, items) {
  for (let i = 0; i < items.length; i++) {
    const para = body.appendParagraph('[ ] ' + items[i]);
    p_(para, {
      font: CONFIG.F.body,
      sz: CONFIG.SZ.table,
      color: CONFIG.C.ink,
      after: 6
    });
  }
  appendSpacer_(body, 6);
}
