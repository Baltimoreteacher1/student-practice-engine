/******************************************************************
 * NotebookExtractors_Precision.gs
 * EduWonderLab source-fidelity extractor for 5.3 "Determine the Area
 * of Triangles"
 *
 * This file generates the notebook package consumed by:
 *   - notebook_generator.gs
 *   - Notebookgrids.gs
 *
 * Contract:
 *   extractNotebookPackageFromDeck_(SlidesApp.Presentation) -> {
 *     meta: {...},
 *     sessions: [session1Spec, session2Spec]
 *   }
 ******************************************************************/

var NOTEBOOK_BUILD_DEFAULTS_ = {
  MIN_SLIDES_PER_SESSION: 12,
  MAX_SLIDES_PER_SESSION: 14,
  MAX_SLIDES_WITHOUT_DISCUSSION: 3
};

function extractNotebookPackageFromDeck_(presentation) {
  var pres = presentation || SlidesApp.getActivePresentation();
  var source = extractTriangleFlagshipSource_(pres);
  var blueprint = buildTriangleFlagshipBlueprint_(source);
  validateTriangleBlueprint_(blueprint);
  var pkg = buildNotebookPackageFromBlueprint_(blueprint);
  validateExtractorPackage_(pkg);
  return pkg;
}

function extractTriangleFlagshipSource_(presentation) {
  var deck = buildNotebookDeckModel_(presentation);
  var lessonTitle = inferNotebookLessonTitle_(deck);
  var learningTargets = extractLearningTargets_(deck);
  var sessions = extractNotebookSessionSlides_(deck);

  assertTriangleDeck_(lessonTitle, deck);

  return {
    meta: {
      title: lessonTitle,
      standard: '',
      sourcePresentationId: presentation.getId(),
      sourcePresentationName: safeNotebookText_(presentation.getName()),
      sourceSlideCount: deck.slides.length
    },
    lessonTitle: lessonTitle,
    learningTargets: learningTargets,
    deck: deck,
    session1: {
      slides: sessions[0],
      mindset: findFirstMatchingLineInSlides_(sessions[0], /positive relationship with a classmate/i) ||
        'How do you establish a positive relationship with a classmate in math class?',
      anchors: {
        launch: buildFlagshipSourceAnchor_(deck, [6], {
          imageSlideNumber: 6,
          fallbackPrompt: 'What is the area of each of the triangles shown?'
        }),
        triangleAInfo: buildFlagshipSourceAnchor_(deck, [7, 8], {
          imageSlideNumber: 8,
          fallbackPrompt: 'What information do you know about Triangle A?'
        }),
        triangleACompose: buildFlagshipSourceAnchor_(deck, [10, 11], {
          imageSlideNumber: 10,
          revealSlideNumbers: [11],
          fallbackPrompt: 'How can you use what you know about finding the area of a parallelogram to find the area of Triangle A?',
          fallbackReveal: 'We can compose the triangle into a parallelogram to determine the area.'
        }),
        triangleB: buildFlagshipSourceAnchor_(deck, [20, 21], {
          imageSlideNumber: 20,
          revealSlideNumbers: [21],
          fallbackPrompt: 'Let’s determine the area of Triangle B using the area formula for triangles. What type of triangle is Triangle B? What does this tell you about the height of the triangle?',
          fallbackReveal: 'It is a right triangle, so the height is the vertical side connecting the base to the opposite vertex.'
        }),
        lucy: buildFlagshipSourceAnchor_(deck, [25], {
          imageSlideNumber: 25,
          fallbackPrompt: 'Lucy found the area of the triangle shown by first taking half of the height and then multiplying the result by the base. Is this an acceptable way to find the area of a triangle? Explain.'
        })
      }
    },
    session2: {
      slides: sessions[1],
      mindset: findFirstMatchingLineInSlides_(sessions[1], /positive relationship with a classmate/i) ||
        'How are you establishing a positive relationship with a classmate in math class?',
      anchors: {
        bermudaIntro: buildFlagshipSourceAnchor_(deck, [31], {
          imageSlideNumber: 31,
          revealSlideNumbers: [34],
          fallbackPrompt: 'The Bermuda Triangle is an area in the Atlantic Ocean between Miami, Florida, the island of Bermuda, and San Juan, Puerto Rico. It covers an area of about 437,409 square miles. What is the approximate distance from Bermuda to San Juan? What tool can you use to represent the problem?'
        }),
        bermudaRepresent: buildFlagshipSourceAnchor_(deck, [33, 34], {
          imageSlideNumber: 33,
          revealSlideNumbers: [34],
          fallbackPrompt: 'Step 1 Let’s create a visual representation of the problem. What is the problem asking you to find? How does this relate to the visual representation?',
          fallbackReveal: 'We need to find the distance from Bermuda to San Juan, which is the base of the triangle.'
        }),
        bermudaRelate: buildFlagshipSourceAnchor_(deck, [35, 36], {
          imageSlideNumber: 35,
          revealSlideNumbers: [36],
          fallbackPrompt: 'How can you relate the base, height, and area to find the missing measure?',
          fallbackReveal: 'We can use the area formula for triangles.'
        }),
        bermudaCollaborate: buildFlagshipSourceAnchor_(deck, [37, 38], {
          imageSlideNumber: 37,
          fallbackPrompt: 'Let’s use an area formula for triangles to find the base. What is the missing measure? What does this value represent?'
        }),
        bermudaSolve: buildFlagshipSourceAnchor_(deck, [39, 40], {
          imageSlideNumber: 39,
          fallbackPrompt: 'Let’s substitute the given values for the area and the height for A and h in the formula, then solve for b.',
          fallbackReveal: 'The approximate distance from Bermuda to San Juan, Puerto Rico, is 954 miles.'
        }),
        extension: buildFlagshipSourceAnchor_(deck, [42], {
          imageSlideNumber: 42,
          fallbackPrompt: 'What other information would you need to calculate the distance from San Juan to Miami using the area formula?'
        }),
        cabin: buildFlagshipSourceAnchor_(deck, [44], {
          imageSlideNumber: 44,
          fallbackPrompt: 'Use the cabin measurements to determine the paintable triangular area and the least amount Francisco should budget for paint.'
        })
      }
    }
  };
}

function buildFlagshipSourceAnchor_(deck, slideNumbers, options) {
  options = options || {};
  var promptLines = collectPromptLinesFromSlides_(deck, slideNumbers);
  return {
    slideNumbers: slideNumbers.slice(),
    slideLabel: formatSlideRangeLabel_(slideNumbers),
    prompt: combinePromptLinesFromSlides_(deck, slideNumbers, safeNotebookText_(options.fallbackPrompt)),
    promptLines: promptLines,
    questionLines: collectQuestionLinesFromSlides_(deck, slideNumbers),
    factLines: collectFactLinesFromSlides_(deck, slideNumbers),
    reveal: extractRevealTextFromSlides_(deck, options.revealSlideNumbers || slideNumbers, safeNotebookText_(options.fallbackReveal)),
    lines: promptLines,
    imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(deck, options.imageSlideNumber || slideNumbers[0]))
  };
}

function buildTriangleFlagshipBlueprint_(source) {
  return {
    meta: source.meta,
    sessions: [
      {
        sessionLabel: 'Session 1',
        lessonLabel: 'Session 1',
        sessionTitle: source.lessonTitle + ' — Session 1',
        objectives: {
          content: source.learningTargets[0] ||
            'I can find the area of a triangle by composing it into a parallelogram and by using a formula.',
          language: 'I can explain how base, height, and composing a triangle into a parallelogram prove the triangle area formula.'
        },
        mindset: source.session1.mindset,
        pages: buildTriangleSession1PagesFromSource_(source)
      },
      {
        sessionLabel: 'Session 2',
        lessonLabel: 'Session 2',
        sessionTitle: source.lessonTitle + ' — Session 2',
        objectives: {
          content: source.learningTargets[1] ||
            'I can attend to precision to find the missing dimension of a triangle by using the area formula.',
          language: 'I can explain how the area formula helps me find a missing dimension and justify whether my answer is reasonable in context.'
        },
        mindset: source.session2.mindset,
        pages: buildTriangleSession2PagesFromSource_(source)
      }
    ]
  };
}

function buildNotebookPackageFromBlueprint_(blueprint) {
  return {
    meta: blueprint.meta,
    sessions: blueprint.sessions.map(function(session) {
      return {
        sessionLabel: session.sessionLabel,
        lessonLabel: session.lessonLabel,
        sessionTitle: session.sessionTitle,
        objectives: session.objectives,
        mindset: session.mindset,
        sections: session.pages
      };
    })
  };
}

function validateTriangleBlueprint_(blueprint) {
  if (!blueprint || !blueprint.sessions || blueprint.sessions.length !== 2) {
    throw new Error('Flagship blueprint must contain exactly two sessions.');
  }

  for (var i = 0; i < blueprint.sessions.length; i++) {
    var session = blueprint.sessions[i];
    if (!session.pages || !session.pages.length) {
      throw new Error('Blueprint session ' + (i + 1) + ' is missing pages.');
    }
    if (session.pages.length < NOTEBOOK_BUILD_DEFAULTS_.MIN_SLIDES_PER_SESSION ||
        session.pages.length > NOTEBOOK_BUILD_DEFAULTS_.MAX_SLIDES_PER_SESSION) {
      throw new Error('Blueprint session ' + (i + 1) + ' must contain between ' +
        NOTEBOOK_BUILD_DEFAULTS_.MIN_SLIDES_PER_SESSION + ' and ' +
        NOTEBOOK_BUILD_DEFAULTS_.MAX_SLIDES_PER_SESSION + ' pages.');
    }
    for (var j = 0; j < session.pages.length; j++) {
      var page = session.pages[j];
      if (!page.type || !page.title) {
        throw new Error('Blueprint session ' + (i + 1) + ' has a page missing type or title.');
      }
      if ((page.type === 'sourceTask' || page.type === 'extensionChallenge' || page.type === 'collaborateDeep') && !page.problem) {
        throw new Error('Blueprint page "' + page.title + '" is missing its problem text.');
      }
      if (page.requiredFactsCount) {
        var factCount = (page.sourceLines || []).length + (page.factCards || []).length;
        if (factCount < page.requiredFactsCount) {
          throw new Error('Blueprint page "' + page.title + '" does not include enough visible source facts.');
        }
      }
    }
  }
}

function buildTriangleSession1PagesFromSource_(source) {
  var session = source.session1;
  var launch = session.anchors.launch;
  var triangleAInfo = session.anchors.triangleAInfo;
  var triangleACompose = session.anchors.triangleACompose;
  var lucy = session.anchors.lucy;
  var contentObjective = source.learningTargets[0] ||
    'I can find the area of a triangle by composing it into a parallelogram and by using a formula.';
  var languageObjective = 'I can explain how base, height, and composing a triangle into a parallelogram prove the triangle area formula.';

  return [
    {
      type: 'cover',
      title: source.lessonTitle,
      kicker: 'Session 1 Student Notebook',
      contentObjective: contentObjective,
      languageObjective: languageObjective,
      mindset: session.mindset
    },
    {
      type: 'sourceTask',
      title: 'Objectives + Success Criteria',
      kicker: 'Set the purpose',
      stageLabel: 'Learning Targets',
      sourceRefLabel: launch.slideLabel,
      problem: contentObjective,
      supportLabel: 'Language Objective',
      supportQuestion: languageObjective,
      sourceLines: [
        'Identify the base and the perpendicular height in each triangle.',
        'Explain how a triangle can be composed into a parallelogram.',
        'Use A = 1/2 × b × h and include precise square units.'
      ],
      visualKey: 'formula',
      strategyTitle: 'Track your goal',
      discussionQs: [
        'Which words in the objective tell you what mathematical action you must take today?',
        'How will you show that your answer is precise, not just correct?'
      ]
    },
    {
      type: 'beCurious',
      title: 'Be Curious',
      kicker: 'Notice + Wonder',
      imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(source.deck, 4)) || firstImageDataUriFromSlide_(getSlideByNumber_(source.deck, 5)),
      imageCaption: 'Use the image, sentence kernels, and vocabulary to start your thinking.',
      noticeKernels: [
        'I notice ___ because ___.',
        'One measurement that looks important is ___.'
      ],
      wonderKernels: [
        'I wonder ___ because ___.',
        'I predict the base or height will matter because ___.'
      ],
      vocabBox: [
        { term: 'triangle', definition: 'a polygon with three sides' },
        { term: 'base', definition: 'the side paired with a perpendicular height' },
        { term: 'height', definition: 'the perpendicular distance to the opposite vertex' },
        { term: 'area', definition: 'the number of square units inside a figure' }
      ]
    },
    {
      type: 'vocabulary',
      title: 'Academic Vocabulary',
      kicker: 'Keep these words visible',
      terms: [
        {
          term: 'Triangle',
          definition: 'A polygon with three sides and three angles.',
          studentFriendly: 'A triangle is a closed shape with three straight sides.',
          example: 'Triangle A and Triangle B are the figures we analyze in the lesson.',
          visualKey: 'triangle'
        },
        {
          term: 'Base',
          definition: 'A chosen side of the triangle used with the height to find area.',
          studentFriendly: 'The base is the side you decide to measure with a matching height.',
          example: 'The base is one side of Triangle A that matches the perpendicular height.',
          visualKey: 'base'
        },
        {
          term: 'Height',
          definition: 'The perpendicular distance from the base to the opposite vertex.',
          studentFriendly: 'The height must meet the base at a right angle.',
          example: 'In a right triangle, a vertical leg can serve as the height.',
          visualKey: 'height'
        },
        {
          term: 'Parallelogram',
          definition: 'A quadrilateral with two pairs of parallel sides.',
          studentFriendly: 'Two matching triangles can be combined to form a parallelogram.',
          example: 'The lesson composes Triangle A into a parallelogram to justify the formula.',
          visualKey: 'parallelogram'
        }
      ]
    },
    {
      type: 'vocabActivity',
      title: 'Interactive Vocabulary Rotation',
      kicker: 'Match the math language',
      sourceRefLabel: triangleAInfo.slideLabel,
      activityFamily: 'match_pair',
      activityInstructions: 'Match each clue to the correct academic word before you solve Triangle A.',
      activityBankTitle: 'Moveable Word Cards',
      bankIntro: 'Use each math word once. Then explain which word matters most for Triangle A.',
      dropzoneHint: 'Place the best word here',
      movablePieces: [
        'triangle',
        'base',
        'height',
        'parallelogram'
      ],
      answerCheck: 'Students should match triangle to the three-sided figure, base to the chosen side, height to the perpendicular measure, and parallelogram to the composed shape.',
      activities: [
        {
          label: 'Clue 1',
          prompt: 'The chosen side that will be paired with the perpendicular measure.'
        },
        {
          label: 'Clue 2',
          prompt: 'The perpendicular distance from the base to the opposite vertex.'
        },
        {
          label: 'Clue 3',
          prompt: 'The figure made when two congruent triangles are composed together.'
        },
        {
          label: 'Clue 4',
          prompt: 'The three-sided polygon students are analyzing in this lesson.'
        }
      ],
      discussionQs: [
        'Which word is easiest to picture, and which word needs the most careful explanation?',
        'How do base, height, and parallelogram work together in the Triangle A example?'
      ]
    },
    {
      type: 'sourceTask',
      title: 'First Problem: Study Triangle A',
      kicker: 'Anchor problem',
      stageLabel: 'Problem Launch',
      sourceRefLabel: launch.slideLabel,
      problem: launch.prompt || 'What is the area of each of the triangles shown?',
      supportLabel: 'Look Closely',
      supportQuestion: 'Before solving, what is the problem asking you to find and what information do you already notice about Triangle A?',
      sourceLines: [
        'The first source problem asks for the area of the triangles shown.',
        'Stay focused on Triangle A first so you can build one clear strategy.',
        'Use the diagram to identify a base and a perpendicular height before calculating.'
      ],
      factCards: [
        'Start with Triangle A first',
        'Goal: find the area',
        'Look for a base-height pair',
        'Use the source diagram to gather givens'
      ],
      requiredFactsCount: 4,
      imageDataUri: launch.imageDataUri,
      visualKey: 'triangle',
      strategyTitle: 'Read the problem first',
      discussionQs: [
        'What is the first problem actually asking you to find?',
        'Why is it helpful to focus on one triangle before trying to solve both?'
      ]
    },
    {
      type: 'sourceTask',
      title: 'Guided Notes: Triangle A',
      kicker: 'Read the givens',
      stageLabel: 'Guided Notes',
      sourceRefLabel: triangleAInfo.slideLabel,
      problem: triangleAInfo.prompt,
      supportLabel: 'Think About It',
      supportQuestion: 'Which measurement is the base and which measurement is the perpendicular height in Triangle A?',
      sourceLines: [
        'Triangle A has a base of 30 centimeters.',
        'The dashed perpendicular segment shows a height of 15 centimeters.',
        'Use the source diagram to match the height to the chosen base.'
      ],
      factCards: [
        'Triangle A: base = 30 cm',
        'Triangle A: height = 15 cm',
        'The dashed segment is perpendicular.',
        'Area will be written in square centimeters.'
      ],
      requiredFactsCount: 5,
      revealLabel: 'Reveal',
      reveal: 'Triangle A uses a base of 30 centimeters and a perpendicular height of 15 centimeters.',
      imageDataUri: triangleAInfo.imageDataUri || firstImageDataUriFromSlide_(getSlideByNumber_(source.deck, 8)),
      visualKey: 'triangle',
      strategyTitle: 'Label the base and height',
      discussionQs: [
        'What information do you know about Triangle A before you use a formula?',
        'Why does the height have to be perpendicular to the base?'
      ]
    },
    {
      type: 'vocabActivity',
      title: 'Interactive Activity: Sort the Givens',
      kicker: 'Given now or needed later?',
      sourceRefLabel: triangleAInfo.slideLabel,
      activityFamily: 'sort_classify',
      activityInstructions: 'Sort each card into the correct category so Triangle A is ready to solve.',
      activityBankTitle: 'Given + Need Sort',
      bankIntro: 'Use the source diagram and your notes to decide which facts are already given and which ideas help you solve.',
      dropzoneHint: 'Sort the matching cards here',
      movablePieces: [
        '30 cm base',
        '15 cm height',
        'perpendicular height',
        'A = 1/2 × b × h',
        'square centimeters',
        'unknown area of Triangle A'
      ],
      answerCheck: 'The given measurements are the 30 cm base and 15 cm height. The needed idea is the unknown area, supported by the formula and square-centimeter unit.',
      activities: [
        {
          label: 'Already Given',
          prompt: 'Sort the cards that come directly from the source diagram or the labeled figure.'
        },
        {
          label: 'Needed to Solve',
          prompt: 'Sort the cards that help you set up, calculate, or label the area of Triangle A.'
        }
      ],
      discussionQs: [
        'Which cards came directly from the source diagram, and which ones came from your math knowledge?',
        'Why is it important to separate the givens from the tools you will use to solve?'
      ]
    },
    {
      type: 'formulaDerivation',
      title: 'Build the Formula with Triangle A',
      kicker: 'Compose the triangle into a parallelogram',
      sourceRefLabel: triangleACompose.slideLabel,
      steps: [
        {
          title: 'Identify the givens',
          visualKey: 'height',
          explanation: 'Triangle A shows a base of 30 centimeters and a perpendicular height of 15 centimeters.'
        },
        {
          title: 'Compose a new figure',
          visualKey: 'compose',
          explanation: triangleACompose.prompt + ' ' + triangleACompose.reveal
        },
        {
          title: 'Connect the model to the formula',
          visualKey: 'formula',
          explanation: 'Two congruent triangles make one parallelogram, so one triangle has half the area. That is why A = 1/2 × b × h.'
        }
      ],
      discussionQs: [
        'Why do two copies of the triangle create a parallelogram with the same base and height?',
        'How does the composed shape prove that a triangle is half of b × h?'
      ]
    },
    {
      type: 'vocabActivity',
      title: 'Interactive Activity: Build the Triangle A Formula',
      kicker: 'Givens, model, formula',
      sourceRefLabel: triangleACompose.slideLabel,
      activityFamily: 'build_construct',
      activityInstructions: 'Use the cards to build the full story of Triangle A: the givens, the composed model, and the formula.',
      activityBankTitle: 'Triangle A Builder',
      bankIntro: 'Start with the givens, then connect the composed parallelogram to the one-half formula.',
      dropzoneHint: 'Place two matching cards here',
      movablePieces: [
        '30 cm base',
        '15 cm height',
        'copy and rotate',
        'make a parallelogram',
        'triangle is half',
        'A = 1/2 × b × h'
      ],
      answerCheck: 'A strong arrangement identifies the two givens, shows that two copies make a parallelogram, and connects that model to A = 1/2 × b × h.',
      activities: [
        {
          label: 'Givens',
          prompt: 'Place the cards that show the two measurements given for Triangle A.'
        },
        {
          label: 'Model',
          prompt: 'Place the cards that show how Triangle A is composed into a parallelogram.'
        },
        {
          label: 'Formula',
          prompt: 'Place the cards that explain why the triangle area formula uses one-half.'
        }
      ],
      discussionQs: [
        'Which card helps you prove the one-half in the formula most clearly?',
        'How do the given measurements and the composed model work together to justify A = 1/2 × b × h?'
      ]
    },
    {
      type: 'sourceTask',
      title: 'Solve Triangle A',
      kicker: 'Apply the formula',
      stageLabel: 'Solve',
      sourceRefLabel: triangleACompose.slideLabel,
      problem: 'Use Triangle A from the first problem. What is the area of Triangle A?',
      supportLabel: 'Show Your Setup',
      supportQuestion: 'How will you substitute the 30-centimeter base and 15-centimeter height into A = 1/2 × b × h?',
      sourceLines: [
        'Substitute b = 30 cm and h = 15 cm.',
        'Multiply 30 × 15 first, then take half of the product.',
        'Write the final answer in square centimeters.',
        'Use the model from the previous slides to justify why the formula works.'
      ],
      factCards: [
        'b = 30 cm',
        'h = 15 cm',
        'A = 1/2 × b × h',
        'Answer in square centimeters'
      ],
      requiredFactsCount: 7,
      revealLabel: 'Reveal',
      reveal: 'Triangle A has an area of 225 square centimeters.',
      imageDataUri: triangleAInfo.imageDataUri || launch.imageDataUri,
      visualKey: 'formula',
      strategyTitle: 'Set up, solve, label',
      discussionQs: [
        'Why is 225 the correct area for Triangle A?',
        'How does the model help you trust the formula answer?'
      ]
    },
    {
      type: 'extensionChallenge',
      title: 'Error Analysis: Lucy’s Strategy',
      kicker: 'Let’s Explore More',
      sourceRefLabel: lucy.slideLabel,
      problem: lucy.prompt,
      hint: 'Compare (1/2 × h) × b with 1/2 × b × h. Decide whether the order changes the product and justify your conclusion with the formula or a model.',
      discussionQs: [
        'When would Lucy’s method produce the same result as the standard triangle formula?',
        'How could you prove your answer with numbers, labels, or a diagram?'
      ]
    },
    {
      type: 'summary',
      title: 'Summary: Triangle A Strategy',
      kicker: 'Wrap up the big idea',
      summaryStem: 'A strong triangle-area strategy starts with the givens, uses the model, and ends with a precise labeled answer.',
      bullets: [
        'Triangle A uses a base of 30 centimeters and a perpendicular height of 15 centimeters.',
        'Composing two copies into a parallelogram explains why the formula includes one-half.',
        'The area of Triangle A is 225 square centimeters.'
      ],
      formulaTitle: 'Triangle A Takeaway',
      mathLines: [
        'A = 1/2 × b × h',
        'A = 1/2 × 30 × 15',
        'A = 225 cm²'
      ],
      comparePrompt: 'How did the source diagram, the composed model, and the formula all work together to solve Triangle A?',
      discussionQs: [
        'Which part of the Triangle A strategy do you want to remember next time?',
        'How does the model make the formula easier to explain?'
      ]
    },
    {
      type: 'reflection',
      title: 'Session 1 Reflection + Goal Tracker',
      kicker: 'Write about your learning',
      prompts: [
        'I can now solve a triangle-area problem by...',
        'One part of Triangle A that helped me most was...',
        'My goal for the next lesson is...'
      ]
    }
  ];
}

function buildTriangleSession2PagesFromSource_(source) {
  var session = source.session2;
  var intro = session.anchors.bermudaIntro;
  var represent = session.anchors.bermudaRepresent;
  var relate = session.anchors.bermudaRelate;
  var collaborate = session.anchors.bermudaCollaborate;
  var solve = session.anchors.bermudaSolve;
  var cabin = session.anchors.cabin;
  var extension = session.anchors.extension;
  var contentObjective = source.learningTargets[1] ||
    'I can attend to precision to find the missing dimension of a triangle by using the area formula.';
  var languageObjective = 'I can explain how the area formula helps me find a missing dimension and justify whether my answer is reasonable in context.';

  return [
    {
      type: 'cover',
      title: source.lessonTitle,
      kicker: 'Session 2 Student Notebook',
      contentObjective: contentObjective,
      languageObjective: languageObjective,
      mindset: session.mindset
    },
    {
      type: 'sourceTask',
      title: 'Objectives + Success Criteria',
      kicker: 'Set the purpose',
      stageLabel: 'Learning Targets',
      problem: contentObjective,
      supportLabel: 'Language Objective',
      supportQuestion: languageObjective,
      sourceLines: [
        'Decide whether the problem needs an estimate or an exact answer.',
        'Represent the triangle and identify the unknown dimension clearly.',
        'Use the area formula, solve for the missing measure, and interpret the units.'
      ],
      factCards: [
        'Missing dimension',
        'Use A = 1/2 × b × h',
        'Interpret the answer in context',
        'Keep units visible'
      ],
      requiredFactsCount: 4,
      visualKey: 'missingdimension',
      strategyTitle: 'Track your goal',
      discussionQs: [
        'What does the phrase missing dimension tell you about today’s work?',
        'How will you know whether an answer is reasonable in context?'
      ]
    },
    {
      type: 'beCurious',
      title: 'Be Curious',
      kicker: 'What could the question be?',
      imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(source.deck, 27)) || firstImageDataUriFromSlide_(getSlideByNumber_(source.deck, 28)),
      imageCaption: 'Use the image, sentence kernels, and vocabulary to predict the question.',
      noticeKernels: [
        'I notice ___ because ___.',
        'A measurement that looks important is ___.'
      ],
      wonderKernels: [
        'I wonder what the question could be because ___.',
        'I predict the missing measure is ___ because ___.'
      ],
      vocabBox: [
        { term: 'approximate', definition: 'close to the exact value, not perfectly exact' },
        { term: 'precision', definition: 'careful and accurate use of numbers, labels, and units' },
        { term: 'base', definition: 'the side being solved for in the Bermuda Triangle task' },
        { term: 'height', definition: 'the perpendicular measure paired with the base' }
      ]
    },
    {
      type: 'vocabulary',
      title: 'Vocabulary for Precision',
      kicker: 'Keep these words visible',
      terms: [
        {
          term: 'Approximate',
          definition: 'Close to an exact value but not perfectly exact.',
          studentFriendly: 'An approximate answer is a sensible estimate based on the information given.',
          example: 'The distance from Bermuda to San Juan is about 954 miles.',
          visualKey: 'approximate'
        },
        {
          term: 'Precision',
          definition: 'Carefully using the correct numbers, units, and labels.',
          studentFriendly: 'Precision means your work is clearly labeled and mathematically accurate.',
          example: 'The source asks students to attend to precision when finding a missing dimension.',
          visualKey: 'precision'
        },
        {
          term: 'Base',
          definition: 'The side of the triangle paired with the height in the area formula.',
          studentFriendly: 'In the Bermuda Triangle task, the unknown distance is the base.',
          example: 'The distance from Bermuda to San Juan is treated as the base.',
          visualKey: 'base'
        },
        {
          term: 'Missing Dimension',
          definition: 'A side length or measure that must be solved for using known information.',
          studentFriendly: 'A missing dimension is the part of the figure you do not know yet.',
          example: 'Session 2 uses the area formula to find the missing base.',
          visualKey: 'missingdimension'
        }
      ]
    },
    {
      type: 'sourceTask',
      title: 'Guided Notes: Represent the Bermuda Triangle',
      kicker: 'Set up the problem',
      stageLabel: 'Guided Notes',
      sourceRefLabel: represent.slideLabel,
      problem: intro.prompt,
      supportLabel: 'Context + Precision',
      supportQuestion: 'Does the problem call for an estimate or an exact solution?',
      sourceLines: [
        'Use a visual representation of the triangle before you solve.',
        'A = 437,409 square miles.',
        'h = 917 miles.',
        'b = ? represents the distance from Bermuda to San Juan.'
      ],
      factCards: [
        'A = 437,409 sq mi',
        'h = 917 mi',
        'b = ?',
        'Answer is approximate'
      ],
      requiredFactsCount: 7,
      revealLabel: 'Reveal',
      reveal: represent.reveal,
      imageDataUri: intro.imageDataUri || represent.imageDataUri,
      visualKey: 'triangle',
      strategyTitle: 'Represent the problem',
      discussionQs: [
        'Why is the missing distance the base in this context?',
        'Which values are known before you start solving with the formula?'
      ]
    },
    {
      type: 'formulaDerivation',
      title: 'Build the Missing-Base Strategy',
      kicker: 'Guided practice',
      sourceRefLabel: relate.slideLabel,
      steps: [
        {
          title: 'Create the visual representation',
          visualKey: 'triangle',
          explanation: represent.prompt + ' ' + represent.reveal
        },
        {
          title: 'Use the area formula for triangles',
          visualKey: 'formula',
          explanation: relate.prompt + ' ' + relate.reveal
        },
        {
          title: 'Interpret the missing measure',
          visualKey: 'missingdimension',
          explanation: solve.reveal || 'The approximate distance from Bermuda to San Juan, Puerto Rico, is 954 miles.'
        }
      ],
      discussionQs: [
        'Why do we multiply the area by 2 before dividing by the height?',
        'What does the solved value represent in the real-world situation?'
      ]
    },
    {
      type: 'vocabActivity',
      title: 'Interactive Activity: Build the Bermuda Equation',
      kicker: 'Facts, equation, meaning',
      sourceRefLabel: solve.slideLabel,
      activityFamily: 'build_construct',
      activityInstructions: 'Build the path from the Bermuda Triangle facts to the missing base.',
      activityBankTitle: 'Equation Builder',
      bankIntro: 'Use each card to show the givens, the equation, and the meaning of the solution.',
      dropzoneHint: 'Place two matching cards here',
      movablePieces: [
        'A = 437,409 sq mi',
        'h = 917 mi',
        '437,409 = 1/2 × b × 917',
        'multiply by 2, then divide by 917',
        'b = 954',
        '954 miles from Bermuda to San Juan'
      ],
      answerCheck: 'A strong build uses the area and height as the givens, writes the substitution equation correctly, and connects 954 to the Bermuda-to-San-Juan distance in miles.',
      activities: [
        {
          label: 'Givens',
          prompt: 'Place the cards that show the two values taken directly from the source problem.'
        },
        {
          label: 'Equation',
          prompt: 'Place the cards that show the substitution equation and the algebra move used to isolate b.'
        },
        {
          label: 'Meaning',
          prompt: 'Place the cards that show the solved value and what it represents in the context of the map.'
        }
      ],
      discussionQs: [
        'Which card shows the actual substitution step, and which card shows the algebra move after substitution?',
        'Why is 954 not enough by itself until you explain what it represents?'
      ]
    },
    {
      type: 'sourceTask',
      title: 'Substitute and Solve for b',
      kicker: 'Think About It',
      stageLabel: 'Think About It',
      sourceRefLabel: solve.slideLabel,
      problem: solve.prompt,
      supportLabel: 'Explain',
      supportQuestion: 'Why can you write the triangle area formula as A = 1/2 × b × h or A = (b × h) ÷ 2?',
      sourceLines: [
        'Substitute A = 437,409 square miles into the formula.',
        'Use h = 917 miles from the source diagram.',
        'Solve for b, the missing base.',
        'Write the final distance in miles and connect it back to the map.'
      ],
      factCards: [
        '437,409 = 1/2 × b × 917',
        'Multiply the area by 2',
        'Divide by 917',
        'b is measured in miles'
      ],
      requiredFactsCount: 7,
      revealLabel: 'Reveal',
      reveal: solve.reveal || 'The approximate distance from Bermuda to San Juan, Puerto Rico, is 954 miles.',
      imageDataUri: solve.imageDataUri,
      visualKey: 'formula',
      strategyTitle: 'Substitute and solve',
      discussionQs: [
        'Why do you substitute the known values before solving for b?',
        'How do you know the final answer should be written in miles?'
      ]
    },
    {
      type: 'collaborateDeep',
      title: 'Collaborate and Connect',
      kicker: 'Explain the meaning',
      sourceRefLabel: collaborate.slideLabel,
      problem: collaborate.prompt,
      partnerRoles: [
        'Partner A: Solve the formula',
        'Partner B: Interpret the context'
      ],
      partnerTasks: [
        'Use A = 437,409 and h = 917 to explain the steps that isolate b in A = 1/2 × b × h and identify the missing measure.',
        'Explain why the solution represents miles and why the answer is approximate, not exact.'
      ],
      discussionQs: [
        'How does the algebra show that you solved for the base instead of the area?',
        'Why is a numerical answer incomplete until you explain what it represents?'
      ]
    },
    {
      type: 'sourceTask',
      title: 'Application: Painting a Cabin',
      kicker: 'Apply the strategy',
      stageLabel: 'Independent Practice',
      sourceRefLabel: cabin.slideLabel,
      problem: cabin.prompt,
      supportLabel: 'Use the Source Information',
      supportQuestion: 'How do you find the triangular area first, then subtract the window areas, then decide how many full gallons are needed?',
      sourceLines: [
        'The triangular section has a base of 34 feet and a height of 23 3/4 feet.',
        'The top window is 5 feet by 2 feet.',
        'The bottom window is 13 1/2 feet by 5 feet.',
        'One gallon costs $24.95 and covers 200 square feet.',
        'Subtract both window areas before you determine the paint cost.'
      ],
      factCards: [
        'Triangle: b = 34 ft, h = 23 3/4 ft',
        'Top window: 5 ft by 2 ft',
        'Bottom window: 13 1/2 ft by 5 ft',
        '$24.95 per gallon; 200 sq ft coverage'
      ],
      requiredFactsCount: 8,
      imageDataUri: cabin.imageDataUri,
      visualKey: 'formula',
      strategyTitle: 'Make a plan',
      discussionQs: [
        'Why do the window areas need to be removed before calculating the amount of paint?',
        'How will you decide whether Francisco needs one gallon or more than one gallon of paint?'
      ]
    },
    {
      type: 'collaborateDeep',
      title: 'Application Planner: Compare and Decide',
      kicker: 'Reason through both problems',
      sourceRefLabel: cabin.slideLabel,
      problem: 'Compare the Bermuda Triangle problem and the cabin problem. How does the same triangle-area formula lead to two different kinds of answers?',
      partnerRoles: [
        'Partner A: Bermuda Triangle',
        'Partner B: Painting a Cabin'
      ],
      partnerTasks: [
        'Explain how the Bermuda Triangle task uses the area formula to solve for a missing base and interpret the answer in miles.',
        'Explain how the cabin task uses triangle area, subtraction, and paint coverage to make a budgeting decision.'
      ],
      discussionQs: [
        'How is solving for a missing base different from using area to make a decision about paint?',
        'What do both problems have in common even though the final answers mean different things?'
      ]
    },
    {
      type: 'extensionChallenge',
      title: 'Extension Challenge: Go Further',
      kicker: 'Let’s Explore More',
      sourceRefLabel: extension.slideLabel,
      problem: extension.prompt,
      hint: 'Think about which measurement is still unknown and why multiplying by one-half and dividing by two are equivalent operations.',
      discussionQs: [
        'Which missing measure stops you from finding the distance from San Juan to Miami right now?',
        'How would you explain the equivalence of multiplying by 1/2 and dividing by 2 to another student?'
      ]
    },
    {
      type: 'summary',
      title: 'Summary: Missing Dimension + Application',
      kicker: 'Wrap up the big idea',
      summaryStem: 'We can use the triangle area formula to solve for a missing dimension and to make decisions in real situations.',
      bullets: [
        'Use A = 1/2 × b × h to connect the area, base, and height.',
        'When the base is missing, isolate b and interpret the answer in context.',
        'In application problems, you may need extra steps after finding area, such as subtracting windows or deciding how many gallons to buy.'
      ],
      formulaTitle: 'Key Structure',
      mathLines: [
        'A = 1/2 × b × h',
        '437,409 = 1/2(b)(917)',
        'b = (2A) ÷ h'
      ],
      comparePrompt: 'How do the Bermuda Triangle and cabin tasks use the same formula in different ways?',
      discussionQs: [
        'When do you solve for a missing dimension, and when do you use area to make a decision?',
        'Why does context matter when you interpret the final answer?'
      ]
    },
    {
      type: 'reflection',
      title: 'Session 2 Exit Ticket + Goal Tracker',
      kicker: 'Explain and reflect',
      prompts: [
        'I can find a missing dimension of a triangle by...',
        'The value 954 represents ... because ...',
        'My confidence with triangle area today is ... and my next step is ...'
      ]
    }
  ];
}

function buildTriangleSession1Spec_LEGACY_UNUSED_(deck, sessionSlides, learningTargets) {
  var sessionLabel = 'Session 1';
  var contentObjective = learningTargets[0] ||
    'I can find the area of a triangle by composing it into a parallelogram and by using a formula.';
  var languageObjective =
    'I can explain how base, height, and composing a triangle into a parallelogram prove the triangle area formula.';
  var mindset = findFirstMatchingLineInSlides_(sessionSlides, /positive relationship with a classmate/i) ||
    'How do you establish a positive relationship with a classmate in math class?';
  var triangleAInfoPrompt = combinePromptLinesFromSlides_(deck, [7, 8],
    'What information do you know about Triangle A?');
  var triangleAComposePrompt = combinePromptLinesFromSlides_(deck, [10],
    'How can you use what you know about finding the area of a parallelogram to find the area of Triangle A?');
  var triangleAComposeReveal = extractRevealTextFromSlides_(deck, [11],
    'We can compose the triangle into a parallelogram to determine the area.');
  var triangleBPrompt = combinePromptLinesFromSlides_(deck, [20],
    'Let’s determine the area of Triangle B using the area formula for triangles. What type of triangle is Triangle B? What does this tell you about the height of the triangle?');
  var triangleBReveal = extractRevealTextFromSlides_(deck, [21],
    'It is a right triangle, so the height of the triangle is the same as the vertical side connecting the base to the opposite vertex.');
  var symbolsPrompt = combinePromptLinesFromSlides_(deck, [23, 24],
    'What is the area of each of the triangles shown? The area of Triangle A is 225 square centimeters. The area of Triangle B is 70 square inches. What units of measure are needed to be precise?');
  var lucyPrompt = combinePromptLinesFromSlides_(deck, [25],
    'Lucy found the area of the triangle shown by first taking half of the height and then multiplying the result by the base. Is this an acceptable way to find the area of a triangle? Explain.');

  return {
    sessionLabel: sessionLabel,
    lessonLabel: sessionLabel,
    sessionTitle: 'Determine the Area of Triangles — Session 1',
    objectives: {
      content: contentObjective,
      language: languageObjective
    },
    mindset: mindset,
    sections: [
      {
        type: 'cover',
        title: deck.lessonTitle,
        kicker: 'Session 1 Student Notebook',
        contentObjective: contentObjective,
        languageObjective: languageObjective,
        mindset: mindset
      },
      {
        type: 'sourceTask',
        title: 'Objectives + Success Criteria',
        kicker: 'Set the purpose',
        stageLabel: 'Learning Targets',
        problem: contentObjective,
        supportLabel: 'Language Objective',
        supportQuestion: languageObjective,
        sourceLines: [
          'Identify the base and the perpendicular height in each triangle.',
          'Explain how a triangle can be composed into a parallelogram.',
          'Use A = 1/2 × b × h and include precise square units.'
        ],
        visualKey: 'formula',
        strategyTitle: 'Track your goal',
        discussionQs: [
          'Which words in the objective tell you what mathematical action you must take today?',
          'How will you show that your answer is precise, not just correct?'
        ]
      },
      {
        type: 'beCurious',
        title: 'Be Curious',
        kicker: 'Notice + Wonder',
        imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(deck, 4)) || firstImageDataUriFromSlide_(getSlideByNumber_(deck, 5)),
        imageCaption: 'Use the image, sentence kernels, and vocabulary to start your thinking.',
        noticeKernels: [
          'I notice ___ because ___.',
          'One measurement that looks important is ___.'
        ],
        wonderKernels: [
          'I wonder ___ because ___.',
          'I predict the base or height will matter because ___.'
        ],
        vocabBox: [
          { term: 'triangle', definition: 'a polygon with three sides' },
          { term: 'base', definition: 'the side paired with a perpendicular height' },
          { term: 'height', definition: 'the perpendicular distance to the opposite vertex' },
          { term: 'area', definition: 'the number of square units inside a figure' }
        ]
      },
      {
        type: 'vocabulary',
        title: 'Academic Vocabulary',
        kicker: 'Keep these words visible',
        terms: [
          {
            term: 'Triangle',
            definition: 'A polygon with three sides and three angles.',
            studentFriendly: 'A triangle is a closed shape with three straight sides.',
            example: 'Triangle A and Triangle B are the figures we analyze in the lesson.',
            visualKey: 'triangle'
          },
          {
            term: 'Base',
            definition: 'A chosen side of the triangle used with the height to find area.',
            studentFriendly: 'The base is the side you decide to measure with a matching height.',
            example: 'The base is one side of Triangle A that matches the perpendicular height.',
            visualKey: 'base'
          },
          {
            term: 'Height',
            definition: 'The perpendicular distance from the base to the opposite vertex.',
            studentFriendly: 'The height must meet the base at a right angle.',
            example: 'In a right triangle, a vertical leg can serve as the height.',
            visualKey: 'height'
          },
          {
            term: 'Parallelogram',
            definition: 'A quadrilateral with two pairs of parallel sides.',
            studentFriendly: 'Two matching triangles can be combined to form a parallelogram.',
            example: 'The lesson composes Triangle A into a parallelogram to justify the formula.',
            visualKey: 'parallelogram'
          }
        ]
      },
      {
        type: 'vocabActivity',
        title: 'Interactive Vocabulary Rotation',
        kicker: 'Rotate, match, explain',
        activityFamily: 'match_pair',
        activityInstructions: 'Rotate through the four boxes. Drag each card to the best match, then explain one connection aloud.',
        activityBankTitle: 'Vocabulary Rotation',
        bankIntro: 'Match the movable cards, then use the vocabulary in a complete math sentence.',
        movablePieces: [
          'triangle',
          'base',
          'height',
          'three-sided polygon',
          'chosen side used with height',
          'perpendicular distance'
        ],
        answerCheck: 'A strong match connects triangle to three-sided polygon, base to the chosen side used with height, and height to perpendicular distance.',
        activities: [
          {
            label: 'Word',
            prompt: 'Place the vocabulary word cards that name the key ideas from the lesson.'
          },
          {
            label: 'Meaning',
            prompt: 'Match each word to the meaning that fits it best.'
          },
          {
            label: 'Use It',
            prompt: 'Use one matched pair in a complete sentence about triangle area.'
          }
        ],
        discussionQs: [
          'Which vocabulary word is easiest to picture in the diagram, and which one needs the most explanation?',
          'How do base, height, and parallelogram connect to the triangle area formula?'
        ]
      },
      {
        type: 'sourceTask',
        title: 'Guided Notes: Triangle A',
        kicker: 'Read the diagram',
        stageLabel: 'Guided Notes',
        problem: triangleAInfoPrompt,
        supportLabel: 'Think About It',
        supportQuestion: 'Which measurement is the base and which measurement is the perpendicular height in Triangle A?',
        sourceLines: [
          'Triangle A has a base of 30 centimeters.',
          'The dashed perpendicular segment shows a height of 15 centimeters.',
          'Use the source diagram to match the height to the chosen base.',
          'This slide focuses on the given information you need before you calculate the area.'
        ],
        revealLabel: 'Reveal',
        reveal: 'Triangle A uses a base of 30 centimeters and a perpendicular height of 15 centimeters.',
        imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(deck, 6)) || firstImageDataUriFromSlide_(getSlideByNumber_(deck, 8)) || firstImageDataUriFromSlide_(getSlideByNumber_(deck, 9)),
        visualKey: 'triangle',
        strategyTitle: 'Label the base and height',
        discussionQs: [
          'What information do you know about Triangle A before you use a formula?',
          'Why does the height have to be perpendicular to the base?'
        ]
      },
      {
        type: 'formulaDerivation',
        title: 'Build the Formula with Triangle A',
        kicker: 'Compose the triangle into a parallelogram',
        steps: [
          {
            title: 'What information do you know about Triangle A?',
            visualKey: 'height',
            explanation: 'Triangle A shows a base of 30 centimeters and a perpendicular height of 15 centimeters. Those are the measurements you use to determine the area.'
          },
          {
            title: 'How can you compose the triangle into a parallelogram?',
            visualKey: 'compose',
            explanation: triangleAComposePrompt + ' ' + triangleAComposeReveal
          },
          {
            title: 'How does this help you develop the formula?',
            visualKey: 'formula',
            explanation: 'A parallelogram built from two copies of Triangle A has area 30 × 15. One triangle is half of that area, which is why the formula is A = 1/2 × b × h.'
          }
        ],
        discussionQs: [
          'Why do two copies of the triangle create a parallelogram with the same base and height?',
          'How does the composed shape prove that a triangle is half of b × h?'
        ]
      },
      {
        type: 'vocabActivity',
        title: 'Interactive Activity: Build the Triangle A Formula',
        kicker: 'Given, model, compare, explain',
        activityFamily: 'build_construct',
        activityInstructions: 'Drag the cards into the matching boxes to show the givens, the model, the area relationship, and the formula for Triangle A.',
        activityBankTitle: 'Drag + Build',
        bankIntro: 'Use the movable cards to reconstruct the logic of the Triangle A example.',
        dropzoneHint: 'Place two matching cards here',
        movablePieces: [
          '30 cm base',
          '15 cm height',
          'copy and rotate',
          'make a parallelogram',
          'triangle is half',
          'same base and height',
          'A = 1/2 × b × h'
        ],
        answerCheck: 'A strong arrangement identifies the 30-centimeter base and 15-centimeter height, shows that two copies make a parallelogram, and connects that model to A = 1/2 × b × h.',
        activities: [
          {
            label: 'Givens',
            prompt: 'Place the cards that show the two measurements given for Triangle A.'
          },
          {
            label: 'Model',
            prompt: 'Place the cards that show how Triangle A is composed into a parallelogram.'
          },
          {
            label: 'Formula',
            prompt: 'Place the cards that connect the model to the triangle area formula.'
          }
        ],
        discussionQs: [
          'Which card helps you prove the one-half in the formula most clearly?',
          'How do the given measurements and the model work together to justify A = 1/2 × b × h?'
        ]
      },
      {
        type: 'sourceTask',
        title: 'Guided Practice: Triangle B',
        kicker: 'Use the formula',
        stageLabel: 'Guided Practice',
        problem: triangleBPrompt,
        supportLabel: 'Think About It',
        supportQuestion: 'If Triangle B is a right triangle, which measurement in the diagram already gives the height?',
        sourceLines: [
          'Triangle B is a right triangle.',
          'The base is 10 inches.',
          'The vertical height is 14 inches.',
          'Use A = 1/2 × 10 × 14 and write the area in square inches.'
        ],
        revealLabel: 'Reveal',
        reveal: triangleBReveal,
        imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(deck, 20)),
        visualKey: 'rightTriangle',
        strategyTitle: 'Show your work',
        discussionQs: [
          'How does the right angle help you identify the height immediately?',
          'What would you label first before substituting values into the formula?'
        ]
      },
      {
        type: 'collaborateDeep',
        title: 'Strategy Studio: Triangle B',
        kicker: 'Compare and defend',
        problem: 'After solving Triangle B, compare how you used the 10-inch base and 14-inch height. Which step mattered most: identifying the right triangle, using the vertical height, or substituting into the formula?',
        partnerRoles: [
          'Partner A: Solve and justify',
          'Partner B: Check and compare'
        ],
        partnerTasks: [
          'Show how you identified the 10-inch base and the 14-inch height, then explain why that choice is valid for Triangle B.',
          'Listen for precision. Check whether the explanation includes the right angle, the formula, the correct multiplication, and the final square-inch unit.'
        ],
        discussionQs: [
          'Which step in the Triangle B solution would be easiest to skip by accident, and why?',
          'How can two correct strategies still sound different when students explain them?'
        ]
      },
      {
        type: 'collaborateDeep',
        title: 'Precision Clinic',
        kicker: 'Units + explanation',
        problem: symbolsPrompt,
        partnerRoles: [
          'Partner A: Match the units',
          'Partner B: Defend the units'
        ],
        partnerTasks: [
          'Match 225 to Triangle A and 70 to Triangle B, then attach the correct square unit to each answer.',
          'Explain why centimeters become square centimeters and inches become square inches when you find area.'
        ],
        discussionQs: [
          'Why does area need square units instead of linear units?',
          'How can precise units keep two correct calculations from becoming confusing?'
        ]
      },
      {
        type: 'extensionChallenge',
        title: 'Error Analysis: Lucy’s Strategy',
        kicker: 'Let’s Explore More',
        problem: lucyPrompt,
        hint: 'Compare (1/2 × h) × b with 1/2 × b × h. Decide whether the order changes the product and justify your conclusion with the formula or a model.',
        discussionQs: [
          'When would Lucy’s method produce the same result as the standard triangle formula?',
          'How could you prove your answer with numbers, labels, or a diagram?'
        ]
      },
      {
        type: 'summary',
        title: 'Summary: Triangle Area',
        kicker: 'Wrap up the big idea',
        summaryStem: 'We can write the formula that reflects half the area of a parallelogram.',
        bullets: [
          'The area of a triangle is half the area of the parallelogram made from two congruent copies.',
          'The formula for triangle area is A = 1/2 × b × h.',
          'Use precise square units when you write the final area.'
        ],
        formulaTitle: 'Area of Triangles',
        mathLines: [
          'A = 1/2 × b × h',
          'Triangle area = 1/2(parallelogram area)',
          'Use square units'
        ],
        comparePrompt: 'How does composing a triangle into a parallelogram justify the triangle area formula?',
        discussionQs: [
          'Why is a triangle’s area half of the composed parallelogram’s area?',
          'How does the model help you remember the formula?'
        ]
      },
      {
        type: 'reflection',
        title: 'Session 1 Reflection + Goal Tracker',
        kicker: 'Write about your learning',
        prompts: [
          'I can now find the area of a triangle by...',
          'One way I showed precision today was...',
          'My goal for the next lesson is...'
        ]
      }
    ]
  };
}

function buildTriangleSession2Spec_LEGACY_UNUSED_(deck, sessionSlides, learningTargets) {
  var sessionLabel = 'Session 2';
  var contentObjective = learningTargets[1] ||
    'I can attend to precision to find the missing dimension of a triangle by using the area formula.';
  var languageObjective =
    'I can explain how the area formula helps me find a missing dimension and justify whether my answer is reasonable in context.';
  var mindset = findFirstMatchingLineInSlides_(sessionSlides, /positive relationship with a classmate/i) ||
    'How are you establishing a positive relationship with a classmate in math class?';
  var bermudaIntroPrompt = combinePromptLinesFromSlides_(deck, [31],
    'The Bermuda Triangle is an area in the Atlantic Ocean between Miami, Florida, the island of Bermuda, and San Juan, Puerto Rico. It covers an area of about 437,409 square miles. What is the approximate distance from Bermuda to San Juan? What tool can you use to represent the problem?');
  var bermudaEstimatePrompt = 'Does the problem call for an estimate or an exact solution?';
  var bermudaStep1Prompt = combinePromptLinesFromSlides_(deck, [33],
    'Step 1 Let’s create a visual representation of the problem. The Bermuda Triangle covers an area of about 437,409 square miles. What is the approximate distance from Bermuda to San Juan? What is the problem asking you to find? How does this relate to the visual representation?');
  var bermudaStep1Reveal = extractRevealTextFromSlides_(deck, [34],
    'We need to find the distance from Bermuda to San Juan, which is the base of the triangle.');
  var bermudaRelatePrompt = combinePromptLinesFromSlides_(deck, [35],
    'We need to find the distance from Bermuda to San Juan, which is the base of the triangle. How can you relate the base, height, and area to find the missing measure?');
  var bermudaRelateReveal = extractRevealTextFromSlides_(deck, [36],
    'We can use the area formula for triangles.');
  var bermudaCollaboratePrompt = combinePromptLinesFromSlides_(deck, [37, 38],
    'Let’s use an area formula for triangles to find the base. What is the missing measure? What does this value represent?');
  var bermudaSubstitutePrompt = combinePromptLinesFromSlides_(deck, [39, 40],
    'Let’s substitute the given values for the area and the height for A and h in the formula, then solve for b.');
  var bermudaAnswerReveal = 'The approximate distance from Bermuda to San Juan, Puerto Rico, is 954 miles.';
  var extensionPrompt = combinePromptLinesFromSlides_(deck, [42],
    'What other information would you need to calculate the distance from San Juan to Miami using the area formula?');

  return {
    sessionLabel: sessionLabel,
    lessonLabel: sessionLabel,
    sessionTitle: 'Determine the Area of Triangles — Session 2',
    objectives: {
      content: contentObjective,
      language: languageObjective
    },
    mindset: mindset,
    sections: [
      {
        type: 'cover',
        title: deck.lessonTitle,
        kicker: 'Session 2 Student Notebook',
        contentObjective: contentObjective,
        languageObjective: languageObjective,
        mindset: mindset
      },
      {
        type: 'sourceTask',
        title: 'Objectives + Success Criteria',
        kicker: 'Set the purpose',
        stageLabel: 'Learning Targets',
        problem: contentObjective,
        supportLabel: 'Language Objective',
        supportQuestion: languageObjective,
        sourceLines: [
          'Decide whether the problem needs an estimate or an exact answer.',
          'Represent the triangle and identify the unknown dimension clearly.',
          'Use the area formula, solve for the missing measure, and interpret the units.'
        ],
        visualKey: 'missingdimension',
        strategyTitle: 'Track your goal',
        discussionQs: [
          'What does the phrase missing dimension tell you about today’s work?',
          'How will you know whether an answer is reasonable in context?'
        ]
      },
      {
        type: 'beCurious',
        title: 'Be Curious',
        kicker: 'What could the question be?',
        imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(deck, 27)) || firstImageDataUriFromSlide_(getSlideByNumber_(deck, 28)),
        imageCaption: 'Use the image, sentence kernels, and vocabulary to predict the question.',
        noticeKernels: [
          'I notice ___ because ___.',
          'A measurement that looks important is ___.'
        ],
        wonderKernels: [
          'I wonder what the question could be because ___.',
          'I predict the missing measure is ___ because ___.'
        ],
        vocabBox: [
          { term: 'approximate', definition: 'close to the exact value, not perfectly exact' },
          { term: 'precision', definition: 'careful and accurate use of numbers, labels, and units' },
          { term: 'base', definition: 'the side being solved for in the Bermuda Triangle task' },
          { term: 'height', definition: 'the perpendicular measure paired with the base' }
        ]
      },
      {
        type: 'vocabulary',
        title: 'Vocabulary for Precision',
        kicker: 'Keep these words visible',
        terms: [
          {
            term: 'Approximate',
            definition: 'Close to an exact value but not perfectly exact.',
            studentFriendly: 'An approximate answer is a sensible estimate based on the information given.',
            example: 'The distance from Bermuda to San Juan is about 954 miles.',
            visualKey: 'approximate'
          },
          {
            term: 'Precision',
            definition: 'Carefully using the correct numbers, units, and labels.',
            studentFriendly: 'Precision means your work is clearly labeled and mathematically accurate.',
            example: 'The source asks students to attend to precision when finding a missing dimension.',
            visualKey: 'precision'
          },
          {
            term: 'Base',
            definition: 'The side of the triangle paired with the height in the area formula.',
            studentFriendly: 'In the Bermuda Triangle task, the unknown distance is the base.',
            example: 'The distance from Bermuda to San Juan is treated as the base.',
            visualKey: 'base'
          },
          {
            term: 'Missing Dimension',
            definition: 'A side length or measure that must be solved for using known information.',
            studentFriendly: 'A missing dimension is the part of the figure you do not know yet.',
            example: 'Session 2 uses the area formula to find the missing base.',
            visualKey: 'missingdimension'
          }
        ]
      },
      {
        type: 'sourceTask',
        title: 'Guided Notes: Represent the Bermuda Triangle',
        kicker: 'Set up the problem',
        stageLabel: 'Guided Notes',
        problem: bermudaIntroPrompt,
        supportLabel: 'Context + Precision',
        supportQuestion: bermudaEstimatePrompt,
        sourceLines: [
          'Use a visual representation of the triangle before you solve.',
          'A = 437,409 square miles.',
          'h = 917 miles.',
          'b = ? represents the distance from Bermuda to San Juan.'
        ],
        revealLabel: 'Reveal',
        reveal: bermudaStep1Reveal,
        imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(deck, 29)) || firstImageDataUriFromSlide_(getSlideByNumber_(deck, 31)),
        visualKey: 'triangle',
        strategyTitle: 'Represent the problem',
        discussionQs: [
          'Why is the missing distance the base in this context?',
          'Which values are known before you start solving with the formula?'
        ]
      },
      {
        type: 'formulaDerivation',
        title: 'Build the Missing-Base Strategy',
        kicker: 'Guided practice',
        steps: [
          {
            title: 'Create the visual representation',
            visualKey: 'triangle',
            explanation: 'Step 1 asks you to create a visual representation of the problem. The missing distance from Bermuda to San Juan is the base of the triangle.'
          },
          {
            title: 'Use the area formula for triangles',
            visualKey: 'formula',
            explanation: bermudaRelatePrompt + ' ' + bermudaRelateReveal
          },
          {
            title: 'Interpret the missing measure',
            visualKey: 'missingdimension',
            explanation: bermudaAnswerReveal + ' This value is the missing base, so it tells the distance from Bermuda to San Juan in miles.'
          }
        ],
        discussionQs: [
          'Why do we multiply the area by 2 before dividing by the height?',
          'What does the solved value represent in the real-world situation?'
        ]
      },
      {
        type: 'vocabActivity',
        title: 'Equation Builder: Bermuda Triangle',
        kicker: 'Given, equation, meaning',
        activityFamily: 'sort_classify',
        activityInstructions: 'Drag the cards into the best box to show the Bermuda Triangle givens, the matching equation setup, and the meaning of the solved value.',
        activityBankTitle: 'Equation Builder',
        bankIntro: 'Sort the information cards, then explain how the equation connects the givens to the missing base.',
        dropzoneHint: 'Place two matching cards here',
        movablePieces: [
          'A = 437,409 sq mi',
          'h = 917 mi',
          'b = ?',
          'distance from Bermuda to San Juan',
          '437,409 = 1/2(b)(917)',
          'solve for b'
        ],
        answerCheck: 'A strong sort places the area and height with the givens, puts the matching equation setup together, and explains that solving for b gives the approximate distance from Bermuda to San Juan.',
        activities: [
          {
            label: 'Given',
            prompt: 'Place the cards that show the known measurements and the unknown in the Bermuda Triangle problem.'
          },
          {
            label: 'Equation',
            prompt: 'Place the cards that show how the area formula is set up to solve for the missing base.'
          },
          {
            label: 'Meaning',
            prompt: 'Place the cards that explain what the solved value means in the map context.'
          }
        ],
        discussionQs: [
          'Which card tells you what the variable b actually represents in the map?',
          'How does the equation connect the givens to the distance you want to find?'
        ]
      },
      {
        type: 'sourceTask',
        title: 'Substitute and Solve for b',
        kicker: 'Think About It',
        stageLabel: 'Think About It',
        problem: bermudaSubstitutePrompt,
        supportLabel: 'Explain',
        supportQuestion: 'Why can you write the triangle area formula as A = 1/2 × b × h or A = (b × h) ÷ 2?',
        sourceLines: [
          'Substitute A = 437,409 square miles into the formula.',
          'Use h = 917 miles from the source diagram.',
          'Solve for b, the missing base.',
          'Write the final distance in miles and connect it back to the map.'
        ],
        revealLabel: 'Reveal',
        reveal: bermudaAnswerReveal,
        imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(deck, 39)) || firstImageDataUriFromSlide_(getSlideByNumber_(deck, 41)),
        visualKey: 'formula',
        strategyTitle: 'Substitute and solve',
        discussionQs: [
          'Why do you substitute the known values before solving for b?',
          'How do you know the final answer should be written in miles?'
        ]
      },
      {
        type: 'sourceTask',
        title: 'Application: Painting a Cabin',
        kicker: 'Apply the strategy',
        stageLabel: 'Independent Practice',
        problem: 'Francisco is painting the triangular section of the cabin. Use the source measurements to determine the paintable triangular area and the least amount he should budget for paint.',
        supportLabel: 'Use the Source Information',
        supportQuestion: 'How do you find the triangular area first, then subtract the window areas, then decide how many full gallons are needed?',
        sourceLines: [
          'The triangular section has a base of 34 feet and a height of 23 3/4 feet.',
          'The top window is 5 feet by 2 feet.',
          'The bottom window is 13 1/2 feet by 5 feet.',
          'One gallon costs $24.95 and covers 200 square feet.',
          'Subtract both window areas before you determine the paint cost.'
        ],
        imageDataUri: firstImageDataUriFromSlide_(getSlideByNumber_(deck, 44)),
        visualKey: 'formula',
        strategyTitle: 'Make a plan',
        discussionQs: [
          'Why do the window areas need to be removed before calculating the amount of paint?',
          'How will you decide whether Francisco needs one gallon or more than one gallon of paint?'
        ]
      },
      {
        type: 'collaborateDeep',
        title: 'Application Planner: Compare and Decide',
        kicker: 'Reason through both problems',
        problem: 'Compare the Bermuda Triangle problem and the cabin problem. How does the same triangle-area formula lead to two different kinds of answers?',
        partnerRoles: [
          'Partner A: Bermuda Triangle',
          'Partner B: Painting a Cabin'
        ],
        partnerTasks: [
          'Explain how the Bermuda Triangle task uses the area formula to solve for a missing base and interpret the answer in miles.',
          'Explain how the cabin task uses triangle area, subtraction, and paint coverage to make a budgeting decision.'
        ],
        discussionQs: [
          'How is solving for a missing base different from using area to make a decision about paint?',
          'What do both problems have in common even though the final answers mean different things?'
        ]
      },
      {
        type: 'extensionChallenge',
        title: 'Extension Challenge: Go Further',
        kicker: 'Let’s Explore More',
        problem: extensionPrompt,
        hint: 'Think about which measurement is still unknown and why multiplying by one-half and dividing by two are equivalent operations.',
        discussionQs: [
          'Which missing measure stops you from finding the distance from San Juan to Miami right now?',
          'How would you explain the equivalence of multiplying by 1/2 and dividing by 2 to another student?'
        ]
      },
      {
        type: 'summary',
        title: 'Summary: Missing Dimension + Application',
        kicker: 'Wrap up the big idea',
        summaryStem: 'We can use the triangle area formula to solve for a missing dimension and to make decisions in real situations.',
        bullets: [
          'Use A = 1/2 × b × h to connect the area, base, and height.',
          'When the base is missing, isolate b and interpret the answer in context.',
          'In application problems, you may need extra steps after finding area, such as subtracting windows or deciding how many gallons to buy.'
        ],
        formulaTitle: 'Key Structure',
        mathLines: [
          'A = 1/2 × b × h',
          '437,409 = 1/2(b)(917)',
          'b = (2A) ÷ h'
        ],
        comparePrompt: 'How do the Bermuda Triangle and cabin tasks use the same formula in different ways?',
        discussionQs: [
          'When do you solve for a missing dimension, and when do you use area to make a decision?',
          'Why does context matter when you interpret the final answer?'
        ]
      },
      {
        type: 'reflection',
        title: 'Session 2 Exit Ticket + Goal Tracker',
        kicker: 'Explain and reflect',
        prompts: [
          'I can find a missing dimension of a triangle by...',
          'The value 954 represents ... because ...',
          'My confidence with triangle area today is ... and my next step is ...'
        ]
      }
    ]
  };
}

function buildNotebookDeckModel_(presentation) {
  var rawSlides = presentation.getSlides();
  var slides = [];
  for (var i = 0; i < rawSlides.length; i++) {
    slides.push(buildNotebookSlideModel_(rawSlides[i], i + 1));
  }
  return {
    lessonTitle: inferNotebookLessonTitleFromSlides_(slides),
    slides: slides
  };
}

function buildNotebookSlideModel_(slide, number) {
  var lines = extractSlideLines_(slide);
  var joined = lines.join(' ').replace(/\s+/g, ' ').trim();
  return {
    number: number,
    slide: slide,
    lines: lines,
    text: joined,
    lowerText: joined.toLowerCase(),
    title: lines.length ? lines[0] : '',
    imageDataUri: largestSlideImageDataUri_(slide)
  };
}

function inferNotebookLessonTitle_(deck) {
  return deck.lessonTitle || 'Determine the Area of Triangles';
}

function inferNotebookLessonTitleFromSlides_(slides) {
  var first = slides.length ? safeNotebookText_(slides[0].text) : '';
  var match = first.match(/^(.*?)(?:\s+Session\s+\d+)?$/i);
  var cleaned = match ? match[1] : first;
  cleaned = cleaned.replace(/\s+/g, ' ').trim();
  return cleaned || 'Determine the Area of Triangles';
}

function extractNotebookSessionSlides_(deck) {
  var session2Index = -1;
  for (var i = 0; i < deck.slides.length; i++) {
    if (/\bsession\s*2\b/i.test(deck.slides[i].text)) {
      session2Index = i;
      break;
    }
  }

  if (session2Index < 0) {
    session2Index = Math.max(1, Math.floor(deck.slides.length / 2));
  }

  return [
    deck.slides.slice(0, session2Index),
    deck.slides.slice(session2Index)
  ];
}

function extractLearningTargets_(deck) {
  var targetSlide = getSlideByNumber_(deck, 45) || findFirstSlideByPattern_(deck.slides, /learning targets?/i);
  var targets = [];
  if (!targetSlide) return targets;

  var matches = safeNotebookText_(targetSlide.text).match(/I can[^.?!]*[.?!]?/gi);
  if (!matches) return targets;

  for (var i = 0; i < matches.length; i++) {
    var cleaned = matches[i].replace(/\s+/g, ' ').trim();
    if (cleaned && targets.indexOf(cleaned) === -1) {
      targets.push(cleaned.replace(/\.$/, '') + '.');
    }
  }
  return targets;
}

function assertTriangleDeck_(lessonTitle, deck) {
  var joined = safeNotebookText_(lessonTitle + ' ' + (deck.slides.length ? deck.slides[0].text : ''));
  if (!/determine the area of triangles/i.test(joined)) {
    throw new Error('NotebookExtractors_Precision.gs is configured for the "Determine the Area of Triangles" deck.');
  }
}

function validateExtractorPackage_(pkg) {
  if (!pkg || !pkg.meta || !pkg.sessions || pkg.sessions.length !== 2) {
    throw new Error('Extractor package must return meta plus exactly two sessions.');
  }

  for (var i = 0; i < pkg.sessions.length; i++) {
    validateExtractorSession_(pkg.sessions[i], i + 1);
  }
}

function validateExtractorSession_(sessionSpec, index) {
  if (!sessionSpec || !sessionSpec.sections || !sessionSpec.sections.length) {
    throw new Error('Session ' + index + ' is missing sections.');
  }

  var slideCount = sessionSpec.sections.length;
  if (slideCount < NOTEBOOK_BUILD_DEFAULTS_.MIN_SLIDES_PER_SESSION ||
      slideCount > NOTEBOOK_BUILD_DEFAULTS_.MAX_SLIDES_PER_SESSION) {
    throw new Error(
      'Session ' + index + ' must contain between ' +
      NOTEBOOK_BUILD_DEFAULTS_.MIN_SLIDES_PER_SESSION + ' and ' +
      NOTEBOOK_BUILD_DEFAULTS_.MAX_SLIDES_PER_SESSION + ' slides. Found ' + slideCount + '.'
    );
  }

  var hasCover = false;
  var hasBeCurious = false;
  var hasVocabulary = false;
  var hasReflection = false;
  var hasObjectiveSlide = false;
  var slidesSinceDiscussion = 0;

  for (var i = 0; i < sessionSpec.sections.length; i++) {
    var section = sessionSpec.sections[i] || {};
    if (section.type === 'cover') hasCover = true;
    if (section.type === 'beCurious') hasBeCurious = true;
    if (section.type === 'vocabulary') hasVocabulary = true;
    if (section.type === 'reflection') hasReflection = true;
    if (/objective/i.test(safeNotebookText_(section.title))) hasObjectiveSlide = true;

    if (i > 0) {
      slidesSinceDiscussion++;
      if (hasStructuredDiscussion_(section)) {
        slidesSinceDiscussion = 0;
      }
      if (slidesSinceDiscussion > NOTEBOOK_BUILD_DEFAULTS_.MAX_SLIDES_WITHOUT_DISCUSSION) {
        throw new Error(
          'Session ' + index + ' must include a structured discussion opportunity at least every ' +
          NOTEBOOK_BUILD_DEFAULTS_.MAX_SLIDES_WITHOUT_DISCUSSION + ' slides.'
        );
      }
    }
  }

  if (!hasCover) throw new Error('Session ' + index + ' is missing a cover slide.');
  if (!hasObjectiveSlide) throw new Error('Session ' + index + ' is missing an objectives slide.');
  if (!hasBeCurious) throw new Error('Session ' + index + ' is missing a Be Curious slide.');
  if (!hasVocabulary) throw new Error('Session ' + index + ' is missing a vocabulary slide.');
  if (!hasReflection) throw new Error('Session ' + index + ' is missing a reflection slide.');
}

function hasStructuredDiscussion_(section) {
  var prompts = section && section.discussionQs;
  return Array.isArray(prompts) && prompts.filter(function(prompt) {
    return !!safeNotebookText_(prompt);
  }).length > 0;
}

function getSlideByNumber_(deck, number) {
  if (!deck || !deck.slides) return null;
  for (var i = 0; i < deck.slides.length; i++) {
    if (deck.slides[i].number === number) return deck.slides[i];
  }
  return null;
}

function findFirstSlideByPattern_(slides, pattern) {
  for (var i = 0; i < slides.length; i++) {
    if (pattern.test(slides[i].text)) return slides[i];
  }
  return null;
}

function findFirstMatchingLineInSlides_(slides, pattern) {
  for (var i = 0; i < slides.length; i++) {
    var lines = slides[i].lines || [];
    for (var j = 0; j < lines.length; j++) {
      if (pattern.test(lines[j])) return lines[j];
    }
  }
  return '';
}

function combinePromptLinesFromSlides_(deck, slideNumbers, fallback) {
  var lines = collectPromptLinesFromSlides_(deck, slideNumbers);
  var selected = selectPromptBundleLines_(lines);
  return selected.length ? selected.join(' ') : fallback;
}

function extractRevealTextFromSlides_(deck, slideNumbers, fallback) {
  for (var i = 0; i < slideNumbers.length; i++) {
    var slide = getSlideByNumber_(deck, slideNumbers[i]);
    if (!slide) continue;
    var lines = slide.lines || [];
    for (var j = 0; j < lines.length; j++) {
      var normalized = normalizeNotebookLine_(lines[j]);
      if (/^Reveal:\s*/i.test(normalized)) {
        return normalized.replace(/^Reveal:\s*/i, '').trim() || fallback;
      }
    }
  }
  return fallback;
}

function collectPromptLinesFromSlides_(deck, slideNumbers) {
  var collected = [];
  var seen = {};

  for (var i = 0; i < slideNumbers.length; i++) {
    var slide = getSlideByNumber_(deck, slideNumbers[i]);
    if (!slide) continue;
    var lines = slide.lines || [];
    for (var j = 0; j < lines.length; j++) {
      var normalized = normalizeNotebookLine_(lines[j]);
      if (!normalized) continue;
      if (isNotebookPromptNoiseLine_(normalized)) continue;
      if (/^Reveal:\s*/i.test(normalized)) continue;
      if (seen[normalized]) continue;
      seen[normalized] = true;
      collected.push(normalized);
    }
  }

  return collected;
}

function collectQuestionLinesFromSlides_(deck, slideNumbers) {
  var lines = collectPromptLinesFromSlides_(deck, slideNumbers);
  return lines.filter(isNotebookQuestionLine_);
}

function collectFactLinesFromSlides_(deck, slideNumbers) {
  var lines = collectPromptLinesFromSlides_(deck, slideNumbers);
  return lines.filter(isNotebookFactLine_);
}

function selectPromptBundleLines_(lines) {
  if (!lines || !lines.length) return [];

  var firstQuestionIndex = -1;
  var lastQuestionIndex = -1;
  for (var i = 0; i < lines.length; i++) {
    if (isNotebookQuestionLine_(lines[i])) {
      if (firstQuestionIndex < 0) firstQuestionIndex = i;
      lastQuestionIndex = i;
    }
  }

  if (firstQuestionIndex >= 0) {
    var start = Math.max(0, firstQuestionIndex - 2);
    var end = Math.min(lines.length, lastQuestionIndex + 1);
    return lines.slice(start, end);
  }

  return lines.slice(0, 3);
}

function isNotebookQuestionLine_(line) {
  return /[?]$/.test(line) || /^(What|How|Why|Which|Is|Does|Can|Should|Would|Will|Where|When)\b/i.test(line);
}

function isNotebookFactLine_(line) {
  return /\d/.test(line) ||
    /\b(base|height|area|square|miles|inches|centimeters|feet|gallon|window|formula|triangle|parallelogram|right triangle)\b/i.test(line);
}

function isNotebookPromptNoiseLine_(line) {
  return /^(Triangles|The Bermuda Triangle|Collaborate and Connect|Let’s Explore More|Let's Explore More|Precision|Using Symbols Appropriately|Workspace|Learning Targets|Apply: Painting a Cabin|Summarize: Determine the Area of Triangles|Reveal:|Think About It:)$/i.test(line);
}

function formatSlideRangeLabel_(slideNumbers) {
  if (!slideNumbers || !slideNumbers.length) return '';
  if (slideNumbers.length === 1) return 'Slide ' + slideNumbers[0];
  return 'Slides ' + slideNumbers[0] + '-' + slideNumbers[slideNumbers.length - 1];
}

function firstImageDataUriFromSlide_(slideModel) {
  return slideModel && slideModel.imageDataUri ? slideModel.imageDataUri : '';
}

function extractSlideLines_(slide) {
  var lines = [];
  var elements = slide.getPageElements();
  for (var i = 0; i < elements.length; i++) {
    collectTextFromPageElement_(elements[i], lines);
  }

  var cleaned = [];
  for (var j = 0; j < lines.length; j++) {
    var line = normalizeNotebookLine_(lines[j]);
    if (!line) continue;
    if (cleaned.indexOf(line) === -1) cleaned.push(line);
  }
  return cleaned;
}

function collectTextFromPageElement_(element, bucket) {
  if (!element) return;

  var type = element.getPageElementType();
  if (type === SlidesApp.PageElementType.SHAPE) {
    var shape = element.asShape();
    if (shape && shape.getText) {
      pushTextLines_(shape.getText().asString(), bucket);
    }
    return;
  }

  if (type === SlidesApp.PageElementType.GROUP) {
    var children = element.asGroup().getChildren();
    for (var i = 0; i < children.length; i++) {
      collectTextFromPageElement_(children[i], bucket);
    }
    return;
  }

  if (type === SlidesApp.PageElementType.TABLE) {
    var table = element.asTable();
    var rows = table.getNumRows();
    var cols = table.getNumColumns();
    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        try {
          pushTextLines_(table.getCell(row, col).getText().asString(), bucket);
        } catch (err) {}
      }
    }
  }
}

function pushTextLines_(text, bucket) {
  var raw = safeNotebookText_(text).replace(/\u200B/g, '');
  if (!raw) return;
  var parts = raw.split(/\r?\n/);
  for (var i = 0; i < parts.length; i++) {
    if (parts[i]) bucket.push(parts[i]);
  }
}

function normalizeNotebookLine_(line) {
  return safeNotebookText_(line)
    .replace(/\s+/g, ' ')
    .replace(/\u00A0/g, ' ')
    .trim();
}

function largestSlideImageDataUri_(slide) {
  var images;
  try {
    images = slide.getImages();
  } catch (err) {
    images = [];
  }

  if (!images || !images.length) return '';

  var best = null;
  var bestArea = -1;
  for (var i = 0; i < images.length; i++) {
    var area = Number(images[i].getWidth()) * Number(images[i].getHeight());
    if (area > bestArea) {
      best = images[i];
      bestArea = area;
    }
  }

  if (!best) return '';
  return imageBlobToDataUri_(best.getBlob());
}

function imageBlobToDataUri_(blob) {
  if (!blob) return '';
  var contentType = safeNotebookText_(blob.getContentType()) || 'image/png';
  return 'data:' + contentType + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function safeNotebookText_(value) {
  return value === null || value === undefined ? '' : String(value);
}
