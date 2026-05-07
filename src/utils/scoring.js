export function normalizeAnswer(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

export function getPercent(pointsEarned, totalPoints) {
  return totalPoints === 0 ? 0 : Math.round((pointsEarned / totalPoints) * 100);
}

export function gradeAnswer(question, rawAnswer) {
  const possibleAnswers = Array.isArray(question.correctAnswer)
    ? question.correctAnswer
    : [question.correctAnswer];
  const studentAnswer = String(rawAnswer ?? "").trim();
  const isCorrect = possibleAnswers
    .map((answer) => normalizeAnswer(answer))
    .includes(normalizeAnswer(studentAnswer));

  return {
    studentAnswer,
    isCorrect,
    pointsEarned: isCorrect ? question.points : 0,
    pointsPossible: question.points,
  };
}

export function getAllQuestions(sections) {
  return sections.flatMap((section) =>
    section.questions.map((question) => ({
      ...question,
      sectionId: section.id,
      sectionTitle: section.title,
    })),
  );
}

export function getSectionScore(section, answers) {
  const questionResults = section.questions.map((question) => {
    const savedAnswer = answers[question.id];
    const gradedAnswer = savedAnswer?.pointsPossible ? savedAnswer : gradeAnswer(question, savedAnswer?.studentAnswer ?? savedAnswer);

    return {
      question,
      savedAnswer,
      gradedAnswer,
    };
  });

  const answeredCount = questionResults.filter((result) => result.savedAnswer !== undefined).length;
  const correctCount = questionResults.filter((result) => result.savedAnswer && result.gradedAnswer.isCorrect).length;
  const pointsEarned = questionResults.reduce(
    (total, result) => total + (result.savedAnswer ? result.gradedAnswer.pointsEarned : 0),
    0,
  );
  const totalPoints = section.questions.reduce((total, question) => total + question.points, 0);

  return {
    id: section.id,
    title: section.title,
    answeredCount,
    totalQuestions: section.questions.length,
    correctCount,
    pointsEarned,
    totalPoints,
    percent: getPercent(pointsEarned, totalPoints),
  };
}

export function calculateActivityScore(sections, answers) {
  const sectionBreakdown = sections.map((section) => getSectionScore(section, answers));
  const totals = sectionBreakdown.reduce(
    (current, section) => ({
      answeredCount: current.answeredCount + section.answeredCount,
      totalQuestions: current.totalQuestions + section.totalQuestions,
      correctCount: current.correctCount + section.correctCount,
      pointsEarned: current.pointsEarned + section.pointsEarned,
      totalPoints: current.totalPoints + section.totalPoints,
    }),
    { answeredCount: 0, totalQuestions: 0, correctCount: 0, pointsEarned: 0, totalPoints: 0 },
  );

  return {
    ...totals,
    allComplete: totals.answeredCount === totals.totalQuestions,
    percent: getPercent(totals.pointsEarned, totals.totalPoints),
    sectionBreakdown,
  };
}

export function getReviewSkills(sections, answers) {
  const missedSkills = getAllQuestions(sections)
    .filter((question) => {
      const savedAnswer = answers[question.id];
      return savedAnswer && !gradeAnswer(question, savedAnswer.studentAnswer ?? savedAnswer).isCorrect;
    })
    .map((question) => question.skill);

  return [...new Set(missedSkills)];
}
