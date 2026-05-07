export function getAllQuestions(levels) {
  return levels.flatMap((level) =>
    level.questions.map((question) => ({
      ...question,
      levelTitle: level.title,
    })),
  );
}

export function getLevelScore(level, answers) {
  const answeredCount = level.questions.filter((question) => answers[question.id]).length;
  const correctCount = level.questions.filter((question) => answers[question.id]?.isCorrect).length;

  return {
    answeredCount,
    correctCount,
    totalQuestions: level.questions.length,
    isComplete: answeredCount === level.questions.length,
  };
}

export function getPercent(correctCount, totalQuestions) {
  return totalQuestions === 0 ? 0 : Math.round((correctCount / totalQuestions) * 100);
}

export function calculateTotals(levels, answers) {
  const questions = getAllQuestions(levels);
  const answeredCount = questions.filter((question) => answers[question.id]).length;
  const correctCount = questions.filter((question) => answers[question.id]?.isCorrect).length;

  return {
    answeredCount,
    correctCount,
    totalQuestions: questions.length,
    allComplete: answeredCount === questions.length,
    percent: getPercent(correctCount, questions.length),
  };
}
