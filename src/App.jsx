import { useEffect, useMemo, useState } from "react";
import ActivityShell from "./components/ActivityShell.jsx";
import QuestionCard from "./components/QuestionCard.jsx";
import ScoreReport from "./components/ScoreReport.jsx";
import { questionBanks } from "./data/questionBanks.js";
import { calculateTotals } from "./utils/scoring.js";
import { clearProgress, emptyProgress, loadProgress, saveProgress } from "./utils/storage.js";

function App() {
  const [progress, setProgress] = useState(loadProgress);

  const currentLevel = questionBanks.find((level) => level.id === progress.currentLevelId);
  const currentQuestion = currentLevel?.questions[progress.currentQuestionIndex];

  const totals = useMemo(
    () => calculateTotals(questionBanks, progress.answers),
    [progress.answers],
  );

  useEffect(() => {
    saveProgress(progress);
  }, [progress]);

  function startPractice(studentName) {
    setProgress({
      ...emptyProgress,
      studentName,
    });
  }

  function chooseLevel(levelId) {
    const level = questionBanks.find((item) => item.id === levelId);
    if (!level) return;

    const firstUnansweredIndex = level.questions.findIndex((question) => !progress.answers[question.id]);

    setProgress((current) => ({
      ...current,
      currentLevelId: levelId,
      currentQuestionIndex: firstUnansweredIndex === -1 ? 0 : firstUnansweredIndex,
      showResults: false,
    }));
  }

  function saveAnswer(question, selectedIndex) {
    setProgress((current) => ({
      ...current,
      answers: {
        ...current.answers,
        [question.id]: {
          selectedIndex,
          isCorrect: selectedIndex === question.answerIndex,
        },
      },
    }));
  }

  function goToNextQuestion() {
    if (!currentLevel) return;

    const isLastQuestion = progress.currentQuestionIndex === currentLevel.questions.length - 1;
    if (isLastQuestion) {
      setProgress((current) => ({
        ...current,
        currentLevelId: "",
        currentQuestionIndex: 0,
        showResults: totals.allComplete,
      }));
      return;
    }

    setProgress((current) => ({
      ...current,
      currentQuestionIndex: current.currentQuestionIndex + 1,
    }));
  }

  function goToLevelMap() {
    setProgress((current) => ({
      ...current,
      currentLevelId: "",
      currentQuestionIndex: 0,
      showResults: false,
    }));
  }

  function showResults() {
    setProgress((current) => ({
      ...current,
      currentLevelId: "",
      currentQuestionIndex: 0,
      showResults: true,
    }));
  }

  function resetPractice() {
    clearProgress();
    setProgress({ ...emptyProgress });
  }

  if (progress.showResults) {
    return (
      <ScoreReport
        answers={progress.answers}
        levels={questionBanks}
        onBack={goToLevelMap}
        onReset={resetPractice}
        studentName={progress.studentName}
      />
    );
  }

  if (currentLevel && currentQuestion) {
    return (
      <QuestionCard
        answer={progress.answers[currentQuestion.id]}
        level={currentLevel}
        onAnswer={saveAnswer}
        onBack={goToLevelMap}
        onNext={goToNextQuestion}
        question={currentQuestion}
        questionNumber={progress.currentQuestionIndex + 1}
        totalQuestions={currentLevel.questions.length}
      />
    );
  }

  return (
    <ActivityShell
      answers={progress.answers}
      levels={questionBanks}
      onChooseLevel={chooseLevel}
      onReset={resetPractice}
      onShowResults={showResults}
      onStart={startPractice}
      studentName={progress.studentName}
      totals={totals}
    />
  );
}

export default App;
