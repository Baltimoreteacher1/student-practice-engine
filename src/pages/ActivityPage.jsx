import { useEffect, useMemo, useState } from "react";
import ActivityShell from "../components/ActivityShell.jsx";
import QuestionCard from "../components/QuestionCard.jsx";
import ScoreReport from "../components/ScoreReport.jsx";
import { getActivityBySlug } from "../data/activityCatalog.js";
import { questionBanks } from "../data/activities/index.js";
import { calculateActivityScore, getReviewSkills, gradeAnswer } from "../utils/scoring.js";
import { clearProgress, loadProgress, saveProgress } from "../utils/storage.js";
import NotFoundPage from "./NotFoundPage.jsx";

function ActivityPage({ activitySlug, onNavigate }) {
  const activity = getActivityBySlug(activitySlug);
  const questionBank = activity ? questionBanks[activity.questionBankId] : null;
  const [progress, setProgress] = useState(() => loadProgress(activitySlug));

  useEffect(() => {
    setProgress(loadProgress(activitySlug));
  }, [activitySlug]);

  useEffect(() => {
    saveProgress(activitySlug, progress);
  }, [activitySlug, progress]);

  const score = useMemo(
    () => calculateActivityScore(questionBank?.sections ?? [], progress.answers),
    [progress.answers, questionBank],
  );

  if (!activity || !questionBank) {
    return <NotFoundPage onNavigate={onNavigate} />;
  }

  function updateProgress(nextProgress) {
    setProgress((current) => ({ ...current, ...nextProgress }));
  }

  function handleAnswer(question, rawAnswer) {
    if (rawAnswer === null || rawAnswer === undefined || String(rawAnswer).trim() === "") return;
    const gradedAnswer = gradeAnswer(question, rawAnswer);
    updateProgress({
      answers: {
        ...progress.answers,
        [question.id]: gradedAnswer,
      },
    });
  }

  function handleRestart() {
    clearProgress(activity.slug);
    setProgress(loadProgress(activity.slug));
  }

  function showReport() {
    updateProgress({ showReport: true });
  }

  if (progress.showReport) {
    return (
      <ScoreReport
        activity={activity}
        onHome={() => onNavigate("/")}
        onRestart={handleRestart}
        reviewSkills={getReviewSkills(questionBank.sections, progress.answers)}
        score={score}
        studentName={progress.studentName}
      />
    );
  }

  return (
    <ActivityShell
      activity={activity}
      onBackHome={() => onNavigate("/")}
      onRestart={handleRestart}
      onStudentNameChange={(studentName) => updateProgress({ studentName })}
      score={score}
      studentName={progress.studentName}
    >
      <section className="directions-panel">
        <h2>{questionBank.title}</h2>
        <p>{questionBank.directions}</p>
      </section>

      {questionBank.sections.map((section) => (
        <section className="question-section" key={section.id}>
          <div className="section-heading">
            <div>
              <p className="eyebrow">Section</p>
              <h2>{section.title}</h2>
            </div>
            <p>{section.directions}</p>
          </div>
          <div className="question-list">
            {section.questions.map((question) => (
              <QuestionCard
                answer={progress.answers[question.id]}
                key={question.id}
                onAnswer={handleAnswer}
                question={question}
              />
            ))}
          </div>
        </section>
      ))}

      <section className="report-callout">
        <div>
          <h2>Ready for your report?</h2>
          <p>You can view your score now. Unanswered questions stay saved so you can come back.</p>
        </div>
        <button className="button" onClick={showReport} type="button">
          View Final Report
        </button>
      </section>
    </ActivityShell>
  );
}

export default ActivityPage;
