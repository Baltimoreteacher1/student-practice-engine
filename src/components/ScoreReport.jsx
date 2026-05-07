import { calculateTotals, getAllQuestions, getLevelScore, getPercent } from "../utils/scoring.js";

function ScoreReport({ answers, levels, onBack, onReset, studentName }) {
  const totals = calculateTotals(levels, answers);
  const questions = getAllQuestions(levels);
  const missedQuestions = questions.filter((question) => answers[question.id]?.isCorrect === false);

  return (
    <main className="app-shell">
      <section className="results-hero">
        <div className="score-circle">
          <strong>{totals.percent}%</strong>
          <span>
            {totals.correctCount}/{totals.totalQuestions}
          </span>
        </div>
        <div>
          <p className="eyebrow">Final Report</p>
          <h1>{studentName}'s MCAP Practice Score</h1>
          <p>
            You answered {totals.answeredCount} of {totals.totalQuestions} questions.
            Use the section scores and review list to decide what to practice next.
          </p>
        </div>
      </section>

      <section className="breakdown-grid" aria-label="Section scores">
        {levels.map((level) => {
          const levelScore = getLevelScore(level, answers);
          const levelPercent = getPercent(levelScore.correctCount, levelScore.totalQuestions);

          return (
            <article className={`breakdown-card ${level.color}`} key={level.id}>
              <h2>{level.title}</h2>
              <strong>{levelPercent}%</strong>
              <p>
                {levelScore.correctCount} correct out of {levelScore.totalQuestions}
              </p>
            </article>
          );
        })}
      </section>

      <section className="review-list">
        <h2>Questions to Review</h2>
        {missedQuestions.length === 0 ? (
          <p>No missed answered questions yet. Nice work.</p>
        ) : (
          missedQuestions.map((question) => {
            const answer = answers[question.id];

            return (
              <article key={question.id}>
                <p className="eyebrow">{question.levelTitle}</p>
                <h3>{question.prompt}</h3>
                <p>
                  Student answer: <strong>{question.choices[answer.selectedIndex]}</strong>
                </p>
                <p>
                  Correct answer: <strong>{question.choices[question.answerIndex]}</strong>
                </p>
                <p>{question.explanation}</p>
              </article>
            );
          })
        )}

        <div className="result-actions">
          <button className="ghost-button" onClick={onBack} type="button">
            Back to Sections
          </button>
          <button className="danger-button" onClick={onReset} type="button">
            Reset Practice
          </button>
        </div>
      </section>
    </main>
  );
}

export default ScoreReport;
