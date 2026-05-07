function ScoreReport({ activity, onHome, onRestart, reviewSkills, score, studentName }) {
  return (
    <main className="activity-shell">
      <section className="results-hero">
        <div className="score-circle">
          <strong>{score.percent}%</strong>
          <span>
            {score.pointsEarned}/{score.totalPoints} points
          </span>
        </div>
        <div>
          <p className="eyebrow">Final Report</p>
          <h1>{activity.title}</h1>
          <p>
            {studentName ? `${studentName}, y` : "Y"}ou answered {score.answeredCount} of {score.totalQuestions}{" "}
            questions.
          </p>
        </div>
      </section>

      <section className="breakdown-grid" aria-label="Section scores">
        {score.sectionBreakdown.map((section) => (
          <article className="breakdown-card" key={section.id}>
            <h2>{section.title}</h2>
            <strong>{section.percent}%</strong>
            <p>
              {section.pointsEarned} of {section.totalPoints} points | {section.answeredCount} answered
            </p>
          </article>
        ))}
      </section>

      <section className="review-list">
        <h2>Suggested Review</h2>
        {reviewSkills.length === 0 ? (
          <p>No missed answered questions yet. Keep practicing or restart when you are ready.</p>
        ) : (
          <div className="skill-list">
            {reviewSkills.map((skill) => (
              <span key={skill}>{skill}</span>
            ))}
          </div>
        )}
      </section>

      <div className="result-actions">
        <button className="button button-danger" onClick={onRestart} type="button">
          Restart Activity
        </button>
        <button className="button" onClick={onHome} type="button">
          Home
        </button>
      </div>
    </main>
  );
}

export default ScoreReport;
