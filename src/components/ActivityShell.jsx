import ProgressBar from "./ProgressBar.jsx";

function ActivityShell({ activity, children, onBackHome, onRestart, score, studentName, onStudentNameChange }) {
  return (
    <main className="activity-shell">
      <header className="activity-topbar">
        <button className="button button-secondary" onClick={onBackHome} type="button">
          Home
        </button>
        <button className="button button-danger" onClick={onRestart} type="button">
          Restart
        </button>
      </header>

      <section className="activity-hero">
        <div>
          <p className="eyebrow">
            {activity.unit} | {activity.activityType}
          </p>
          <h1>{activity.title}</h1>
          <p>{activity.description}</p>
        </div>
        <div className="activity-summary">
          <span>{activity.grade}</span>
          <strong>{activity.estimatedMinutes} min</strong>
        </div>
      </section>

      <section className="student-panel">
        <label htmlFor="student-name">Student name</label>
        <input
          id="student-name"
          onChange={(event) => onStudentNameChange(event.target.value)}
          placeholder="Type your name"
          type="text"
          value={studentName}
        />
      </section>

      <section className="score-strip" aria-label="Activity progress">
        <div className="score-tile">
          <strong>{score.answeredCount}</strong>
          <span>Answered</span>
        </div>
        <div className="score-tile">
          <strong>{score.pointsEarned}</strong>
          <span>Points</span>
        </div>
        <div className="score-tile">
          <strong>{score.percent}%</strong>
          <span>Score</span>
        </div>
      </section>

      <ProgressBar current={score.answeredCount} label="Activity progress" total={score.totalQuestions} />

      {children}
    </main>
  );
}

export default ActivityShell;
