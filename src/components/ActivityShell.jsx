import { useState } from "react";
import { getLevelScore } from "../utils/scoring.js";
import ProgressBar from "./ProgressBar.jsx";

function ActivityShell({
  answers = {},
  levels = [],
  onChooseLevel,
  onReset,
  onShowResults,
  onStart,
  studentName = "",
  totals,
}) {
  const [name, setName] = useState("");

  function handleSubmit(event) {
    event.preventDefault();
    const cleanName = name.trim();

    if (cleanName) {
      onStart(cleanName);
    }
  }

  if (!studentName) {
    return (
      <main className="start-screen">
        <section className="start-panel">
          <p className="eyebrow">Grade 6 MCAP Math</p>
          <h1>Practice Arena</h1>
          <p className="start-copy">
            Choose a section, answer MCAP-style questions, and build a score
            report you can review at the end.
          </p>

          <form className="name-form" onSubmit={handleSubmit}>
            <label htmlFor="student-name">Student name</label>
            <div className="name-row">
              <input
                autoComplete="name"
                id="student-name"
                onChange={(event) => setName(event.target.value)}
                placeholder="Type your name"
                type="text"
                value={name}
              />
              <button type="submit">Start Practice</button>
            </div>
          </form>
        </section>
      </main>
    );
  }

  return (
    <main className="app-shell">
      <header className="top-bar">
        <div>
          <p className="eyebrow">Welcome, {studentName}</p>
          <h1>Choose a Practice Section</h1>
          <p>Work through the five sections in any order. Your progress saves on this device.</p>
        </div>
        <button className="danger-button" onClick={onReset} type="button">
          Reset
        </button>
      </header>

      <section className="score-strip" aria-label="Practice summary">
        <div className="score-tile">
          <strong>{totals.answeredCount}</strong>
          <span>Answered</span>
        </div>
        <div className="score-tile">
          <strong>{totals.correctCount}</strong>
          <span>Correct</span>
        </div>
        <div className="score-tile">
          <strong>{totals.totalQuestions}</strong>
          <span>Total</span>
        </div>
      </section>

      <ProgressBar
        current={totals.answeredCount}
        label="Overall progress"
        total={totals.totalQuestions}
      />

      <section className="level-grid" aria-label="Practice sections">
        {levels.map((level, index) => {
          const levelScore = getLevelScore(level, answers);

          return (
            <article className={`level-card ${level.color}`} key={level.id}>
              <div className="level-number">Level {index + 1}</div>
              <h2>{level.title}</h2>
              <p>{level.description}</p>
              <div className="level-meta">
                <span>{levelScore.answeredCount} answered</span>
                <span>{levelScore.correctCount} correct</span>
              </div>
              <button onClick={() => onChooseLevel(level.id)} type="button">
                {levelScore.isComplete ? "Review Level" : "Practice"}
              </button>
            </article>
          );
        })}
      </section>

      <section className="report-callout">
        <div>
          <h2>Ready to check your score?</h2>
          <p>You can view your report anytime. Unanswered questions count as not finished.</p>
        </div>
        <button onClick={onShowResults} type="button">
          View Report
        </button>
      </section>
    </main>
  );
}

export default ActivityShell;
