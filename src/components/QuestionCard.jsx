import ProgressBar from "./ProgressBar.jsx";

function QuestionCard({ answer, level, onAnswer, onBack, onNext, question, questionNumber, totalQuestions }) {
  const hasAnswered = Boolean(answer);
  const answeredInLevel = questionNumber - 1 + (hasAnswered ? 1 : 0);
  const isLastQuestion = questionNumber === totalQuestions;

  return (
    <main className="question-shell">
      <header className="question-top">
        <button className="ghost-button" onClick={onBack} type="button">
          Level Map
        </button>
        <p>
          Question {questionNumber} of {totalQuestions}
        </p>
      </header>

      <ProgressBar current={answeredInLevel} label={level.title} total={totalQuestions} />

      <section className={`question-card ${level.color}`}>
        <p className="eyebrow">{level.title}</p>
        <h1>{question.prompt}</h1>
        <p className="skill-tag">{question.skill}</p>

        <div className="choice-list">
          {question.choices.map((choice, index) => {
            const isSelected = answer?.selectedIndex === index;
            const isCorrectChoice = question.answerIndex === index;
            let className = "choice-button";

            if (hasAnswered && isCorrectChoice) className += " correct";
            if (hasAnswered && isSelected && !isCorrectChoice) className += " incorrect";

            return (
              <button
                className={className}
                disabled={hasAnswered}
                key={choice}
                onClick={() => onAnswer(question, index)}
                type="button"
              >
                <span>{String.fromCharCode(65 + index)}</span>
                {choice}
              </button>
            );
          })}
        </div>

        {hasAnswered && (
          <div className={answer.isCorrect ? "feedback correct" : "feedback incorrect"} role="status">
            <strong>{answer.isCorrect ? "Correct" : "Try this idea"}</strong>
            <p>{question.explanation}</p>
          </div>
        )}

        <div className="question-actions">
          <button className="ghost-button" onClick={onBack} type="button">
            Pause
          </button>
          <button disabled={!hasAnswered} onClick={onNext} type="button">
            {isLastQuestion ? "Finish Level" : "Next"}
          </button>
        </div>
      </section>
    </main>
  );
}

export default QuestionCard;
