function QuestionCard({ answer, onAnswer, question }) {
  const hasAnswered = Boolean(answer);
  const currentValue = answer?.studentAnswer ?? "";

  function handleShortAnswerSubmit(event) {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    onAnswer(question, formData.get("shortAnswer"));
  }

  return (
    <article className="question-card">
      <div className="question-header">
        <span>{question.skill}</span>
        <strong>
          {question.points} point{question.points === 1 ? "" : "s"}
        </strong>
      </div>
      <h3>{question.prompt}</h3>

      {question.type === "multipleChoice" ? (
        <div className="choice-list">
          {question.choices.map((choice) => {
            const isSelected = currentValue === choice;
            const isCorrect = hasAnswered && choice === question.correctAnswer;
            const isIncorrect = hasAnswered && isSelected && !isCorrect;
            let className = "choice-button";
            if (isCorrect) className += " correct";
            if (isIncorrect) className += " incorrect";

            return (
              <button
                className={className}
                disabled={hasAnswered}
                key={choice}
                onClick={() => onAnswer(question, choice)}
                type="button"
              >
                {choice}
              </button>
            );
          })}
        </div>
      ) : (
        <form className="short-answer-form" onSubmit={handleShortAnswerSubmit}>
          <input
            aria-label="Short answer"
            defaultValue={currentValue}
            disabled={hasAnswered}
            name="shortAnswer"
            placeholder="Type your answer"
            type="text"
          />
          <button className="button" disabled={hasAnswered} type="submit">
            Check
          </button>
        </form>
      )}

      {hasAnswered && (
        <div className={answer.isCorrect ? "feedback correct" : "feedback incorrect"} role="status">
          <strong>{answer.isCorrect ? "Correct" : "Review this"}</strong>
          <p>{question.explanation}</p>
        </div>
      )}
    </article>
  );
}

export default QuestionCard;
