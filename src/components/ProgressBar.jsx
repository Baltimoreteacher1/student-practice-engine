function ProgressBar({ current, label = "Progress", total }) {
  const percent = total === 0 ? 0 : Math.round((current / total) * 100);

  return (
    <div className="progress-block" aria-label={`${label}: ${percent}%`}>
      <div className="progress-meta">
        <span>{label}</span>
        <strong>{percent}%</strong>
      </div>
      <div className="progress-track">
        <div className="progress-fill" style={{ width: `${percent}%` }} />
      </div>
    </div>
  );
}

export default ProgressBar;
