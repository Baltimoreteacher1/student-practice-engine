import { activityPath } from "../utils/routing.js";

function ActivityCard({ activity, onNavigate }) {
  return (
    <article className="activity-card">
      <div className="card-topline">
        <span>{activity.grade}</span>
        <span>{activity.estimatedMinutes} min</span>
      </div>
      <h3>{activity.title}</h3>
      <p>{activity.description}</p>
      <div className="activity-meta">
        <span>{activity.activityType}</span>
        <span className="status-pill">{activity.status}</span>
      </div>
      <div className="skill-list" aria-label="Skills">
        {activity.standardsOrSkills.map((skill) => (
          <span key={skill}>{skill}</span>
        ))}
      </div>
      <button className="button button-full" onClick={() => onNavigate(activityPath(activity.slug))} type="button">
        Open Activity
      </button>
    </article>
  );
}

export default ActivityCard;
