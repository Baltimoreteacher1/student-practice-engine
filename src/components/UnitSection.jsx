import { unitPath } from "../utils/routing.js";
import ActivityCard from "./ActivityCard.jsx";

function UnitSection({ activities, onNavigate, unit, unitSlug }) {
  return (
    <section className="unit-section">
      <div className="unit-header">
        <div>
          <p className="eyebrow">Unit</p>
          <h2>{unit}</h2>
        </div>
        <button className="button button-secondary" onClick={() => onNavigate(unitPath(unitSlug))} type="button">
          View Unit
        </button>
      </div>
      <div className="activity-grid">
        {activities.map((activity) => (
          <ActivityCard activity={activity} key={activity.slug} onNavigate={onNavigate} />
        ))}
      </div>
    </section>
  );
}

export default UnitSection;
