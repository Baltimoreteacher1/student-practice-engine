import ActivityCard from "../components/ActivityCard.jsx";
import { getActivitiesByUnit, getUnits } from "../data/activityCatalog.js";

function UnitPage({ onNavigate, unitSlug }) {
  const activities = getActivitiesByUnit(unitSlug);
  const unit = getUnits().find((item) => item.unitSlug === unitSlug);

  return (
    <main className="page-shell">
      <button className="button button-secondary" onClick={() => onNavigate("/")} type="button">
        Back Home
      </button>
      <section className="page-heading">
        <p className="eyebrow">Unit</p>
        <h1>{unit?.unit ?? "Unit Not Found"}</h1>
        <p>
          {activities.length} activity{activities.length === 1 ? "" : "ies"} ready for students.
        </p>
      </section>

      {activities.length === 0 ? (
        <section className="empty-state">
          <h2>No activities here yet</h2>
          <p>Add an activity to the catalog with this unit slug when you are ready.</p>
        </section>
      ) : (
        <div className="activity-grid">
          {activities.map((activity) => (
            <ActivityCard activity={activity} key={activity.slug} onNavigate={onNavigate} />
          ))}
        </div>
      )}
    </main>
  );
}

export default UnitPage;
