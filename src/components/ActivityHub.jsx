import { getUnits } from "../data/activityCatalog.js";
import UnitSection from "./UnitSection.jsx";

function ActivityHub({ onNavigate }) {
  const units = getUnits();

  return (
    <div className="activity-hub">
      {units.map((unit) => (
        <UnitSection
          activities={unit.activities}
          key={unit.unitSlug}
          onNavigate={onNavigate}
          unit={unit.unit}
          unitSlug={unit.unitSlug}
        />
      ))}
    </div>
  );
}

export default ActivityHub;
