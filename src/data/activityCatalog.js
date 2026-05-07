export const activityCatalog = [
  {
    slug: "mean-median-mode-review",
    title: "Mean, Median, and Mode Review",
    unit: "Statistics",
    unitSlug: "statistics",
    grade: "Grade 6",
    activityType: "Mission Review",
    estimatedMinutes: 45,
    description: "Practice mean, median, mode, and mixed data questions.",
    standardsOrSkills: ["Mean", "Median", "Mode", "Data Displays"],
    questionBankId: "meanMedianModeReview",
    status: "ready",
  },
  {
    slug: "exponents-warmup",
    title: "Exponents Warmup",
    unit: "Expressions",
    unitSlug: "expressions",
    grade: "Grade 6",
    activityType: "Warmup",
    estimatedMinutes: 15,
    description: "Review exponent vocabulary and simple powers before class practice.",
    standardsOrSkills: ["Exponents", "Repeated Multiplication", "Numerical Expressions"],
    questionBankId: "exponentsWarmup",
    status: "ready",
  },
  {
    slug: "geometry-area-review",
    title: "Geometry Area Review",
    unit: "Geometry",
    unitSlug: "geometry",
    grade: "Grade 6",
    activityType: "Skill Review",
    estimatedMinutes: 20,
    description: "Practice finding area of rectangles and triangles with friendly numbers.",
    standardsOrSkills: ["Area", "Rectangles", "Triangles"],
    questionBankId: "sampleActivity",
    status: "ready",
  },
];

export function getActivityBySlug(activitySlug) {
  return activityCatalog.find((activity) => activity.slug === activitySlug);
}

export function getActivitiesByUnit(unitSlug) {
  return activityCatalog.filter((activity) => activity.unitSlug === unitSlug);
}

export function getUnits() {
  const unitsBySlug = new Map();

  activityCatalog.forEach((activity) => {
    if (!unitsBySlug.has(activity.unitSlug)) {
      unitsBySlug.set(activity.unitSlug, {
        title: activity.unit,
        slug: activity.unitSlug,
        activities: [],
      });
    }

    unitsBySlug.get(activity.unitSlug).activities.push(activity);
  });

  return [...unitsBySlug.values()];
}
