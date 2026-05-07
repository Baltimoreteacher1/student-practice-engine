export const emptyProgress = {
  studentName: "",
  answers: {},
  showReport: false,
};

export function getProgressKey(activitySlug) {
  return `studentPracticeEngine:${activitySlug}:progress`;
}

export function loadProgress(activitySlug) {
  try {
    const savedProgress = localStorage.getItem(getProgressKey(activitySlug));
    return savedProgress ? { ...emptyProgress, ...JSON.parse(savedProgress) } : { ...emptyProgress };
  } catch {
    return { ...emptyProgress };
  }
}

export function saveProgress(activitySlug, progress) {
  try {
    localStorage.setItem(getProgressKey(activitySlug), JSON.stringify(progress));
  } catch {
    // The activity still works for the current session if browser storage is blocked.
  }
}

export function clearProgress(activitySlug) {
  try {
    localStorage.removeItem(getProgressKey(activitySlug));
  } catch {
    // Nothing else is needed if browser storage is blocked.
  }
}
