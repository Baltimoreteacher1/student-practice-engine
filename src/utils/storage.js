export const STORAGE_KEY = "mcap-grade-6-practice-progress";

export const emptyProgress = {
  studentName: "",
  currentLevelId: "",
  currentQuestionIndex: 0,
  answers: {},
  showResults: false,
};

export function loadProgress() {
  try {
    const savedProgress = localStorage.getItem(STORAGE_KEY);
    return savedProgress ? { ...emptyProgress, ...JSON.parse(savedProgress) } : { ...emptyProgress };
  } catch {
    return { ...emptyProgress };
  }
}

export function saveProgress(progress) {
  try {
    if (!progress.studentName) {
      localStorage.removeItem(STORAGE_KEY);
      return;
    }

    localStorage.setItem(STORAGE_KEY, JSON.stringify(progress));
  } catch {
    // The app still works for the current session if browser storage is blocked.
  }
}

export function clearProgress() {
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // Nothing else is needed if browser storage is blocked.
  }
}
