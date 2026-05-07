import assert from "node:assert/strict";
import test from "node:test";

import { activityCatalog } from "../src/data/activityCatalog.js";
import { questionBanks } from "../src/data/activities/index.js";
import { parsePath } from "../src/utils/routing.js";
import { calculateActivityScore, gradeAnswer } from "../src/utils/scoring.js";
import { getProgressKey } from "../src/utils/storage.js";

test("parsePath supports home, unit, activity, and not found routes", () => {
  assert.deepEqual(parsePath("/"), { page: "home" });
  assert.deepEqual(parsePath("/units/statistics"), {
    page: "unit",
    unitSlug: "statistics",
  });
  assert.deepEqual(parsePath("/activities/mean-median-mode-review"), {
    page: "activity",
    activitySlug: "mean-median-mode-review",
  });
  assert.deepEqual(parsePath("/not-a-real-page"), { page: "notFound" });
});

test("activity catalog maps the review activity to a question bank", () => {
  const activity = activityCatalog.find(
    (item) => item.slug === "mean-median-mode-review",
  );

  assert.equal(activity?.questionBankId, "meanMedianModeReview");
  assert.equal(questionBanks[activity.questionBankId].id, "meanMedianModeReview");
});

test("progress keys are separate for each activity", () => {
  assert.equal(
    getProgressKey("mean-median-mode-review"),
    "studentPracticeEngine:mean-median-mode-review:progress",
  );
  assert.notEqual(
    getProgressKey("mean-median-mode-review"),
    getProgressKey("exponents-warmup"),
  );
});

test("gradeAnswer handles multiple choice and short answer questions", () => {
  assert.equal(
    gradeAnswer(
      {
        id: "q1",
        type: "multipleChoice",
        choices: ["7", "8", "9"],
        correctAnswer: "8",
        points: 1,
      },
      "8",
    ).isCorrect,
    true,
  );

  assert.equal(
    gradeAnswer(
      {
        id: "q2",
        type: "shortAnswer",
        correctAnswer: "11",
        points: 2,
      },
      " 11 ",
    ).pointsEarned,
    2,
  );
});

test("calculateActivityScore returns totals and section breakdowns", () => {
  const sections = [
    {
      id: "practice",
      title: "Practice",
      questions: [
        {
          id: "q1",
          type: "multipleChoice",
          correctAnswer: "A",
          skill: "Mean",
          points: 1,
        },
        {
          id: "q2",
          type: "shortAnswer",
          correctAnswer: "12",
          skill: "Median",
          points: 2,
        },
      ],
    },
  ];

  const result = calculateActivityScore(sections, { q1: "A", q2: "10" });

  assert.equal(result.answeredCount, 2);
  assert.equal(result.totalQuestions, 2);
  assert.equal(result.correctCount, 1);
  assert.equal(result.pointsEarned, 1);
  assert.equal(result.totalPoints, 3);
  assert.equal(result.percent, 33);
  assert.equal(result.allComplete, true);
  assert.deepEqual(result.sectionBreakdown[0], {
    id: "practice",
    title: "Practice",
    answeredCount: 2,
    totalQuestions: 2,
    correctCount: 1,
    pointsEarned: 1,
    totalPoints: 3,
    percent: 33,
  });
});
