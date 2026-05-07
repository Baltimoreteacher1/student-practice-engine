export const meanMedianModeQuestions = [
  {
    id: "mmm-1",
    skill: "Mean",
    prompt: "The quiz scores are 6, 8, 8, 10, and 13. What is the mean score?",
    choices: ["8", "9", "10", "13"],
    answerIndex: 1,
    explanation: "Add the scores to get 45, then divide by 5 scores. The mean is 9.",
  },
  {
    id: "mmm-2",
    skill: "Median",
    prompt: "The data set is 12, 15, 19, 21, 28. What is the median?",
    choices: ["15", "19", "21", "28"],
    answerIndex: 1,
    explanation: "The median is the middle value when the numbers are in order. The middle value is 19.",
  },
  {
    id: "mmm-3",
    skill: "Mode",
    prompt: "The number of books read by students is 3, 4, 4, 5, 6, 6, 6. What is the mode?",
    choices: ["4", "5", "6", "No mode"],
    answerIndex: 2,
    explanation: "The mode is the value that appears most often. The number 6 appears three times.",
  },
];
