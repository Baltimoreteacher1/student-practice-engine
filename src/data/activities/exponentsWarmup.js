export const exponentsWarmup = {
  id: "exponentsWarmup",
  title: "Exponents Warmup",
  directions: "Use repeated multiplication to think about each exponent.",
  sections: [
    {
      id: "exponent-basics",
      title: "Exponent Basics",
      directions: "Review what the base and exponent mean.",
      questions: [
        {
          id: "exp-1",
          type: "multipleChoice",
          prompt: "What does 3^2 mean?",
          choices: ["3 + 2", "3 x 2", "3 x 3", "2 x 2 x 2"],
          correctAnswer: "3 x 3",
          explanation: "The exponent 2 means use 3 as a factor two times.",
          skill: "Exponents",
          points: 1,
        },
        {
          id: "exp-2",
          type: "shortAnswer",
          prompt: "What is 2^4?",
          correctAnswer: "16",
          explanation: "2 x 2 x 2 x 2 = 16.",
          skill: "Exponents",
          points: 1,
        },
      ],
    },
  ],
};
