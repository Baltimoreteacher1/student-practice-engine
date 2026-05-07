export const sampleActivity = {
  id: "sampleActivity",
  title: "Geometry Area Review",
  directions: "Find the area. Remember to use square units.",
  sections: [
    {
      id: "area-review",
      title: "Area Review",
      directions: "Use friendly numbers and check your formula.",
      questions: [
        {
          id: "area-1",
          type: "shortAnswer",
          prompt: "A rectangle is 5 units long and 4 units wide. What is its area?",
          correctAnswer: "20",
          explanation: "Area of a rectangle is length x width. 5 x 4 = 20 square units.",
          skill: "Rectangle Area",
          points: 1,
        },
        {
          id: "area-2",
          type: "multipleChoice",
          prompt: "A triangle has a base of 8 units and a height of 3 units. What is its area?",
          choices: ["11", "12", "24", "48"],
          correctAnswer: "12",
          explanation: "Triangle area is one-half of base x height. Half of 8 x 3 is 12.",
          skill: "Triangle Area",
          points: 1,
        },
      ],
    },
  ],
};
