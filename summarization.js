const pptxParser = require('pptx-parser');
const Tesseract = require('tesseract.js');
const axios = require('axios');
const fs = require('fs');
const OpenAI = require('openai');
const openai = new OpenAI(process.env.OPENAI_API_KEY);

async function extractTextFromPPTX(filePath, imageOutputDir = 'images') {
  const pptx = await pptxParser.parse(filePath);
  const textContent = [];

  if (!fs.existsSync(imageOutputDir)) {
    fs.mkdirSync(imageOutputDir);
  }

  for (let i = 0; i < pptx.slides.length; i++) {
    const slide = pptx.slides[i];
    let slideText = [];

    // Extract text from shapes
    slide.shapes.forEach(shape => {
      if (shape.text) {
        slideText.push(shape.text);
      }
    });

    // Extract text from images (using Tesseract.js)
    for (let j = 0; j < slide.images.length; j++) {
      const image = slide.images[j];
      const imagePath = `${imageOutputDir}/slide_${i + 1}_image_${j + 1}.png`;
      fs.writeFileSync(imagePath, image.buffer);

      const { data: { text } } = await Tesseract.recognize(imagePath, 'eng');
      slideText.push(`Image OCR: ${text}`);
    }

    textContent.push(slideText.join(' '));
  }

  return textContent;
}

async function summarizeText(text, maxLength = 40) {
  const response = await axios.post(
    'https://api-inference.huggingface.co/models/facebook/bart-large-cnn',
    {
      inputs: text,
      parameters: { max_length: maxLength, min_length: 30 }
    },
    {
      headers: { Authorization: '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' }
    }
  );

  return response.data[0].summary_text;
}

async function summarizeSlides(textData, maxLength = 40) {
  const summaries = [];
  for (const text of textData) {
    const summary = await summarizeText(text, maxLength);
    summaries.push(summary);
  }
  return summaries;
}

async function generateQuizQuestions(textData, maxLength = 50) {
  // Create the string of summarized data
  const combinedText = textData.join(" ");

  const messages = [
    { role: "system", content: "You are a helpful assistant." },
    { role: "user", content: `Based on the following summarized content, generate a list of quiz questions:\n\n${combinedText}\n\nPlease make the questions varied and engaging, and ensure they are suitable for a quiz.` }
  ];

  try {
    // Prompt the model to create the quiz
    const response = await openai.createChatCompletion({
      model: "gpt-3.5-turbo",
      messages: messages,
      max_tokens: 1000,
      temperature: 0.7 
    });

    const generatedText = response.data.choices[0].message.content.trim();
    const questions = generatedText.split("\n").filter(q => q.trim() !== "");

    console.log(questions);
    return questions;
  } catch (error) {
    console.error("Error generating quiz questions:", error);
  }
}

// Example usage
(async () => {
  const filePath = '8 - CAP 4611 - Logistic Regression.pptx';
  const textData = await extractTextFromPPTX(filePath);
  console.log(textData);

  const summaries = await summarizeSlides(textData);
  console.log(summaries);

  const quiz = await generateQuizQuestions(summaries);
  console.log(quiz);
})();
