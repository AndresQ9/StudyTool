const { JSDOM } = require('jsdom');
const { window } = new JSDOM(``).window;
global.window = window;
const officeParser = require('officeparser');
const fs = require('fs');
//const pipe = require('@huggingface/transformers');
const { pipeline } = require('@huggingface/transformers');
const { OpenAI } = require('openai');
const { SummarizationPipeline } = require('@huggingface/transformers');

const openai = new OpenAI({
  apiKey: 'YOUR_API_KEY HERE', dangerouslyAllowBrowser: true
});

async function extractTextFromPPTX(filePath) {
  return new Promise((resolve, reject) => {
    officeParser.parseOffice(filePath, (data, err) => {
      if (err) {
        console.log("Error is : " + err);
        //reject(err);
      } else {
        //console.log("Extracted data:\n " + data);
        resolve(data);
      }
    });
  });
}


async function summarizeText(text, maxLength = 1024) {
  try {
    const response = await openai.chat.completions.create({
        model: "gpt-4o-mini",
        messages: [
            {
                role: "system",
                content: "You are a study assistant used to summarize large amounts of information presented in class. Please summarize the following text while maintaining all important and key bits of information.:"
            },
            {
                role: "user",
                content: text
            }
        ],
  
    });
    return response.choices[0].message.content;
} catch (error) {
    console.error('Error in summarization:', error);
    throw error;
}
  
}
/* Initially used to summarize slides but we are no longer processing slide by slide
async function summarizeSlides(textData, maxLength = 40) {
  const summaries = [];
  for (const text of textData) {
    const summary = await summarizeText(text, maxLength);
    summaries.push(summary);
  }
  return summaries;
}
*/
async function generateQuizQuestions(textData, maxLength = 50) {
  // Create the string of summarized data
  //const combinedText = textData.join(" ");

  const messages = [
    { role: "system", content: "You are a helpful assistant." },
    { role: "user", content: `Based on the following summarized content, generate a list of quiz questions:\n\n${textData}\n\nPlease make the questions varied and engaging, and ensure they are suitable for a quiz. Also provide multiple choice answers to the questions.` }
  ];

  try {
    // Prompt the model to create the quiz
    const response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: messages,
      max_tokens: 1000,
      temperature: 0.7 
    });
    const generatedText = response.choices[0].message.content.trim();
    const questions = generatedText.split("\n").filter(q => q.trim() !== "");

    console.log(questions);
    return questions;
  } catch (error) {
    console.error("Error generating quiz questions:", error);
  }
}

// Example usage
(async () => {
  try{
    //const filePath = "C:/Users/andre/Downloads/COP 4331 - Gathering Requirements.pptx";
    const filePath = "C:/Users/andre/Downloads/Projects.docx";
    const textData = await extractTextFromPPTX(filePath);
    console.log("Printing the text:\n" +textData);

    const summaries = await summarizeText(textData);
    console.log("Starting summaries:\n" + summaries);

    const quiz = await generateQuizQuestions(summaries);
    //console.log("Here is your quiz:\n" + quiz);
    for (const question of quiz) {
      console.log(question);
    }
  }
  catch(error){
    console.log(error);
  }
})();
