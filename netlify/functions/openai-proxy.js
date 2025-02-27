const fetch = require('node-fetch');

exports.handler = async function(event, context) {
  // Only allow POST requests
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  try {
    const body = JSON.parse(event.body);
    const { text, prompt, model } = body;
    
    // Access API key from environment (set in Netlify dashboard)
    const openaiKey = process.env.OPENAI_API_KEY;
    
    if (!openaiKey) {
      return {
        statusCode: 500,
        body: JSON.stringify({ error: "API key not configured" })
      };
    }
    
    // Default system message for Malaysian labor law
    const systemMessage = "You are a labor law advisor who is specialized in labor law in Malaysia.";
    
    // Default prompt for document review
    const defaultPrompt = `please review the attached employment contract and suggest sections that need to be amended so that it conforms to Malaysia employment law in following text, output them in Json object that consists of original text and recommended text; the key for original text is "original_text"., and key for recommended text is "recommended_text"; you should avoid capturing original texts that are title of a section, typically short wordings, less than 10 words, that ended with 2 new lines "\\n\\n" or colon or dash- "${text}"`;
    
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${openaiKey}`
      },
      body: JSON.stringify({
        model: model || "gpt-4",
        messages: [
          { role: "system", content: prompt?.systemMessage || systemMessage },
          { role: "user", content: prompt?.userMessage || defaultPrompt }
        ],
        max_tokens: 2500
      })
    });
    
    const data = await response.json();
    
    return {
      statusCode: 200,
      body: JSON.stringify(data)
    };
  } catch (error) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.message })
    };
  }
};