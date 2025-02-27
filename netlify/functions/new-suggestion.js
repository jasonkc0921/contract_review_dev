const fetch = require('node-fetch');

exports.handler = async function(event, context) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  try {
    const body = JSON.parse(event.body);
    const { originalText } = body;
    
    const openaiKey = process.env.OPENAI_API_KEY;
    
    if (!openaiKey) {
      return {
        statusCode: 500,
        body: JSON.stringify({ error: "API key not configured" })
      };
    }
    
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${openaiKey}`
      },
      body: JSON.stringify({
        model: "gpt-3.5-turbo",
        messages: [
          { 
            role: "system", 
            content: "You are a labor law advisor specialized in Malaysian labor law. You're being asked to revise a specific clause from an employment contract to ensure it conforms to Malaysian employment law. Provide ONLY the revised text without any explanations or preamble."
          },
          { 
            role: "user", 
            content: `Please review and revise the following employment contract clause to fully comply with Malaysian labor law. Return ONLY the revised text: "${originalText}"`
          }
        ],
        max_tokens: 1000,
        temperature: 0.5
      })
    });

    const data = await response.json();
    
    if (data.choices && data.choices.length > 0) {
      return {
        statusCode: 200,
        body: JSON.stringify({ text: data.choices[0].message.content })
      };
    } else {
      return {
        statusCode: 500,
        body: JSON.stringify({ error: "No suggestion received from AI" })
      };
    }
  } catch (error) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.message })
    };
  }
};