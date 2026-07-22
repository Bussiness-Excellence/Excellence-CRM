export const config = {
  runtime: 'edge',
};

export default async function handler(req) {
  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), {
      status: 405,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  try {
    const { messages, contextData } = await req.json();

    if (!messages || !Array.isArray(messages)) {
      return new Response(JSON.stringify({ error: 'Missing or invalid messages array' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
      });
    }

    const systemPrompt = `You are a highly intelligent CRM data analyst for EXCELLENCE CRM.
Your job is to answer questions about the sales, field activity, doctor calls, and product/specialty data provided to you.
Be extremely brief, concise, and professional. Only output markdown.
Do not hallucinate data. If the answer is not in the context data, say you cannot find it in the current view.
Here is the raw context data for the currently filtered view on the user's dashboard:
${JSON.stringify(contextData).slice(0, 8000)} // Truncated to avoid token limits
`;

    const apiMessages = [
      { role: 'system', content: systemPrompt },
      ...messages.map(m => ({ role: m.role, content: m.content }))
    ];

    // Using Pollinations.ai for completely free, keyless AI text generation
    const response = await fetch('https://text.pollinations.ai/openai', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'openai', // uses the default free model (often Llama 3 or GPT-4o-mini equivalent)
        messages: apiMessages,
        temperature: 0.1,
        jsonMode: false
      })
    });

    const data = await response.json();
    if (!response.ok) throw new Error(data.error?.message || 'AI API Error');
    
    return new Response(JSON.stringify({ reply: data.choices[0].message.content }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });

  } catch (err) {
    console.error('AI Error:', err);
    return new Response(JSON.stringify({ error: 'Failed to process AI request' }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
    });
  }
}
