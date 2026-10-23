/**
 * External AI Service Module (Practical 13)
 * Integrates OpenAI and Gemini APIs with Rate Limiting, Responsible AI Prompts,
 * and Graceful Fallback Handling.
 */

const OpenAI = require('openai');

// Rate limiting map: tracks timestamp of last AI request per client/IP
// Supplementary Problem 2: Max 1 AI request per minute (60,000 ms)
const rateLimitCache = new Map();
const RATE_LIMIT_WINDOW_MS = 60 * 1000; // 60 seconds

/**
 * Checks and updates rate limit for a client/task
 * @param {string} key Identifier (IP or User ID or Task Title)
 * @returns {{ allowed: boolean, waitSeconds?: number }}
 */
function checkRateLimit(key) {
  const now = Date.now();
  if (rateLimitCache.has(key)) {
    const lastRequestTime = rateLimitCache.get(key);
    const elapsed = now - lastRequestTime;
    if (elapsed < RATE_LIMIT_WINDOW_MS) {
      const waitSeconds = Math.ceil((RATE_LIMIT_WINDOW_MS - elapsed) / 1000);
      return { allowed: false, waitSeconds };
    }
  }
  rateLimitCache.set(key, now);
  return { allowed: true };
}

/**
 * Intelligent Contextual Fallback Generator (Step 6 & Problem Definition)
 * Activated when API key is missing, invalid, or external API times out/fails.
 * Guarantees the application never crashes and continues normal operation.
 */
function generateFallbackSuggestion(title) {
  const lower = title.toLowerCase();
  let suggestedPriority = 'medium';
  let suggestedDesc = `Complete the activities required for "${title}". Ensure all deliverables and guidelines are verified.`;

  if (lower.includes('exam') || lower.includes('urgent') || lower.includes('bug') || lower.includes('critical') || lower.includes('deadline')) {
    suggestedPriority = 'high';
    suggestedDesc = `High-priority requirement for "${title}". Investigate root causes, execute necessary remediations, and verify resolution promptly.`;
  } else if (lower.includes('submit') || lower.includes('assignment') || lower.includes('practical') || lower.includes('lab') || lower.includes('report')) {
    suggestedPriority = 'high';
    suggestedDesc = `Prepare, compile, and double-check all required documentation and source files for "${title}" prior to official submission.`;
  } else if (lower.includes('read') || lower.includes('explore') || lower.includes('optional') || lower.includes('idea')) {
    suggestedPriority = 'low';
    suggestedDesc = `Review background documentation and reference materials for "${title}" when core priorities have been addressed.`;
  } else if (lower.includes('docker') || lower.includes('deploy') || lower.includes('ci') || lower.includes('test')) {
    suggestedPriority = 'medium';
    suggestedDesc = `Configure, run tests, and validate container and workflow execution for "${title}" ensuring zero regressions.`;
  }

  return {
    description: suggestedDesc,
    priority: suggestedPriority,
    fallback: true,
    provider: 'Contextual Fallback Engine',
    disclaimer: 'AI suggestion generated via local heuristic fallback. Please review and edit before saving.',
  };
}

/**
 * Generate Task Description & Priority using OpenAI (or Gemini) with Fallback
 * @param {string} title Task title provided by user
 * @param {string} clientId Client IP or User ID for rate limiting
 * @returns {Promise<object>} Generated description, priority, and metadata
 */
async function generateTaskDescription(title, clientId = 'default') {
  // 1. Enforce Rate Limit (Supplementary Problem 2)
  const rateCheck = checkRateLimit(clientId);
  if (!rateCheck.allowed) {
    return {
      error: 'RateLimitExceeded',
      status: 429,
      message: `Rate limit reached: Max 1 AI request per minute. Please wait ${rateCheck.waitSeconds}s.`,
      waitSeconds: rateCheck.waitSeconds,
    };
  }

  const openAiKey = process.env.OPENAI_API_KEY;
  const isKeyPlaceholder = !openAiKey || openAiKey === 'your_openai_api_key_here' || openAiKey.trim() === '';

  // 2. If OpenAI key is not configured, gracefully trigger contextual fallback
  if (isKeyPlaceholder) {
    console.log(`[AI Service] No valid external API key configured. Utilizing Contextual Fallback for title: "${title}".`);
    return generateFallbackSuggestion(title);
  }

  // 3. Attempt live OpenAI API call
  try {
    const openai = new OpenAI({ apiKey: openAiKey });
    const prompt = `You are a helpful task management assistant.
Given the task title: "${title}", provide:
1. A concise, professional, actionable 1-2 sentence description.
2. A suggested priority level (low, medium, or high).
Format your response exactly as:
Description: [Your 1-2 sentence description]
Priority: [low|medium|high]`;

    const completion = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 150,
      temperature: 0.7,
    });

    const rawContent = completion.choices[0]?.message?.content || '';
    
    // Parse description and priority from model response
    let description = rawContent;
    let priority = 'medium';

    const descMatch = rawContent.match(/Description:\s*(.+)/i);
    const prioMatch = rawContent.match(/Priority:\s*(low|medium|high)/i);

    if (descMatch) {
      description = descMatch[1].trim();
    }
    if (prioMatch) {
      priority = prioMatch[1].toLowerCase().trim();
    }

    return {
      description,
      priority,
      fallback: false,
      provider: 'OpenAI (gpt-4o-mini)',
      disclaimer: 'AI-generated suggestion powered by OpenAI. Please review and modify before saving.',
    };
  } catch (err) {
    // 4. Graceful Error Handling & Fallback (Problem Definition & Step 6)
    console.warn(`[AI Service Warning] External AI API failed: ${err.message}. Gracefully falling back.`);
    const fallback = generateFallbackSuggestion(title);
    fallback.apiError = err.message;
    return fallback;
  }
}

module.exports = {
  generateTaskDescription,
  checkRateLimit,
  generateFallbackSuggestion,
};
