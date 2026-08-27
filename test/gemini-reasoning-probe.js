/**
 * Diagnostic: call Gemini's OpenAI-compatible endpoint and dump every
 * delta field seen in the SSE stream, so we can see how thinking/reasoning
 * is actually surfaced.
 *
 * Usage:
 *   set GEMINI_KEY=AIza...
 *   node test/gemini-reasoning-probe.js
 *
 * Or pass the key as an argument:
 *   node test/gemini-reasoning-probe.js AIza...
 */
const key = process.argv[2] || process.env.GEMINI_KEY;
if (!key) {
  console.error('Usage: node test/gemini-probe.js <GEMINI_API_KEY>');
  console.error('   or: set GEMINI_KEY=AIza... && node test/gemini-probe.js');
  process.exit(1);
}

const MODEL = process.argv[3] || 'gemini-3.6-flash';

async function probe() {
  const noThinking = process.argv.includes('--no-thinking');
  const direct = process.argv.includes('--direct');

  // Three modes:
  //  default: extra_body.google.thinking_config (what the add-in sends)
  //  --direct: google.thinking_config at top level (no extra_body wrapper)
  //  --no-thinking: no thinking config at all (baseline)
  const thinkingConfig = { include_thoughts: true };

  const body = {
    model: MODEL,
    messages: [
      { role: 'system', content: 'You are a helpful assistant. Think step by step.' },
      { role: 'user', content: 'What is 17 * 23? Show your reasoning.' }
    ],
    stream: true,
    stream_options: { include_usage: true }
  };

  if (!noThinking) {
    if (direct) {
      body.google = { thinking_config: thinkingConfig };
    } else {
      body.extra_body = { google: { thinking_config: thinkingConfig } };
    }
  }

  const mode = noThinking ? 'NO thinking config' : direct ? 'direct google.thinking_config' : 'extra_body.google.thinking_config';
  console.log(`\n=== Probe: ${MODEL} (${mode}) ===`);
  const { messages, ...rest } = body;
  console.log('Request body (excluding messages):');
  console.log(JSON.stringify(rest, null, 2));
  console.log('');

  const resp = await fetch('https://generativelanguage.googleapis.com/v1beta/openai/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${key}`
    },
    body: JSON.stringify(body)
  });

  if (!resp.ok) {
    const text = await resp.text().catch(() => '');
    console.error(`HTTP ${resp.status}: ${text.slice(0, 500)}`);
    process.exit(1);
  }

  const reader = resp.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let chunkCount = 0;
  const allDeltaKeys = new Set();
  let reasoningFound = '';
  let contentFound = '';
  const otherFields = [];

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop() || '';

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) continue;

      if (!trimmed.startsWith('data:')) {
        console.log(`[non-data line] ${trimmed}`);
        continue;
      }

      const jsonStr = trimmed.slice(5).trim();
      if (!jsonStr || jsonStr === '[DONE]') continue;

      let chunk;
      try { chunk = JSON.parse(jsonStr); } catch (e) { continue; }
      chunkCount++;

      const choice = chunk.choices && chunk.choices[0];
      if (!choice) {
        console.log(`[chunk ${chunkCount}] no choices, keys: ${Object.keys(chunk).join(', ')}`);
        continue;
      }

      const delta = choice.delta;
      if (!delta) {
        console.log(`[chunk ${chunkCount}] no delta, finish_reason: ${choice.finish_reason}`);
        continue;
      }

      const deltaKeys = Object.keys(delta);
      deltaKeys.forEach(k => allDeltaKeys.add(k));

      // Log every chunk that has non-content fields, or the first 20
      if (chunkCount <= 20 || deltaKeys.some(k => k !== 'content')) {
        const summary = {};
        for (const k of deltaKeys) {
          const v = delta[k];
          if (typeof v === 'string') {
            summary[k] = v.length > 80 ? v.slice(0, 80) + '...' : v;
          } else if (v === null || v === undefined) {
            summary[k] = String(v);
          } else if (typeof v === 'object') {
            summary[k] = JSON.stringify(v).slice(0, 120);
          } else {
            summary[k] = String(v);
          }
        }
        console.log(`[chunk ${chunkCount}] delta: ${JSON.stringify(summary)}`);
      }

      if (delta.content) contentFound += delta.content;

      // Check every possible reasoning field name
      for (const field of ['reasoning', 'reasoning_content', 'thought', 'thinking', 'thought_summary', 'thoughtSignature']) {
        if (delta[field] !== undefined) {
          const v = delta[field];
          const text = typeof v === 'string' ? v : (v && v.text ? v.text : JSON.stringify(v));
          reasoningFound += `[${field}] ${text}\n`;
        }
      }

      for (const k of deltaKeys) {
        if (!['content', 'tool_calls', 'role'].includes(k) && !otherFields.includes(k)) {
          otherFields.push(k);
        }
      }
    }
  }

  console.log('\n=== Summary ===');
  console.log(`Total chunks: ${chunkCount}`);
  console.log(`All delta keys seen: ${[...allDeltaKeys].join(', ')}`);
  console.log(`Non-standard fields: ${otherFields.join(', ') || '(none)'}`);
  console.log(`\nContent (${contentFound.length} chars):\n${contentFound.slice(0, 300)}`);
  console.log(`\nReasoning found:\n${reasoningFound || '(none)'}`);

  if (!reasoningFound && !noThinking) {
    console.log('\n! No reasoning/thinking fields found in any delta!');
    console.log('  Try: node test/gemini-reasoning-probe.js <key> <model> --direct');
    console.log('  And: node test/gemini-reasoning-probe.js <key> <model> --no-thinking');
  }
}

probe().catch(e => { console.error(e); process.exit(1); });
