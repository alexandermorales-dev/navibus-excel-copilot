/* Tests for the transport layer: tolerant JSON extraction and SSE parsing.
   Run with: node --test test/ */

const test = require('node:test');
const assert = require('node:assert');
const { loadCore } = require('./helper.js');

const { LLM } = loadCore();

test('extractJSON: plain object', () => {
  const r = LLM.extractJSON('{"a":1}');
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(r.value, { a: 1 });
});

test('extractJSON: fenced code block with trailing comma', () => {
  const r = LLM.extractJSON('```json\n{"a":1,"b":[1,2,]}\n```');
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(r.value, { a: 1, b: [1, 2] });
});

test('extractJSON: unfenced fence marker without language', () => {
  const r = LLM.extractJSON('```\n{"ops":[]}\n```');
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(r.value, { ops: [] });
});

test('extractJSON: JSON embedded in prose', () => {
  const r = LLM.extractJSON('Sure! Here is the plan:\n{"ops":[{"op":"add_sheet"}]}\nHope that helps.');
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(r.value, { ops: [{ op: 'add_sheet' }] });
});

test('extractJSON: truncated output gets closed', () => {
  const r = LLM.extractJSON('{"ops":[{"op":"add_sheet","name":"X"}');
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(r.value, { ops: [{ op: 'add_sheet', name: 'X' }] });
});

test('extractJSON: braces inside string values are not treated as delimiters', () => {
  const r = LLM.extractJSON('{"answer":"he said \\"hi\\" {not json}"}');
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.value.answer, 'he said "hi" {not json}');
});

test('extractJSON: trailing comma before closing brace', () => {
  const r = LLM.extractJSON('{"a":1,}');
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(r.value, { a: 1 });
});

test('extractJSON: preamble plus fenced block prefers the fence', () => {
  const r = LLM.extractJSON('Thinking... I will build it.\n```json\n{"intent":"build","ops":[]}\n```\nDone.');
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.value.intent, 'build');
});

test('extractJSON: no JSON present fails cleanly', () => {
  const r = LLM.extractJSON('I cannot do that.');
  assert.strictEqual(r.ok, false);
  assert.ok(r.error);
});

test('extractJSON: empty input fails cleanly', () => {
  assert.strictEqual(LLM.extractJSON('').ok, false);
  assert.strictEqual(LLM.extractJSON(null).ok, false);
});

test('extractJSON: nested objects survive', () => {
  const src = '{"ops":[{"op":"recipe.dashboard","source":{"sheet":"Datos","range":"A1:F10"}}]}';
  const r = LLM.extractJSON(src);
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.value.ops[0].source.sheet, 'Datos');
});

test('extractJSON: smart quotes are normalized', () => {
  const r = LLM.extractJSON('{\u201cintent\u201d:\u201cqa\u201d}');
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.value.intent, 'qa');
});

/* ---------- SSE stream parsing ---------- */

function fakeResponse(chunks) {
  const encoder = new TextEncoder();
  let i = 0;
  return {
    body: {
      getReader() {
        return {
          read() {
            if (i >= chunks.length) return Promise.resolve({ done: true });
            return Promise.resolve({ done: false, value: encoder.encode(chunks[i++]) });
          }
        };
      }
    }
  };
}

test('parseStream: accumulates content across chunks', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"content":"Hel"}}]}\n\n',
    'data: {"choices":[{"delta":{"content":"lo"}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.ok, true);
  assert.strictEqual(out.text, 'Hello');
});

test('parseStream: handles a JSON object split mid-chunk', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"con',
    'tent":"ok"}}]}\n\ndata: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.text, 'ok');
});

test('parseStream: captures reasoning and usage', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"reasoning":"thinking..."}}]}\n\n',
    'data: {"choices":[{"delta":{"content":"answer"}}]}\n\n',
    'data: {"choices":[],"usage":{"total_tokens":123}}\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.reasoning, 'thinking...');
  assert.strictEqual(out.usage.total_tokens, 123);
});

test('parseStream: assembles tool calls by index', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"c1","function":{"name":"read_","arguments":"{\\"a\\":"}}]}}]}\n\n',
    'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"name":"range","arguments":"1}"}}]}}]}\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.functionCalls.length, 1);
  assert.strictEqual(out.functionCalls[0].name, 'read_range');
  assert.deepStrictEqual(out.functionCalls[0].args, { a: 1 });
});

test('parseStream: reports truncation via finish_reason length', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"content":"partial"},"finish_reason":"length"}]}\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.truncated, true);
});

test('parseStream: handles object-shaped reasoning_content (Gemini)', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"reasoning_content":{"text":"Let me think"}}}]}\n\n',
    'data: {"choices":[{"delta":{"reasoning_content":{"text":" about this"}}}]}\n\n',
    'data: {"choices":[{"delta":{"content":"answer"}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.ok, true);
  assert.strictEqual(out.reasoning, 'Let me think about this');
  assert.strictEqual(out.text, 'answer');
});

test('parseStream: handles reasoning_content with signature-only object', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"reasoning_content":{"signature":"abc123"}}}]}\n\n',
    'data: {"choices":[{"delta":{"content":"answer"}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.ok, true);
  assert.strictEqual(out.reasoning, '');   // no text field → no reasoning text
  assert.strictEqual(out.text, 'answer');
});

test('parseStream: onThinking callback fires with object reasoning_content', async () => {
  let thinkingCalls = [];
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"reasoning_content":{"text":"step 1"}}}]}\n\n',
    'data: {"choices":[{"delta":{"reasoning_content":{"text":" step 2"}}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp, null, (full) => thinkingCalls.push(full), null);
  assert.strictEqual(thinkingCalls.length, 2);
  assert.strictEqual(thinkingCalls[0], 'step 1');
  assert.strictEqual(thinkingCalls[1], 'step 1 step 2');
});

/* ---------- Gemini <thought> tag parsing ---------- */

test('parseStream: routes Gemini <thought> content to reasoning, not text', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"content":"<thought>Let me think","extra_content":"{\\"google\\":{\\"thought\\":true}}","role":"assistant"}}]}\n\n',
    'data: {"choices":[{"delta":{"content":" about this</thought>The answer is 391.","role":"assistant"}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.ok, true);
  assert.strictEqual(out.reasoning, 'Let me think about this');
  assert.strictEqual(out.text, 'The answer is 391.');
});

test('parseStream: handles <thought> tag split across chunks', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"content":"<thought>step 1","extra_content":"{\\"google\\":{\\"thought\\":true}}","role":"assistant"}}]}\n\n',
    'data: {"choices":[{"delta":{"content":" step 2","extra_content":"{\\"google\\":{\\"thought\\":true}}","role":"assistant"}}]}\n\n',
    'data: {"choices":[{"delta":{"content":"</thought>Result","role":"assistant"}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.ok, true);
  assert.strictEqual(out.reasoning, 'step 1 step 2');
  assert.strictEqual(out.text, 'Result');
});

test('parseStream: handles thought without extra_content marker', async () => {
  // Some chunks may not include extra_content but still have <thought> tags
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"content":"<thought>thinking","role":"assistant"}}]}\n\n',
    'data: {"choices":[{"delta":{"content":" more</thought>answer","role":"assistant"}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.ok, true);
  assert.strictEqual(out.reasoning, 'thinking more');
  assert.strictEqual(out.text, 'answer');
});

test('parseStream: handles text before <thought> tag', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"content":"Before thought <thought>the thinking</thought> after thought","role":"assistant"}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.ok, true);
  assert.strictEqual(out.text, 'Before thought  after thought');
  assert.strictEqual(out.reasoning, 'the thinking');
});

test('parseStream: onThinking fires for Gemini thought blocks', async () => {
  let thinkingCalls = [];
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"content":"<thought>part 1","extra_content":"{\\"google\\":{\\"thought\\":true}}","role":"assistant"}}]}\n\n',
    'data: {"choices":[{"delta":{"content":" part 2</thought>done","role":"assistant"}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  await LLM.parseStream(resp, null, (full) => thinkingCalls.push(full), null);
  assert.ok(thinkingCalls.length >= 2);
  assert.strictEqual(thinkingCalls[0], 'part 1');
  assert.strictEqual(thinkingCalls[thinkingCalls.length - 1], 'part 1 part 2');
});

test('parseStream: thought_signature chunk does not corrupt output', async () => {
  const resp = fakeResponse([
    'data: {"choices":[{"delta":{"content":"<thought>thinking</thought>answer","role":"assistant"}}]}\n\n',
    'data: {"choices":[{"delta":{"extra_content":"{\\"google\\":{\\"thought_signature\\":\\"abc123\\"}}","role":"assistant"}}]}\n\n',
    'data: [DONE]\n\n'
  ]);
  const out = await LLM.parseStream(resp);
  assert.strictEqual(out.ok, true);
  assert.strictEqual(out.reasoning, 'thinking');
  assert.strictEqual(out.text, 'answer');
});
