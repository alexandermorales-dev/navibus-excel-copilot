/* ============================================
   repair.js — Failed-action repair round-trip
   Sends failed actions + errors back to Gemini for one repair attempt.
   ============================================ */

const Repair = {
  /**
   * Ask Gemini to repair failed actions.
   * @param {string} systemPrompt - The original system prompt
   * @param {Array} conversation - Full conversation history
   * @param {Array} failed - Failed action objects with errors
   * @param {Object} schemaSnapshot - Current workbook snapshot
   * @returns { ok: true, plan } | { ok: false, error }
   */
  async repairActions(systemPrompt, conversation, failed, schemaSnapshot, originalPlan) {
    const repairMessage = this.buildRepairMessage(failed, schemaSnapshot, originalPlan);

    const contents = [...conversation, {
      role: 'user',
      parts: [{ text: repairMessage }]
    }];

    const result = await Gemini.generateWithFallback(systemPrompt, contents);

    if (!result.ok) {
      return { ok: false, error: result.error };
    }

    // Parse the repaired plan
    const parsed = this.parsePlan(result.text);
    if (!parsed.ok) {
      return { ok: false, error: `Repair returned invalid JSON: ${parsed.error}` };
    }

    // Validate the repaired plan
    const validation = Actions.validatePlan(parsed.plan);
    if (!validation.valid) {
      return { ok: false, error: `Repair plan validation failed: ${validation.errors.join('; ')}` };
    }

    return { ok: true, plan: parsed.plan, explanation: parsed.explanation };
  },

  buildRepairMessage(failed, schemaSnapshot, originalPlan) {
    const lines = [
      'The previous action plan failed during execution. Here is the FULL original plan and the errors:',
      '',
      'ORIGINAL PLAN (that failed):',
      JSON.stringify(originalPlan, null, 2),
      '',
      'ERRORS:',
      ''
    ];

    for (const f of failed) {
      lines.push(`- Action ${f.index + 1} (${f.op}): ${f.error}`);
      if (f.action) {
        lines.push(`  Failed action details: ${JSON.stringify(f.action)}`);
      }
    }

    lines.push('');
    lines.push('Current workbook state (after rollback — should be clean):');
    lines.push(Schema.toText(schemaSnapshot));
    lines.push('');
    lines.push('IMPORTANT INSTRUCTIONS FOR REPAIR:');
    lines.push('1. Return a COMPLETE, CORRECTED action plan (not just the fixed action).');
    lines.push('2. Start from scratch — the workbook was rolled back to its original state.');
    lines.push('3. Fix the specific errors above.');
    lines.push('4. Make sure writeRange values array dimensions match the range size.');
    lines.push('5. Make sure all sheet names and column references are correct.');
    lines.push('6. Return ONLY the JSON array. No markdown, no explanation.');
    lines.push('');

    return lines.join('\n');
  },

  parsePlan(text) {
    try {
      let cleaned = text.trim();

      // Strip ```json ... ``` fences
      const fenceMatch = cleaned.match(/```(?:json)?\s*([\s\S]*?)```/);
      if (fenceMatch) {
        cleaned = fenceMatch[1].trim();
      }

      // Only treat as an action plan if the text STARTS with '[' (a JSON array).
      // If the model wrote a conversational response with brackets somewhere in
      // the middle, that's not an action plan — it's a text response.
      if (!cleaned.startsWith('[')) {
        return { ok: false, error: 'Response does not start with JSON array — treating as text.' };
      }

      const plan = JSON.parse(cleaned);
      if (!Array.isArray(plan)) {
        return { ok: false, error: 'Parsed JSON is not an array.' };
      }

      // Verify it looks like an action plan: every element must have an "op" field
      if (plan.length > 0 && plan.every(item => item && typeof item === 'object' && typeof item.op === 'string')) {
        return { ok: true, plan: plan, explanation: '' };
      }

      return { ok: false, error: 'JSON array does not contain valid action objects.' };
    } catch (e) {
      return { ok: false, error: e.message };
    }
  }
};
