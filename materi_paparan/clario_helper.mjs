#!/usr/bin/env node
/**
 * Clario Cloud API Helper
 * Base URL: https://clario.apicloud.my.id/v1
 * Default Model: clario/glm-5.3-flash
 */

const API_KEY = process.env.CLARIO_API_KEY || 'sk-clario-8a0483a00917b34dab1b663416cf14242849cea57bd92482';
const BASE_URL = 'https://clario.apicloud.my.id/v1';

const args = process.argv.slice(2);
if (args.length === 0) {
  console.log(`
Penggunaan:
  node clario_helper.mjs "<prompt>" [--model <model_id>]

Contoh:
  node clario_helper.mjs "Buatkan ide visual untuk slide 6 GAP STP"
  node clario_helper.mjs "Beri rekomendasi layout" --model clario/gpt-5.6-sol

Model Populer Tersedia:
  - clario/glm-5.3-flash (Default, Super Cepat & Cerdas)
  - clario/deepseek-v4-flash (Cepat, Efisien)
  - clario/gpt-5.6-sol (Coding & Refactoring Kuat)
  - clario/claude-opus-5 (Analisis Kompleks & Arsitektur)
`);
  process.exit(0);
}

let model = 'clario/glm-5.3-flash';
let prompt = '';

for (let i = 0; i < args.length; i++) {
  if (args[i] === '--model' && args[i + 1]) {
    model = args[i + 1];
    i++;
  } else {
    prompt += (prompt ? ' ' : '') + args[i];
  }
}

async function askClario() {
  console.log(`\n⏳ Mengirim permintaan ke Clario (${model})...\n`);
  try {
    const response = await fetch(`${BASE_URL}/chat/completions`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: model,
        messages: [
          { role: 'system', content: 'Kamu adalah asisten profesional ahli strategi bisnis, teknologi, dan komunikasi kepresidenan/pemerintahan daerah.' },
          { role: 'user', content: prompt }
        ],
        max_tokens: 4096
      })
    });

    if (!response.ok) {
      const errText = await response.text();
      console.error(`❌ HTTP Error ${response.status}:`, errText);
      return;
    }

    const data = await response.json();
    const reply = data.choices?.[0]?.message?.content;
    console.log('--- JAWABAN CLARIO ---');
    console.log(reply);
    console.log('----------------------');
  } catch (err) {
    console.error('❌ Gagal menghubungi Clario:', err.message);
  }
}

askClario();
