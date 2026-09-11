---
name: clario-ai
description: Guide for utilizing Clario Cloud AI models (GLM-5.3-flash, DeepSeek V4, GPT-5.6 Sol, Claude Opus 5) for brainstorming, executive drafting, coding, and slide enrichment via OpenAI-compatible endpoints.
---

# Clario Cloud AI Workflow & Reference

Use this skill when needing additional AI capabilities from Clario Cloud for drafting, coding, brainstorming, or enriching presentations.

## 1. Quick Config
- **Base URL**: `https://clario.apicloud.my.id/v1`
- **Active Key**: `sk-clario-8a0483a00917b34dab1b663416cf14242849cea57bd92482`
- **Default Model**: `clario/glm-5.3-flash`

## 2. Pilihan Model Populer
- `clario/glm-5.3-flash`: Cepat, cerdas, reasoning bawaan, ideal untuk brainstorming dan slide drafting.
- `clario/deepseek-v4-flash`: Super cepat, latensi rendah, ideal untuk tugas rutin dan parsing kode.
- `clario/gpt-5.6-sol`: Kuat untuk arsitektur, debugging mendalam, dan pemrograman tingkat lanjut.
- `clario/claude-opus-5`: Analisis mendalam dokumen kebijakan, hukum daerah, dan narasi pimpinan eksekutif.

## 3. Eksekusi Cepat via Node.js
Gunakan skrip pembantu di workspace:
```bash
node materi_paparan/clario_helper.mjs "<prompt>" [--model <model_id>]
```

## 4. Format cURL Langsung
```bash
curl.exe -s -X POST https://clario.apicloud.my.id/v1/chat/completions \
  -H "Authorization: Bearer sk-clario-8a0483a00917b34dab1b663416cf14242849cea57bd92482" \
  -H "Content-Type: application/json" \
  -d '{"model":"clario/glm-5.3-flash","messages":[{"role":"user","content":"halo"}]}'
```
