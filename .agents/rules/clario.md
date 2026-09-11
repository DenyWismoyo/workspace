# Clario Cloud AI Integration Guidelines

## Konfigurasi Kredensial & Endpoint
- **Base URL**: `https://clario.apicloud.my.id/v1`
- **Default Model**: `clario/glm-5.3-flash`
- **Protocol**: OpenAI-compatible (`/v1/chat/completions`) & Anthropic-compatible (`/v1/messages`).
- **Default Auth**: `Authorization: Bearer sk-clario-8a0483a00917b34dab1b663416cf14242849cea57bd92482`

## Rekomendasi Penggunaan Model
1. **Brainstorming & Drafting Slide**: Gunakan `clario/glm-5.3-flash` (cepat, hemat biaya, dan penalaran tajam).
2. **Coding Harian & Regex/Query**: Gunakan `clario/deepseek-v4-flash` (respons ~1.2 detik).
3. **Arsitektur & Refactoring Kompleks**: Gunakan `clario/gpt-5.6-sol`.
4. **Naskah Strategis & Dokumen Eksekutif**: Gunakan `clario/claude-opus-5`.

## Pemanggilan Mandiri
- Gunakan skrip pembantu yang tersedia di `materi_paparan/clario_helper.mjs` untuk menjalankan instruksi atau konsultasi ke Clario.
