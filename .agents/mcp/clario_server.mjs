import readline from 'readline';

const API_KEY = process.env.CLARIO_API_KEY || 'sk-clario-8a0483a00917b34dab1b663416cf14242849cea57bd92482';
const BASE_URL = process.env.CLARIO_BASE_URL || 'https://clario.apicloud.my.id/v1';

const TOOLS = [
  {
    name: 'clario_chat',
    description: 'Ask Clario Cloud AI models anything: general chat, business advice, summaries, brainstorming, and answering questions.',
    inputSchema: {
      type: 'object',
      properties: {
        prompt: {
          type: 'string',
          description: 'User prompt or question'
        },
        model: {
          type: 'string',
          description: 'Model ID (e.g. clario/glm-5.3-flash, clario/deepseek-v4-flash, clario/gpt-5.6-sol, clario/claude-opus-5, clario/minimax-m3)',
          default: 'clario/glm-5.3-flash'
        },
        system_prompt: {
          type: 'string',
          description: 'Optional system instructions'
        }
      },
      required: ['prompt']
    }
  },
  {
    name: 'clario_code_expert',
    description: 'Solve programming, debugging, architecture, refactoring, and code review tasks using Clario AI models.',
    inputSchema: {
      type: 'object',
      properties: {
        task: {
          type: 'string',
          description: 'Coding task or question'
        },
        code_context: {
          type: 'string',
          description: 'Code snippet, file content, or error trace'
        },
        language: {
          type: 'string',
          description: 'Language (typescript, python, go, etc.)'
        },
        model: {
          type: 'string',
          description: 'Model ID (defaults to clario/glm-5.3-flash or clario/gpt-5.6-sol)',
          default: 'clario/glm-5.3-flash'
        }
      },
      required: ['task']
    }
  }
];

async function handleClarioChat({ prompt, model = 'clario/glm-5.3-flash', system_prompt }) {
  const messages = [];
  if (system_prompt) {
    messages.push({ role: 'system', content: system_prompt });
  }
  messages.push({ role: 'user', content: prompt });

  const response = await fetch(`${BASE_URL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: model || 'clario/glm-5.3-flash',
      messages: messages,
      max_tokens: 4096
    })
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Clario API error (${response.status}): ${errText}`);
  }

  const data = await response.json();
  return data.choices?.[0]?.message?.content || 'No response generated.';
}

async function handleClarioCode({ task, code_context, language, model = 'clario/glm-5.3-flash' }) {
  let prompt = `Task: ${task}\n`;
  if (language) prompt += `Language: ${language}\n`;
  if (code_context) prompt += `\nContext:\n\`\`\`${language || ''}\n${code_context}\n\`\`\`\n`;

  const messages = [
    { role: 'system', content: 'You are an elite coding expert, software architect, and debugging specialist.' },
    { role: 'user', content: prompt }
  ];

  const response = await fetch(`${BASE_URL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: model || 'clario/glm-5.3-flash',
      messages: messages,
      max_tokens: 4096
    })
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Clario API error (${response.status}): ${errText}`);
  }

  const data = await response.json();
  return data.choices?.[0]?.message?.content || 'No response generated.';
}

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
  terminal: false
});

function sendResponse(id, result, error = null) {
  const payload = { jsonrpc: '2.0', id };
  if (error) {
    payload.error = error;
  } else {
    payload.result = result;
  }
  process.stdout.write(JSON.stringify(payload) + '\n');
}

rl.on('line', async (line) => {
  line = line.trim();
  if (!line) return;

  let message;
  try {
    message = JSON.parse(line);
  } catch (e) {
    return;
  }

  const { id, method, params } = message;

  try {
    switch (method) {
      case 'initialize':
        sendResponse(id, {
          protocolVersion: '2024-11-05',
          capabilities: {
            tools: {}
          },
          serverInfo: {
            name: 'clario-ai',
            version: '1.0.0'
          }
        });
        break;

      case 'notifications/initialized':
        break;

      case 'ping':
        sendResponse(id, {});
        break;

      case 'tools/list':
        sendResponse(id, { tools: TOOLS });
        break;

      case 'tools/call': {
        const { name, arguments: args } = params || {};
        let textResult = '';

        if (name === 'clario_chat') {
          textResult = await handleClarioChat(args || {});
        } else if (name === 'clario_code_expert') {
          textResult = await handleClarioCode(args || {});
        } else {
          sendResponse(id, null, { code: -32601, message: `Tool not found: ${name}` });
          return;
        }

        sendResponse(id, {
          content: [
            {
              type: 'text',
              text: textResult
            }
          ]
        });
        break;
      }

      default:
        if (id !== undefined) {
          sendResponse(id, null, { code: -32601, message: `Method not found: ${method}` });
        }
        break;
    }
  } catch (err) {
    if (id !== undefined) {
      sendResponse(id, {
        isError: true,
        content: [
          {
            type: 'text',
            text: `Error executing ${method}: ${err.message}`
          }
        ]
      });
    }
  }
});
