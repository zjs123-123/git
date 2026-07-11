// DeepSeek API 调用服务
// API密钥由服务端管理，绝不暴露到前端

const https = require('https');
const http = require('http');

const DEEPSEEK_BASE_URL = process.env.DEEPSEEK_BASE_URL || 'https://api.deepseek.com/v1';
const DEEPSEEK_API_KEY = process.env.DEEPSEEK_API_KEY || '';
const DEEPSEEK_MODEL = process.env.DEEPSEEK_MODEL || 'deepseek-chat';

function chatCompletion(messages, options) {
  return new Promise((resolve, reject) => {
    const url = new URL('/chat/completions', DEEPSEEK_BASE_URL);
    const body = JSON.stringify({
      model: DEEPSEEK_MODEL,
      messages: messages,
      max_tokens: (options && options.maxTokens) || 300,
      temperature: (options && options.temperature) || 0.7,
      stream: false
    });

    const req = https.request({
      hostname: url.hostname,
      port: 443,
      path: url.pathname,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + DEEPSEEK_API_KEY,
        'Content-Length': Buffer.byteLength(body)
      },
      timeout: 30000
    }, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.error) {
            reject(new Error(json.error.message || 'DeepSeek API error'));
            return;
          }
          const content = json.choices && json.choices[0] && json.choices[0].message
            ? json.choices[0].message.content
            : '';
          resolve(content.trim());
        } catch (e) {
          reject(new Error('Failed to parse DeepSeek response: ' + data.slice(0, 200)));
        }
      });
    });

    req.on('error', (e) => { reject(e); });
    req.on('timeout', () => { req.destroy(); reject(new Error('DeepSeek API timeout')); });
    req.write(body);
    req.end();
  });
}

module.exports = { chatCompletion };
