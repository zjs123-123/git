// ============================================
// 云岭村乡村共创平台 · AI智能体后端服务器
// 设计依据：Vibe Coding后端全流程指南 + 数据库设计规范指南
// ============================================

require('dotenv').config();

const express = require('express');
const cors = require('cors');
const crypto = require('crypto');
const rateLimit = require('express-rate-limit');
const path = require('path');

const db = require('./db');
const { chatCompletion } = require('./services/deepseek');
const { getFallbackResponse } = require('./services/fallback');

// ============================================
// 1. 应用初始化
// ============================================
const app = express();
const PORT = process.env.PORT || 3456;

// 中间件
app.use(cors());
app.use(express.json({ limit: '10kb' }));

// 静态文件：serve agent HTML pages
app.use(express.static(path.join(__dirname, '..')));

// ============================================
// 2. 安全防线：速率限制
// 设计依据：后端安全指南 - 第6道防线（防御冗余）
// ============================================
const chatLimiter = rateLimit({
  windowMs: 60 * 1000,                                // 1分钟窗口
  max: parseInt(process.env.RATE_LIMIT_MAX) || 20,    // 每窗口最多20次
  message: { code: 429, message: '请求过于频繁，请稍后再试', data: null },
  standardHeaders: true,
  legacyHeaders: false,
  keyGenerator: (req) => hashIp(req.ip)               // 用IP哈希做key，不存原始IP
});

// ============================================
// 3. 工具函数
// ============================================
function hashIp(ip) {
  return crypto.createHash('sha256').update(ip || 'unknown').digest('hex').slice(0, 16);
}

function ok(data) {
  return { code: 0, message: 'ok', data };
}

function err(code, message) {
  return { code, message, data: null };
}

// ============================================
// 4. 输入校验中间件
// 设计依据：后端安全指南 - 第2道防线（接口输入）
// ============================================
function validateChatInput(req, res, next) {
  const { agent_id, message } = req.body;

  if (!agent_id || !['huangwenxiu', 'zhukezhen'].includes(agent_id)) {
    return res.status(400).json(err(400, '无效的智能体ID'));
  }
  if (!message || typeof message !== 'string' || message.trim().length === 0) {
    return res.status(400).json(err(400, '消息不能为空'));
  }
  if (message.length > 500) {
    return res.status(400).json(err(400, '消息长度不能超过500字'));
  }

  // 净化输入：去除HTML标签
  req.body.message = message.replace(/<[^>]*>/g, '').trim();
  next();
}

// ============================================
// 5. API路由
// ============================================

// 5a. 健康检查
// 设计依据：后端骨架指南 - 跑通启动线
app.get('/api/health', (req, res) => {
  res.json(ok({ status: 'running', timestamp: new Date().toISOString() }));
});

// 5b. 聊天接口（核心）
// POST /api/chat
// Body: { agent_id, conversation_id?, message }
// Response: { code, data: { conversation_id, reply, fallback_used } }
app.post('/api/chat', chatLimiter, validateChatInput, async (req, res) => {
  const { agent_id, message } = req.body;
  let { conversation_id } = req.body;
  const ipHash = hashIp(req.ip);

  try {
    // 获取智能体配置
    const agent = db.getAgent(agent_id);
    if (!agent) {
      return res.status(400).json(err(400, '智能体不存在'));
    }

    // 创建或恢复会话
    if (!conversation_id) {
      conversation_id = db.createConversation(agent_id, ipHash);
    } else {
      const conv = db.getConversation(conversation_id);
      if (!conv) {
        conversation_id = db.createConversation(agent_id, ipHash);
      }
    }

    // 保存用户消息
    db.saveMessage(conversation_id, 'user', message, false);

    // 自动设置会话标题（取首条用户消息前30字）
    db.updateConversationTitle(conversation_id, message.slice(0, 30));

    // 构建对话上下文
    const history = db.getMessages(conversation_id, 20);
    const messages = [
      { role: 'system', content: agent.system_prompt },
      ...history.map(m => ({ role: m.role, content: m.content }))
    ];

    // 调用DeepSeek API，失败则兜底
    let reply;
    let fallbackUsed = false;

    try {
      reply = await chatCompletion(messages, { maxTokens: agent.max_tokens });
    } catch (apiErr) {
      console.error('DeepSeek API failed, using fallback:', apiErr.message);
      reply = getFallbackResponse(agent_id, message);
      fallbackUsed = true;
    }

    // 保存AI回复
    db.saveMessage(conversation_id, 'assistant', reply, fallbackUsed);

    // 返回结果
    res.json(ok({
      conversation_id,
      reply,
      fallback_used: fallbackUsed
    }));

  } catch (e) {
    console.error('Chat error:', e);
    res.status(500).json(err(500, '服务器内部错误，请稍后再试'));
  }
});

// 5c. 获取会话列表
// GET /api/conversations/:agentId
app.get('/api/conversations/:agentId', (req, res) => {
  const { agentId } = req.params;
  const ipHash = hashIp(req.ip);

  try {
    const conversations = db.getRecentConversations(agentId, ipHash, 10);
    res.json(ok(conversations));
  } catch (e) {
    res.status(500).json(err(500, '获取会话列表失败'));
  }
});

// 5d. 获取会话消息
// GET /api/conversations/:id/messages
app.get('/api/conversations/:id/messages', (req, res) => {
  try {
    const messages = db.getMessages(req.params.id, 100);
    res.json(ok(messages));
  } catch (e) {
    res.status(500).json(err(500, '获取消息失败'));
  }
});

// ============================================
// 5e. 任务提交平台 API
// 设计依据：数据库设计规范指南 - 全流程
// ============================================

const taskLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 30,
  message: { code: 429, message: '请求过于频繁，请稍后再试', data: null },
  standardHeaders: true,
  legacyHeaders: false,
  keyGenerator: (req) => hashIp(req.ip)
});

// 任务提交校验
function validateTaskInput(req, res, next) {
  const { task_type, group_name, members, answers } = req.body;
  if (!task_type || !['task1', 'task2', 'task3'].includes(task_type)) {
    return res.status(400).json(err(400, '无效的任务类型'));
  }
  if (!group_name || typeof group_name !== 'string' || group_name.trim().length === 0) {
    return res.status(400).json(err(400, '小组名称不能为空'));
  }
  if (!members || typeof members !== 'string' || members.trim().length === 0) {
    return res.status(400).json(err(400, '组员姓名不能为空'));
  }
  if (!Array.isArray(answers) || answers.length === 0) {
    return res.status(400).json(err(400, '答案不能为空'));
  }
  for (let i = 0; i < answers.length; i++) {
    if (typeof answers[i] !== 'string' || answers[i].trim().length === 0) {
      return res.status(400).json(err(400, '第' + (i+1) + '题答案不能为空'));
    }
    if (answers[i].length > 3000) {
      return res.status(400).json(err(400, '第' + (i+1) + '题答案超过3000字限制'));
    }
    answers[i] = answers[i].replace(/<[^>]*>/g, '').trim();
  }
  next();
}

// AI评分系统提示词
const TASK1_QUESTIONS = [
  '当前群众问题分析', '群众不信任的原因',
  '黄文秀哪些做法体现了群众观点', '黄文秀哪些做法体现了群众路线',
  '作为实践团成员，你会怎样真正发动群众参与建设', '人民群众为什么能够推动乡村建设'
];
const TASK2_QUESTIONS = [
  '云岭村面临哪些自然条件限制', '地形如何影响交通与产业',
  '为什么山区乡村发展风险更高', '自然环境与乡村发展之间的关系'
];
const TASK3_QUESTIONS = [
  '群众需求分析', '地理条件分析', '产业发展路径',
  '群众参与方式', '可持续发展考虑'
];

function buildAnalyzePrompt(taskType, question, answer) {
  if (taskType === 'task1') {
    return `你是乡村治理课程AI助教。请对以下学生回答评分（满分100）并给出30字以内评语。

题目：${question}
学生回答：${answer}

评分标准：结合云岭村实际数据（60%滞销率、0.6元/斤等）、引用黄文秀具体做法、分析有深度（区分表层与深层原因）。

请严格按此JSON格式回复：{"score": 数字, "feedback": "评语"}`;
  } else if (taskType === 'task2') {
    return `你是乡村治理课程AI助教。请对以下学生回答评分（满分100）并给出30字以内评语。

题目：${question}
学生回答：${answer}

评分标准：引用云岭村地理数据（海拔620-940m、坡度23°、喀斯特等）、建立因果传导链、体现竺可桢因地制宜理念。

请严格按此JSON格式回复：{"score": 数字, "feedback": "评语"}`;
  } else {
    return `你是乡村治理课程AI助教。请对以下学生回答评分（满分100）并给出30字以内评语。

题目：${question}
学生回答：${answer}

评分标准：融合黄文秀群众工作法与竺可桢地理分析、方案具体可行有步骤、兼顾生态经济社会可持续。

请严格按此JSON格式回复：{"score": 数字, "feedback": "评语"}`;
  }
}

// 前端关键词评分（DeepSeek不可用时的兜底）
function fallbackAnalyze(taskType, qIndex, text) {
  const len = text.length;
  let score = Math.min(85, 40 + Math.floor(len / 15));
  let feedback = '';

  if (taskType === 'task1') {
    const keywords = [['信任','顾虑','风险'], ['历史','过去','产业','赔'], ['走访','入户','人民'], ['从群众中来','调研','倾听'], ['一起','共同','参与'], ['历史唯物','创造','云岭','百坭']];
    for (const kw of keywords[qIndex] || []) { if (text.indexOf(kw) !== -1) score += 4; }
    feedback = len > 60 ? '分析有一定深度，可进一步结合云岭村具体数据。' : '建议展开论述，引用黄文秀具体做法。';
  } else if (taskType === 'task2') {
    const keywords = [['地形','气候','土地','水文'], ['道路','运输','产业','成本'], ['叠加','滑坡','风险','脆弱'], ['因地制宜','顺应','约束','和谐']];
    for (const kw of keywords[qIndex] || []) { if (text.indexOf(kw) !== -1) score += 4; }
    feedback = len > 60 ? '地理分析有据，建议加强因果传导链的论述。' : '建议引用云岭村具体地理数据展开分析。';
  } else {
    const keywords = [['走访','信任','需求','群众'], ['地形','气候','喀斯特','机遇'], ['产业','品牌','步骤','风险'], ['合作社','入股','分红','决策'], ['生态','市场','人才','制度']];
    for (const kw of keywords[qIndex] || []) { if (text.indexOf(kw) !== -1) score += 3; }
    feedback = len > 80 ? '方案设计具体，建议补充量化目标和时间节点。' : '建议从多个维度展开，提出可操作的具体措施。';
  }
  return { score: Math.min(95, Math.max(30, score)), feedback };
}

// POST /api/tasks/submit — 提交方案并获取AI分析
app.post('/api/tasks/submit', taskLimiter, validateTaskInput, async (req, res) => {
  const { task_type, group_name, members, answers, self_score } = req.body;
  const ipHash = hashIp(req.ip);

  try {
    const questions = task_type === 'task1' ? TASK1_QUESTIONS : (task_type === 'task2' ? TASK2_QUESTIONS : TASK3_QUESTIONS);
    const analysisItems = [];
    let totalScore = 0;

    for (let i = 0; i < answers.length; i++) {
      let item = { qIndex: i, score: 0, feedback: '' };

      try {
        const prompt = buildAnalyzePrompt(task_type, questions[i], answers[i]);
        const reply = await chatCompletion([
          { role: 'system', content: '你是严格的课程AI助教，只输出JSON格式评分结果。' },
          { role: 'user', content: prompt }
        ], { maxTokens: 200, temperature: 0.3 });

        const json = JSON.parse(reply.replace(/```json|```/g, '').trim());
        item.score = Math.min(95, Math.max(25, parseInt(json.score) || 60));
        item.feedback = json.feedback || '';
      } catch (apiErr) {
        console.error('AI analysis failed, using fallback:', apiErr.message);
        const fallback = fallbackAnalyze(task_type, i, answers[i]);
        item.score = fallback.score;
        item.feedback = fallback.feedback + ' [本地评分]';
      }

      analysisItems.push(item);
      totalScore += item.score;
    }

    const avgScore = Math.round(totalScore / answers.length);
    const scores = analysisItems.map(a => a.score);

    const id = db.saveTaskSubmission(task_type, group_name.trim(), members.trim(), answers, scores, avgScore, analysisItems, self_score || null, ipHash);

    res.json(ok({
      id,
      total_score: avgScore,
      scores,
      analysis: analysisItems
    }));

  } catch (e) {
    console.error('Task submit error:', e);
    res.status(500).json(err(500, '提交失败，请稍后再试'));
  }
});

// GET /api/tasks/submissions/:taskType — 获取提交列表
app.get('/api/tasks/submissions/:taskType', (req, res) => {
  const { taskType } = req.params;
  if (!['task1', 'task2', 'task3'].includes(taskType)) {
    return res.status(400).json(err(400, '无效的任务类型'));
  }
  try {
    const subs = db.getAllTaskSubmissions(taskType);
    const parsed = subs.map(s => ({
      ...s,
      answers: JSON.parse(s.answers || '[]'),
      scores: JSON.parse(s.scores || '[]'),
      ai_analysis: JSON.parse(s.ai_analysis || '[]')
    }));
    res.json(ok(parsed));
  } catch (e) {
    res.status(500).json(err(500, '获取提交列表失败'));
  }
});

// GET /api/tasks/my-submissions/:taskType — 获取我的提交
app.get('/api/tasks/my-submissions/:taskType', (req, res) => {
  const { taskType } = req.params;
  if (!['task1', 'task2', 'task3'].includes(taskType)) {
    return res.status(400).json(err(400, '无效的任务类型'));
  }
  try {
    const subs = db.getTaskSubmissions(taskType, hashIp(req.ip));
    const parsed = subs.map(s => ({
      ...s,
      answers: JSON.parse(s.answers || '[]'),
      scores: JSON.parse(s.scores || '[]'),
      ai_analysis: JSON.parse(s.ai_analysis || '[]')
    }));
    res.json(ok(parsed));
  } catch (e) {
    res.status(500).json(err(500, '获取提交列表失败'));
  }
});

// DELETE /api/tasks/submissions — 清空我的提交
app.delete('/api/tasks/submissions', (req, res) => {
  try {
    db.deleteAllTaskSubmissions(hashIp(req.ip));
    res.json(ok(null));
  } catch (e) {
    res.status(500).json(err(500, '清空失败'));
  }
});

// ============================================
// 6. 统一错误处理（404）
// ============================================
app.use((req, res) => {
  res.status(404).json(err(404, '接口不存在'));
});

// ============================================
// 7. 启动服务器
// ============================================
app.listen(PORT, () => {
  console.log('========================================');
  console.log('  云岭村 AI智能体后端 v2.0');
  console.log('========================================');
  console.log(`  服务地址: http://localhost:${PORT}`);
  console.log(`  健康检查: http://localhost:${PORT}/api/health`);
  console.log('');
  console.log('  已注册智能体:');
  console.log(`    黄文秀AI助手:   http://localhost:${PORT}/agent-huangwenxiu.html`);
  console.log(`    竺可桢地理智库:  http://localhost:${PORT}/agent-zhukezhen.html`);
  console.log('');
  console.log('  架构: Express + SQLite + DeepSeek API');
  console.log('  安全: 速率限制 | 输入校验 | XSS防护 | 密钥服务端管理');
  console.log('========================================');
});
