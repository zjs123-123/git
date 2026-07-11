// 数据库连接与初始化模块
// 使用 Node.js 内置 SQLite (v22+)，无需编译
// 设计依据：AI辅助开发场景下的数据库设计规范指南 - 第六步

const path = require('path');
const fs = require('fs');
const { DatabaseSync } = require('node:sqlite');

const dbPath = path.join(__dirname, 'data.db');
const schemaPath = path.join(__dirname, 'schema.sql');

// 打开数据库（自动创建）
const db = new DatabaseSync(dbPath);
db.exec('PRAGMA journal_mode = WAL');
db.exec('PRAGMA foreign_keys = ON');

// 初始化Schema
function initSchema() {
  const schema = fs.readFileSync(schemaPath, 'utf-8');
  db.exec(schema);
}
initSchema();

// === 智能体查询 ===
function getAgent(agentId) {
  const stmt = db.prepare('SELECT * FROM agents WHERE id = ?');
  return stmt.get(agentId);
}

// === 会话操作 ===
function createConversation(agentId, ipHash) {
  const { v4: uuidv4 } = require('uuid');
  const id = uuidv4();
  db.prepare(
    'INSERT INTO conversations (id, agent_id, ip_hash) VALUES (?, ?, ?)'
  ).run(id, agentId, ipHash);
  return id;
}

function getConversation(id) {
  return db.prepare('SELECT * FROM conversations WHERE id = ?').get(id);
}

function updateConversationTitle(id, title) {
  db.prepare(
    "UPDATE conversations SET title = ?, updated_at = datetime('now','localtime') WHERE id = ? AND title IS NULL"
  ).run(title, id);
}

// === 消息操作 ===
function saveMessage(conversationId, role, content, fallbackUsed) {
  const result = db.prepare(
    'INSERT INTO messages (conversation_id, role, content, fallback_used) VALUES (?, ?, ?, ?)'
  ).run(conversationId, role, content, fallbackUsed ? 1 : 0);

  db.prepare(
    "UPDATE conversations SET updated_at = datetime('now','localtime') WHERE id = ?"
  ).run(conversationId);

  return result.lastInsertRowid;
}

function getMessages(conversationId, limit) {
  return db.prepare(
    'SELECT role, content, created_at FROM messages WHERE conversation_id = ? ORDER BY created_at ASC LIMIT ?'
  ).all(conversationId, limit || 50);
}

function getRecentConversations(agentId, ipHash, limit) {
  return db.prepare(
    'SELECT id, title, created_at, updated_at FROM conversations WHERE agent_id = ? AND ip_hash = ? ORDER BY updated_at DESC LIMIT ?'
  ).all(agentId, ipHash, limit || 10);
}

// === 任务提交操作 ===
function saveTaskSubmission(taskType, groupName, members, answers, scores, totalScore, aiAnalysis, selfScore, ipHash) {
  const result = db.prepare(
    'INSERT INTO task_submissions (task_type, group_name, members, answers, scores, total_score, ai_analysis, self_score, ip_hash) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'
  ).run(taskType, groupName, members, JSON.stringify(answers), JSON.stringify(scores), totalScore, JSON.stringify(aiAnalysis), selfScore || null, ipHash);
  return result.lastInsertRowid;
}

function getTaskSubmissions(taskType, ipHash) {
  return db.prepare(
    'SELECT id, task_type, group_name, members, answers, scores, total_score, ai_analysis, self_score, created_at FROM task_submissions WHERE task_type = ? AND ip_hash = ? ORDER BY created_at DESC LIMIT 50'
  ).all(taskType, ipHash);
}

function getAllTaskSubmissions(taskType) {
  return db.prepare(
    'SELECT id, task_type, group_name, members, answers, scores, total_score, ai_analysis, self_score, created_at FROM task_submissions WHERE task_type = ? ORDER BY created_at DESC LIMIT 100'
  ).all(taskType);
}

function getTaskSubmission(id) {
  return db.prepare('SELECT * FROM task_submissions WHERE id = ?').get(id);
}

function deleteAllTaskSubmissions(ipHash) {
  return db.prepare('DELETE FROM task_submissions WHERE ip_hash = ?').run(ipHash);
}

module.exports = {
  getAgent,
  createConversation, getConversation, updateConversationTitle,
  saveMessage, getMessages, getRecentConversations,
  saveTaskSubmission, getTaskSubmissions, getAllTaskSubmissions, getTaskSubmission, deleteAllTaskSubmissions
};
