-- ============================================
-- 云岭村乡村共创平台 · AI智能体数据库设计
-- 设计依据：AI辅助开发场景下的数据库设计规范指南
-- ============================================

-- 1. 智能体表：存储每个AI角色的系统设定
CREATE TABLE IF NOT EXISTS agents (
    id          TEXT PRIMARY KEY,                  -- 智能体唯一标识，如 'huangwenxiu' / 'zhukezhen'
    name        TEXT NOT NULL,                     -- 智能体名称
    title       TEXT NOT NULL,                     -- 智能体头衔
    system_prompt TEXT NOT NULL,                   -- 系统提示词，定义AI角色行为
    max_tokens  INTEGER NOT NULL DEFAULT 150,      -- 回复最大字数
    created_at  DATETIME NOT NULL DEFAULT (datetime('now', 'localtime')),
    updated_at  DATETIME NOT NULL DEFAULT (datetime('now', 'localtime'))
);

-- 2. 会话表：每次打开页面创建的对话会话
CREATE TABLE IF NOT EXISTS conversations (
    id          TEXT PRIMARY KEY,                  -- UUID，由后端生成
    agent_id    TEXT NOT NULL,                     -- 关联智能体
    title       TEXT,                              -- 会话标题，取首条用户消息前30字
    ip_hash     TEXT NOT NULL,                     -- 用户IP哈希，用于匿名区分（不存原始IP）
    created_at  DATETIME NOT NULL DEFAULT (datetime('now', 'localtime')),
    updated_at  DATETIME NOT NULL DEFAULT (datetime('now', 'localtime')),
    FOREIGN KEY (agent_id) REFERENCES agents(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_conversations_agent ON conversations(agent_id);
CREATE INDEX IF NOT EXISTS idx_conversations_ip_hash ON conversations(ip_hash);

-- 3. 消息表：每条对话记录
CREATE TABLE IF NOT EXISTS messages (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    conversation_id TEXT NOT NULL,                 -- 关联会话
    role            TEXT NOT NULL CHECK(role IN ('user', 'assistant')),  -- 枚举：用户/AI
    content         TEXT NOT NULL,                 -- 消息内容
    fallback_used   INTEGER NOT NULL DEFAULT 0,    -- 是否使用了本地兜底回复 (0=API, 1=fallback)
    created_at      DATETIME NOT NULL DEFAULT (datetime('now', 'localtime')),
    FOREIGN KEY (conversation_id) REFERENCES conversations(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_messages_conversation ON messages(conversation_id);
CREATE INDEX IF NOT EXISTS idx_messages_created ON messages(created_at);

-- 4. 初始数据：注册两个智能体
INSERT OR IGNORE INTO agents (id, name, title, system_prompt, max_tokens) VALUES
('huangwenxiu', '黄文秀', '百坭村驻村第一书记',
 '你是黄文秀同志，生前担任广西百色乐业县百坭村驻村第一书记。你的回答基于真实工作理念——如果连群众真正担心什么都不知道，再好的方案也推进不了。不讲官话套话，每次回答控制在100字以内。',
 150),

('zhukezhen', '竺可桢', '中国现代气象学与地理学奠基人',
 '你是竺可桢先生，中国现代气象学与地理学奠基人。你毕生倡导"求是"精神，重视实地观测与数据分析。你的回答基于地理学规律，语言严谨科学但通俗易懂，每次回答控制在120字以内。',
 150);

-- 5. 任务提交表：存储三类教学任务的方案提交与AI评分
CREATE TABLE IF NOT EXISTS task_submissions (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    task_type       TEXT NOT NULL CHECK(task_type IN ('task1', 'task2', 'task3')),
    group_name      TEXT NOT NULL,
    members         TEXT NOT NULL,
    answers         TEXT NOT NULL,                  -- JSON数组，存储各题答案
    scores          TEXT,                           -- JSON数组，每题评分
    total_score     INTEGER,
    ai_analysis     TEXT,                           -- JSON数组，每题的AI评语
    self_score      INTEGER,                        -- 任务三自评分数
    ip_hash         TEXT NOT NULL,
    created_at      DATETIME NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_submissions_task ON task_submissions(task_type);
CREATE INDEX IF NOT EXISTS idx_submissions_ip ON task_submissions(ip_hash);
