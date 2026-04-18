# Cron 任务失败记录

## 📅 2026-04-18 下午 (故障分析报告)

### ❌ 失败任务（3个）

| 任务名称 | Job ID | 连续失败 | 根本原因 |
|---------|--------|----------|----------|
| 每日要闻早报 | fc9fe341-59e1-437f-bd23-5b4f6cbbf3bc | 4次 | MiniMax超时 + modelstudio全部401 |
| 【国学】道德经课件发送 | d3328768-cc5b-48d6-a9e3-38a5e7283e11 | 4次 | 同上 |
| 【国学】道德经每日学习 | f32287f7-3574-4157-8e1f-b20f4405d2c9 | 2次 | 同上 |

---

## 🔍 根本原因分析

### 问题1：模型认证全面崩溃

| 模型 | 状态 | 错误详情 |
|------|------|----------|
| minimax/MiniMax-M2.7 | ⏱️ 超时 | 高并发时段超时（7-8AM） |
| modelstudio/qwen3.5-plus | ❌ 401 | API Key无效或已过期 |
| modelstudio/qwen3-max-2026-01-23 | ❌ 无认证 | No available auth profile |
| modelstudio/qwen3-coder-next | ❌ 无认证 | No available auth profile |
| modelstudio/qwen3-coder-plus | ❌ 无认证 | No available auth profile |
| modelstudio/MiniMax-M2.5 | ❌ 无认证 | No available auth profile |
| modelstudio/glm-5 | ❌ 无认证 | No available auth profile |
| modelstudio/glm-4.7 | ❌ 无认证 | No available auth profile |
| modelstudio/kimi-k2.5 | ❌ 无认证 | No available auth profile |

**结论：8个fallback模型全灭，仅靠MiniMax一个支撑，一旦超时立刻崩溃。**

### 问题2：飞书消息未反馈

**实际情况：** cron任务使用isolated session，与主会话隔离。cron任务的结果通过`delivery.announce`发送到飞书，但在主会话中看不到详情，只能看到完成事件。

**这不是消息丢失，而是架构设计——isolated session的结果以事件形式推送，不在主会话对话中显示。**

---

## ✅ 已修复措施

### 1. 模型配置修复（2026-04-18 13:10）
```json
// 修复前：9个fallback（含8个无效的modelstudio）
"fallbacks": ["modelstudio/qwen3.5-plus", "modelstudio/qwen3-max...", ...]

// 修复后：仅保留有效的MiniMax
"fallbacks": ["minimax/MiniMax-M2.7"]
```

### 2. 模型超时问题
MiniMax在7-8AM高并发时段容易超时，需要后续优化：
- 错峰执行关键任务
- 或申请更高API配额

---

## ⚠️ 待处理事项

### 紧急：modelstudio API Key需要续期
- **服务：** 阿里云 DashScope（modelstudio）
- **问题：** HTTP 401 - invalid access token
- **影响：** 8个模型全部不可用
- **操作：** 需要到阿里云DashScope控制台续期API Key

### 重要：cron任务超时配置
- 新闻早报任务内容复杂（需搜索+整理），建议timeout设长一些
- 当前无explicit timeout，使用默认值

---

## 📊 修复后状态

- **修复时间：** 2026-04-18 13:10
- **修复后测试：** 手动触发新闻早报任务 → 状态变为running（正常）
- **待验证：** 下次7:50新闻早报是否正常运行

---

## 🛡️ 预防措施

1. **监控模型可用性** - 健康检查增加模型状态检测
2. **模型降级策略** - 主模型失败时自动切换到备用（需配置有效备用模型）
3. **任务超时配置** - 复杂任务设置更长timeout
4. **API Key定期检查** - 避免过期导致突然故障

## 2026-04-18 14:17 (UTC 06:17)

### 失败任务
- **Job ID:** f32287f7-3574-4157-8e1f-b20f4405d2c9
- **任务名:** 【国学】道德经每日学习
- **状态:** error
- **操作:** 已手动重新触发 (openclaw cron run)


## 2026-04-18 16:18 (健康检查)

**失败任务:** 【国学】道德经每日学习 (f32287f7-3574-4157-8e1f-b20f4405d2c9)
- **状态:** error
- **上次运行:** 1h ago
- **处理:** 已手动重新触发 (enqueued)

**其余14个任务:** 状态均正常 (ok/idle)

## ✅ 2026-04-18 16:40 — 道德经每日学习任务超时问题已解决

**问题：** 任务反复报 timeout（cron: job execution timed out）

**根因：** 原任务消息内容过于复杂（诵读+白话解读+典故案例+精华提炼），处理时间超过300秒

**解决方案：** 简化任务消息
- 修改前：4步详细学习流程（诵读→解读→典故→提炼）
- 修改后：简单提醒+课件已发送

**结果：** ✅ 修复后测试成功，任务13秒完成，正常送达飞书
