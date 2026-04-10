# 💾 OpenClaw 备份方案

## 📋 备份概述

**备份目的：** 确保工作和聊天记录安全，防止数据丢失
**备份位置：** ~/openclaw-backups/
**备份内容：**
- 工作区文件（/Users/mac/.openclaw/workspace/）
- 配置文件（/Users/mac/.openclaw/openclaw.json）
- 记忆数据
- 个人兴趣数据库（美食、房源、书单等）

---

## 📁 备份文件结构

```
~/openclaw-backups/
└── openclaw_backup_YYYYMMDD_HHMMSS.tar.gz
```

**自动保留：** 最近10个备份（自动清理旧备份）

---

## 🔧 备份脚本

**位置：** /Users/mac/.openclaw/scripts/backup.sh

**使用方法：**
```bash
# 手动执行备份
/Users/mac/.openclaw/scripts/backup.sh

# 查看备份列表
ls -lh ~/openclaw-backups/

# 查看备份大小
du -sh ~/openclaw-backups/*
```

---

## ⏰ 定时备份计划

### 待设置
| 时间 | 说明 | 状态 |
|------|------|------|
| 早上 10:00 | 每日第一次备份 | ⚠️ 待配对解决 |
| 晚上 22:00 | 每日第二次备份 | ⚠️ 待配对解决 |

---

## 📊 备份记录

| 时间 | 文件 | 大小 | 状态 |
|------|------|------|------|
| 2026-04-09 21:32 | openclaw_backup_20260409_213247.tar.gz | 516K | ✅ 成功 |

---

## 🔄 恢复方法

如果需要恢复备份：

```bash
# 1. 进入备份目录
cd ~/openclaw-backups/

# 2. 查看可用备份
ls -lt *.tar.gz

# 3. 解压恢复（会覆盖现有文件，请谨慎）
tar -xzf openclaw_backup_YYYYMMDD_HHMMSS.tar.gz -C ~/
```

---

## ⚠️ 待办事项

- [ ] 配对问题解决后，设置定时备份任务（10:00、22:00）
- [ ] 建议同步备份到云存储（如有需要）

---

_方案创建时间：2026-04-09_
