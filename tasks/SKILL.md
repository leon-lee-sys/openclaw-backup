# 任务管理技能

## 功能
管理三先生安排的所有任务，记录执行状态，定期自检。

## 核心文件
- `/Users/mac/.openclaw/workspace/tasks/tasks.md` - 任务列表
- `/Users/mac/.openclaw/workspace/tasks/log.md` - 执行日志

## 任务记录格式
```
## 任务名称
- 创建日期：YYYY-MM-DD
- 创建人：三先生
- 任务内容：...
- 执行周期：每日/每周/一次性
- cron ID：xxx（如果有）
- 状态：✅进行中/❌已停止/🔄暂停
- 最后执行：YYYY-MM-DD
- 执行记录：...
- 备注：...
```

## 自我检查规则
每天晚10点检查一次：
1. 读取 tasks.md
2. 对比 cron 列表
3. 发现任务缺失或遗漏，立即补救
4. 记录检查结果到 log.md

## 添加新任务
1. 在 tasks.md 追加任务记录
2. 设置对应 cron
3. 更新本技能说明

## 现有任务清单
见 tasks.md
