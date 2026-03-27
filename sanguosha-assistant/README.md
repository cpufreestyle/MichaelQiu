# 三国杀助手 (Sanguosha Assistant)

一个智能的三国杀问答助手，可以回答武将技能、卡牌效果、游戏规则等问题。

## 功能特性

- 🤖 智能问答：回答三国杀相关问题
- 🎴 武将查询：查询武将技能、体力、势力等信息
- 🃏 卡牌查询：查询基本牌、锦囊牌、装备牌效果
- 📜 规则解答：解答游戏规则相关问题

## 使用方法

```python
from assistant import SanguoshaAssistant

assistant = SanguoshaAssistant()
response = assistant.ask("诸葛亮的技能是什么？")
print(response)
```

## 项目结构

```
sanguosha-assistant/
├── data/               # 数据文件
│   ├── heroes.json     # 武将数据
│   ├── cards.json      # 卡牌数据
│   └── rules.json      # 规则数据
├── assistant.py        # 主程序
└── README.md          # 说明文档
```
