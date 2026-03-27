# 三国杀助手 - 开源发布指南

## 项目信息

- **项目名称**: Sanguosha Assistant (三国杀问答助手)
- **版本**: 1.0.0
- **简介**: 随时为你解答三国杀规则、武将技能、卡牌用法的贴心助手
- **技术栈**: Python
- **许可证**: MIT License

## 快速开始

### 1. 克隆项目
```bash
git clone https://github.com/你的用户名/sanguosha-assistant.git
cd sanguosha-assistant
```

### 2. 安装依赖
```bash
pip install -r requirements.txt
```

### 3. 运行
```bash
python assistant.py
```

### 或者直接运行 EXE
```bash
# Windows 用户可以直接运行
./dist/三国杀助手.exe
```

## 功能特性

- 🎴 **武将查询** - 收录10+经典武将，查询技能效果、体力值、势力信息
- 🃏 **卡牌查询** - 基本牌、锦囊牌、装备牌全覆盖
- 📜 **规则解答** - 游戏规则、判定机制、回合流程全面解答
- 💬 **智能问答** - 自然语言交互，随时随地解答问题

## 目录结构

```
sanguosha-assistant/
├── assistant.py        # 主程序入口
├── data/
│   ├── heroes.json     # 武将数据
│   ├── cards.json      # 卡牌数据
│   └── rules.json      # 规则数据
├── dist/
│   └── 三国杀助手.exe  # Windows 可执行文件
├── README.md          # 项目说明
├── QUICKSTART.md       # 快速开始指南
└── start.bat           # Windows 启动脚本
```

## 使用示例

```python
from assistant import SanguoshaAssistant

assistant = SanguoshaAssistant()

# 查询武将
print(assistant.ask("诸葛亮的技能是什么？"))

# 查询卡牌
print(assistant.ask("南蛮入侵怎么用？"))

# 查询规则
print(assistant.ask("判定是什么？"))
```

## 武将列表

- 刘备 - 仁德之君 (蜀, 4体力)
- 关羽 - 美髯公 (蜀, 4体力)
- 张飞 - 万夫不当 (蜀, 4体力)
- 诸葛亮 - 迟暮的丞相 (蜀, 3体力)
- 赵云 - 少年将军 (蜀, 4体力)
- 曹操 - 魏武帝 (魏, 4体力)
- 司马懿 - 狼顾之鬼 (魏, 3体力)
- 孙权 - 年轻的贤君 (吴, 4体力)
- 吕布 - 无双 (群, 4体力)
- 貂蝉 - 绝世的舞姬 (群, 3体力)

## 贡献指南

欢迎提交 Issue 和 Pull Request！

## 许可证

MIT License