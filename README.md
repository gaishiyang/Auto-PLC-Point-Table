# Auto-PLC-Point-Table 🤖⚡

**AI-Driven PLC Automation Engineering Toolchain**

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

> 将工业 PLC 自动化工程的点位设计、模块排序、端子表生成、EPLAN 导出全链路 AI 化，让工程师从重复劳动中解放。

---

## 解决的问题 / The Problem

PLC 工程项目中，点位表编制、模块排序、端子表生成等环节长期依赖人工 Excel 操作，效率低、易出错、标准不统一。跨团队协作时，不同工程师输出的格式各异，电气设计与 PLC 编程之间缺乏自动化的桥梁。

## 核心功能 / Features

| 模块 | 功能 | 技术亮点 |
|------|------|----------|
| **点位表生成** | 从模板自动分配 I/Q/IW/QW 地址，支持别名映射与优先级排序 | 长链推理: 设备类型 → IO 需求 → 最优地址分配 |
| **模块排序生成** | 支持 S7-1500 / 200SMART 混合 8 点/4 点模块配置 | 多 PLC 型号自适应排序策略 |
| **端子表生成** | 自动生成端子接线表，含 24V 电源分配、KA 继电器编号 | 优先级驱动的最优排布算法 |
| **EPLAN 导出** | 生成标准 EPLAN 电气设计点位表 | 偶数/奇数位交错规则，KA 编号自动映射 |

## 快速开始 / Quick Start

### 方法一：直接运行可执行文件

从 [Releases](https://github.com/gaishiyang/Auto-PLC-Point-Table/releases) 下载 `CC点位.exe`，双击运行即可。

### 方法二：源码运行

```bash
# 克隆仓库
git clone https://github.com/gaishiyang/Auto-PLC-Point-Table.git
cd Auto-PLC-Point-Table

# 安装依赖
pip install openpyxl

# 运行
python plc_generator_ui.py
```

## 项目结构 / Project Structure

```
Auto-PLC-Point-Table/
├── plc_generator_ui.py       # 主程序 (Tkinter GUI)
├── generate_plc.py           # 点位表生成核心逻辑
├── generate_module.py        # 模块排序生成
├── generate_terminal.py      # 端子表生成
├── generate_eplan.py         # EPLAN 点位表导出
├── priority_config.json      # 优先级配置文件
├── README.md
└── LICENSE
```

## 技术架构 / Architecture

```
[用户界面] plc_generator_ui.py (Tkinter)
       │
       ├── generate_plc.py      ──→ PLC 点位地址分配引擎
       ├── generate_module.py   ──→ 模块排序优化引擎
       ├── generate_terminal.py ──→ 端子接线生成引擎
       └── generate_eplan.py   ──→ EPLAN 格式导出引擎
```

支持 **多 Agent 级联推理**: 输出点位 → 模块排序 → 端子生成 → EPLAN 导出，全流程自动衔接。

## 效果 / Impact

- ⏱ **设计效率**: 点位表编制从 3 天 → 2 小时
- ✅ **准确率**: 人工平均 8% 错误率 → 接近零误差
- 🔄 **标准化**: 团队输出格式 100% 统一

## 路线图 / Roadmap

- [ ] 支持更多 PLC 品牌（三菱、西门子、罗克韦尔）
- [ ] Web 版本（FastAPI + Vue3）
- [ ] AI 辅助控制逻辑代码生成

## 许可证 / License

[MIT](LICENSE) © gaishiyang

---

**愿景: 让工业自动化更智能、更高效。**
