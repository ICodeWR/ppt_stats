# ppt-stats

<p align="center">
  <img src="assets/logo.svg" alt="ppt-stats logo" width="128">
</p>

<p align="center">
  <strong>PPT 批量统计工具</strong> — 一条命令扫描整个文件夹，统计页数、尺寸、形状数量
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License"></a>
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/python-3.9+-blue.svg" alt="Python"></a>
</p>

## 功能

| 功能 | 说明 |
| --- | --- |
| 批量统计页数 | 扫描文件夹下所有 .pptx 文件，输出每份的页数 |
| 幻灯片尺寸 | 输出宽度、高度（cm 和 px 两种单位） |
| 形状分类统计 | 区分文本框、图片、表格、图表、组合等 |
| 多格式输出 | 支持表格、Markdown、CSV、JSON 四种格式 |
| 递归搜索 | 自动搜索子目录 |

## 安装

本项目使用 [uv](https://docs.astral.sh/uv/) 管理依赖。

```bash
git clone <repo-url>
cd ppt-stats
uv sync
```

## 使用

```bash
# 统计当前目录下所有 PPT
uv run python ppt_stats.py

# 统计指定文件夹
uv run python ppt_stats.py "D:\交付文档" -r

# 输出为 Markdown
uv run python ppt_stats.py "D:\交付文档" -r -m -o 统计结果.md

# 输出为 CSV（Excel 可直接打开）
uv run python ppt_stats.py "D:\交付文档" -r -c -o 统计结果.csv

# 输出为 JSON
uv run python ppt_stats.py "D:\交付文档" -r -j -o 统计结果.json
```

### 命令行参数

| 参数 | 说明 |
| --- | --- |
| `path` | PPT 文件或目录路径（默认: 当前目录） |
| `-r, --recursive` | 递归搜索子目录 |
| `-d, --detail` | 显示每页详细信息 |
| `-m, --markdown` | Markdown 表格格式输出 |
| `-c, --csv` | CSV 格式输出 |
| `-j, --json` | JSON 格式输出 |
| `-o, --output FILE` | 将结果保存到文件 |

## 开源协议

本项目基于 [MIT License](LICENSE) 开源，版权归 **码上工坊** 所有。

## 关注我们

<p align="center">
  <img src="assets/msgf.jpg" alt="码上工坊公众号" width="200">
</p>

<p align="center">
  微信扫码关注「<strong>码上工坊</strong>」，获取更多实用工具
</p>

<p align="center">
  回复「<strong>PPT统计</strong>」获取完整源码和安装包
</p>