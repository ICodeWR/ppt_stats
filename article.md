# 几百个PPT要统计页数，甲方要求每份不能少于N页——我写了个工具一键搞定

> 🏗️ 造个工具 · 第5篇 | 项目验收被卡在PPT页数上？批量统计工具来了


## 一、故事是这样开始的

项目验收前一天，甲方在群里丢了一句话：

> "所有交付的PPT，页数不能少于30页。你们自己先检查一遍，不合格的补上。"

我打开文件夹一看——好家伙，**187个PPT**。

手动打开每一个看页数？那今晚别想回家了。用Python写个脚本吧，10分钟的事。

于是就有了这个工具：一条命令，扫描整个文件夹，把所有PPT的页数、幻灯片尺寸、图片数量、表格数量全部统计出来，还能导出成表格直接发给甲方。

这就是 **ppt-stats** 的由来。


## 二、工具能做什么

| 功能 | 说明 |
|------|------|
| **批量统计页数** | 扫描文件夹下所有PPT，输出每份的页数 |
| **幻灯片尺寸** | 输出宽度、高度（cm和px两种单位） |
| **形状分类统计** | 区分文本框、图片、表格、图表、组合等 |
| **多格式输出** | 支持表格、JSON、Markdown、CSV四种格式 |
| **递归搜索** | 自动搜索子目录，一次性统计整个项目文件夹 |
| **单文件详情** | 可查看每个PPT每一页的详细形状信息 |


## 三、安装与使用

### 3.1 安装

```bash
pip install ppt-stats
```

或者直接运行脚本（源码在文末获取）。

### 3.2 基础用法

```bash
# 统计当前目录下所有PPT
ppt-stats .

# 统计指定文件夹，递归搜索子目录
ppt-stats ./交付文档 -r

# 输出为CSV，方便在Excel里筛选
ppt-stats ./交付文档 -r --csv -o 统计结果.csv
```

### 3.3 输出示例

```bash
$ ppt-stats ./ppts -r
```

```
文件名                        | 页数 | 宽(cm) | 高(cm) | 宽(px) | 高(px) | 形状总数 | 图片数 | 表格数 | 图表数
-----------------------------|------|--------|--------|--------|--------|---------|--------|--------|--------
项目方案_v1.pptx             | 32   | 25.40  | 19.05  | 960    | 720    | 156     | 23     | 2      | 3
项目方案_v2.pptx             | 28   | 25.40  | 19.05  | 960    | 720    | 142     | 18     | 1      | 2
技术架构图.pptx              | 15   | 33.87  | 19.05  | 1280   | 720    | 89      | 45     | 0      | 0
验收报告_终版.pptx           | 45   | 25.40  | 19.05  | 960    | 720    | 234     | 31     | 5      | 4
...
--------------------------------------------------------------------------------
文件总数: 187,  幻灯片总页数: 5234
```

一眼就能看出哪些PPT页数不够。

### 3.4 甲方验收场景：筛选页数不足的PPT

```bash
# 导出CSV，在Excel里筛选页数<30的文件
ppt-stats ./交付文档 -r --csv -o 页数统计.csv
```

打开Excel，用筛选功能快速找出页数不足的PPT，挨个补页。验收前把所有文件过一遍，发给甲方时心里有底。


## 四、核心代码解析

### 4.1 EMU单位转换

PPT内部使用EMU（英制计量单位），需要转换成厘米和像素才能看懂：

```python
def emu_to_cm(emu: int) -> float:
    """将 EMU 转换为厘米"""
    return round(emu / 914400 * 2.54, 2)

def emu_to_px(emu: int, dpi: int = 96) -> int:
    """将 EMU 转换为像素"""
    return int(emu / 914400 * dpi)
```

### 4.2 形状分类统计

区分不同类型的形状，让统计更有用：

```python
def count_shape(shape):
    if shape.has_text_frame and shape.text.strip():
        text_boxes += 1
    shape_type = shape.shape_type
    if shape_type == 13:      # 图片
        pictures += 1
    elif shape_type == 19:    # 表格
        tables += 1
    elif shape_type == 3:     # 图表
        charts += 1
    elif shape_type == 6:     # 组合
        groups += 1
```

### 4.3 递归搜索

一次性统计整个项目文件夹：

```python
def find_ppt_files(directory: str, recursive: bool = False) -> list:
    path = Path(directory)
    pattern = "**/*.pptx" if recursive else "*.pptx"
    return [str(p) for p in path.glob(pattern) if p.is_file()]
```

### 4.4 多格式输出

支持四种输出格式，适配不同场景：

| 格式 | 适用场景 |
|------|---------|
| 表格 | 终端直接查看 |
| Markdown | 写入文档/README |
| JSON | 程序间数据交换 |
| CSV | Excel打开/筛选/排序 |


## 五、代码结构

完整代码约 280 行，主要包含以下模块：

| 模块 | 函数 | 说明 |
|------|------|------|
| 单位转换 | `emu_to_cm` / `emu_to_px` | PPT 内部 EMU 单位转厘米/像素 |
| 形状统计 | `get_shape_info` | 分类统计每页形状（文本/图片/表格/图表等） |
| 文件分析 | `analyze_ppt` | 解析单个 PPT 文件，返回完整统计信息 |
| 文件搜索 | `find_ppt_files` | 递归搜索目录下的 .pptx 文件 |
| 输出格式 | `print_table` / `print_csv` / `print_markdown_table` | 多格式输出 |
| 入口 | `main` | 命令行参数解析与调度 |

> 完整代码已开源，扫码关注公众号发送「**PPT统计**」即可获取。项目基于 MIT 协议开源，欢迎 Star 和 PR。


## 六、安装为命令行工具

将代码保存为 `ppt_stats.py`，在同目录创建 `pyproject.toml`：

```toml
[project]
name = "ppt-stats"
version = "0.1.0"
description = "统计 PPT 页数、尺寸、形状数量等信息的工具"
requires-python = ">=3.9"
dependencies = [
    "python-pptx>=0.6.23",
]

[project.scripts]
ppt-stats = "ppt_stats:main"

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"
```

安装：

```bash
pip install -e .
```

之后就可以在任意目录直接运行 `ppt-stats` 命令了。


## 七、获取源码

📦 回复「**PPT统计**」获取完整源码和安装包。


## 八、系列导航

📍 本文属于「造个工具」系列

**已发布**：
- [x] 第1篇：文件自动整理工具（UI版）
- [x] 第2篇：批量图片压缩工具
- [x] 第3篇：文件批量生成器
- [x] 第4篇：AI证件照工具
- [x] 第5篇：PPT批量统计工具（本文）

📮 关注「码上工坊」，每周四早8点更新
🛠️ 回复「工具」查看效率工具实战合集


## 九、互动

> 你遇到过类似的情况吗？批量处理PPT、Word、PDF的统计需求？
>
> 或者你还想要什么批量统计工具？留言告诉我 👇