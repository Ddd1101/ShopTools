# AGENTS.md

本文件为 AI 编码助手（Agent）提供 ShopTools 工程的背景、架构与强制约束。
开始任何任务前，请先阅读本文件与 [CHANGELOG.md](CHANGELOG.md) 了解最近改动。

## 概述

ShopTools 是一个基于 **PySide2（Qt）** 的桌面 GUI 工具，用于把 **1688 商家订单** 导出/生成为工厂的「备货单」Excel。主要面向童装/饰品制衣厂。

- 入口：[main.py](main.py)（`if __name__ == "__main__"` 创建 `QApplication` → `Window()` → `show()`）
- Python 环境：conda 的 `ShopTools`
- 主要依赖：PySide2、requests、xlrd、xlsxwriter、PIL、yaml

## 目录结构

```
main.py                # 入口 + Window 类（业务逻辑、1688 API、Excel 生成、线程）
Main.ui / QtUi.ui      # 旧版 .ui（已被 View/MainWindow.py 取代，勿再使用）
View/
  MainWindow.py        # 纯代码构建主窗口（3 Tab + 日志框）
  MenuSettingView.py   # 「设置 → Key设置」对话框
Common/
  Utils.py             # 1688 API 签名、下单/交易数据、价格表查询
  ImageHandler.py      # 商品图片存取
ShopBackData/          # git 子模块（github.com/Ddd1101/ShopBackData.git）
config.yaml            # 店铺配置（实际密钥以 main.py 内硬编码字典为准）
CHANGELOG.md           # 变更历史（每次改动后追加）
```

## 数据文件

- `price.xlsx` — 童装价目表
- `price_accessor.xlsx` — 饰品价目表
- `factory.xlsx` — 工厂/地址表
- 店铺密钥硬编码在 [main.py](main.py) 顶部的 `AppKey` / `AppSecret` / `access_token` 字典（联球/朝雄/朝瑞/万盈/睿得/茜阳）。

## 核心业务流程

`OrderList` → `GetOrderBill` → `GetOrderHistory` / `GetSingleOrder` → `GetBeihuoJson` → `GetTable` → `PrintSumReporter`（写 Excel）。
另有 `CheckDelivery`（异常物流检测）、`AliAccessor*`（饰品厂流程，与童装平行但未复用）。

## 强制约束（Hard Constraints）

1. **日志必须走 LogEmitter 信号跨线程投递**，禁止在工作线程直接操作 Qt 日志控件。
2. **禁止在工作线程直接访问任何 Qt 界面控件**（含读取文本框内容、调用 `QDesktopServices` 等）。所有跨线程界面操作必须通过 Qt 信号槽。
3. **任务入口需加防重入锁**（`_is_running`），避免重复点击导致多任务并发堆积。
4. **任务运行期间禁用所有启动按钮，任务结束（正常/异常/手动中断）后恢复**。

## 工程约定（Engineering Conventions）

- 日志方法 `Logout` 统一处理文件写入与界面输出；debug 级别默认仅写文件，由界面「debug日志输出到输出框」复选框控制。
- 终端 `stdout`/`stderr` 通过 `_TeeStream` 同时转发到界面日志框，保留终端原始输出。
- 任务状态变更通过自定义 Qt 信号在主线程更新界面，避免跨线程操作 Qt。

## 踩坑教训（Lessons Learned）

- 非 GUI 线程直接操作 Qt 控件会导致日志无法实时刷新、低配设备易闪退。
- 工作线程访问 Qt 控件必须在主线程预先读取所有界面参数。
- 界面优化方案若不符合用户「专业高级」预期会被取消，界面改动需谨慎。
- 任务异常中断时若未正确复位 `_is_running` 锁，会导致无法再次启动，须在 `finally` 中统一释放。

## 提交约定（Commit Conventions）

- commit message 分两段：**第一行写摘要**（Conventional Commits 前缀 + 一句话概括），**换行后的正文写改造内容**，分点说明改了什么、为什么改、影响范围、涉及的文件/模块。
- **bugfix 必须写明「问题」和「解决方案」**两部分。
- 涉及多个独立改动时，拆分为多个聚焦单一主题的提交，便于回溯。

格式示例：

    feat: 新增打开备货单文件夹与文件选择功能

    - 左侧工具条新增「打开备货单文件夹」「备货单文件下拉框」「打开文件」三个控件
    - 下拉框列出备货单文件夹内所有 xlsx，按修改时间倒序，过滤 ~$ 锁文件
    - 产出文件后经 latestFileChanged 信号跨线程回填并自动选中最新文件

    fix: 货号归类错误

    问题：原按「取最长段」启发式归类会破坏 TZ- 前缀，导致 TZ-240312-17 归并错误
    解决方案：改为剥离末尾 -数字(1~4位) 变体后缀，统一到基础货号

## 变更记录约定

- 每次重要改动完成后，在 [CHANGELOG.md](CHANGELOG.md) 顶部追加一条记录（日期 + 摘要）。
- 每次新对话处理任务前，先读 [CHANGELOG.md](CHANGELOG.md) 了解最近改动。