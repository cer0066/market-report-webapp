# 电力市场报表自动填报工具

这是一个网页工具，用于把 Excel 文件自动计算并回填到最终报送表模板。目前支持 `水电` sheet 和 `光伏` sheet。

## 本地启动

双击 `start.bat`，浏览器访问：

```text
http://127.0.0.1:8765
```

## 支持的文件

模板文件必填，其余来源文件按当天实际资料上传。没上传的类别不会回填，也不会清空模板。

水电支持：

- 中长期合同明细
- 转让后优先电量
- 转让后留存电量
- 转让后外送

光伏支持：

- 腾龙、凤舞：省内优先、留存、中长期合同明细、转让后外送
- 赛日：省内优先、留存、中长期合同明细、转让后外送

每个光伏资料格都可以选择文件，或勾选“今日无数据”。如果既没有上传也没有勾选，网页处理状态会提示未确认。

网页支持选择文件，也支持把 `.xlsx` 文件直接拖到对应框里。

## 计算逻辑

- 省内市场：合同电量 * (合同电价 + 绿证价格)，按交易单元/售方主体和交易时段汇总。
- 省内优先：优先计划电量 * 优先计划电价。
- 留存：留存电量 * 留存电价。
- 省间外送：外送电量 * (外送电价 + 绿证价格)，绿证价格为空按 0。
- 每小时加权电价 = 汇总电费 / 汇总电量。
- 每小时电量拆成 4 个 15 分钟点，电价四个点保持一致。
- 水电保护 `L:P`、`AC:AG`。
- 光伏保护 `L:P`、`AC:AG`、`AT:AX`。

## 输出位置

生成的 Excel 会保存在：

```text
outputs
```

## 上传 GitHub

上传这些文件和目录：

```text
server.py
README.md
requirements.txt
runtime.txt
Procfile
render.yaml
.gitignore
start.bat
static/
```

不要上传：

```text
uploads/
outputs/
__pycache__/
*.xlsx
```

## 云端部署

这个项目需要 Python 后端处理 Excel，所以不能只用 GitHub Pages。可以把代码放到 GitHub，然后用 Render、Railway 等平台部署。

Render 常用配置：

```text
Build Command: pip install -r requirements.txt
Start Command: python server.py
```

云平台会自动提供 `PORT`，程序会自动读取它；本地运行时仍默认使用 `127.0.0.1:8765`。
