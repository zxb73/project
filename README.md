# 股票数据分析系统 - 打包成单文件 .exe 指南

此仓库包含一个基于 PyQt5 的桌面应用脚本 `20251003.py`，用于分析本地股票数据并将分析结果输出为桌面上的 Word 文档。

目标：生成一个可在任意 Windows 机器上双击运行的单文件 `.exe`，不需要用户安装 Python 或依赖。

重要：程序需要一个 DeepSeek API 密钥，运行时通过环境变量 `DEEPSEEK_API_KEY` 提供，或在程序启动后在界面中输入。

- 不要在源码中明文写入密钥。

## 本地在 Windows 上构建（推荐）

1. 在 Windows 机器上安装 Python 3.10/3.11，并在 `PATH` 中可用。
2. 将本仓库复制到本地目录（或直接把 `20251003.py` 放到一个文件夹）。
3. 打开 PowerShell 或 CMD，进入项目目录，执行：

```powershell
python -m pip install -r requirements.txt
.\\build_exe_windows.bat
```

4. 构建完成后，`dist` 目录下会生成单文件 `.exe`（文件名与脚本名对应）。将该 `.exe` 复制到其他 Windows 机器即可双击运行。

注意：如果需要显示控制台输出或调试，请将 `build_exe_windows.bat` 中的 `--windowed` 选项移除。

## 使用 GitHub Actions 自动构建并下载 .exe（无 Windows 机器）

仓库已包含一个 GitHub Actions workflow：`.github/workflows/build-windows.yml`。

操作步骤：
1. 将代码推送到 GitHub 仓库。
2. 在仓库页面的 "Actions" 里手动触发或等待 push 触发构建。
3. 构建成功后，在 workflow 运行页面下载生成的 artifact（`stock-analyzer-exe`），里面包含 `dist/*.exe` 文件。

## 运行时说明

- 运行前请确保目标机器已下载并放置好 `DEEPSEEK_API_KEY` 环境变量（可在系统环境变量中设置），或者在程序首次启动时手动输入密钥。
- 生成的 `.exe` 是独立的单文件程序，但某些杀毒软件可能误报，请在分发前做好签名与信任链处理（可选）。

## 常见问题

- PyInstaller 打包 PyQt5 时若缺少插件显示异常，可以尝试在打包命令中添加 `--add-data` 指定 PyQt5 的 Qt 平台插件目录，或在 Windows 上直接使用该 workflow 的构建产物。

## 安全

切勿在仓库中提交或公开你的 API 密钥。

