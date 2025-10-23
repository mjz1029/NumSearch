操作步骤

使用 开箱即用.py 改名为 phone_location_app.py

### 步骤1：确保依赖已正确安装
首先在终端中重新安装所有依赖（确保在打包的Python环境中）：
```bash
pip install flask flask-cors requests pyinstaller
```


### 步骤2：生成.spec配置文件（关键）
1. 生成基础配置文件：
   ```bash
   pyi-makespec --onefile --name "手机号归属地查询工具" phone_location_app.py
   ```
   会生成一个 `手机号归属地查询工具.spec` 文件

2. 编辑这个.spec文件，在`hiddenimports`中添加所有依赖：
   ```python
   # 打开.spec文件后，找到hiddenimports行，修改为：
   hiddenimports=['flask', 'flask_cors', 'requests', 'socket', 'threading', 'webbrowser'],
   ```


### 步骤3：基于.spec文件重新打包
```bash
pyinstaller --onefile --windowed "手机号归属地查询工具.spec"
```


### 关键说明
- **hiddenimports的作用**：强制PyInstaller将这些库包含到最终的可执行文件中
- **验证依赖**：可以通过`pip list`确认`flask`、`flask-cors`、`requests`确实已安装在当前环境中


按照以上步骤重新打包后，生成的可执行文件会包含所有必要的依赖，如果有问题，可以检查.spec文件中的`hiddenimports`是否遗漏了其他库，或尝试在打包前创建一个干净的虚拟环境重新安装依赖。