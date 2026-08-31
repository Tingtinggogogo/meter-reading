# 杭州商场每日水电用量抄表

面向 STO401 的移动端抄表应用。网页、API、PostgreSQL 集中存储、月度 Excel 导出和二维码由同一个 Zeabur 项目提供。

## 功能

- 18 个水电点位，服务器自动记录中国标准时间
- 同一站点每天仅保留一条记录，再次提交自动更新
- 填写端使用商场验证码，管理员端使用独立密码
- 管理员按月份查询、删除和导出 Excel
- Excel 匹配业务模板，高压进线明细乘以 3000；光伏和充电桩按原始读数导出
- Excel 使用浏览器原生附件下载，兼容 Android、鸿蒙和 iOS
- Android 导出先获取两分钟有效的签名链接，再交给系统下载管理器处理
- `/api/qr.png` 自动生成当前公网 HTTPS 网站二维码
- `/qr` 提供可打印二维码页面，并提示小米、红米用户使用系统相机和系统浏览器
- `/api/health` 提供 Zeabur 健康检查

## 项目结构

```text
.
├── app/main.py          # FastAPI、PostgreSQL 和 Excel 导出
├── public/index.html    # 移动端网页
├── tests/test_app.py    # 数据校验和报表测试
├── scripts/audit_project.py # 项目交付护栏审计
├── .github/workflows/ci.yml # PR 和 main 的自动测试与镜像构建
├── Dockerfile           # Zeabur 自动识别并构建
├── requirements.txt
├── requirements-dev.txt
└── .env.example
```

## Zeabur 部署

### 1. 创建项目和数据库

1. 登录 [Zeabur](https://zeabur.com/) 并新建 Project。
2. 在项目中选择 **Add Service → Marketplace → PostgreSQL**。
3. 数据库创建完成后保留默认私有网络配置，不需要为数据库绑定公网域名。

### 2. 部署 GitHub 服务

1. 选择 **Add Service → Git → GitHub**。
2. 选择 `Tingtinggogogo/meter-reading` 仓库。
3. Zeabur 会自动识别根目录的 `Dockerfile`。
4. 在 Web 服务的 **Variables** 中设置以下变量：

| 变量 | 必填 | 说明 |
|---|---|---|
| `DATABASE_URL` | 是 | 引用 PostgreSQL 服务暴露的连接字符串；在变量输入框选择数据库服务提供的连接变量 |
| `SUBMISSION_CODE` | 是 | 第三方填写时使用的商场验证码，至少 8 位 |
| `ADMIN_PASSWORD` | 是 | 查询、删除和导出使用的管理员密码，至少 8 位且不能与填写验证码相同 |
| `PUBLIC_URL` | 否 | 无法从请求取得域名时的二维码备用 HTTPS 地址，例如 `https://meter.example.com` |
| `PORT` | 否 | Zeabur 自动注入；Dockerfile 默认使用 `8080` |

不要把真实密码或数据库连接字符串写入 GitHub。

### 3. 绑定域名

1. 打开 Web 服务的 **Networking / Domains**。
2. 先生成 Zeabur 提供的 HTTPS 域名进行测试。
3. 正式使用时绑定自有 HTTPS 域名；服务会优先用当前访问域名生成二维码。
4. 打开 `/api/qr.png`，保存二维码后打印张贴。

### 4. 发布检查

1. 访问 `/api/health`，应返回 `{"status":"ok"}`。
2. 用填写验证码提交一组测试记录。
3. 用管理员密码查询当月记录并导出 Excel。
4. 核对合计公式后删除测试月份记录。

## 本地运行

需要 Python 3.12 和 PostgreSQL：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
$env:DATABASE_URL = "postgresql://postgres:password@localhost:5432/postgres"
$env:SUBMISSION_CODE = "replace-with-a-long-submission-code"
$env:ADMIN_PASSWORD = "replace-with-a-different-long-admin-password"
uvicorn app.main:app --reload
```

打开 <http://127.0.0.1:8000>。

运行测试：

```powershell
python scripts\audit_project.py
python -m pytest -q
```

## 开发与发布流程

1. 从最新 `main` 创建功能分支，禁止直接向 `main` 提交。
2. 修改业务规则时同步检查后端配置、页面文案、Excel、测试和 README。
3. 本地运行项目护栏审计和测试。
4. 推送分支并创建 Pull Request。
5. GitHub Actions 的 `test-and-build` 通过后才能合并；它会重复运行审计、14 项回归测试并构建正式 Docker 镜像。
6. 风险变更先部署到使用独立 PostgreSQL 的 staging 环境验收，再合并到由 Zeabur 监听的生产分支。

仓库长期约束记录在 `.github/copilot-instructions.md`。GitHub 仓库是唯一源码真相，禁止从运行中的容器提取文件重新打包或改变部署方式。

## 数据与安全

- PostgreSQL 表会在服务启动时自动创建，无需手工执行 SQL。
- 数据库不开放公网，只允许同一 Zeabur 项目内的 Web 服务访问。
- 所有写入使用服务器时间，避免用户修改手机日期。
- 普通填写者无法读取历史数据；查询、导出和删除均需管理员密码。
- 建议开启 PostgreSQL 定期备份，并至少每月下载一份 Excel。
- 当前为共享验证码模式，适合低敏感度内部运营数据；验证码泄露后应立即在 Zeabur 中更换。
