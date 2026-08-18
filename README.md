# 可转债双低每日筛选（网页自动版）

把双低策略筛选器做成**每日自动跑 + 网页展示 + 微信推送**的成品。照搬 QQQ 信号页那套 GitHub Pages + Actions 工作流。

## 功能
- 每日收盘后（北京时间 22:00）自动抓取东方财富可转债行情 + 基本信息
- 计算并输出：
  - **安全双低候选**：双低值排序 + 评级≥AA- + 规模 2~30 亿 + 价格≤130
  - **剩余本息锚 / 买入上限**（安全债核心，标记"价格超买入上限=偏贵"）
  - **估值水位总开关**：价格中位数/双低值中位数 → 估值结论；双低<130 合格标的 < 15 只触发**空仓信号**
  - **强赎关注区**：价格≥130 的转债清单
- 生成 `index.html`（Apple 风格、移动端优先、含估值水位趋势曲线）
- 每日追加一行到 `history.csv`（形成时间序列，用于判断"估值何时回落"）
- 可选：配置 `SERVERCHAN_KEY` 后，每日把结论推送到微信（Server酱）

## 目录
```
screener.py              主脚本（抓取+计算+输出+推送）
requirements.txt         Python 依赖
.github/workflows/daily.yml   每日定时工作流
index.html              生成的网页（GitHub Pages 源）
output/                 每日 Excel/CSV 快照（存档）
history.csv             估值水位时间序列（存档+趋势图）
```

## 部署到 GitHub（一次性）

```bash
cd cb-screener
git init
git add -A
git commit -m "init: 可转债双低每日筛选"
# 在 GitHub 新建一个空仓库（如 convertible-bond-screener），然后：
git remote add origin git@github.com:<你的用户名>/convertible-bond-screener.git
git branch -M main
git push -u origin main
```

## 开启自动运行
1. 仓库 **Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: main / (root)**，保存。
   几分钟后访问 `https://<用户名>.github.io/convertible-bond-screener/` 即每日更新页。
2. 仓库 **Settings → Secrets and variables → Actions → New repository secret**，名 `SERVERCHAN_KEY`、值填你的 Server酱 SCT 密钥（不填则只出网页不推送）。
3. 工作流默认 **每天 UTC 14:00（北京 22:00）** 自动跑；也可在 **Actions → 可转债双低每日筛选 → Run workflow** 手动触发。

## 本地手动跑（调试用）
```bash
pip install -r requirements.txt
python screener.py
```
> 本地无公网或抓取失败时，脚本会自动回退读取 `output/` 下已有的缓存 JSON（quote_cache_*.json / info_cache_*.json / q2_p*.json / kzz_info*.json），保证网页与存档照常生成。

## 免责声明
本仓库所有内容均为基于公开数据的量化筛选结果，**仅供研究参考，不构成任何投资建议**。可转债存在信用与价格风险，历史表现不预示未来，实盘决策需独立判断并严格控制仓位。
