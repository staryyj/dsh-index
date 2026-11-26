# DSH 插件索引（自建）

把本仓库推到 GitHub，然后用 raw 地址作为 DSH 的**索引源**：

```
https://raw.githubusercontent.com/<账号>/<仓库>/main/index.json
```

## 发布三步

1. 在 GitHub 建一个公开仓库（例如 `dsh-index`），把本目录推上去；
2. 在仓库根目录运行：

   ```bash
   python make_index.py --repo https://github.com/<账号>/<仓库>
   git commit -am "index: dsh-bid-docs 0.1.0" && git push
   ```

   （`make_index.py` 会把真实 commit SHA、发布时间写进 `index.json`——**索引里的 git 源必须是 40 位 commit**，不能写分支名。）
3. 打开 DSH → **Plugins** → **Discover** → **Add index source**，粘贴上面的 raw 地址。

## 仓库结构

```
index.json                     # format 1 索引
icons/bid-docs.svg             # 图标（相对路径会按索引 URL 解析）
packages/dsh-bid-docs/         # 可安装的 bundle 源码（技能 + 脚本 + 模板）
make_index.py                  # 推送后生成正式索引
```

## 说明

- 市场索引里的一行 = 一个**可安装 bundle**；技能（skill）不能单独上架，所以“投标方案生成”已打成 bundle `dsh-bid-docs`；
- 条目来源支持 `npm` 与 `git`；git 源必须固定 40 位 commit，且仓库里要能直接 `pnpm add` 该目录；
- 索引是第三方内容，安装前请自行确认来源。
