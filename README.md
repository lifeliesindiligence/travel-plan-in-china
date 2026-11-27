# travel plan In China

从**人数、地点、日期**三个输入，制作有来源、有路线、有时间安排的国内游玩 PPT。名称保持为 **travel plan In China**；仓库使用便于安装的目录名 `travel-plan-in-china`。

本仓库是一份可持续迭代的 Codex skill。它规定如何核查景区开放、天气、项目、地图和小红书经验，再用自带脚本验证资料并生成 PowerPoint。**生成器不会自动联网搜索；AI 助手必须先完成取证和内容整理。**

## 安装

将本仓库克隆到 Codex 的 skills 目录，或将 `travel-plan-in-china` 文件夹复制到该目录。重启 Codex 后，可以说：

> 用 travel plan In China 给 2 个人做 2027 年 5 月 12 日去某景区的游玩 PPT。

仅需给人数、地点和日期；兴趣、关系、饮食限制等有则补充，不能擅自推断。依赖 Python 3.10+，生成 PPT 时安装：

```bash
python -m pip install -r requirements.txt
```

## 产出流程

1. 核对目的地及目标日期；优先查官方售票、营业公告和项目停运信息，再查天气、地图及交通。
2. 按 `SKILL.md` 搜集小红书攻略与注意事项，记录检索时间、排序方式、样本量和评论实际可见量。视频可选转录；不可用时仍需如实制作条件版。
3. 整理为 `references/plan-schema.md` 所述 JSON，区分已核实、待核实、关闭。路线图与照片只放在本地工作目录，确认使用权。
4. 校验并生成 PPT：

```bash
python scripts/validate_plan.py path/to/plan.json
python scripts/build_ppt.py path/to/plan.json path/to/travel-plan.pptx
```

非 Windows 用户可以指定本机已安装的中文字体；优先级为 `--font`、环境变量 `TRAVEL_PLAN_FONT`、默认 `Microsoft YaHei`：

```bash
python scripts/build_ppt.py path/to/plan.json path/to/travel-plan.pptx --font "PingFang SC"
```

生成器会同时声明西文和东亚字体；默认东亚字体名为 `微软雅黑`。它不会安装或嵌入字体，跨电脑使用仍需选择可用字体并检查渲染结果。

`0.2.0` 保留原暖橙装饰色 `#E7774F`，将橙底白字及白色卡片上的橙字改用 `#BC5230`（对比度约 4.79:1；不用于米白底小字）；关闭状态独立使用 `#B3261E`，页脚为 9.5 磅。时间表和来源列表均衡分页并居中排布，10 行时间表为 5+5、8 条来源为 4+4，无需删减资料来迁就版式。长段落仍需检查是否溢出。

`examples/example-plan.json` 是**虚构演示数据**，仅供测试脚本；不能当成真实攻略。可运行：

```bash
python scripts/build_ppt.py examples/example-plan.json output/example.pptx
```

## 准确性边界

- 没查到目标日期的官方逐日开放信息时，写“待核实”和出发前核查动作；不把往年营业时间写成明天确定营业。
- 排队时间是计划容错，不是实时队列预测。当天公告和现场广播优先于提前制作的 PPT。
- 小红书检索结果不等于全站热度榜。只汇总实际可见的评论；不要声称每篇都读了 30 条，除非记录确实如此。
- 视频、评论、转录、个人凭据和最终旅行者私人计划不提交到公开仓库。

## 版本迭代

当前版本为 `0.2.0`，变更记录见 [CHANGELOG.md](CHANGELOG.md)。保留三项基础输入和事实核查门槛。
