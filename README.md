# FEL光源：13.5 nm EUV光刻研究

本项目评估FEL在13.5 nm波段实现高平均功率并匹配EUV光刻照明系统的可行性，供学习、汇报和研发决策。项目定义见 [MISSION.md](MISSION.md)，实际进度见 [STATUS.md](STATUS.md)。

## 当前成果

- [首个里程碑：EUV FEL功率扩展瓶颈与路线初判](outputs/2026-11-04_EUV-FEL功率扩展与照明匹配初判.md)。
- 长期决策及证据边界见 [NOTES.md](NOTES.md)，资料索引见 [RESOURCES.md](RESOURCES.md)。

## 目录

`references/` 存原始资料；`assets/` 存正式素材；`outputs/` 存交付物；`work/` 存可编辑源稿；`tmp/` 存可再生临时文件；`lessons/` 存验证后的经验。

## 安全与跨电脑交接

每轮工作先核对 Git 分支、工作树及 `AGENTS.md`、`MISSION.md`、`STATUS.md`、`NOTES.md`、`RESOURCES.md`。离开电脑前保存关键修改并更新状态；换电脑时先通过 GitHub Desktop 拉取远端及 Git LFS 文件，再只读恢复项目状态。若本地有未提交修改，不得强制拉取、重置或强制推送覆盖。

仓库内部使用相对路径；不得提交凭证。未经用户明确授权不创建提交、不推送或发布。详见 `AGENTS.md`。
