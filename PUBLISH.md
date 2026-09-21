# First GitHub Release

建议仓库名：`med-ppt-editable`。

建议描述：`Editable medical PowerPoint skill with native charts, tables, and source traceability.`

本包已经包含 README、MIT 许可证、skill 指令、生成器、检查器、测试、虚构案例与真实 PowerPoint 预览。首个公开版本标为 `v0.1.1` 预发布版，让使用者清楚这是小范围验证的初版。

## 注册后发布

1. 在 GitHub 完成个人账号注册、邮箱验证和正常登录。密码与验证码由账号本人处理。
2. 新建公共仓库 `med-ppt-editable`。
3. 将本目录中的文件作为仓库根目录内容发布。`SKILL.md` 应位于根目录。上传源文件和目录，不能只放一个 ZIP 就把它当作可直接读取的 skill 仓库。
4. 确认 README 图片、样稿链接、源码与许可证正常显示。
5. 创建 `v0.1.1` 预发布，附发布 ZIP 与以下说明。

本包不包含真实病例。本地 `work/`、患者资料、密钥和其他个人文件不应随代码上传。项目是否公开、仓库归属和最终地址在注册后确认。

## Release Text

```text
Medical Editable PPT v0.1.1

首个面向中文医学病例、文献汇报和学术报告的可编辑 PPT skill 初版。

提供六种基础布局、原生文本与表格、带内嵌工作簿的折线图/柱状图、
可替换图片、来源备注、结构检查器，以及 7 页虚构病例样稿。

补充图形排版和自定义随访图指导，涵盖留白与层级、百分数、重合点标签、
独立内嵌工作簿和追加页面时的数据保留。通用 JSON 生成器仍为六种基础布局。

8 项行为测试通过；样稿已在 Windows PowerPoint 中打开并导出预览。
WPS 尚未验证；不支持任意模板无损转换、可编辑公式或手工修改自动合并。
生成器不调用付费 API，宿主 AI 产品费用另计。

全部示例仅用于演示，不构成医学证据或临床建议。
```
