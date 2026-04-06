# TOOLS.md - Local Notes

Skills define _how_ tools work. This file is for _your_ specifics — the stuff that's unique to your setup.

## What Goes Here

Things like:

- Camera names and locations
- SSH hosts and aliases
- Preferred voices for TTS
- Speaker/room names
- Device nicknames
- Anything environment-specific

## Examples

```markdown
### Cameras

- living-room → Main area, 180° wide angle
- front-door → Entrance, motion-triggered

### SSH

- home-server → 192.168.1.100, user: admin

### TTS

- Preferred voice: "Nova" (warm, slightly British)
- Default speaker: Kitchen HomePod
```

## Why Separate?

Skills are shared. Your setup is yours. Keeping them apart means you can update skills without losing your notes, and share skills without leaking your infrastructure.

---

## 邮件发送

### 163 邮箱配置
- 邮箱: skygentwu@163.com
- SMTP: smtp.163.com:465 (SSL)
- 发信脚本: `./tools/send_mail.py`

### 使用方法
```bash
python3 tools/send_mail.py <收件人> <主题> <内容文件>
```

或让我直接帮你发：告诉我收件人、主题、内容即可。

---

## PPT模板规范

### 浪潮集团模板（默认模板）
**参考文件**: `深老非气井开采与采气智能体-汇报PPT-浪潮模板版.pptx`

**字体规范**:
| 元素 | 字体 | 字号 | 样式 |
|------|------|------|------|
| 封面标题 | 微软雅黑 | 42号 | 加粗 |
| 副标题 | 微软雅黑 | 28号 | 常规 |
| 目录标题 | 微软雅黑 | 32号 | 加粗 |
| 目录项 | 微软雅黑 | 24号 | 常规 |
| 内容页标题 | 微软雅黑 | 32号 | 加粗 |
| 二级标题 | 微软雅黑 | 20号 | 加粗 |
| 正文 | 微软雅黑 | 18号 | 常规 |
| 底部标语 | 微软雅黑 | 10号 | 常规 |

**配色方案**:
- 浪潮蓝: RGB(0, 102, 204) - 用于Logo、标题
- 浪潮橙: RGB(255, 102, 0) - 用于装饰、章节编号
- 正文深灰: RGB(51, 51, 51)
- 副标题灰: RGB(102, 102, 102)

**装饰元素**:
- 左上角: inspur 浪潮 Logo
- 右下角: 橙红色大三角形 + 蓝色小三角形
- 底部: "未来，因潮澎湃 Inspur in Future"

**生成脚本**: `create_inspur_ppt.py`
