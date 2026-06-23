# 合同与文档 PDF 自动生成系统

基于 Word 模板的合同文档 PDF 自动生成系统，支持 7 种文档的在线编辑与 PDF 导出。

## 文档类型

| # | 文档 | 端口 |
|---|------|------|
| 1 | 委托代理协议 | 8088 |
| 2 | 俄英贸易合同 | 8089 |
| 3 | 附件2 | 8091 |
| 4 | 代理出口结算协议 | 8092 |
| 5 | 应付账款垫付协议 | 8093 |
| 6 | 垫付申请书与承诺函 | 8094 |
| 7 | 付款保函 | 8095 |
| **合并** | 全部文档一键生成 ZIP | **8090** |

## 快速启动

```bash
# 启动单个服务（以合并页面为例）
cd combined-generator
python server.py

# 在浏览器打开
# http://127.0.0.1:8090
```

## 系统要求

- **Windows**（需要 Word 桌面版用于 PDF 转换）
- Python 3.8+
- Word COM 组件（Microsoft Office 已安装）

## 技术架构

- 后端：Python HTTP 服务器（ThreadingHTTPServer）
- 前端：纯 HTML + CSS + JavaScript
- 文档处理：python-docx + lxml 对模板 XML 进行后处理
- PDF 转换：Word COM（win32com.client）
- 模板格式：DOCX（Word Open XML）