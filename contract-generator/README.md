# 委托代理协议 PDF 自动生成网站

## 使用方式

1. 打开命令行，进入本文件夹。
2. 运行：

```powershell
python server.py
```

3. 浏览器打开：

```text
http://127.0.0.1:8088
```

4. 填写甲乙公司信息、贸易合同信息和收款账户信息，点击“生成并下载 PDF”。

## 生成逻辑

- `templates/agreement-template.docx` 是原 Word 协议模板副本。
- 网站不会改写固定协议条款，只替换表单字段对应的公司、合同、日期、账户等信息。
- 生成 PDF 时会先基于 Word 模板生成临时 DOCX，再调用 LibreOffice/soffice 转换为 PDF，以尽量保持原 Word 排版。

## 注意

- 需要本机可运行 `python`。
- 需要安装 LibreOffice，或确保命令行中可以执行 `soffice`。
- 如果 Word 模板后续有更新，替换 `templates/agreement-template.docx` 后重启网站即可。
