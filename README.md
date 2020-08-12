# AddFormula2DOCX
Add Any Formula to DOCX (LaTex/MathML)

向docx添加任意类型的公式（包括LaTeX和MathML）
## Installation 安装
just copy the file "addFormula2Word.py", run

只需下载上面的文件即可
## Usage 使用
Please download the file and read the document
请下载文件并详细阅读
```python
doc = Document()
para1 = doc.add_paragraph("Text ... Text")
transform = getTransform("MML2OMML.XSL")
para2 = doc.add_paragraph()

addEq2Word(transform, mathml_string, "mathml", para1)
addEq2Word(transform, mathml_string, "mathml", para2, "non-in-line")
addEq2Word(transform, mathml_string, "mathml", doc.add_paragraph(), "non-in-line")
```
## Start 安装依赖
```
pip install latex2mathml
```
## License 许可证
[MIT License](LICENSE)
