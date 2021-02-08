# Formulas
Add Any Formula to DOCX (LaTex/MathML/OMML)

向docx添加任意类型的公式（包括LaTeX、MathML或者OMML）

### Installation 安装
Clone the repository and import it when needed

只需克隆这个仓库即可，然后直接导入模块

### Usage 使用方法
Please clone it and read the 'formulas/Formulas.py', eg.

请下载文件并详细阅读，用法示例：
```python
from formulas import Formulas, init
from docx import Document
init()
doc = Document()
para = doc.add_paragraph('some text...')
fm1 = Formula('LaTeX|MathML|OMML string')
fm2 = Formula('LaTeX string', 'latex')
fm3 = Formula('MathML string', 'mathml', 'block')
fm1.add_to(doc)
para2 = doc.add_paragraph('some text...') + fm2
para3 = doc + fm3
fm1.get_latex(True) # 'LaTeX string if its available...'
fm1.get_mathml(to_string=True)
fm1.get_omml(to_string=True, safe_mode=True)
```

### Dependencies 安装依赖
```
pip3 install python-docx
pip3 install latex2mathml
```

### Reference 参考
[python-docx:Issues/320](https://github.com/python-openxml/python-docx/issues/320)

Microsoft Office: Word 2016 (MathML2OMML.XSL and OMML2MathML.XSL)

### License 许可证
[MIT License](LICENSE)
