# AddFormula2DOCX
Add Any Formula to DOCX (LaTex/MathML)
向docx添加任意类型的公式（包括LaTeX和MathML）
## Installation
just copy the file "addFormula2Word.py", run
## Usage
Please download the file and read the document
```python
doc = Document()
para1 = doc.add_paragraph("Text ... Text")
transform = getTransform("MML2OMML.XSL")
para2 = doc.add_paragraph()

addEq2Word(transform, mathml_string, "mathml", para1)
addEq2Word(transform, mathml_string, "mathml", para2, "non-in-line")
addEq2Word(transform, mathml_string, "mathml", doc.add_paragraph(), "non-in-line")
```
## Start
```
pip install latex2mathml
```
## License
[MIT License](License)
