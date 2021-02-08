
# _author    = 'Sun-ZhenXing'
# _github    = 'https://github.com/Sun-ZhenXing/addFormula2docx'
# _date      = '2021-2-8'
# About this project
# Add formulas(MathML, LaTeX or OMML) to .docx or .doc

from .Formula import init, Formula, LaTeX, MathML, OMML
from .utils import latex_to_mathml, mathml_to_omml, omml_to_mathml
