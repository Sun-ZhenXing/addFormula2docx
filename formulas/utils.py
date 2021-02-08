from lxml import etree
from latex2mathml.converter import convert

MathPara = '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"></m:oMathPara>'

def _getXSLT(filename):
    xslt = etree.parse(filename)
    return etree.XSLT(xslt)

MML2OMML = _getXSLT('MML2OMML.XSL')
OMML2MML = _getXSLT('OMML2MML.XSL')

def latex_to_mathml(latex, display='inline'):
    style = display if display == 'inline' else 'block'
    return etree.fromstring(convert(latex, display=style))

def mathml_to_omml(mathml, display='inline'):
    if type(mathml) == str:
        output = MML2OMML(etree.fromstring(mathml))
    output = MML2OMML(mathml)
    if display == 'inline':
        return output
    else:
        wrapper = etree.fromstring(MathPara)
        wrapper.append(MML2OMML(output))
        return wrapper

def omml_to_mathml(omml, display='inline'):
    if type(omml) == str:
        return OMML2MML(etree.fromstring(omml))
    return OMML2MML(omml)

