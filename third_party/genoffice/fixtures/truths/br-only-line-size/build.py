# 功能：集成上游合成 DOCX 样本生成器，原始作者归属与许可见 ../../LICENSE；时间：2026-11-03；集成作者：lq。
import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = ('<w:p><w:pPr>' + spacing(after=0, line=240, rule='auto') + rpr('Times New Roman', 28) + '</w:pPr>'
        '<w:r>' + rpr(None, 22) + '<w:br/></w:r><w:r>' + rpr('Times New Roman', 28, extra='<w:b/>') + '<w:t>TNR 14 bold text</w:t></w:r></w:p>'
        + '<w:p><w:r>' + rpr(None, 56) + '<w:br/></w:r><w:r>' + rpr('Times New Roman', 28) + '<w:t>after a 28pt br</w:t></w:r></w:p>')
write(out_path(), body, styles_xml(doc_defaults('Times New Roman', 24)), settings_xml(), sect=A4)
