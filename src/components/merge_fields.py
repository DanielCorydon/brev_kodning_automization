from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def create_merge_field(parent, key):
    """Create a MERGEFIELD in Word XML format"""
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    parent.append(begin)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" MERGEFIELD {key} "
    parent.append(instr)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    parent.append(end)


def create_if_field(parent, condition_key, true_result_key):
    """
    Create a nested IF field in Word XML:
    { IF "J" = "{ MERGEFIELD <condition_key> }" "{ MERGEFIELD <true_result_key> }" }
    """
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    parent.append(begin)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f' IF "J" = "'
    parent.append(instr)

    create_merge_field(parent, condition_key)

    instr2 = OxmlElement("w:instrText")
    instr2.set(qn("xml:space"), "preserve")
    instr2.text = f'" "'
    parent.append(instr2)

    create_merge_field(parent, true_result_key)

    instr3 = OxmlElement("w:instrText")
    instr3.set(qn("xml:space"), "preserve")
    instr3.text = '"'
    parent.append(instr3)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    parent.append(end)


def create_merge_field_with_formatting(parent, key, run=None):
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    parent.append(begin)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" MERGEFIELD {key} "
    parent.append(instr)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    parent.append(end)


def create_if_field_with_formatting(parent, condition_key, true_result_key, run=None):
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    parent.append(begin)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f' IF "J" = "'
    parent.append(instr)

    create_merge_field_with_formatting(parent, condition_key, run)

    instr2 = OxmlElement("w:instrText")
    instr2.set(qn("xml:space"), "preserve")
    instr2.text = f'" "'
    parent.append(instr2)

    create_merge_field_with_formatting(parent, true_result_key, run)

    instr3 = OxmlElement("w:instrText")
    instr3.set(qn("xml:space"), "preserve")
    instr3.text = '"'
    parent.append(instr3)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    parent.append(end)
