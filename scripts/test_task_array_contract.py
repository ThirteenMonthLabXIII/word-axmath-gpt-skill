from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from math_layout_contract import _semantic_tokens, _structure_defects, _tex_structure_signature
from normalize_axmath_tex import canonical_prime_issues
from source_semantic_risks import _source_prime_orders, scan_source_math


def main():
    assert _semantic_tokens(r"20(min)") == _semantic_tokens(r"20(\min)")
    assert _semantic_tokens(r"20(min)") != _semantic_tokens(r"20(\max)")
    assert _semantic_tokens("x\u200b+y") == _semantic_tokens("x+y")
    assert _tex_structure_signature(r"e^x+e^x+e^{\frac{x}{2}}")['superscripts'] == 3
    assert _tex_structure_signature(r"e^x+e^x+e^{\frac{x}{2}}")['superscripts'] != _tex_structure_signature(r"e^x+e^x+e")["superscripts"]
    assert _tex_structure_signature(r"m_{0}^{1-\frac{k}{a}}")["subsuperscripts"] == 1
    assert _tex_structure_signature(r"m_{0}+\frac{k}{a}")["subsuperscripts"] == 0
    assert _semantic_tokens(r"9.8(\mathrm{~{}^{\circ}~C/h})") == _semantic_tokens(r"9.8(\degree\mathrm{C}/\mathrm{h})")
    assert _semantic_tokens(r"9.8(\mathrm{~{}^{\circ}~C/h})") != _semantic_tokens(r"9.8(C/h)")
    assert _semantic_tokens(r"20^{\circ}") == _semantic_tokens(r"20^{\degree}")
    assert _semantic_tokens(r"20^{\circ}") != _semantic_tokens(r"20^{\comp}")
    assert _semantic_tokens(r"20\^{\circ}") != _semantic_tokens(r"20^{\circ}")
    assert _semantic_tokens(r"sin 2x") == _semantic_tokens(r"\sin 2x")
    assert _semantic_tokens(r"\sin x") == _semantic_tokens(r"sin\!\:x")
    assert _semantic_tokens(r"\int_0^1 x\,dx") == _semantic_tokens(r"\int_0^1{x}dx")
    assert _semantic_tokens(r"x,y") != _semantic_tokens(r"xy")
    assert _semantic_tokens(r"x\%y") != _semantic_tokens(r"xy")
    assert _semantic_tokens(r"cos x") != _semantic_tokens(r"\sin x")
    assert _semantic_tokens(r"sinh x") != _semantic_tokens(r"\sin x")
    assert _semantic_tokens(r'y y^{\prime}') == _semantic_tokens(r'yy^{\prime}')
    assert _semantic_tokens(r"y''+y\prime") != _semantic_tokens('y')
    assert _semantic_tokens(r"x\prime") != _semantic_tokens('xPRIME')
    assert _semantic_tokens(r'\left\{x\right.') == _semantic_tokens(r'\{x.')
    assert _semantic_tokens(r'x\leq y') != _semantic_tokens(r'x<y')
    assert _semantic_tokens(r'\begin{matrix}x\\y\end{matrix}') == _semantic_tokens(r'\begin{array}{c}x\\y\end{array}')
    matrix = {'matrices': 1, 'equation_arrays': 0}
    assert not _structure_defects(matrix, r'\begin{array}{c}x\\y\end{array}', exact=True)
    assert _structure_defects(matrix, r'\begin{cases}x\\y\end{cases}', exact=True)
    assert _structure_defects(matrix, r'x+y', exact=True)
    assert _structure_defects({'matrices': 1, 'equation_arrays': 1}, r'\begin{array}{c}x\\y\end{array}', exact=True)
    assert _source_prime_orders('y′′′+y′′+y′') == [3, 2, 1]
    xml = (r'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           r'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><m:oMath>'
           r'<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/></m:dPr>'
           r'<m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>'
           r'<m:nary><m:naryPr/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>'
           r'</m:oMath></w:p></w:body></w:document>')
    with TemporaryDirectory() as temp:
        path = Path(temp) / 'source.docx'
        with ZipFile(path, 'w') as archive:
            archive.writestr('word/document.xml', xml)
        result = scan_source_math(path)
        assert result['math_structure_candidates'][0]['structure']['delimiters'] == [{'begin': '{', 'end': ''}]
        assert result['math_structure_candidates'][0]['structure']['operators'] == {'integral': 1}
        assert result['math_structure_candidates'][0]['structure']['subscripts'] == 0
        bounded = xml.replace(r'<m:naryPr/><m:e>', r'<m:naryPr/><m:sub><m:r><m:t>0</m:t></m:r></m:sub><m:sup><m:r><m:t>1</m:t></m:r></m:sup><m:e>')
        bounded = bounded.replace(r'<m:t>x</m:t></m:r></m:e></m:nary>', r'<m:t>x⋅y</m:t></m:r></m:e></m:nary>')
        with ZipFile(path, 'w') as archive:
            archive.writestr('word/document.xml', bounded)
        structure = scan_source_math(path)['math_structure_candidates'][0]['structure']
        assert (structure['subscripts'], structure['superscripts'], structure['subsuperscripts']) == (1, 1, 1)
        assert structure['relations'] == {'cdot': 1}
        assert not _structure_defects(structure, r'\{\int_0^1 x\cdot y', exact=True)
        assert _structure_defects(structure, r'\{\int_0 x\cdot y', exact=True)
def test_task_array_contract():
    main()


def test_high_order_prime_runs_remain_single_risks():
    assert _source_prime_orders("y′′′′+z⁗+w''''") == [4, 4, 4]
    assert "unverified_prime_order" in canonical_prime_issues("y''''")


def test_degree_mark_can_replace_superscript_circle_without_hiding_exponent_loss():
    degree_source = {'superscripts': 1, 'degree_superscripts': 1}
    assert not _structure_defects(degree_source, r'20\degree', exact=True)
    assert _structure_defects(degree_source, r'20', exact=True)
    assert _structure_defects(degree_source, r'20^x', exact=True)
    assert _structure_defects({'superscripts': 1}, r'20\degree', exact=True)
    mixed_source = {'superscripts': 2, 'degree_superscripts': 1, 'relations': {'plus': 1}}
    assert not _structure_defects(mixed_source, r'20\degree+x^2', exact=True)
    extra_exponent = {'superscripts': 1, 'degree_superscripts': 1, 'relations': {'plus': 1}}
    assert any(x['reason'] == 'superscripts_structure_count_mismatch'
               for x in _structure_defects(extra_exponent, r'20\degree+x^2', exact=True))
    tex_source = _tex_structure_signature(r'20^{\circ}')
    assert not _structure_defects(tex_source, r'20\degree', exact=True)


def test_nary_limit_visibility_respects_omml_onoff_values():
    def inspect(sub_property, sup_property):
        xml = (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><m:oMath>'
            '<m:nary><m:naryPr>' + sub_property + sup_property + '</m:naryPr>'
            '<m:sub><m:r><m:t>0</m:t></m:r></m:sub><m:sup><m:r><m:t>1</m:t></m:r></m:sup>'
            '<m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>'
            '</m:oMath></w:p></w:body></w:document>'
        )
        with TemporaryDirectory() as temp:
            path = Path(temp) / 'source.docx'
            with ZipFile(path, 'w') as archive:
                archive.writestr('word/document.xml', xml)
            return scan_source_math(path)['math_structure_candidates'][0]['structure']

    for value in ('0', 'false', 'off', 'FALSE', 'OFF'):
        row = inspect(f'<m:subHide m:val="{value}"/>', f'<m:supHide m:val="{value}"/>')
        assert (row['subscripts'], row['superscripts'], row['subsuperscripts']) == (1, 1, 1)
    for value in ('1', 'true', 'on'):
        row = inspect(f'<m:subHide m:val="{value}"/>', f'<m:supHide m:val="{value}"/>')
        assert (row['subscripts'], row['superscripts'], row['subsuperscripts']) == (0, 0, 0)
    row = inspect('<m:subHide/>', '<m:supHide m:val="off"/>')
    assert (row['subscripts'], row['superscripts'], row['subsuperscripts']) == (0, 1, 0)
    row = inspect('', '')
    assert (row['subscripts'], row['superscripts'], row['subsuperscripts']) == (1, 1, 1)


if __name__ == '__main__':
    main()
    test_high_order_prime_runs_remain_single_risks()
    test_degree_mark_can_replace_superscript_circle_without_hiding_exponent_loss()
    test_nary_limit_visibility_respects_omml_onoff_values()
    print('TASK_ARRAY_CONTRACT_OK')
