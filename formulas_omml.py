
# OMML XML strings for S-DBPA Report Equations
# Uses strict Office Open XML Math structure.
# Reference: http://www.datypic.com/sc/ooxml/e-m_oMath-1.html

# Namespace map for Reference
# xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
# xmlns:w="http://schemas.openxmlformats.org/officeDocument/2006/wordprocessingml"

# 1. The Sampling Steps Block
# We use an Equation Array (eqArr) inside oMathPara.
# This ensures lines are stacked.
SAMPLING_STEPS_OMML = """
<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/officeDocument/2006/wordprocessingml">
  <m:oMath>
    <m:eqArr>
      <m:e>
        <!-- Line 1: P_raw = {p'_1, ..., p'_N} ~ Generator(p) -->
        <m:r><m:t>1. P</m:t></m:r>
        <m:sSub>
          <m:e><m:r><m:t>raw</m:t></m:r></m:e>
          <m:sub><m:phantom/></m:sub>
        </m:sSub>
        <m:r><m:t>= {p</m:t></m:r>
        <m:sSup><m:e><m:r><m:t>'</m:t></m:r></m:e><m:sup><m:r><m:t>1</m:t></m:r></m:sup></m:sSup>
        <m:r><m:t>, ..., p</m:t></m:r>
        <m:sSup><m:e><m:r><m:t>'</m:t></m:r></m:e><m:sup><m:r><m:t>N</m:t></m:r></m:sup></m:sSup>
        <m:r><m:t>} ~ Generator(p)</m:t></m:r>
        
        <!-- Line 2: P_sem = {x in P_raw | cos(psi(x), psi(p)) > tau} -->
        <m:r><m:br m:type="texting"/></m:r> <!-- Line break -->
        <m:r><m:t>2. P</m:t></m:r>
        <m:sSub>
          <m:e><m:r><m:t>sem</m:t></m:r></m:e>
          <m:sub><m:phantom/></m:sub>
        </m:sSub>
        <m:r><m:t>= {x &#8712; P</m:t></m:r>
        <m:sSub>
          <m:e><m:r><m:t>raw</m:t></m:r></m:e>
          <m:sub><m:phantom/></m:sub>
        </m:sSub>
        <m:r><m:t>| cos(&#968;(x), &#968;(p)) > &#964;}</m:t></m:r>
        
        <!-- Line 3: forall p'_i in P_sem, r'_i ~ f_theta(p'_i) -->
        <m:r><m:br m:type="texting"/></m:r>
        <m:r><m:t>3. &#8704; p</m:t></m:r>
        <m:sSup><m:e><m:r><m:t>'</m:t></m:r></m:e><m:sup><m:r><m:t>i</m:t></m:r></m:sup></m:sSup>
        <m:r><m:t>&#8712; P</m:t></m:r>
        <m:sSub>
          <m:e><m:r><m:t>sem</m:t></m:r></m:e>
          <m:sub><m:phantom/></m:sub>
        </m:sSub>
        <m:r><m:t>, r</m:t></m:r>
        <m:sSup><m:e><m:r><m:t>'</m:t></m:r></m:e><m:sup><m:r><m:t>i</m:t></m:r></m:sup></m:sSup>
        <m:r><m:t>~ f</m:t></m:r>
        <m:sSub>
          <m:e><m:r><m:t>&#952;</m:t></m:r></m:e>
          <m:sub><m:phantom/></m:sub>
        </m:sSub>
        <m:r><m:t>(p</m:t></m:r>
        <m:sSup><m:e><m:r><m:t>'</m:t></m:r></m:e><m:sup><m:r><m:t>i</m:t></m:r></m:sup></m:sSup>
        <m:r><m:t>)</m:t></m:r>
        
        <!-- Line 4: Statistic: T({r'_i}, R_ref) -->
        <m:r><m:br m:type="texting"/></m:r>
        <m:r><m:t>4. Statistic: T({r</m:t></m:r>
        <m:sSup><m:e><m:r><m:t>'</m:t></m:r></m:e><m:sup><m:r><m:t>i</m:t></m:r></m:sup></m:sSup>
        <m:r><m:t>}, R</m:t></m:r>
        <m:sSub>
          <m:e><m:r><m:t>ref</m:t></m:r></m:e>
          <m:sub><m:phantom/></m:sub>
        </m:sSub>
        <m:r><m:t>)</m:t></m:r>
      </m:e>
    </m:eqArr>
  </m:oMath>
</m:oMathPara>
"""

# Theorem 1 Inline Math Definitions
# Using simpler oMath containers for inline insertion.
THEOREM_1_OMML = [
    # S (script S) - using standard S for simplicity to avoid font issues, or unicode &#119982;
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>S</m:t></m:r></m:oMath>""",
    # p_a, p_b in S
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>a</m:t></m:r></m:sub></m:sSub><m:r><m:t>, </m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>b</m:t></m:r></m:sub></m:sSub><m:r><m:t> &#8712; S</m:t></m:r></m:oMath>""",
    # P(r|p_a) = P(r|p_b)
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>P(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>a</m:t></m:r></m:sub></m:sSub><m:r><m:t>) = P(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>b</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r></m:oMath>""",
    # H_0
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSub><m:e><m:r><m:t>H</m:t></m:r></m:e><m:sub><m:r><m:t>0</m:t></m:r></m:sub></m:sSub></m:oMath>""",
    # R_ref
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>ref</m:t></m:r></m:sub></m:sSub></m:oMath>"""
]

# Proof Math Definitions
PROOF_OMML = [
    # 0: H_0
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSub><m:e><m:r><m:t>H</m:t></m:r></m:e><m:sub><m:r><m:t>0</m:t></m:r></m:sub></m:sSub></m:oMath>""",
    # 1: S
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>S</m:t></m:r></m:oMath>""",
    # 2: x_task + x_persona
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t> + </m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>persona</m:t></m:r></m:sub></m:sSub></m:oMath>""",
    # 3: f_theta(r|x_task, x_persona) = f_theta(r|x_task)
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>f</m:t></m:r><m:sSub><m:e><m:r><m:t>&#952;</m:t></m:r></m:e><m:sub><m:phantom/></m:sub></m:sSub><m:r><m:t>(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t>, </m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>persona</m:t></m:r></m:sub></m:sSub><m:r><m:t>) = f</m:t></m:r><m:sSub><m:e><m:r><m:t>&#952;</m:t></m:r></m:e><m:sub><m:phantom/></m:sub></m:sSub><m:r><m:t>(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r></m:oMath>""",
    # 4: R_ref
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>ref</m:t></m:r></m:sub></m:sSub></m:oMath>""",
    # 5: x_task
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub></m:oMath>""",
    # 6: R_sem
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>sem</m:t></m:r></m:sub></m:sSub></m:oMath>""",
    # 7: f_theta(.|x_task)
    """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>f</m:t></m:r><m:sSub><m:e><m:r><m:t>&#952;</m:t></m:r></m:e><m:sub><m:phantom/></m:sub></m:sSub><m:r><m:t>(.|</m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r></m:oMath>"""
]

# Robustness Formula (Centered)
# omega_S = E_{p ~ S} [ E[D(r_p, r_ref)] ]
ROBUSTNESS_OMML = """
<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/officeDocument/2006/wordprocessingml">
  <m:oMath>
    <m:sSub>
      <m:e><m:r><m:t>&#969;</m:t></m:r></m:e>
      <m:sub><m:r><m:t>S</m:t></m:r></m:sub>
    </m:sSub>
    <m:r><m:t> = </m:t></m:r>
    <m:sSub>
      <m:e><m:r><m:t>E</m:t></m:r></m:e>
      <m:sub><m:r><m:t>p~S</m:t></m:r></m:sub>
    </m:sSub>
    <m:d>
      <m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>
      <m:e>
        <m:r><m:t>E</m:t></m:r>
        <m:d>
          <m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>
          <m:e>
            <m:r><m:t>D(</m:t></m:r>
            <m:sSub><m:e><m:r><m:t>r</m:t></m:r></m:e><m:sub><m:r><m:t>p</m:t></m:r></m:sub></m:sSub>
            <m:r><m:t>, </m:t></m:r>
            <m:sSub><m:e><m:r><m:t>r</m:t></m:r></m:e><m:sub><m:r><m:t>ref</m:t></m:r></m:sub></m:sSub>
            <m:r><m:t>)</m:t></m:r>
          </m:e>
        </m:d>
      </m:e>
    </m:d>
  </m:oMath>
</m:oMathPara>
"""
