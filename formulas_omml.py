
# OMML XML strings for S-DBPA Report Equations
# Strictly uses Office Open XML Math (OMML) for all mathematical content.

# --- Helper for Inline Math Wrapper ---
def wrap_inline(inner_xml):
    """Wraps math XML in a standard inline math container."""
    return f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{inner_xml}</m:oMath>'

# --- Unicode Constants for Math ---
# Using direct unicode characters to avoid XML entity parsing issues
PHI = "\u03d5"     # phi (math variant often, or standard 03c6)
PSI = "\u03c8"     # psi
TAU = "\u03c4"     # tau
THETA = "\u03b8"   # theta
OMEGA = "\u03c9"   # omega
EPSILON = "\u2208" # element of
FORALL = "\u2200"  # for all
TILDE = "~"

# --- Inline Math Registry ---
INLINE_MATH = {
    "P": wrap_inline('<m:r><m:t>P</m:t></m:r>'),
    "P_prime": wrap_inline('<m:sSup><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sup><m:r><m:t>\'</m:t></m:r></m:sup></m:sSup>'),
    "P_sem": wrap_inline('<m:sSub><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sub><m:r><m:t>sem</m:t></m:r></m:sub></m:sSub>'),
    "P_raw": wrap_inline('<m:sSub><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sub><m:r><m:t>raw</m:t></m:r></m:sub></m:sSub>'),
    "phi": wrap_inline(f'<m:r><m:t>{PHI}</m:t></m:r>'),
    "psi": wrap_inline(f'<m:r><m:t>{PSI}</m:t></m:r>'),
    "tau": wrap_inline(f'<m:r><m:t>{TAU}</m:t></m:r>'),
    "tau_0_50": wrap_inline(f'<m:r><m:t>{TAU}=0.50</m:t></m:r>'),
    "H_0": wrap_inline('<m:sSub><m:e><m:r><m:t>H</m:t></m:r></m:e><m:sub><m:r><m:t>0</m:t></m:r></m:sub></m:sSub>'),
    "N_200": wrap_inline('<m:r><m:t>N=200</m:t></m:r>'),
    "r_prime_i": wrap_inline('<m:sSup><m:e><m:sSub><m:e><m:r><m:t>r</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub></m:e><m:sup><m:r><m:t>\'</m:t></m:r></m:sup></m:sSup>'),
    "f_theta": wrap_inline(f'<m:sSub><m:e><m:r><m:t>f</m:t></m:r></m:e><m:sub><m:r><m:t>{THETA}</m:t></m:r></m:sub></m:sSub>'),
    "p": wrap_inline('<m:r><m:t>p</m:t></m:r>'),
    "P_base": wrap_inline('<m:sSub><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sub><m:r><m:t>base</m:t></m:r></m:sub></m:sSub>'),
    "V_1": wrap_inline('<m:sSub><m:e><m:r><m:t>V</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>'),
    "V_2": wrap_inline('<m:sSub><m:e><m:r><m:t>V</m:t></m:r></m:e><m:sub><m:r><m:t>2</m:t></m:r></m:sub></m:sSub>'),
    "V_3": wrap_inline('<m:sSub><m:e><m:r><m:t>V</m:t></m:r></m:e><m:sub><m:r><m:t>3</m:t></m:r></m:sub></m:sSub>'),
    "omega": wrap_inline(f'<m:r><m:t>{OMEGA}</m:t></m:r>'),
}

# --- Block Equations ---

# 1. The Sampling Steps Block - SPLIT into 4 distinct separate paragraphs/equations.
# Using Unicode literals directly in the strings.
SAMPLING_STEPS_PARTS = [
    # Line 1: 1. P_raw = {p'_1, ..., p'_N} ~ Generator(p)
f"""<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath>
        <m:r><m:t>1. </m:t></m:r>
        <m:sSub><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sub><m:r><m:t>raw</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> = {{</m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>, ..., </m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>N</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>}} {TILDE} Generator(p)</m:t></m:r>
</m:oMath></m:oMathPara>""",

    # Line 2: 2. P_sem = {x in P_raw | cos(psi(x), psi(p)) > tau}
f"""<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath>
        <m:r><m:t>2. </m:t></m:r>
        <m:sSub><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sub><m:r><m:t>sem</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> = {{x {EPSILON} </m:t></m:r>
        <m:sSub><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sub><m:r><m:t>raw</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> | cos({PSI}(x), {PSI}(p)) > {TAU}}}</m:t></m:r>
</m:oMath></m:oMathPara>""",

    # Line 3: 3. forall p'_i in P_sem, r'_i ~ f_theta(p'_i)
f"""<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath>
        <m:r><m:t>3. {FORALL}</m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> {EPSILON} </m:t></m:r>
        <m:sSub><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sub><m:r><m:t>sem</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>, </m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>r</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> {TILDE} </m:t></m:r>
        <m:sSub><m:e><m:r><m:t>f</m:t></m:r></m:e><m:sub><m:r><m:t>{THETA}</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>(</m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>)</m:t></m:r>
</m:oMath></m:oMathPara>""",

    # Line 4: 4. Statistic: T({r'_i}, R_ref)
f"""<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath>
        <m:r><m:t>4. Statistic: T({{</m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>r</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>}}, </m:t></m:r>
        <m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>ref</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>)</m:t></m:r>
</m:oMath></m:oMathPara>"""
]

# Block: Robustness Formula
ROBUSTNESS_OMML = f"""
<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:oMath>
    <m:sSub>
      <m:e><m:r><m:t>{OMEGA}</m:t></m:r></m:e>
      <m:sub><m:r><m:t>S</m:t></m:r></m:sub>
    </m:sSub>
    <m:r><m:t> = </m:t></m:r>
    <m:sSub>
       <m:e><m:r><m:t>E</m:t></m:r></m:e>
       <m:sub><m:r><m:t>p{TILDE}S</m:t></m:r></m:sub>
    </m:sSub>
    <m:d>
      <m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>
      <m:e>
        <m:r><m:t> E </m:t></m:r>
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

# Theorem and Proof inline text replacements
THEOREM_1_PARTS = [
    wrap_inline('<m:r><m:t>S</m:t></m:r>'), # S
    wrap_inline(f'<m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>a</m:t></m:r></m:sub></m:sSub><m:r><m:t>, </m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>b</m:t></m:r></m:sub></m:sSub><m:r><m:t> {EPSILON} S</m:t></m:r>'), # p_a, p_b in S
    wrap_inline('<m:r><m:t>P(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>a</m:t></m:r></m:sub></m:sSub><m:r><m:t>) = P(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>b</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>'), # P(r|pa) = P(r|pb)
    wrap_inline('<m:sSub><m:e><m:r><m:t>H</m:t></m:r></m:e><m:sub><m:r><m:t>0</m:t></m:r></m:sub></m:sSub>'), # H_0 (cond)
    wrap_inline('<m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>ref</m:t></m:r></m:sub></m:sSub>') # R_ref
]

PROOF_PARTS = [
    wrap_inline('<m:sSub><m:e><m:r><m:t>H</m:t></m:r></m:e><m:sub><m:r><m:t>0</m:t></m:r></m:sub></m:sSub>'), # H_0
    wrap_inline('<m:r><m:t>S</m:t></m:r>'), # S
    wrap_inline('<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t> + </m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>persona</m:t></m:r></m:sub></m:sSub>'), # x_task + x_persona
    wrap_inline(f'<m:sSub><m:e><m:r><m:t>f</m:t></m:r></m:e><m:sub><m:r><m:t>{THETA}</m:t></m:r></m:sub></m:sSub><m:r><m:t>(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t>, </m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>persona</m:t></m:r></m:sub></m:sSub><m:r><m:t>) = </m:t></m:r><m:sSub><m:e><m:r><m:t>f</m:t></m:r></m:e><m:sub><m:r><m:t>{THETA}</m:t></m:r></m:sub></m:sSub><m:r><m:t>(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>'), # f_theta(...) = ...
    wrap_inline('<m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>ref</m:t></m:r></m:sub></m:sSub>'), # R_ref
    wrap_inline('<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub>'), # x_task
    wrap_inline('<m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>sem</m:t></m:r></m:sub></m:sSub>'), # R_sem
    wrap_inline(f'<m:sSub><m:e><m:r><m:t>f</m:t></m:r></m:e><m:sub><m:r><m:t>{THETA}</m:t></m:r></m:sub></m:sSub><m:r><m:t>(.|</m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>') # f_theta(.|x_task)
]
