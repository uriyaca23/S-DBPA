
# OMML XML strings for S-DBPA Report Equations
# Uses m:oMath (not oMathPara) to allow wrapping in standard paragraphs for safety.

# --- Helper for Inline Math Wrapper ---
def wrap_inline(inner_xml):
    """Wraps math XML in a standard inline math container."""
    return f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{inner_xml}</m:oMath>'

# --- Unicode/Entity Constants ---
PHI = "&#981;"   
PSI = "&#968;"   
TAU = "&#964;"   
THETA = "&#952;" 
OMEGA = "&#969;" 
EPSILON = "&#8712;" 
FORALL = "&#8704;" 
TILDE = "~"
QED_CHAR = "&#9632;" # Black square

# --- Helpers for Styles ---
def script(text):
    """Returns OMML for script (calligraphic) text."""
    return f'<m:r><m:rPr><m:scr m:val="script"/></m:rPr><m:t>{text}</m:t></m:r>'

def hat(inner_xml):
    """Returns OMML for hat accent."""
    # m:chr m:val="^" usually works for hat in OMML
    return f'<m:acc><m:accPr><m:chr m:val="^"/></m:accPr><m:e>{inner_xml}</m:e></m:acc>'

# --- Inline Math Registry ---

# We pre-compute these snippets.
P_CAL = script("P")
S_CAL = script("S")

INLINE_MATH = {
    "P": wrap_inline('<m:r><m:t>P</m:t></m:r>'),
    "P_prime": wrap_inline('<m:sSup><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sup><m:r><m:t>\'</m:t></m:r></m:sup></m:sSup>'),
    
    # Use Script P for P_sem, P_raw to match \mathcal{P}
    "P_sem": wrap_inline(f'<m:sSub><m:e>{P_CAL}</m:e><m:sub><m:r><m:t>sem</m:t></m:r></m:sub></m:sSub>'),
    "P_raw": wrap_inline(f'<m:sSub><m:e>{P_CAL}</m:e><m:sub><m:r><m:t>raw</m:t></m:r></m:sub></m:sSub>'),
    
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
    
    # omega hat for "effect size" inline
    "omega_hat_p": wrap_inline(f'<m:sSub><m:e>{hat(f"<m:r><m:t>{OMEGA}</m:t></m:r>")}</m:e><m:sub><m:r><m:t>p</m:t></m:r></m:sub></m:sSub>'),
    "omega": wrap_inline(f'<m:r><m:t>{OMEGA}</m:t></m:r>'),
    
    # QED
    "QED": wrap_inline(f'<m:r><m:t>{QED_CHAR}</m:t></m:r>')
}

# --- Block Equations ---

# 1. The Sampling Steps Block - SPLIT into 4 parts.
SAMPLING_STEPS_PARTS = [
    # Line 1
f"""<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:r><m:t>1. </m:t></m:r>
        <m:sSub><m:e>{P_CAL}</m:e><m:sub><m:r><m:t>raw</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> = {{</m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>, ..., </m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>N</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>}} {TILDE} Generator(p)</m:t></m:r>
</m:oMath>""",

    # Line 2
f"""<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:r><m:t>2. </m:t></m:r>
        <m:sSub><m:e>{P_CAL}</m:e><m:sub><m:r><m:t>sem</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> = {{x {EPSILON} </m:t></m:r>
        <m:sSub><m:e>{P_CAL}</m:e><m:sub><m:r><m:t>raw</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> | cos({PSI}(x), {PSI}(p)) > {TAU}}}</m:t></m:r>
</m:oMath>""",

    # Line 3
f"""<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:r><m:t>3. {FORALL}</m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> {EPSILON} </m:t></m:r>
        <m:sSub><m:e>{P_CAL}</m:e><m:sub><m:r><m:t>sem</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>, </m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>r</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> {TILDE} </m:t></m:r>
        <m:sSub><m:e><m:r><m:t>f</m:t></m:r></m:e><m:sub><m:r><m:t>{THETA}</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>(</m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>)</m:t></m:r>
</m:oMath>""",

    # Line 4
f"""<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:r><m:t>4. Statistic: T({{</m:t></m:r>
        <m:sSub><m:e><m:sSup><m:e><m:r><m:t>r</m:t></m:r></m:e><m:sup><m:r><m:t>'</m:t></m:r></m:sup></m:sSup></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>}}, </m:t></m:r>
        <m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>ref</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>)</m:t></m:r>
</m:oMath>"""
]

# Robustness Formula: omega hat sub S = E ...
ROBUSTNESS_OMML = f"""
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <m:sSub>
      <m:e>{hat(f"<m:r><m:t>{OMEGA}</m:t></m:r>")}</m:e>
      <m:sub>{S_CAL}</m:sub>
    </m:sSub>
    <m:r><m:t> = </m:t></m:r>
    <m:sSub>
       <m:e><m:r><m:t>E</m:t></m:r></m:e>
       <m:sub><m:r><m:t>p{TILDE}</m:t></m:r>{S_CAL}</m:sub>
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
"""

# Theorem and Proof inline text replacements
# S -> S_CAL (Script S)
THEOREM_1_PARTS = [
    wrap_inline(S_CAL), 
    wrap_inline(f'<m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>a</m:t></m:r></m:sub></m:sSub><m:r><m:t>, </m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>b</m:t></m:r></m:sub></m:sSub><m:r><m:t> {EPSILON} </m:t></m:r>{S_CAL}'), 
    wrap_inline('<m:r><m:t>P(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>a</m:t></m:r></m:sub></m:sSub><m:r><m:t>) = P(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:t>b</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>'), 
    wrap_inline('<m:sSub><m:e><m:r><m:t>H</m:t></m:r></m:e><m:sub><m:r><m:t>0</m:t></m:r></m:sub></m:sSub>'), 
    wrap_inline('<m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>ref</m:t></m:r></m:sub></m:sSub>') 
]

PROOF_PARTS = [
    wrap_inline('<m:sSub><m:e><m:r><m:t>H</m:t></m:r></m:e><m:sub><m:r><m:t>0</m:t></m:r></m:sub></m:sSub>'), 
    wrap_inline(S_CAL), 
    wrap_inline('<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t> + </m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>persona</m:t></m:r></m:sub></m:sSub>'), 
    wrap_inline(f'<m:sSub><m:e><m:r><m:t>f</m:t></m:r></m:e><m:sub><m:r><m:t>{THETA}</m:t></m:r></m:sub></m:sSub><m:r><m:t>(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t>, </m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>persona</m:t></m:r></m:sub></m:sSub><m:r><m:t>) = </m:t></m:r><m:sSub><m:e><m:r><m:t>f</m:t></m:r></m:e><m:sub><m:r><m:t>{THETA}</m:t></m:r></m:sub></m:sSub><m:r><m:t>(r|</m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>'), 
    wrap_inline('<m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>ref</m:t></m:r></m:sub></m:sSub>'), 
    wrap_inline('<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub>'), 
    wrap_inline('<m:sSub><m:e><m:r><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:t>sem</m:t></m:r></m:sub></m:sSub>'), 
    wrap_inline(f'<m:sSub><m:e><m:r><m:t>f</m:t></m:r></m:e><m:sub><m:r><m:t>{THETA}</m:t></m:r></m:sub></m:sSub><m:r><m:t>(.|</m:t></m:r><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>task</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r>') 
]
