import os
import sys
import types

# Stub out external dependencies (python-docx and msoffcrypto) when not installed
try:
    import docx
except ImportError:
    docx = types.ModuleType('docx')
    def Document(*args, **kwargs):
        raise ImportError("python-docx is not installed")
    docx.Document = Document
    sys.modules['docx'] = docx
# Stub submodule for shared measurements
if 'docx.shared' not in sys.modules:
    shared = types.ModuleType('docx.shared')
    class _DummyMeasure:
        def __init__(self, *args, **kwargs): pass
    # Stub measurement and color classes
    shared.Inches = _DummyMeasure
    shared.Pt = _DummyMeasure
    shared.RGBColor = _DummyMeasure
    sys.modules['docx.shared'] = shared
try:
    import msoffcrypto
except ImportError:
    msoffcrypto = types.ModuleType('msoffcrypto')
    sys.modules['msoffcrypto'] = msoffcrypto
# Stub docx.enum.style for style-related tooling
if 'docx.enum.style' not in sys.modules:
    enum_pkg = types.ModuleType('docx.enum')
    style_mod = types.ModuleType('docx.enum.style')
    class _DummyStyleEnum:
        PARAGRAPH = 1
    style_mod.WD_STYLE_TYPE = _DummyStyleEnum
    sys.modules['docx.enum'] = enum_pkg
    sys.modules['docx.enum.style'] = style_mod
# Stub docx.enum.text for color formatting
# Stub docx.enum.text for color and alignment formatting
if 'docx.enum.text' not in sys.modules:
    text_enum = types.ModuleType('docx.enum.text')
    class _DummyColorIndex:
        pass
    class _DummyAlign:
        LEFT = 0
        CENTER = 1
        RIGHT = 2
    text_enum.WD_COLOR_INDEX = _DummyColorIndex
    text_enum.WD_ALIGN_PARAGRAPH = _DummyAlign
    sys.modules['docx.enum.text'] = text_enum
# Stub docx.enum.section for section break types
if 'docx.enum.section' not in sys.modules:
    # Ensure 'docx.enum' package exists
    enum_pkg = sys.modules.get('docx.enum') or types.ModuleType('docx.enum')
    sys.modules['docx.enum'] = enum_pkg
    section_mod = types.ModuleType('docx.enum.section')
    class _DummySectionStart:
        NEXT_PAGE = 0
        EVEN_PAGE = 1
        ODD_PAGE = 2
    section_mod.WD_SECTION_START = _DummySectionStart
    sys.modules['docx.enum.section'] = section_mod
# Stub docx.oxml.shared for table manipulation
if 'docx.oxml.shared' not in sys.modules:
    oxml_pkg = types.ModuleType('docx.oxml')
    shared_mod = types.ModuleType('docx.oxml.shared')
    def OxmlElement(*args, **kwargs): return None
    def qn(tag): return tag
    shared_mod.OxmlElement = OxmlElement
    shared_mod.qn = qn
    # Provide parse_xml stub on docx.oxml, and stub OxmlElement
    oxml_pkg.parse_xml = lambda xml: None
    oxml_pkg.OxmlElement = OxmlElement
    sys.modules['docx.oxml'] = oxml_pkg
    sys.modules['docx.oxml.shared'] = shared_mod
# Stub docx.oxml.ns for namespace declarations
if 'docx.oxml.ns' not in sys.modules:
    ns_mod = types.ModuleType('docx.oxml.ns')
    def nsdecls(prefix): return ''
    ns_mod.nsdecls = nsdecls
    sys.modules['docx.oxml.ns'] = ns_mod

# Add project root to sys.path so tests can import the package modules without installation
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)