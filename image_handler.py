import os
import json
import base64
import logging
import re
import threading
import hashlib
import shutil
import urllib.parse
import requests
try:
    import schemdraw
    import schemdraw.elements as elm
    SCHEMDRAW_AVAILABLE = True
except ImportError:
    SCHEMDRAW_AVAILABLE = False
    schemdraw = None
    elm = None

try:
    import docx
    from docx.shared import Inches, Pt, RGBColor, Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
    from docx.enum.table import WD_TABLE_ALIGNMENT
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    docx = None

try:
    from rdkit import Chem
    from rdkit.Chem import Draw
    from rdkit.Chem import rdDepictor
    RDKIT_AVAILABLE = True
except ImportError:
    RDKIT_AVAILABLE = False
    Chem = None
    Draw = None
    rdDepictor = None

logger = logging.getLogger(__name__)
_index_lock = threading.Lock()
_cache_lock = threading.Lock()

# ── NCERT COMMON CHEMISTRY COMPOUNDS LOOKUP DICTIONARY ────────────────────────
NCERT_COMPOUND_SMILES = {
    # Alkanes, Alkenes, Alkynes, Cycloalkanes
    "methane": "C",
    "ethane": "CC",
    "propane": "CCC",
    "butane": "CCCC",
    "n-butane": "CCCC",
    "isobutane": "CC(C)C",
    "2-methylpropane": "CC(C)C",
    "pentane": "CCCCC",
    "n-pentane": "CCCCC",
    "isopentane": "CC(C)CC",
    "neopentane": "CC(C)(C)C",
    "hexane": "CCCCCC",
    "ethene": "C=C",
    "ethylene": "C=C",
    "propene": "CC=C",
    "propylene": "CC=C",
    "but-1-ene": "CCC=C",
    "1-butene": "CCC=C",
    "but-2-ene": "CC=CC",
    "2-butene": "CC=CC",
    "ethyne": "C#C",
    "acetylene": "C#C",
    "propyne": "CC#C",
    "but-1-yne": "CCC#C",
    "but-2-yne": "CC#CC",
    "cyclohexane": "C1CCCCC1",
    "cyclopentane": "C1CCCC1",
    "benzene": "c1ccccc1",
    "toluene": "Cc1ccccc1",
    "methylbenzene": "Cc1ccccc1",
    "naphthalene": "c1ccc2ccccc2c1",
    "anthracene": "c1ccc2cc3ccccc3cc2c1",
    "styrene": "C=Cc1ccccc1",
    "ethylbenzene": "CCc1ccccc1",
    "o-xylene": "Cc1ccccc1C",
    "m-xylene": "Cc1cccc(C)c1",
    "p-xylene": "Cc1ccc(C)cc1",

    # Alcohols, Phenols, Ethers
    "methanol": "CO",
    "methyl alcohol": "CO",
    "ethanol": "CCO",
    "ethyl alcohol": "CCO",
    "propan-1-ol": "CCCO",
    "1-propanol": "CCCO",
    "propan-2-ol": "CC(C)O",
    "2-propanol": "CC(C)O",
    "isopropyl alcohol": "CC(C)O",
    "isopropanol": "CC(C)O",
    "butan-1-ol": "CCCCO",
    "1-butanol": "CCCCO",
    "butan-2-ol": "CCC(C)O",
    "2-butanol": "CCC(C)O",
    "2-methylpropan-1-ol": "CC(C)CO",
    "isobutyl alcohol": "CC(C)CO",
    "2-methylpropan-2-ol": "CC(C)(C)O",
    "tert-butyl alcohol": "CC(C)(C)O",
    "ethylene glycol": "OCCO",
    "ethane-1,2-diol": "OCCO",
    "glycerol": "OCC(O)CO",
    "glycerine": "OCC(O)CO",
    "propane-1,2,3-triol": "OCC(O)CO",
    "phenol": "Oc1ccccc1",
    "carbolic acid": "Oc1ccccc1",
    "o-cresol": "Cc1ccccc1O",
    "2-methylphenol": "Cc1ccccc1O",
    "m-cresol": "Cc1cccc(O)c1",
    "3-methylphenol": "Cc1cccc(O)c1",
    "p-cresol": "Cc1ccc(O)cc1",
    "4-methylphenol": "Cc1ccc(O)cc1",
    "catechol": "Oc1ccccc1O",
    "resorcinol": "Oc1cccc(O)c1",
    "hydroquinone": "Oc1ccc(O)cc1",
    "quinol": "Oc1ccc(O)cc1",
    "picric acid": "Oc1c([N+](=O)[O-])cc([N+](=O)[O-])cc1[N+](=O)[O-]",
    "2,4,6-trinitrophenol": "Oc1c([N+](=O)[O-])cc([N+](=O)[O-])cc1[N+](=O)[O-]",
    "dimethyl ether": "COC",
    "methoxymethane": "COC",
    "diethyl ether": "CCOCC",
    "ethoxyethane": "CCOCC",
    "anisole": "COc1ccccc1",
    "methoxybenzene": "COc1ccccc1",
    "phenetole": "CCOc1ccccc1",
    "ethoxybenzene": "CCOc1ccccc1",

    # Aldehydes & Ketones
    "formaldehyde": "C=O",
    "methanal": "C=O",
    "acetaldehyde": "CC=O",
    "ethanal": "CC=O",
    "propionaldehyde": "CCC=O",
    "propanal": "CCC=O",
    "butyraldehyde": "CCCC=O",
    "butanal": "CCCC=O",
    "benzaldehyde": "O=Cc1ccccc1",
    "cinnamaldehyde": "O=C/C=C/c1ccccc1",
    "salicylaldehyde": "O=Cc1ccccc1O",
    "2-hydroxybenzaldehyde": "O=Cc1ccccc1O",
    "vanillin": "O=Cc1ccc(O)c(OC)c1",
    "acetone": "CC(=O)C",
    "propanone": "CC(=O)C",
    "butan-2-one": "CCC(=O)C",
    "butanone": "CCC(=O)C",
    "ethyl methyl ketone": "CCC(=O)C",
    "pentan-2-one": "CCCC(=O)C",
    "pentan-3-one": "CCC(=O)CC",
    "acetophenone": "CC(=O)c1ccccc1",
    "benzophenone": "O=C(c1ccccc1)c2ccccc2",
    "cyclohexanone": "O=C1CCCCC1",

    # Carboxylic Acids, Esters, Anhydrides & Acyl Halides
    "formic acid": "C(=O)O",
    "methanoic acid": "C(=O)O",
    "acetic acid": "CC(=O)O",
    "ethanoic acid": "CC(=O)O",
    "propanoic acid": "CCC(=O)O",
    "propionic acid": "CCC(=O)O",
    "butanoic acid": "CCCC(=O)O",
    "butyric acid": "CCCC(=O)O",
    "benzoic acid": "O=C(O)c1ccccc1",
    "oxalic acid": "O=C(O)C(=O)O",
    "ethanedioic acid": "O=C(O)C(=O)O",
    "malonic acid": "O=C(O)CC(=O)O",
    "succinic acid": "O=C(O)CCC(=O)O",
    "phthalic acid": "O=C(O)c1ccccc1C(=O)O",
    "terephthalic acid": "O=C(O)c1ccc(cc1)C(=O)O",
    "lactic acid": "CC(O)C(=O)O",
    "2-hydroxypropanoic acid": "CC(O)C(=O)O",
    "salicylic acid": "O=C(O)c1ccccc1O",
    "2-hydroxybenzoic acid": "O=C(O)c1ccccc1O",
    "acetylsalicylic acid": "CC(=O)Oc1ccccc1C(=O)O",
    "aspirin": "CC(=O)Oc1ccccc1C(=O)O",
    "methyl formate": "COC=O",
    "methyl methanoate": "COC=O",
    "ethyl formate": "CCOC=O",
    "ethyl methanoate": "CCOC=O",
    "methyl acetate": "COC(=O)C",
    "methyl ethanoate": "COC(=O)C",
    "ethyl acetate": "CCOC(=O)C",
    "ethyl ethanoate": "CCOC(=O)C",
    "phenyl acetate": "CC(=O)Oc1ccccc1",
    "acetic anhydride": "CC(=O)OC(=O)C",
    "ethanoic anhydride": "CC(=O)OC(=O)C",
    "phthalic anhydride": "O=C1OC(=O)c2ccccc12",
    "acetyl chloride": "CC(=O)Cl",
    "ethanoyl chloride": "CC(=O)Cl",
    "benzoyl chloride": "O=C(Cl)c1ccccc1",

    # Nitrogen Compounds (Amines, Amides, Cyanides, Nitro)
    "methylamine": "CN",
    "methanamine": "CN",
    "dimethylamine": "CNC",
    "trimethylamine": "CN(C)C",
    "ethylamine": "CCN",
    "ethanamine": "CCN",
    "diethylamine": "CCNCC",
    "triethylamine": "CCN(CC)CC",
    "aniline": "Nc1ccccc1",
    "benzenamine": "Nc1ccccc1",
    "n-methylaniline": "CNc1ccccc1",
    "n,n-dimethylaniline": "CN(C)c1ccccc1",
    "formamide": "NC=O",
    "methanamide": "NC=O",
    "acetamide": "CC(N)=O",
    "ethanamide": "CC(N)=O",
    "benzamide": "NC(=O)c1ccccc1",
    "acetanilide": "CC(=O)Nc1ccccc1",
    "urea": "NC(N)=O",
    "carbamide": "NC(N)=O",
    "hydrogen cyanide": "C#N",
    "acetonitrile": "CC#N",
    "methyl cyanide": "CC#N",
    "benzonitrile": "N#Cc1ccccc1",
    "phenyl cyanide": "N#Cc1ccccc1",
    "nitromethane": "C[N+](=O)[O-]",
    "nitroethane": "CC[N+](=O)[O-]",
    "nitrobenzene": "[O-][N+](=O)c1ccccc1",
    "diazomethane": "C=[N+]=[N-]",
    "benzene diazonium chloride": "[Cl-].[N+]#Nc1ccccc1",

    # Biomolecules & Pharmaceuticals
    "glucose": "OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O",
    "d-glucose": "OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O",
    "fructose": "OCC1(O)OC(CO)[C@@H](O)[C@@H]1O",
    "d-fructose": "OCC1(O)OC(CO)[C@@H](O)[C@@H]1O",
    "sucrose": "OC[C@H]1O[C@@](CO)(O[C@H]2O[C@H](CO)[C@@H](O)[C@H](O)[C@H]2O)[C@@H](O)[C@@H]1O",
    "maltose": "OC[C@H]1O[C@H](O[C@H]2[C@H](O)[C@@H](O)C(O)O[C@@H]2CO)[C@H](O)[C@@H](O)[C@@H]1O",
    "glycine": "NCC(=O)O",
    "alanine": "CC(N)C(=O)O",
    "valine": "CC(C)C(N)C(=O)O",
    "paracetamol": "CC(=O)Nc1ccc(O)cc1",
    "acetaminophen": "CC(=O)Nc1ccc(O)cc1",
    "ibuprofen": "CC(C)Cc1ccc(cc1)C(C)C(=O)O",
    "chloroquine": "CCN(CC)CCCC(C)Nc1ccnc2cc(Cl)ccc12",

    # Haloalkanes & Haloarenes
    "chloromethane": "CCl",
    "methyl chloride": "CCl",
    "bromomethane": "CBr",
    "methyl bromide": "CBr",
    "iodomethane": "CI",
    "methyl iodide": "CI",
    "chloroethane": "CCCl",
    "ethyl chloride": "CCCl",
    "bromoethane": "CCBr",
    "ethyl bromide": "CCBr",
    "iodoethane": "CCI",
    "ethyl iodide": "CCI",
    "dichloromethane": "ClCCl",
    "methylene chloride": "ClCCl",
    "chloroform": "ClC(Cl)Cl",
    "trichloromethane": "ClC(Cl)Cl",
    "bromoform": "BrC(Br)Br",
    "iodoform": "IC(I)I",
    "carbon tetrachloride": "ClC(Cl)(Cl)Cl",
    "tetrachloromethane": "ClC(Cl)(Cl)Cl",
    "chlorobenzene": "Clc1ccccc1",
    "bromobenzene": "Brc1ccccc1",
    "iodobenzene": "Ic1ccccc1",
    "benzyl chloride": "ClCc1ccccc1",
    "freon-12": "FC(F)(Cl)Cl",
    "dichlorodifluoromethane": "FC(F)(Cl)Cl",
    "ddt": "Clc1ccc(cc1)C(c2ccc(Cl)cc2)C(Cl)(Cl)Cl"
}


# ── LOCAL CACHE UTILITIES ─────────────────────────────────────────────────────
def _get_cache_dir():
    """Return absolute path to local cache directory for images."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cache_dir = os.path.join(script_dir, 'static', 'cache', 'images')
    os.makedirs(cache_dir, exist_ok=True)
    return cache_dir


def _get_cached_image(cache_key, prefix='img', ext='png'):
    """
    Check if a cached image exists for the given key string (hashed).
    Returns absolute file path if file exists and has content, else None.
    """
    if not cache_key:
        return None
    try:
        cache_dir = _get_cache_dir()
        norm_key = str(cache_key).strip().lower()
        key_hash = hashlib.md5(norm_key.encode('utf-8')).hexdigest()
        cache_filename = f"{prefix}_{key_hash}.{ext}"
        cache_path = os.path.join(cache_dir, cache_filename)
        if os.path.exists(cache_path) and os.path.getsize(cache_path) > 0:
            return cache_path
    except Exception as e:
        logger.debug(f"Cache check error for key '{cache_key}': {e}")
    return None


def _save_to_cache(cache_key, source_path_or_bytes, prefix='img', ext='png'):
    """
    Save file content or copy file to the local cache directory using key hash.
    Returns absolute cached path or None on error.
    """
    if not cache_key or not source_path_or_bytes:
        return None
    try:
        cache_dir = _get_cache_dir()
        norm_key = str(cache_key).strip().lower()
        key_hash = hashlib.md5(norm_key.encode('utf-8')).hexdigest()
        cache_filename = f"{prefix}_{key_hash}.{ext}"
        cache_path = os.path.join(cache_dir, cache_filename)

        with _cache_lock:
            if isinstance(source_path_or_bytes, (bytes, bytearray)):
                with open(cache_path, 'wb') as f:
                    f.write(source_path_or_bytes)
            elif isinstance(source_path_or_bytes, str) and os.path.exists(source_path_or_bytes):
                if os.path.abspath(source_path_or_bytes) != os.path.abspath(cache_path):
                    shutil.copyfile(source_path_or_bytes, cache_path)
            return cache_path
    except Exception as e:
        logger.warning(f"Error saving to cache for '{cache_key}': {e}")
        return None


# ── LICENSE VALIDATOR HELPER ──────────────────────────────────────────────────
def _is_allowed_wikimedia_license(extmetadata):
    """
    Validate if image license from Wikimedia Commons is CC-BY, CC-BY-SA, CC0, or Public Domain.
    Returns True if allowed, False otherwise.
    """
    if not isinstance(extmetadata, dict):
        return False

    lic_short = str(extmetadata.get('LicenseShortName', {}).get('value', '')).strip().lower()
    lic = str(extmetadata.get('License', {}).get('value', '')).strip().lower()
    usage = str(extmetadata.get('UsageTerms', {}).get('value', '')).strip().lower()
    combined = f"{lic_short} {lic} {usage}"

    # Explicitly reject non-free or restricted terms
    restricted_terms = ['all rights reserved', 'fair use', 'non-free', 'copyrighted free use with restriction']
    if any(t in combined for t in restricted_terms):
        return False

    # Allowed licenses: CC-BY, CC-BY-SA, Public Domain, CC0
    allowed_terms = [
        'cc by', 'cc-by', 'cc_by',
        'cc by-sa', 'cc-by-sa', 'cc_by_sa',
        'public domain', 'pd', 'pd-self', 'pd-user', 'pd-author', 'pd-old',
        'cc0', 'cc-0', 'cc zero', 'no rights reserved'
    ]
    return any(term in combined for term in allowed_terms)


# ── SCHEMATIC CIRCUITS ────────────────────────────────────────────────────────
def _format_unit(val, default_unit):
    """Format numeric values with appropriate unit suffix if not already present."""
    if val is None:
        return ""
    val_str = str(val).strip()
    if not val_str:
        return ""
    known_units = ['v', 'volt', 'volts', 'ω', 'ohm', 'ohms', 'kω', 'kohm', 'mω', 'a', 'ma', 'cm', 'm', 'mm']
    lower_val = val_str.lower()
    if any(u in lower_val for u in known_units):
        return val_str
    return f"{val_str} {default_unit}"


def generate_circuit(circuit_type, params=None, output_path="circuit.png"):
    """
    Generate circuit diagram using schemdraw and save as PNG.
    
    Supported circuit_type:
      - 'series_resistors'
      - 'parallel_resistors'
      - 'potentiometer'
      - 'wheatstone_bridge'
      
    params: dict with voltage, resistor values, etc.
    output_path: file path to save the generated PNG.
    """
    if params is None:
        params = {}

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    c_type = str(circuit_type or '').strip().lower().replace('-', '_').replace(' ', '_')

    # 1. SERIES RESISTORS
    if c_type == 'series_resistors':
        v = _format_unit(params.get('voltage', params.get('v', '12V')), 'V') or '12V'
        resistors = params.get('resistors') or [
            params.get('r1', '10Ω'),
            params.get('r2', '20Ω'),
            params.get('r3', params.get('r3_val', ''))
        ]
        res_list = [_format_unit(r, 'Ω') for r in resistors if str(r).strip()]
        if not res_list:
            res_list = ['10 Ω', '20 Ω']

        with schemdraw.Drawing(file=output_path, show=False) as d:
            d.config(fontsize=11)
            batt = d.add(elm.Battery().up().label(v))
            d.add(elm.Line().right(d.unit * 0.5))
            for i, r_val in enumerate(res_list):
                lbl = f"R{i+1} = {r_val}" if not str(r_val).startswith('R') else r_val
                d.add(elm.Resistor().right().label(lbl, loc='top'))
            d.add(elm.Line().down(d.unit))
            d.add(elm.Line().left().to(batt.start))

    # 2. PARALLEL RESISTORS
    elif c_type == 'parallel_resistors':
        v = _format_unit(params.get('voltage', params.get('v', '12V')), 'V') or '12V'
        resistors = params.get('resistors') or [
            params.get('r1', '10Ω'),
            params.get('r2', '20Ω'),
            params.get('r3', params.get('r3_val', ''))
        ]
        res_list = [_format_unit(r, 'Ω') for r in resistors if str(r).strip()]
        if not res_list:
            res_list = ['10 Ω', '20 Ω']

        with schemdraw.Drawing(file=output_path, show=False) as d:
            d.config(fontsize=11)
            batt = d.add(elm.Battery().up().label(v))
            d.add(elm.Line().right(d.unit * 0.8))
            
            top_start = d.here
            
            # Top Branch
            lbl1 = f"R1 = {res_list[0]}" if not str(res_list[0]).startswith('R') else res_list[0]
            d.add(elm.Line().up(d.unit * 0.6))
            d.add(elm.Resistor().right().label(lbl1, loc='top'))
            d.add(elm.Line().down(d.unit * 0.6))
            top_end = d.here

            # Bottom Branch
            d.add(elm.Line().at(top_start).down(d.unit * 0.6))
            lbl2 = f"R2 = {res_list[1]}" if len(res_list) > 1 and not str(res_list[1]).startswith('R') else (res_list[1] if len(res_list) > 1 else '20 Ω')
            d.add(elm.Resistor().right().label(lbl2, loc='bottom'))
            d.add(elm.Line().up(d.unit * 0.6))

            # Middle Branch (if 3 or more resistors)
            if len(res_list) >= 3:
                lbl3 = f"R3 = {res_list[2]}" if not str(res_list[2]).startswith('R') else res_list[2]
                d.add(elm.Resistor().at(top_start).right().label(lbl3, loc='top'))

            d.add(elm.Line().at(top_end).right(d.unit * 0.5))
            d.add(elm.Line().down(d.unit))
            d.add(elm.Line().left().to(batt.start))

    # 3. POTENTIOMETER
    elif c_type == 'potentiometer':
        v_main = _format_unit(params.get('voltage', params.get('v_driver', '6V')), 'V') or '6V'
        v_cell = _format_unit(params.get('v_cell', params.get('cell_voltage', '1.5V')), 'V') or '1.5V'
        r_wire = _format_unit(params.get('r_wire', params.get('wire_resistance', '10Ω')), 'Ω') or '10Ω'
        length = _format_unit(params.get('length', params.get('wire_length', '100 cm')), 'cm') or '100 cm'

        with schemdraw.Drawing(file=output_path, show=False) as d:
            d.config(fontsize=11)
            # Primary / Driver circuit
            d.add(elm.Battery().up().label(f"Driver ({v_main})"))
            d.add(elm.Switch().right().label('K1'))
            d.add(elm.RBox().right().label('Rheostat (Rh)'))
            d.add(elm.Line().down(d.unit * 0.8))
            
            # Potentiometer wire AB
            wire_start = d.here
            d.add(elm.Dot().label('A', loc='left'))
            d.add(elm.Line().left(d.unit * 2.5).label(f"Wire AB ({length}, {r_wire})", loc='top'))
            d.add(elm.Dot().label('B', loc='right'))
            d.add(elm.Line().up(d.unit * 0.8))

            # Secondary circuit with Test Cell & Galvanometer
            d.add(elm.Line().at(wire_start).down(d.unit * 0.9))
            d.add(elm.Battery().left().label(f"Cell ({v_cell})", loc='bottom'))
            d.add(elm.MeterV().left().label('G', loc='bottom'))
            d.add(elm.Line().left(d.unit * 0.4))
            d.add(elm.Arrow().up(d.unit * 0.9).label('Jockey (J)', loc='left'))

    # 4. WHEATSTONE BRIDGE
    elif c_type == 'wheatstone_bridge':
        v = _format_unit(params.get('voltage', params.get('v', '10V')), 'V') or '10V'
        p_val = _format_unit(params.get('p', params.get('r1', '10Ω')), 'Ω') or '10Ω'
        q_val = _format_unit(params.get('q', params.get('r2', '20Ω')), 'Ω') or '20Ω'
        r_val = _format_unit(params.get('r', params.get('r3', '15Ω')), 'Ω') or '15Ω'
        s_val = _format_unit(params.get('s', params.get('r4', '30Ω')), 'Ω') or '30Ω'

        with schemdraw.Drawing(file=output_path, show=False) as d:
            d.config(fontsize=11)
            # Power supply loop at bottom
            batt = d.add(elm.Battery().right(d.unit * 2).label(v))
            d.add(elm.Switch().right().label('Key K'))
            d.add(elm.Line().up(d.unit * 1.5))
            node_c = d.here
            d.add(elm.Dot().label('C', loc='right'))

            # Node A on the left
            d.add(elm.Line().at(batt.start).left(d.unit * 0.5))
            d.add(elm.Line().up(d.unit * 1.5))
            node_a = d.here
            d.add(elm.Dot().label('A', loc='left'))

            # Coordinate offsets for diamond
            top_offset = (node_a[0] + (node_c[0] - node_a[0]) / 2, node_a[1] + d.unit * 1.2)
            bot_offset = (node_a[0] + (node_c[0] - node_a[0]) / 2, node_a[1] - d.unit * 1.2)

            # A -> B (P / R1)
            d.add(elm.Resistor().at(node_a).to(top_offset).label(f"P = {p_val}", loc='top'))
            d.add(elm.Dot().label('B', loc='top'))

            # B -> C (Q / R2)
            d.add(elm.Resistor().at(top_offset).to(node_c).label(f"Q = {q_val}", loc='top'))

            # A -> D (R / R3)
            d.add(elm.Resistor().at(node_a).to(bot_offset).label(f"R = {r_val}", loc='bottom'))
            d.add(elm.Dot().label('D', loc='bottom'))

            # D -> C (S / R4)
            d.add(elm.Resistor().at(bot_offset).to(node_c).label(f"S = {s_val}", loc='bottom'))

            # Galvanometer between B and D
            d.add(elm.MeterV().at(top_offset).to(bot_offset).label('G', loc='right'))

    else:
        raise ValueError(f"Unsupported circuit_type: '{circuit_type}'. Supported types: 'series_resistors', 'parallel_resistors', 'potentiometer', 'wheatstone_bridge'.")

    return output_path


# ── IMAGE BANK INDEX LOOKUP & SAVE ────────────────────────────────────────────
def find_from_bank(keywords, index_path="image_bank_index.json"):
    """
    Search for a pre-tagged NCERT diagram from the JSON index file matching given keywords.
    
    Index format: [{"file": "path.png", "keywords": [...], "topic": "..."}]
    Returns: file path if matched and valid, otherwise None.
    """
    if not keywords:
        return None

    target_index = index_path
    if not os.path.isabs(target_index) and not os.path.exists(target_index):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        alt_path = os.path.join(script_dir, index_path)
        if os.path.exists(alt_path):
            target_index = alt_path

    if not os.path.exists(target_index):
        logger.warning(f"Image bank index not found at: {target_index}")
        return None

    try:
        with open(target_index, 'r', encoding='utf-8') as f:
            index_data = json.load(f)
    except Exception as e:
        logger.error(f"Error reading image bank index: {e}")
        index_data = []

    if not isinstance(index_data, list):
        index_data = []

    # Also automatically include diagrams from NCERT Exemplar Bank if available
    script_dir = os.path.dirname(os.path.abspath(__file__))
    exemplar_index_file = os.path.join(script_dir, 'exemplar_bank_index.json')
    if os.path.exists(exemplar_index_file):
        try:
            with open(exemplar_index_file, 'r', encoding='utf-8') as ef:
                ex_data = json.load(ef)
                if isinstance(ex_data, list):
                    for ch_entry in ex_data:
                        ch_kws = [w.strip().lower() for w in f"{ch_entry.get('subject', '')} {ch_entry.get('chapter', '')}".replace('_', ' ').split() if w.strip()]
                        for img_item in ch_entry.get('images', []):
                            img_f = img_item.get('file')
                            if img_f:
                                index_data.append({
                                    'file': img_f,
                                    'keywords': ch_kws,
                                    'topic': ch_entry.get('chapter', ''),
                                    'source': 'NCERT Exemplar Bank'
                                })
        except Exception as e:
            logger.debug(f"Could not load exemplar bank index in find_from_bank: {e}")

    if isinstance(keywords, str):
        query_words = [w.strip().lower() for w in keywords.replace(',', ' ').split() if w.strip()]

    elif isinstance(keywords, (list, tuple, set)):
        query_words = []
        for item in keywords:
            for w in str(item).replace(',', ' ').split():
                if w.strip():
                    query_words.append(w.strip().lower())
    else:
        query_words = []

    if not query_words:
        return None

    best_match = None
    max_score = 0

    for entry in index_data:
        if not isinstance(entry, dict):
            continue

        file_path = entry.get('file', '')
        entry_keywords = [str(k).strip().lower() for k in entry.get('keywords', [])]
        topic = str(entry.get('topic', '')).strip().lower()

        score = 0
        for qw in query_words:
            if qw in entry_keywords:
                score += 3
            elif any(qw in ek or ek in qw for ek in entry_keywords):
                score += 2
            elif qw in topic:
                score += 1

        if score > max_score:
            max_score = score
            best_match = file_path

    if max_score > 0 and best_match:
        if os.path.exists(best_match):
            return best_match
        script_dir = os.path.dirname(os.path.abspath(__file__))
        resolved = os.path.join(script_dir, best_match)
        if os.path.exists(resolved):
            return resolved
        return best_match

    return None


def save_to_image_bank(file_rel_path, keywords, topic="Educational Diagram", class_="General", subject="General", source="Wikimedia Commons", index_path="image_bank_index.json"):
    """
    Thread-safely append a newly downloaded/generated image to image_bank_index.json.
    Ensures persistent storage so the image is loaded locally in future without re-downloading.
    """
    if not file_rel_path:
        return False

    target_index = index_path
    if not os.path.isabs(target_index) and not os.path.exists(target_index):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        alt_path = os.path.join(script_dir, index_path)
        if os.path.exists(alt_path) or os.path.exists(script_dir):
            target_index = alt_path

    # Clean keywords
    kw_list = []
    if isinstance(keywords, (list, tuple, set)):
        kw_list = [str(k).strip().lower() for k in keywords if str(k).strip()]
    elif isinstance(keywords, str):
        kw_list = [w.strip().lower() for w in keywords.replace(',', ' ').split() if w.strip()]

    with _index_lock:
        index_data = []
        if os.path.exists(target_index):
            try:
                with open(target_index, 'r', encoding='utf-8') as f:
                    index_data = json.load(f)
            except Exception as e:
                logger.error(f"Error reading index before saving: {e}")
                index_data = []

        if not isinstance(index_data, list):
            index_data = []

        norm_rel = file_rel_path.replace('\\', '/')
        # Check if already indexed
        for entry in index_data:
            if isinstance(entry, dict) and entry.get('file', '').replace('\\', '/') == norm_rel:
                # Merge any new keywords
                existing_kws = set(entry.get('keywords', []))
                for k in kw_list:
                    if k not in existing_kws:
                        entry.setdefault('keywords', []).append(k)
                try:
                    with open(target_index, 'w', encoding='utf-8') as f:
                        json.dump(index_data, f, indent=2, ensure_ascii=False)
                except Exception:
                    pass
                return True

        new_entry = {
            "file": norm_rel,
            "keywords": kw_list,
            "topic": topic or "Educational Diagram",
            "class": class_ or "General",
            "subject": subject or "General",
            "source": source
        }
        index_data.append(new_entry)

        try:
            with open(target_index, 'w', encoding='utf-8') as f:
                json.dump(index_data, f, indent=2, ensure_ascii=False)
            logger.info(f"Successfully indexed '{norm_rel}' to image bank index.")
            return True
        except Exception as e:
            logger.error(f"Failed to write image bank index: {e}")
            return False


# ── WIKIMEDIA COMMONS API FETCHING ────────────────────────────────────────────
def fetch_from_wikimedia(keywords, output_path=None, min_width=600):
    """
    Search Wikimedia Commons API (https://commons.wikimedia.org/w/api.php) for images matching keywords.
    Filters top results for resolution (min_width) and CC-BY / CC-BY-SA / Public Domain licenses.
    Strictly filters out scanned PDF/DjVu books, theses, manuscripts, and non-diagram covers.
    Downloads the image via requests and saves to output_path and local cache.
    
    Returns absolute file path if downloaded and valid, otherwise None.
    Handles network/API errors gracefully without raising unhandled exceptions.
    """
    if not keywords:
        return None

    if isinstance(keywords, (list, tuple, set)):
        clean_query = ' '.join([str(k).strip() for k in keywords if str(k).strip()])
    else:
        clean_query = str(keywords).replace('_', ' ').replace('-', ' ').strip()

    if not clean_query:
        return None

    # Step 0: Check if a high-quality pre-verified diagram exists in local bank first
    bank_match = find_from_bank(clean_query)
    if bank_match and os.path.exists(bank_match) and os.path.getsize(bank_match) > 0:
        if output_path and os.path.abspath(bank_match) != os.path.abspath(output_path):
            try:
                os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
                shutil.copyfile(bank_match, output_path)
                return output_path
            except Exception:
                pass
        return bank_match

    # Setup default output path if not provided
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if not output_path:
        sanitized_name = re.sub(r'[^a-zA-Z0-9_-]', '_', clean_query.lower())[:45].strip('_')
        output_path = os.path.join(script_dir, 'static', 'images', 'ncert', f"wiki_{sanitized_name}.png")

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # Check local cache first
    cached_file = _get_cached_image(clean_query, prefix='wiki', ext='png')
    if cached_file and os.path.exists(cached_file) and os.path.getsize(cached_file) > 0:
        if os.path.abspath(cached_file) != os.path.abspath(output_path):
            shutil.copyfile(cached_file, output_path)
        return output_path

    headers = {
        'User-Agent': 'RRB_CBT_Education/1.0 (https://rrbcbt.edu; educational question paper generator) Python-requests'
    }
    api_url = "https://commons.wikimedia.org/w/api.php"

    # Build smart, prioritized query candidates (simplifying verbose AI prompt keywords)
    stop_words = {'diagram', 'labeled', 'structure', 'cross', 'section', 'of', 'and', 'the', 'a', 'an', 'in', 'for', 'with', 'showing', 'representation', 'illustrates', 'study', 'below', 'pathway', 'action'}
    clean_words = [w.lower() for w in re.findall(r'[a-zA-Z0-9]+', clean_query) if w.lower() not in stop_words]

    search_terms = []
    if len(clean_words) >= 2:
        search_terms.append(f"{clean_words[0]} {clean_words[1]} diagram")
        if len(clean_words) >= 3:
            search_terms.append(f"{clean_words[0]} {clean_words[1]} {clean_words[2]} diagram")
        search_terms.append(f"{clean_words[0]} {clean_words[1]}")
    elif len(clean_words) == 1:
        search_terms.append(f"{clean_words[0]} diagram")

    if clean_query not in search_terms:
        search_terms.append(clean_query)
    if f"{clean_query} diagram" not in search_terms and "diagram" not in clean_query.lower():
        search_terms.append(f"{clean_query} diagram")

    # Document & scan file exclusions
    disallowed_exts = ('.pdf', '.djvu', '.tif', '.tiff', '.epub', '.doc', '.docx', '.ogg', '.ogv', '.webm', '.mid', '.midi')
    disallowed_tokens = ('(ia ', 'internet archive', 'manuscript', 'dissertation', 'thesis', 'book cover', 'title page', 'front cover', 'back cover', 'binding', 'spine')

    for q_term in search_terms:
        try:
            search_params = {
                'action': 'query',
                'list': 'search',
                'srsearch': q_term,
                'srnamespace': 6,  # File namespace on Wikimedia Commons
                'srlimit': 10,
                'format': 'json'
            }
            resp = requests.get(api_url, params=search_params, headers=headers, timeout=10)
            if resp.status_code != 200:
                continue

            search_data = resp.json()
            search_results = search_data.get('query', {}).get('search', [])
            if not search_results:
                continue

            # Prioritize SVG / PNG scientific diagrams over random photos or documents
            sorted_results = []
            for item in search_results:
                t_low = str(item.get('title', '')).lower()
                # Skip disallowed documents/books
                if any(t_low.endswith(ext) for ext in disallowed_exts):
                    continue
                if any(tok in t_low for tok in disallowed_tokens):
                    continue
                
                # Priority weight: SVG diagram > PNG diagram > JPG diagram > other images
                score = 0
                if t_low.endswith('.svg'):
                    score += 5
                elif t_low.endswith('.png'):
                    score += 3
                if 'diagram' in t_low or 'schematic' in t_low or 'structure' in t_low or 'anatomy' in t_low:
                    score += 4
                if any(cw in t_low for cw in clean_words):
                    score += 3
                sorted_results.append((score, item))

            sorted_results.sort(key=lambda x: x[0], reverse=True)

            for score, item in sorted_results:
                title = item.get('title')
                if not title:
                    continue

                info_params = {
                    'action': 'query',
                    'titles': title,
                    'prop': 'imageinfo',
                    'iiprop': 'url|size|extmetadata|mime',
                    'iiurlwidth': max(int(min_width), 800),
                    'format': 'json'
                }
                info_resp = requests.get(api_url, params=info_params, headers=headers, timeout=10)
                if info_resp.status_code != 200:
                    continue

                info_data = info_resp.json()
                pages = info_data.get('query', {}).get('pages', {})
                if not pages:
                    continue

                for page_info in pages.values():
                    imageinfo = page_info.get('imageinfo', [])
                    if not imageinfo:
                        continue

                    info = imageinfo[0]
                    orig_width = int(info.get('width') or 0)
                    thumb_width = int(info.get('thumbwidth') or 0)
                    thumb_url = info.get('thumburl')
                    orig_url = info.get('url')
                    mime = str(info.get('mime', '')).lower()
                    extmetadata = info.get('extmetadata', {})

                    # Reject non-visual MIME types
                    if any(bad_mime in mime for bad_mime in ['pdf', 'djvu', 'tiff', 'audio', 'video', 'text', 'octet-stream']):
                        continue

                    # Resolution check
                    if orig_width < min_width and thumb_width < min_width:
                        continue

                    # License check (CC-BY, CC-BY-SA, Public Domain, CC0)
                    if not _is_allowed_wikimedia_license(extmetadata):
                        continue

                    # Choose download URL (prefer high-res rasterized thumbnail for SVGs or very large files)
                    download_url = thumb_url if thumb_url else orig_url
                    if not download_url:
                        continue

                    # Download image
                    img_resp = requests.get(download_url, headers=headers, timeout=15)
                    if img_resp.status_code != 200 or len(img_resp.content) == 0:
                        continue

                    # Write to output_path
                    with open(output_path, 'wb') as out_f:
                        out_f.write(img_resp.content)

                    # Save to local cache
                    _save_to_cache(clean_query, img_resp.content, prefix='wiki', ext='png')

                    # Register in image bank index under clean canonical name if saved in static
                    try:
                        rel_target = os.path.relpath(output_path, script_dir).replace('\\', '/')
                        if 'static/' in rel_target:
                            save_to_image_bank(
                                file_rel_path=rel_target,
                                keywords=clean_words or clean_query.split(),
                                topic=clean_query.title(),
                                source="Wikimedia Commons"
                            )
                    except Exception:
                        pass

                    logger.info(f"Successfully fetched clean Wikimedia image '{title}' for '{clean_query}' to {output_path}")
                    return output_path

        except Exception as e:
            logger.warning(f"Wikimedia search error for query '{q_term}': {e}")
            continue

    return None



def fetch_and_save_wikimedia_diagram(keywords, topic=None, class_=None, subject=None, target_dir=None, index_path="image_bank_index.json"):
    """
    Backward-compatible wrapper for fetching Wikimedia diagram and indexing it.
    """
    if not keywords:
        return None

    script_dir = os.path.dirname(os.path.abspath(__file__))
    if not target_dir:
        target_dir = os.path.join(script_dir, 'static', 'images', 'ncert')
    os.makedirs(target_dir, exist_ok=True)

    if isinstance(keywords, (list, tuple, set)):
        clean_query = ' '.join([str(k).strip() for k in keywords if str(k).strip()])
    else:
        clean_query = str(keywords).replace('_', ' ').replace('-', ' ').strip()

    sanitized_name = re.sub(r'[^a-zA-Z0-9_-]', '_', clean_query.lower())[:45].strip('_')
    output_path = os.path.join(target_dir, f"wiki_{sanitized_name}.png")

    res = fetch_from_wikimedia(keywords, output_path=output_path, min_width=600)
    if res and os.path.exists(res):
        rel_target = os.path.relpath(res, script_dir).replace('\\', '/')
        save_to_image_bank(
            file_rel_path=rel_target,
            keywords=clean_query.split(),
            topic=topic or clean_query.title(),
            class_=str(class_ or 'General'),
            subject=str(subject or 'General'),
            source="Wikimedia Commons",
            index_path=index_path
        )
        return res
    return None


# ── RDKIT CHEMICAL MOLECULAR STRUCTURE GENERATION ─────────────────────────────
def generate_molecule_structure(smiles_or_name, output_path=None):
    """
    Generate 2D chemical structure diagram using RDKit and save as high-resolution PNG (600x600).
    Input can be a SMILES string or common NCERT Class 10-12 chemical compound name.
    
    Returns absolute file path if generated and valid, otherwise None.
    Handles invalid SMILES or missing dependencies gracefully without raising unhandled exceptions.
    """
    if not smiles_or_name:
        return None

    raw_input = str(smiles_or_name).strip()
    if not raw_input:
        return None

    # Check NCERT lookup dictionary first
    clean_name = raw_input.lower().replace('-', ' ').replace('_', ' ').strip()
    smiles = NCERT_COMPOUND_SMILES.get(clean_name) or NCERT_COMPOUND_SMILES.get(raw_input.lower()) or raw_input

    # Setup default output path if not provided
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if not output_path:
        mol_dir = os.path.join(script_dir, 'static', 'images', 'molecules')
        os.makedirs(mol_dir, exist_ok=True)
        safe_name = re.sub(r'[^a-zA-Z0-9_-]', '_', clean_name)[:35].strip('_') or "chem"
        output_path = os.path.join(mol_dir, f"mol_{safe_name}.png")

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # Check local cache first
    cache_key = f"{clean_name}_{smiles}"
    cached_file = _get_cached_image(cache_key, prefix='mol', ext='png')
    if cached_file and os.path.exists(cached_file) and os.path.getsize(cached_file) > 0:
        if os.path.abspath(cached_file) != os.path.abspath(output_path):
            shutil.copyfile(cached_file, output_path)
        return output_path

    if not RDKIT_AVAILABLE or Chem is None or Draw is None:
        logger.warning(f"RDKit is not installed or available. Cannot generate structure for '{smiles_or_name}'.")
        return None

    try:
        # Build Mol object from SMILES
        mol = Chem.MolFromSmiles(smiles)
        if mol is None:
            # Try sanitization recovery or alternate parsing
            mol = Chem.MolFromSmiles(smiles, sanitize=False)
            if mol is not None:
                try:
                    Chem.SanitizeMol(mol)
                except Exception:
                    pass

        if mol is None:
            logger.warning(f"Could not parse SMILES or compound name: '{smiles_or_name}' (resolved: '{smiles}')")
            return None

        # Compute 2D coordinates for clear layout
        if rdDepictor is not None:
            try:
                rdDepictor.Compute2DCoords(mol)
            except Exception:
                pass

        # Generate crisp 600x600 PNG diagram
        Draw.MolToFile(mol, output_path, size=(600, 600), fitImage=True, imageType='png')

        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            _save_to_cache(cache_key, output_path, prefix='mol', ext='png')
            logger.info(f"Successfully generated molecular structure for '{smiles_or_name}' at {output_path}")
            return output_path

    except Exception as e:
        logger.warning(f"Error generating molecular structure for '{smiles_or_name}': {e}")
        return None

    return None


# ── DOCX PLACEHOLDER HELPER ───────────────────────────────────────────────────
def insert_placeholder(doc, description):
    """
    Insert a highlighted (yellow) paragraph in a python-docx Document object:
    '[IMAGE NEEDED: {description}]'
    """
    try:
        p = doc.add_paragraph()
        run = p.add_run(f"[IMAGE NEEDED: {description}]")
        run.bold = True
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return p
    except Exception as e:
        logger.error(f"Error inserting docx placeholder: {e}")
        return None


# ── PROCESS QUESTION IMAGES PIPELINE ──────────────────────────────────────────
def process_question_images(paper_data, output_dir=None):
    """
    Process image requirements for all questions in paper_data.
    Supports:
      - 'circuit': Generates circuit diagram via schemdraw.
      - 'chemistry': Generates 2D molecular structure diagram via RDKit.
      - 'wikimedia': Fetches scientific/educational diagram from Wikimedia Commons API.
      - 'ncert_diagram': Fallback chain (Wikimedia Commons -> Local Bank Index -> Placeholder).
      - 'none': Purely text/theory question.
      
    Attaches image_path, image_filename, image_base64, or image_placeholder to each question dict.
    """
    if not isinstance(paper_data, dict) or 'sections' not in paper_data:
        return paper_data

    script_dir = os.path.dirname(os.path.abspath(__file__))
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    else:
        output_dir = os.path.join(script_dir, 'static', 'images', 'generated')
        os.makedirs(output_dir, exist_ok=True)

    img_counter = 1

    for sec in paper_data.get('sections', []):
        sec_label = str(sec.get('section_label', 'Sec')).replace(' ', '_')
        for q in sec.get('questions', []):
            needs_img = q.get('needs_image', False)
            category = str(q.get('image_category', 'none')).strip().lower()

            if not needs_img or category in ('none', '', 'null'):
                continue

            q_num = q.get('number', img_counter)
            img_counter += 1

            # 1. ELECTRICAL CIRCUIT
            if category == 'circuit':
                c_type = q.get('circuit_type') or 'series_resistors'
                params = q.get('params') or {}
                out_filename = f"circuit_{sec_label}_q{q_num}.png"
                out_path = os.path.join(output_dir, out_filename)

                try:
                    generate_circuit(c_type, params, out_path)
                    if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                        q['image_path'] = out_path
                        q['image_filename'] = out_filename
                        with open(out_path, 'rb') as f:
                            q['image_base64'] = f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
                    else:
                        q['image_placeholder'] = f"Circuit diagram for {str(c_type).replace('_', ' ').title()}"
                except Exception as e:
                    logger.error(f"Error generating circuit for Q{q_num}: {e}")
                    q['image_placeholder'] = f"Circuit diagram for {str(c_type).replace('_', ' ').title()}"

            # 2. CHEMISTRY MOLECULE STRUCTURE
            elif category == 'chemistry':
                chem_identifier = (
                    q.get('compound_name') or
                    q.get('smiles') or
                    q.get('params', {}).get('smiles') or
                    q.get('params', {}).get('compound_name') or
                    q.get('image_prompt')
                )
                if not chem_identifier and q.get('image_keywords'):
                    chem_identifier = q.get('image_keywords')[0] if isinstance(q.get('image_keywords'), list) else str(q.get('image_keywords'))

                out_filename = f"chem_{sec_label}_q{q_num}.png"
                out_path = os.path.join(output_dir, out_filename)

                res_mol = generate_molecule_structure(chem_identifier, out_path)
                if res_mol and os.path.exists(res_mol) and os.path.getsize(res_mol) > 0:
                    q['image_path'] = res_mol
                    q['image_filename'] = os.path.basename(res_mol)
                    try:
                        with open(res_mol, 'rb') as f:
                            q['image_base64'] = f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
                    except Exception as e:
                        logger.warning(f"Could not encode molecular image base64: {e}")
                else:
                    chem_label = str(chem_identifier or q.get('question', '')[:50])
                    q['image_placeholder'] = f"Chemical Structure: {chem_label}"

            # 3. WIKIMEDIA COMMONS DIAGRAM
            elif category == 'wikimedia':
                keywords = q.get('image_keywords') or [q.get('image_prompt', '')] or [q.get('question', '')[:60]]
                out_filename = f"wiki_{sec_label}_q{q_num}.png"
                out_path = os.path.join(output_dir, out_filename)

                # Step 1: Try Wikimedia Commons API
                res_img = fetch_from_wikimedia(keywords, output_path=out_path, min_width=600)

                # Step 2: Fallback to local image bank
                if not res_img or not os.path.exists(res_img):
                    matched_bank = find_from_bank(keywords)
                    if matched_bank and os.path.exists(matched_bank):
                        res_img = matched_bank

                if res_img and os.path.exists(res_img) and os.path.getsize(res_img) > 0:
                    q['image_path'] = res_img
                    q['image_filename'] = os.path.basename(res_img)
                    try:
                        ext = os.path.splitext(res_img)[1].lower().replace('.', '')
                        mime = f"image/{ext}" if ext != 'jpg' else 'image/jpeg'
                        with open(res_img, 'rb') as f:
                            q['image_base64'] = f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
                    except Exception as e:
                        logger.warning(f"Could not encode wikimedia image base64: {e}")
                else:
                    kw_desc = ', '.join([str(k) for k in keywords]) if isinstance(keywords, list) else str(keywords)
                    q['image_placeholder'] = f"Diagram: {kw_desc}"

            # 4. NCERT DIAGRAM (WITH WIKIMEDIA + LOCAL BANK FALLBACK CHAIN)
            elif category == 'ncert_diagram':
                keywords = q.get('image_keywords') or [q.get('image_prompt', '')] or [q.get('question', '')[:60]]
                out_filename = f"ncert_{sec_label}_q{q_num}.png"
                out_path = os.path.join(output_dir, out_filename)

                # Step 1: Check local image bank first
                res_img = find_from_bank(keywords)

                # Step 2: If not in bank, search & fetch from Wikimedia Commons API
                if not res_img or not os.path.exists(res_img):
                    res_img = fetch_from_wikimedia(keywords, output_path=out_path, min_width=600)

                if res_img and os.path.exists(res_img) and os.path.getsize(res_img) > 0:
                    q['image_path'] = res_img
                    q['image_filename'] = os.path.basename(res_img)
                    try:
                        ext = os.path.splitext(res_img)[1].lower().replace('.', '')
                        mime = f"image/{ext}" if ext != 'jpg' else 'image/jpeg'
                        with open(res_img, 'rb') as f:
                            q['image_base64'] = f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
                    except Exception as e:
                        logger.warning(f"Could not encode ncert diagram base64: {e}")
                else:
                    kw_desc = ', '.join([str(k) for k in keywords]) if isinstance(keywords, list) else str(keywords)
                    q['image_placeholder'] = f"NCERT Diagram: {kw_desc}"

    return paper_data


def _add_formatted_runs_docx(paragraph, text, base_font_size=10):
    """
    Renders text containing inline code `...` and LaTeX / Chemistry expressions into Word Paragraph
    using proper fonts, Greek/Unicode symbols, and superscript/subscript runs.
    """
    if not text:
        return

    # Split by inline code: `code`
    tokens = re.split(r'(`[^`\n]+`)', str(text))
    for token in tokens:
        if not token:
            continue
        if token.startswith('`') and token.endswith('`') and len(token) >= 2:
            code_val = token[1:-1]
            r = paragraph.add_run(code_val)
            r.font.name = 'Consolas'
            r.font.size = Pt(base_font_size - 0.5)
            r.font.color.rgb = RGBColor(30, 27, 75)
            continue

        # Non-code text: process LaTeX & Chemistry notation
        clean = token
        clean = re.sub(r'\\ce\{((?:[^{}]|\{[^{}]*\})+)\}', r'\1', clean)
        clean = clean.replace('<=>', ' ⇌ ').replace('->', ' → ').replace('<-', ' ← ')
        clean = clean.replace('^', '↑').replace(' v ', ' ↓ ')

        # Greek letters & math symbols
        greek_math = [
            (r'\\Omega', 'Ω'), (r'\\mu_0', 'μ₀'), (r'\\mu', 'μ'), (r'\\lambda', 'λ'), (r'\\theta', 'θ'),
            (r'\\alpha', 'α'), (r'\\beta', 'β'), (r'\\gamma', 'γ'), (r'\\delta', 'δ'),
            (r'\\pi', 'π'), (r'\\sigma', 'σ'), (r'\\tau', 'τ'), (r'\\phi', 'φ'),
            (r'\\omega', 'ω'), (r'\\Delta', 'Δ'), (r'\\Phi', 'Φ'), (r'\\Psi', 'Ψ'),
            (r'\\varepsilon_0|\\epsilon_0', 'ε₀'), (r'\\varepsilon|\\epsilon', 'ε'),
            (r'\\pm', '±'), (r'\\times', '×'), (r'\\cdot', '·'), (r'\\div', '÷'),
            (r'\\approx', '≈'), (r'\\leq?|<=', '≤'), (r'\\geq?|>=', '≥'), (r'\\neq|!=', '≠'),
            (r'\\infty', '∞'), (r'\\degree', '°'), (r'\\propto', '∝'), (r'\\sqrt', '√'),
            (r'\\int', '∫'), (r'\\sum', '∑'), (r'\\rightarrow|\\to', '→'), (r'\\leftarrow', '←'),
            (r'\\text\{([^{}]+)\}', r'\1'), (r'\\mathbf\{([^{}]+)\}', r'\1'),
            (r'\\vec\{([^{}]+)\}', r'\1⃗'), (r'\\hat\{([^{}]+)\}', r'\1̂'),
            (r'\\frac\{([^{}]+)\}\{([^{}]+)\}', r'(\1 / \2)'),
            (r'\\\(|\\\)|\\\[|\\\]|\$\$|\$', '')
        ]
        for pat, rep in greek_math:
            clean = re.sub(pat, rep, clean)

        # Parse superscripts & subscripts: e.g. x^{2}, 10^{-6}, H_{2}O, Fe^{2+}
        sub_sup_pat = re.compile(r'(\^\{[^}]+\}|\^[0-9+\-a-zA-Z]|_\{[^}]+\}|_[0-9a-zA-Z])')
        sub_tokens = sub_sup_pat.split(clean)
        for st in sub_tokens:
            if not st:
                continue
            if st.startswith('^'):
                sup_val = st[2:-1] if st.startswith('^{') else st[1:]
                r = paragraph.add_run(sup_val)
                r.font.superscript = True
                r.font.size = Pt(base_font_size)
            elif st.startswith('_'):
                sub_val = st[2:-1] if st.startswith('_{') else st[1:]
                r = paragraph.add_run(sub_val)
                r.font.subscript = True
                r.font.size = Pt(base_font_size)
            else:
                r = paragraph.add_run(st)
                r.font.size = Pt(base_font_size)

# ── BUILD NATIVE DOCX QUESTION PAPER ──────────────────────────────────────────
def _render_question_text_docx(doc, text_content, q_prefix=None, marks_badge=None, base_indent_mm=0):
    """
    Renders question/subquestion text into Word docx with strict preservation
    of whitespace and monospace font for Computer Programming code (Python, Java, C++, SQL, HTML)
    and formatting for Physics and Chemistry equations.
    """
    if not text_content:
        return

    text_str = str(text_content).strip()
    code_block_pattern = re.compile(r'```(?:[a-zA-Z0-9_+-]+)?\s*\n?(.*?)\n?```', re.DOTALL)
    
    parts = []
    last_end = 0
    for match in code_block_pattern.finditer(text_str):
        s_idx, e_idx = match.span()
        if s_idx > last_end:
            parts.append(('text', text_str[last_end:s_idx]))
        parts.append(('code', match.group(1)))
        last_end = e_idx
    if last_end < len(text_str):
        parts.append(('text', text_str[last_end:]))
    if not parts:
        parts = [('text', text_str)]

    is_first_para = True
    for p_type, content in parts:
        if p_type == 'text':
            # Check if raw text itself contains formatted lines
            lines = content.split('\n')
            for line in lines:
                line_str = line.rstrip()
                if not line_str and not is_first_para:
                    continue
                
                p = doc.add_paragraph()
                if base_indent_mm > 0:
                    p.paragraph_format.left_indent = Mm(base_indent_mm)
                p.paragraph_format.space_before = Pt(4 if is_first_para else 1)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15

                if is_first_para and q_prefix:
                    r_pfx = p.add_run(q_prefix)
                    r_pfx.bold = True
                    r_pfx.font.size = Pt(10)
                    is_first_para = False

                _add_formatted_runs_docx(p, line_str, base_font_size=10)
        else:
            # Code block - preserve exact whitespace, indent, and Consolas monospace font
            code_lines = content.split('\n')
            for c_line in code_lines:
                p_code = doc.add_paragraph()
                p_code.paragraph_format.left_indent = Mm(base_indent_mm + 8)
                p_code.paragraph_format.space_before = Pt(0)
                p_code.paragraph_format.space_after = Pt(0)
                p_code.paragraph_format.line_spacing = 1.0

                r_c = p_code.add_run(c_line if c_line else " ")
                r_c.font.name = 'Consolas'
                r_c.font.size = Pt(9.5)
                r_c.font.color.rgb = RGBColor(15, 23, 42)

    if marks_badge:
        p_badge = doc.add_paragraph()
        if base_indent_mm > 0:
            p_badge.paragraph_format.left_indent = Mm(base_indent_mm)
        p_badge.paragraph_format.space_before = Pt(0)
        p_badge.paragraph_format.space_after = Pt(2)
        r_mb = p_badge.add_run(marks_badge)
        r_mb.bold = True
        r_mb.font.size = Pt(8.5)
        r_mb.font.color.rgb = RGBColor(100, 116, 139)


def create_paper_docx(paper_data, meta, output_docx_path):
    """
    Build a professionally formatted official DOCX question paper using python-docx.
    Includes: Logo, Institute Name, Address, Academic Session, Class, Subject,
    Topic (conditionally omitted if full/whole syllabus), MM, Time, and exact code formatting.
    """
    doc = docx.Document()

    # A4 Page Setup
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(14)
        section.bottom_margin = Mm(14)
        section.left_margin = Mm(16)
        section.right_margin = Mm(16)

    meta_info = meta or {}
    school_name = meta_info.get('school_name') or 'RRB Group of Schools'
    school_address = meta_info.get('school_address') or ''
    academic_session = meta_info.get('academic_session') or ''
    exam_type = meta_info.get('exam_type') or 'Unit Test'
    paper_type = meta_info.get('paper_type') or 'Descriptive Paper'
    is_graded = meta_info.get('is_graded') if 'is_graded' in meta_info else (paper_type in ('Descriptive Paper', 'DPP'))
    class_ = meta_info.get('class') or 'General'
    subject = meta_info.get('subject') or 'Subject'
    topic = meta_info.get('topic') or meta_info.get('topics') or meta_info.get('chapter') or ''
    is_full_syllabus = meta_info.get('full_syllabus') in (True, 'yes', 'true', '1')
    duration = meta_info.get('duration') or '3 Hours'
    max_marks = meta_info.get('max_marks') or '100'
    teacher_name = meta_info.get('teacher_name') or 'Faculty'

    # A4 Page Setup and Running Header for multi-page papers (100-200 questions)
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(14)
        section.bottom_margin = Mm(14)
        section.left_margin = Mm(16)
        section.right_margin = Mm(16)
        section.different_first_page_header_footer = True

        # Running header for pages 2+
        hdr = section.header
        if hdr.paragraphs:
            p_hdr = hdr.paragraphs[0]
            p_hdr.text = ""
            p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_h = p_hdr.add_run(f"{school_name}  |  {paper_type.upper()} — Class {class_} {subject.upper()}")
            r_h.font.size = Pt(8.5)
            r_h.font.color.rgb = RGBColor(140, 140, 140)

        # Running footer
        ftr = section.footer
        if ftr.paragraphs:
            p_ftr = ftr.paragraphs[0]
            p_ftr.text = ""
            p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_f = p_ftr.add_run(f"{paper_type}  ·  RRB CBT School System")
            r_f.font.size = Pt(8)
            r_f.font.color.rgb = RGBColor(150, 150, 150)

    # 1. Logo (Optional)
    logo_file = meta_info.get('logo_path')
    if logo_file:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        cand_paths = [
            logo_file,
            os.path.join(script_dir, logo_file),
            os.path.join(script_dir, 'static', logo_file),
            os.path.join(script_dir, 'static', 'logo.png')
        ]
        for cp in cand_paths:
            if cp and os.path.exists(cp) and os.path.getsize(cp) > 0:
                try:
                    p_logo = doc.add_paragraph()
                    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_logo.paragraph_format.space_after = Pt(2)
                    doc.add_picture(cp, height=Inches(0.85))
                    break
                except Exception:
                    pass

    # 2. Institute Name
    p_school = doc.add_paragraph()
    p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_school = p_school.add_run(school_name.upper())
    r_school.bold = True
    r_school.font.size = Pt(16)
    r_school.font.color.rgb = RGBColor(30, 27, 75)

    # 3. Address
    if school_address:
        p_addr = doc.add_paragraph()
        p_addr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_addr = p_addr.add_run(school_address)
        r_addr.font.size = Pt(9.5)
        r_addr.font.color.rgb = RGBColor(100, 116, 139)

    # 4. Paper Type / Exam Type & Academic Session
    p_exam = doc.add_paragraph()
    p_exam.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if paper_type and paper_type.lower() != exam_type.lower():
        exam_line = f"{paper_type.upper()} — {exam_type.upper()}"
    else:
        exam_line = f"{(paper_type or exam_type).upper()}"
    if academic_session:
        exam_line += f" (Academic Session: {academic_session})"
    r_exam = p_exam.add_run(exam_line)
    r_exam.bold = True
    r_exam.font.size = Pt(11)

    # 5. Class & Subject
    p_cls = doc.add_paragraph()
    p_cls.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cls = p_cls.add_run(f"CLASS: {class_}  |  SUBJECT: {subject.upper()}")
    r_cls.bold = True
    r_cls.font.size = Pt(11)
    r_cls.font.color.rgb = RGBColor(30, 27, 75)

    # 6. Topic (Only included if NOT whole/full syllabus)
    norm_topic = str(topic).strip().lower()
    if not is_full_syllabus and topic and norm_topic not in ('full syllabus', 'complete syllabus', 'all chapters', 'whole syllabus', 'none', ''):
        p_top = doc.add_paragraph()
        p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_top = p_top.add_run(f"TOPIC / CHAPTER: {topic}")
        r_top.bold = True
        r_top.font.size = Pt(10)
        r_top.font.color.rgb = RGBColor(71, 85, 105)

    # 7. Meta Table (Time Allowed, Maximum Marks / MM)
    tbl = doc.add_table(rows=2, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Mm(88)
    tbl.columns[1].width = Mm(88)

    cell_00 = tbl.cell(0, 0)
    p_00 = cell_00.paragraphs[0]
    r_00 = p_00.add_run(f"Time Allowed: {duration}")
    r_00.bold = True
    r_00.font.size = Pt(9.5)

    cell_01 = tbl.cell(0, 1)
    p_01 = cell_01.paragraphs[0]
    p_01.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if is_graded and max_marks:
        r_01 = p_01.add_run(f"Maximum Marks (MM): {max_marks}")
        r_01.bold = True
        r_01.font.size = Pt(9.5)
    else:
        p_01.add_run("").font.size = Pt(9.5)

    cell_10 = tbl.cell(1, 0)
    p_10 = cell_10.paragraphs[0]
    p_10.add_run(f"Teacher: {teacher_name}").font.size = Pt(9)

    cell_11 = tbl.cell(1, 1)
    p_11 = cell_11.paragraphs[0]
    p_11.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    import datetime
    p_11.add_run(f"Date: {datetime.datetime.now().strftime('%d/%m/%Y')}").font.size = Pt(9)

    # Divider Line
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(8)
    r_div = p_div.add_run("―" * 60)
    r_div.font.color.rgb = RGBColor(200, 200, 200)

    # Sections and Questions
    if isinstance(paper_data, dict) and 'sections' in paper_data:
        for sec in paper_data.get('sections', []):
            label = sec.get('section_label', 'Section')
            title = sec.get('section_title', '')
            instruction = sec.get('instruction', '')
            marks_per_q = sec.get('marks_per_question', 1)

            # Section Banner (keep_with_next prevents orphaned header at page bottom)
            p_sec = doc.add_paragraph()
            p_sec.paragraph_format.space_before = Pt(12)
            p_sec.paragraph_format.space_after = Pt(2)
            p_sec.paragraph_format.keep_with_next = True
            r_sec = p_sec.add_run(f"{label}: {title}")
            r_sec.bold = True
            r_sec.font.size = Pt(11)
            r_sec.font.color.rgb = RGBColor(30, 27, 75)

            if instruction:
                p_inst = doc.add_paragraph()
                p_inst.paragraph_format.space_after = Pt(6)
                p_inst.paragraph_format.keep_with_next = True
                r_inst = p_inst.add_run(f"Note: {instruction}")
                r_inst.italic = True
                r_inst.font.size = Pt(9)
                r_inst.font.color.rgb = RGBColor(100, 116, 139)

            # Questions
            for q in sec.get('questions', []):
                q_num = q.get('number', '')
                q_text = q.get('question', '')

                marks_badge_val = f"[{marks_per_q}M]" if is_graded else None
                _render_question_text_docx(
                    doc=doc,
                    text_content=q_text,
                    q_prefix=f"Q{q_num}. ",
                    marks_badge=marks_badge_val,
                    base_indent_mm=0
                )

                # Question Image or Placeholder
                img_path = q.get('image_path')
                placeholder = q.get('image_placeholder')

                if img_path and os.path.exists(img_path) and os.path.getsize(img_path) > 0:
                    try:
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_before = Pt(4)
                        p_img.paragraph_format.space_after = Pt(6)
                        doc.add_picture(img_path, width=Inches(3.2))
                    except Exception as e:
                        logger.error(f"Error inserting picture in docx: {e}")
                        insert_placeholder(doc, f"Diagram for Q{q_num}")
                elif placeholder:
                    insert_placeholder(doc, placeholder)
                elif q.get('needs_image'):
                    insert_placeholder(doc, f"Diagram for Q{q_num}")

                # Options (for MCQ)
                opts = q.get('options', [])
                if opts:
                    for opt in opts:
                        _render_question_text_docx(
                            doc=doc,
                            text_content=opt,
                            base_indent_mm=8
                        )

                # Sub-questions
                sub_qs = q.get('sub_questions', [])
                if sub_qs:
                    for i, sq in enumerate(sub_qs):
                        sq_text = sq if isinstance(sq, str) else sq.get('question', '')
                        _render_question_text_docx(
                            doc=doc,
                            text_content=sq_text,
                            q_prefix=f"({chr(97+i)}) ",
                            base_indent_mm=8
                        )

    # Footer

    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(24)
    r_end = p_end.add_run("*** End of Question Paper ***\nGenerated by RRB CBT  ·  Developed by Gaurav Shukla & Team")
    r_end.italic = True
    r_end.font.size = Pt(8.5)
    r_end.font.color.rgb = RGBColor(140, 140, 140)


    out_dir = os.path.dirname(output_docx_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc.save(output_docx_path)
    return output_docx_path
