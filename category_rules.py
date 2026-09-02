from __future__ import annotations

import re
from difflib import SequenceMatcher
from typing import Iterable, Mapping

# V7.12: the ONLY valid RTO Pro categories.  These are the unique non-empty
# values from column D (D1:D170) of the user's category.xlsx workbook.
APPROVED_CATEGORIES = [
    'A-Frame',
    'Carolina',
    'Cottage',
    'Chicken Coop',
    'Cabin',
    'Dog Kennel',
    'Economy',
    'Garden Shed',
    'High Barn',
    'Highwall',
    'Barn',
    'Lofted Barn',
    'Garage',
    'Metal Utility',
    'Mini Barn',
    'Highland',
    'Side Utility',
    'Utility',
    'Gazebo',
    'Porch',
    'Quaker',
    'Gable',
    'Metal',
    'Signature',
    'The Westwood',
    'Ultra',
    'Value',
    'Metal Value',
    'Woodsman',
]

_CANON_BY_NORM: dict[str, str] = {}


def _norm(value: object) -> str:
    s = str(value or '').upper().strip()
    s = s.replace('&', ' AND ')
    # Remove common dimensions so 10x12 Side Utility compares as Side Utility.
    s = re.sub(r"\b\d+(?:\.\d+)?\s*['\"]?\s*[X×]\s*\d+(?:\.\d+)?\s*['\"]?\b", ' ', s)
    s = re.sub(r'\bNO\s+OPTIONS?\s+AVA(?:ILABLE)?\b', ' ', s)
    s = re.sub(r'\bSMART\s*SIDE\b|\bSMART\s*LAP\b', ' ', s)
    s = re.sub(r'[^A-Z0-9]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


for _cat in APPROVED_CATEGORIES:
    _CANON_BY_NORM[_norm(_cat)] = _cat


# Confirmed aliases from the green/allowed mapping rows in category.xlsx.
# The values below are all members of APPROVED_CATEGORIES.
CONFIRMED_ALIASES = {
    'A FRAME': 'A-Frame',
    'CAROLINA, SMARTLAP': 'Carolina',
    'COLONIAL COTTAGE': 'Cottage',
    'COOP': 'Chicken Coop',
    'COT': 'Cottage',
    'COTTAGE PORCH': 'Cottage',
    'DELUXE': 'Cabin',
    'DELUXE CABIN': 'Cabin',
    'DELUXE LBC': 'Cabin',
    'DELUXE LOFTED BARN CABIN': 'Cabin',
    'DLB': 'Cabin',
    'DOG HOUSE': 'Dog Kennel',
    'DUTCH CABIN': 'Cabin',
    'ECONO': 'Economy',
    'GARDENER': 'Garden Shed',
    'HIGH BARN': 'High Barn',
    'HIGHWALL': 'Highwall',
    'HORSE BARN': 'Barn',
    'KENNEL': 'Dog Kennel',
    'LOFTED': 'Lofted Barn',
    'LOFTED BARN CABIN': 'Cabin',
    'LOFTED BARN, SMARTSIDE': 'Lofted Barn',
    'LOFTED CABIN': 'Cabin',
    'LOFTED GARAGE': 'Garage',
    'LOW BARN': 'Barn',
    'LOW WALL': 'Barn',
    'METAL UTILITY - NO OPTIONS AVA': 'Metal Utility',
    'MINIBARN': 'Mini Barn',
    'MODERN HIGHLAND': 'Highland',
    'MODERN UTILITY': 'Utility',
    'OCTAGON GAZEBO, STAINED': 'Gazebo',
    'OCTAGON GAZEBO, STAINED VALUE': 'Gazebo',
    'PORCH BUILDING': 'Porch',
    'PORCH NOOK': 'Porch',
    'QUAKER SHED, SMARTSIDE': 'Quaker',
    'SIDE GABLE': 'Gable',
    'SIDE LOFTED': 'Lofted Barn',
    'SIDE LOFTED BARN': 'Lofted Barn',
    'SIDE UTILITY': 'Utility',
    'SLB': 'Lofted Barn',
    'SM': 'Metal',
    'STANDARD LOFT': 'Lofted Barn',
    'STANDARD SIGNATURE': 'Signature',
    'SUT': 'Utility',
    'TEA HOUSE': 'Gazebo',
    'THE WESTWOOD, SMARTSIDE': 'The Westwood',
    'ULTRA CABIN': 'Ultra',
    'ULTRA SERIES': 'Ultra',
    'UTILITY SHED | SMARTSIDE': 'Utility',
    'UTILITY SHED, METAL': 'Metal Utility',
    'UTILITY SHED, SMARTSIDE': 'Utility',
    'VALUE SERIES': 'Value',
    'VALUE SHED, METAL': 'Metal Value',
    'VALUE SHED, SMARTSIDE': 'Value',
    'WOODSMEN': 'Woodsman',
}

_CONFIRMED_BY_NORM = {_norm(k): v for k, v in CONFIRMED_ALIASES.items()}


def canonical_category(value: object) -> str:
    """Return the exact approved spelling for an approved category, else ''."""
    return _CANON_BY_NORM.get(_norm(value), '')


def normalize_color_words(value: object) -> str:
    """Normalize color wording used by the user's RTO descriptions.

    Burnished -> B, and the coating word Urethane is removed completely.
    This helper is safe for full descriptions as well as individual colors.
    """
    s = str(value or '').strip()
    s = re.sub(r'\bURETHANE\b', ' ', s, flags=re.I)
    s = re.sub(r'\bBURNISHED\b', 'B', s, flags=re.I)
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'\s+([,/])', r'\1', s)
    s = re.sub(r'([,/])\s+', r'\1 ', s)
    return s.strip(' ,-/')


def _contains(text: str, *terms: str) -> bool:
    return any(_norm(t) in text for t in terms if _norm(t))


def _keyword_category(text: str) -> str:
    """High-confidence semantic rules before fuzzy fallback."""
    if not text:
        return ''

    # Common abbreviations / otherwise ambiguous labels from the workbook.
    # These prevent short codes from being sent to an unrelated fuzzy match.
    if text == 'GAZ':
        return 'Gazebo'
    if text in {'LBC', 'LC', 'DLB'}:
        return 'Cabin'
    if text == 'LBG':
        return 'Garage'
    if text == 'MB':
        return 'Mini Barn'
    if text in {'DB', 'MDB'}:
        return 'Barn'
    if text in {'NONE', 'MISCELLANEOUS', 'SPECIAL', 'SPECIALTY', 'STANDARD', 'TEST BUILDING', 'ASD', 'BVM', 'BRD AND BTN'}:
        return 'Utility'
    if text == 'DECK':
        return 'Porch'

    # Very specific branded / specialty styles first.
    if 'WESTWOOD' in text:
        return 'The Westwood'
    if 'WOODSMAN' in text or 'WOODSMEN' in text:
        return 'Woodsman'
    if 'ULTRA' in text:
        return 'Ultra'
    if 'SIGNATURE' in text or 'ELITE' in text:
        return 'Signature'
    if 'CAROLINA' in text:
        return 'Carolina'
    if 'QUAKER' in text:
        return 'Quaker'

    # Animal / outdoor structures.
    if 'DOG HOUSE' in text or 'DOG KENNEL' in text or re.search(r'\bKENNEL\b', text):
        return 'Dog Kennel'
    if 'CHICKEN' in text or re.search(r'\bCOOP\b', text):
        return 'Chicken Coop'
    if any(x in text for x in ('GAZEBO', 'PAVILION', 'TEA HOUSE')):
        return 'Gazebo'
    if 'GREENHOUSE' in text or 'POTTING' in text or 'GARDENER' in text or 'GARDEN SHED' in text:
        return 'Garden Shed'

    # Building-family rules where wording strongly identifies the category.
    if 'HIGH WALL' in text or 'HIGHWALL' in text:
        return 'Highwall'
    if 'HIGH BARN' in text or 'HIGHBARN' in text:
        return 'High Barn'
    if 'MINI BARN' in text or 'MINIBARN' in text:
        return 'Mini Barn'
    if 'HIGHLAND' in text or 'HIGH COUNTRY' in text:
        return 'Highland'

    # Value/metal combinations must beat the generic VALUE or METAL rules.
    if 'VALUE' in text and 'METAL' in text:
        return 'Metal Value'
    if 'METAL' in text and ('UTILITY' in text or 'SHED' in text):
        return 'Metal Utility'

    # Lofted cabin/garage are not Lofted Barn in the approved mapping.
    if 'LOFTED' in text and 'GARAGE' in text:
        return 'Garage'
    if ('LOFTED' in text or 'LBC' in text or 'DLB' in text) and 'CABIN' in text:
        return 'Cabin'

    if 'GARAGE' in text or 'CARPORT' in text:
        return 'Garage'
    if any(x in text for x in ('CABIN', 'CHALET', 'HUNTING BLIND')) or re.search(r'\bLBC\b|\bDLB\b', text):
        return 'Cabin'
    if any(x in text for x in ('COTTAGE', 'CABANA', 'CAPE', 'VILLA', 'STUDIO', 'COLONIAL', 'SALTBOX')) or re.search(r'\bCOT\b', text):
        return 'Cottage'
    if 'PORCH' in text:
        return 'Porch'
    if re.search(r'\bA FRAME\b', text):
        return 'A-Frame'
    if 'GABLE' in text:
        return 'Gable'

    if ('LOFTED BARN' in text or re.search(r'\bSLB\b|\bEMLB\b|\bLB\b', text)
            or text == 'LOFTED' or 'STANDARD LOFT' in text):
        return 'Lofted Barn'

    if any(x in text for x in ('ECONO', 'ECONOMY', 'BASELINE', 'BASIC')):
        return 'Economy'

    if any(x in text for x in ('HORSE BARN', 'DUTCH BARN', 'GAMBREL', 'RUN IN', 'LOAFING', 'ANIMAL SHELTER')):
        return 'Barn'
    if re.search(r'\bBARN\b', text):
        return 'Barn'

    if 'VALUE' in text:
        return 'Value'
    if re.search(r'\bMETAL\b', text) or text in {'SM', 'CONTAINER', 'STEEL FRAME', 'ALUMINUM'}:
        return 'Metal'

    # Generic sheds/styles with no stronger signal are safest as Utility.
    if any(x in text for x in (
        'UTILITY', 'STORAGE BUILDING', 'WORKSHOP', 'SHED', 'MONO', 'ONE SLOPE',
        'SINGLE SLOPE', 'THE WEDGE', 'LITE UTILITY', 'DUTCHLAP', 'LAPSIDE',
        'LAP EAVE', 'PINE LAP', 'CEDAR SIDING', 'PAINTED', 'TREATED', 'VINYL',
    )) or re.search(r'\bSUT\b|\bEMUT\b|\bMU\b', text):
        return 'Utility'

    # Play structures are visually closest to a small cottage/cabin shell.
    if any(x in text for x in ('PLAYHOUSE', 'PLAYSET', 'TOTS TOWER')):
        return 'Cottage'

    return ''


def smart_category(*values: object, aliases: Mapping[str, str] | None = None) -> str:
    """Map arbitrary style/description text to one approved RTO Pro category.

    Exact confirmed spreadsheet mappings win, then strong semantic rules, then
    a conservative fuzzy match against confirmed aliases and approved names.
    The function ALWAYS returns one of APPROVED_CATEGORIES.
    """
    texts = [_norm(v) for v in values if str(v or '').strip()]
    text = ' '.join(t for t in texts if t).strip()
    if not text:
        return 'Utility'

    # If the user already supplied a valid category, preserve it exactly.
    for v in values:
        canon = canonical_category(v)
        if canon:
            return canon

    # Exact confirmed spreadsheet alias.
    for t in texts:
        if t in _CONFIRMED_BY_NORM:
            return _CONFIRMED_BY_NORM[t]

    # Existing DataMaps aliases are allowed only when their output itself is
    # one of the approved column-D categories.
    alias_candidates: dict[str, str] = dict(_CONFIRMED_BY_NORM)
    if aliases:
        for raw_alias, raw_category in aliases.items():
            canon = canonical_category(raw_category)
            if canon:
                alias_candidates[_norm(raw_alias)] = canon

    # Longest confirmed phrase contained in the text gets first chance.
    contained = [
        (len(alias.split()), len(alias), alias, cat)
        for alias, cat in alias_candidates.items()
        if alias and alias in text and (len(alias) >= 3)
    ]
    if contained:
        contained.sort(reverse=True)
        return contained[0][3]

    rule = _keyword_category(text)
    if rule:
        return rule

    # Conservative fuzzy fallback. Score confirmed aliases plus canonical names.
    candidates = list(alias_candidates.items()) + [(_norm(c), c) for c in APPROVED_CATEGORIES]
    text_tokens = set(text.split())
    best_cat = 'Utility'
    best_score = -1.0
    for alias, cat in candidates:
        if not alias:
            continue
        seq = SequenceMatcher(None, text, alias).ratio()
        alias_tokens = set(alias.split())
        overlap = len(text_tokens & alias_tokens) / max(1, len(alias_tokens))
        # Token hits matter more than character similarity for multi-word styles.
        score = (seq * 0.58) + (overlap * 0.42)
        if alias_tokens and alias_tokens.issubset(text_tokens):
            score += 0.18
        if score > best_score:
            best_score = score
            best_cat = cat

    # Do not make a confident-looking random choice from weak similarity.
    # Unknown/vague descriptions fall back to the broad approved Utility
    # category; meaningful descriptions are normally handled above.
    if best_score < 0.58:
        return 'Utility'
    return canonical_category(best_cat) or 'Utility'
