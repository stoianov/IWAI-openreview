import re

def extract_field(content, key):
    """  OpenReview V2 stores values as:  content[key]["value"] or content[key]  """
    if key not in content:  return None
    value = content[key]
    if isinstance(value, dict) and "value" in value:
        return value["value"]
    return value

def normalize_text(x):
    if x is None:
        return ""
    return re.sub(r"\s+", " ", x.strip().lower())

def normalize_dict(dic): # 0-1 score and 0.5 as default
    if not dic: return {}
    vals = list(dic.values())
    vmin, vmax = min(vals), max(vals)
    if vmax == vmin: return {r: 0.5 for r in dic} # Some default score
    denom = vmax - vmin
    return { r: (v-vmin)/denom  for r, v in dic.items()}