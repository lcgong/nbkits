from matplotlib.colors import BASE_COLORS, CSS4_COLORS, XKCD_COLORS, TABLEAU_COLORS
from matplotlib.colors import to_hex as _to_hex


def to_hex_color(c: str):
    v = BASE_COLORS.get(c, None)
    if v is not None:
        v = _to_hex(v).upper()
        return v[1:]

    v = CSS4_COLORS.get(c, None)
    if v is not None:
        return v[1:].upper()

    if c.startswith("tab:"):
        v = TABLEAU_COLORS.get(c, None)
        if v is not None:
            print(v)
            return v[1:].upper()
        else:
            items = ", ".join(c[4:] for c in TABLEAU_COLORS.keys())
            raise ValueError(f"tableau color must be one of these: {items}")

    if c.startswith("xkcd:"):
        v = XKCD_COLORS.get(c, None)
        if v is not None:
            return v[1:].upper()
        else:
            items = ", ".join(c[5:] for c in XKCD_COLORS.keys())
            raise ValueError(f"xkcd color must be one of these: {items}")

    return c
