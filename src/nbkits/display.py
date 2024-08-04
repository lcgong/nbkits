from IPython.core.interactiveshell import InteractiveShell
from IPython.display import display_html


def hdisplay(*dfs, titles=None, gap=50, justify="center", alignment="left"):
    if alignment == "left":
        justify = "flex-start"
    elif alignment == "center":
        justify = "center"
    elif alignment == "right":
        justify = "end"
    else:
        raise ValueError("alignment: left/center/right")

    display_format = InteractiveShell.instance().display_formatter.format
    html = ""
    for i, df in enumerate(dfs):
        title = titles[i] if titles and i < len(titles) else None
        title = f"<h4>{title}</h4>" if title else ""
        format_dict, md_dict = display_format(df)
        if "text/html" in format_dict:
            html += f"<div>{title}{format_dict['text/html']}</div>"
        elif "text/plain" in format_dict:
            html += f"<div>{title}<pre>{format_dict['text/plain']}</pre></div>"
        else:
            format_types = ", ".join(format_dict.keys())
            raise ValueError(f"Unknown display format: {format_types}")

    style = f"display:flex; gap:{gap}px; justify-content:{justify};"
    display_html(f'<div style="{style}">{html}</div>', raw=True)
