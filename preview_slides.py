#!/usr/bin/env python3
"""Generate visual preview images of each slide from the PPTX."""

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

W, H = 1920, 1080  # Preview resolution

def emu_to_px(emu_val, slide_dim, px_dim):
    return int(emu_val / slide_dim * px_dim)

def get_color(color_obj):
    try:
        if color_obj and color_obj.rgb:
            r = color_obj.rgb
            return (r[0] << 16 | r[1] << 8 | r[2])
    except:
        pass
    return None

prs = Presentation('notebooklm-training-slides.pptx')
sw = prs.slide_width
sh = prs.slide_height

os.makedirs('slide_previews', exist_ok=True)

# Try to find a font
font_paths = [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
    '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
]

font_path = None
for fp in font_paths:
    if os.path.exists(fp):
        font_path = fp
        break

# Also find a bold font
bold_font_paths = [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
]
bold_font_path = None
for fp in bold_font_paths:
    if os.path.exists(fp):
        bold_font_path = fp
        break

def get_font(size, bold=False):
    try:
        path = bold_font_path if (bold and bold_font_path) else font_path
        if path:
            return ImageFont.truetype(path, size)
    except:
        pass
    return ImageFont.load_default()

for slide_idx, slide in enumerate(prs.slides):
    img = Image.new('RGB', (W, H), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    # Draw background
    bg = slide.background
    try:
        if bg.fill.type is not None:
            c = bg.fill.fore_color.rgb
            img = Image.new('RGB', (W, H), (c[0], c[1], c[2]))
            draw = ImageDraw.Draw(img)
    except:
        pass

    # Draw shapes
    for shape in slide.shapes:
        # Position
        x = emu_to_px(shape.left, sw, W)
        y = emu_to_px(shape.top, sh, H)
        w = emu_to_px(shape.width, sw, W)
        h = emu_to_px(shape.height, sh, H)

        # Draw filled shapes
        try:
            if shape.shape_type is not None and hasattr(shape, 'fill'):
                fill = shape.fill
                if fill.type is not None:
                    try:
                        c = fill.fore_color.rgb
                        color = (c[0], c[1], c[2])
                        if shape.shape_type == 9:  # OVAL
                            draw.ellipse([x, y, x+w, y+h], fill=color)
                        else:
                            draw.rounded_rectangle([x, y, x+w, y+h], radius=8, fill=color)
                    except:
                        pass
        except:
            pass

        # Draw text
        if shape.has_text_frame:
            tf = shape.text_frame
            ty = y + 4
            for para in tf.paragraphs:
                text = para.text.strip()
                if not text:
                    ty += 8
                    continue

                # Get font properties
                fs = 16
                bold = False
                text_color = (51, 51, 51)
                align = 'left'

                try:
                    if para.font.size:
                        fs = int(para.font.size / 12700)
                except:
                    pass
                try:
                    if para.font.bold:
                        bold = True
                except:
                    pass
                try:
                    if para.font.color and para.font.color.rgb:
                        c = para.font.color.rgb
                        text_color = (c[0], c[1], c[2])
                except:
                    pass
                try:
                    if para.alignment == PP_ALIGN.CENTER:
                        align = 'center'
                    elif para.alignment == PP_ALIGN.RIGHT:
                        align = 'right'
                except:
                    pass

                # Scale font size
                render_size = max(10, int(fs * 0.85))
                font = get_font(render_size, bold)

                # Handle text positioning
                try:
                    bbox = draw.textbbox((0, 0), text, font=font)
                    tw = bbox[2] - bbox[0]
                    th = bbox[3] - bbox[1]
                except:
                    tw, th = len(text) * render_size, render_size

                if align == 'center':
                    tx = x + (w - tw) // 2
                elif align == 'right':
                    tx = x + w - tw - 8
                else:
                    tx = x + 8

                # Word wrap for long text
                if tw > w - 16:
                    # Simple word wrap
                    chars_per_line = max(1, int(len(text) * (w - 16) / max(tw, 1)))
                    lines = []
                    while text:
                        if len(text) <= chars_per_line:
                            lines.append(text)
                            break
                        cut = text.rfind(' ', 0, chars_per_line)
                        if cut <= 0:
                            cut = chars_per_line
                        lines.append(text[:cut])
                        text = text[cut:].lstrip()
                    for line in lines:
                        try:
                            bbox = draw.textbbox((0, 0), line, font=font)
                            ltw = bbox[2] - bbox[0]
                        except:
                            ltw = len(line) * render_size
                        if align == 'center':
                            lx = x + (w - ltw) // 2
                        else:
                            lx = tx
                        draw.text((lx, ty), line, fill=text_color, font=font)
                        ty += th + 4
                else:
                    draw.text((tx, ty), text, fill=text_color, font=font)
                    ty += th + 6

    # Save
    path = f'slide_previews/slide_{slide_idx + 1:02d}.png'
    img.save(path)
    print(f'Saved {path}')

print(f'\nAll {len(prs.slides)} slide previews generated in slide_previews/')
