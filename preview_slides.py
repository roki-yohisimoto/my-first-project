#!/usr/bin/env python3
"""Generate visual preview images of each slide from the PPTX with Japanese font support."""

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os, textwrap

W, H = 1920, 1080

FONT_JP = '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'
FONT_BOLD = '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'  # same for now
FONT_EMOJI = '/usr/share/fonts/truetype/NotoColorEmoji.ttf'

def emu_to_px(emu_val, slide_dim, px_dim):
    return int(emu_val / slide_dim * px_dim)

def get_font(size, bold=False):
    try:
        return ImageFont.truetype(FONT_JP, size)
    except:
        return ImageFont.load_default()

prs = Presentation('notebooklm-training-slides.pptx')
sw = prs.slide_width
sh = prs.slide_height

os.makedirs('slide_previews', exist_ok=True)

for slide_idx, slide in enumerate(prs.slides):
    img = Image.new('RGB', (W, H), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    # Background
    try:
        bg = slide.background
        if bg.fill.type is not None:
            c = bg.fill.fore_color.rgb
            img = Image.new('RGB', (W, H), (c[0], c[1], c[2]))
            draw = ImageDraw.Draw(img)
    except:
        pass

    # Collect and sort shapes by z-order (order in XML)
    shapes = list(slide.shapes)

    for shape in shapes:
        x = emu_to_px(shape.left, sw, W)
        y = emu_to_px(shape.top, sh, H)
        w = emu_to_px(shape.width, sw, W)
        h = emu_to_px(shape.height, sh, H)

        # Draw fill
        try:
            if hasattr(shape, 'fill'):
                fill = shape.fill
                if fill.type is not None:
                    try:
                        c = fill.fore_color.rgb
                        color = (c[0], c[1], c[2])
                        try:
                            st = shape.shape_type
                            if st == 9:  # OVAL
                                draw.ellipse([x, y, x+w, y+h], fill=color)
                            else:
                                draw.rounded_rectangle([x, y, x+w, y+h], radius=8, fill=color)
                        except:
                            draw.rectangle([x, y, x+w, y+h], fill=color)
                    except:
                        pass
        except:
            pass

        # Draw border
        try:
            line = shape.line
            if line.fill.type is not None:
                lc = line.color.rgb
                lw = max(1, int((line.width or 12700) / 12700))
                border_color = (lc[0], lc[1], lc[2])
                draw.rounded_rectangle([x, y, x+w, y+h], radius=8, outline=border_color, width=lw)
        except:
            pass

        # Draw text
        if shape.has_text_frame:
            tf = shape.text_frame
            ty = y + 6
            for para in tf.paragraphs:
                text = para.text.strip()
                if not text:
                    ty += 10
                    continue

                # Font properties
                fs = 16
                bold = False
                text_color = (51, 51, 51)
                align = 'left'

                # Get from first run if available
                if para.runs:
                    run = para.runs[0]
                    try:
                        if run.font.size:
                            fs = max(8, int(run.font.size / 12700))
                    except:
                        pass
                    try:
                        if run.font.bold:
                            bold = True
                    except:
                        pass
                    try:
                        if run.font.color and run.font.color.rgb:
                            c = run.font.color.rgb
                            text_color = (c[0], c[1], c[2])
                    except:
                        pass
                else:
                    try:
                        if para.font.size:
                            fs = max(8, int(para.font.size / 12700))
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

                render_size = max(10, int(fs * 0.82))
                font = get_font(render_size, bold)

                # Measure text
                try:
                    bbox = draw.textbbox((0, 0), text, font=font)
                    tw = bbox[2] - bbox[0]
                    th = bbox[3] - bbox[1]
                except:
                    tw, th = len(text) * render_size, render_size

                max_w = w - 16

                # Word wrap if needed
                if tw > max_w and max_w > 0:
                    chars_per_line = max(1, int(len(text) * max_w / max(tw, 1)))
                    lines = []
                    remaining = text
                    while remaining:
                        if len(remaining) <= chars_per_line:
                            lines.append(remaining)
                            break
                        cut = chars_per_line
                        # Try to find a good break point
                        for sep in [' ', '、', '。', '・', '）', '」']:
                            pos = remaining.rfind(sep, 0, chars_per_line + 1)
                            if pos > 0:
                                cut = pos + 1
                                break
                        lines.append(remaining[:cut])
                        remaining = remaining[cut:].lstrip()

                    for line in lines:
                        try:
                            bbox = draw.textbbox((0, 0), line, font=font)
                            ltw = bbox[2] - bbox[0]
                        except:
                            ltw = len(line) * render_size
                        if align == 'center':
                            lx = x + (w - ltw) // 2
                        elif align == 'right':
                            lx = x + w - ltw - 8
                        else:
                            lx = x + 8
                        draw.text((lx, ty), line, fill=text_color, font=font)
                        ty += th + 4
                else:
                    if align == 'center':
                        tx = x + (w - tw) // 2
                    elif align == 'right':
                        tx = x + w - tw - 8
                    else:
                        tx = x + 8
                    draw.text((tx, ty), text, fill=text_color, font=font)
                    ty += th + 6

    path = f'slide_previews/slide_{slide_idx + 1:02d}.png'
    img.save(path)
    print(f'Saved {path}')

print(f'\nAll {len(prs.slides)} slides generated.')
