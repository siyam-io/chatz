"""
Generate a premium 10-slide PowerPoint presentation for ChatZ
Uses the Bookreview.pptx theme as a base for consistent styling
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# ─── Color Palette (Premium Dark + Rose Gold) ────────────────────────────────
DARK_BG       = RGBColor(0x0F, 0x0F, 0x1A)   # Deep navy-black
DARK_CARD     = RGBColor(0x1A, 0x1A, 0x2E)   # Card background
ROSE_GOLD     = RGBColor(0xC4, 0x8B, 0x9F)   # App's signature rose
ROSE_LIGHT    = RGBColor(0xE8, 0xC1, 0xD0)   # Lighter rose
ACCENT_PURPLE = RGBColor(0x6C, 0x5C, 0xE7)   # Purple accent
ACCENT_BLUE   = RGBColor(0x00, 0xCE, 0xC9)   # Teal accent
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY    = RGBColor(0xBB, 0xBB, 0xCC)
DIM_TEXT      = RGBColor(0x88, 0x88, 0xAA)
GRADIENT_START= RGBColor(0x2D, 0x1B, 0x69)   # Dark purple
GRADIENT_END  = RGBColor(0x11, 0x99, 0x8E)   # Teal green
SUCCESS_GREEN = RGBColor(0x00, 0xE6, 0x76)
WARN_AMBER    = RGBColor(0xFF, 0xBE, 0x0B)

# Image paths
IMG_DIR = r"d:\pg\chat-z\thesis\image"
IMAGES = {
    "login": os.path.join(IMG_DIR, "image.png"),
    "feed": os.path.join(IMG_DIR, "00-26-58.jpg"),
    "friends": os.path.join(IMG_DIR, "00-27-18.jpg"),
    "chat_list": os.path.join(IMG_DIR, "00-27-22.jpg"),
    "create_group": os.path.join(IMG_DIR, "00-27-26.jpg"),
    "chat_list_group": os.path.join(IMG_DIR, "00-27-30.jpg"),
    "profile": os.path.join(IMG_DIR, "00-27-33.jpg"),
    "group_chat": os.path.join(IMG_DIR, "1783537073967.jpg"),
    "dm_encrypted": os.path.join(IMG_DIR, "1783537073970.jpg"),
}

OUTPUT_PATH = r"d:\pg\chat-z\thesis\pp\ChatZ_Presentation.pptx"

# ─── Helper Functions ─────────────────────────────────────────────────────────

def set_slide_bg(slide, color):
    """Set a solid background color for a slide."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_gradient_bg(slide):
    """Add a gradient rectangle covering entire slide as background."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, Emu(12192000), Emu(6858000)
    )
    shape.fill.gradient()
    shape.fill.gradient_stops[0].color.rgb = RGBColor(0x0A, 0x0A, 0x1E)
    shape.fill.gradient_stops[0].position = 0.0
    shape.fill.gradient_stops[1].color.rgb = RGBColor(0x1E, 0x0A, 0x2E)
    shape.fill.gradient_stops[1].position = 1.0
    shape.line.fill.background()
    return shape

def add_dark_rect(slide, left, top, width, height, color=None):
    """Add a dark card rectangle."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    fill_color = color or DARK_CARD
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    shape.adjustments[0] = 0.05
    return shape

def add_accent_line(slide, left, top, width, color=None):
    """Add a thin accent line / divider."""
    line_color = color or ROSE_GOLD
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, Pt(3)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = line_color
    shape.line.fill.background()
    return shape

def add_circle(slide, left, top, size, color):
    """Add a colored circle."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, left, top, size, size
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def add_text_box(slide, left, top, width, height, text, 
                  font_size=18, color=WHITE, bold=False, alignment=PP_ALIGN.LEFT,
                  font_name="Segoe UI"):
    """Add a text box with formatting."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox

def add_multiline_text(slide, left, top, width, height, lines,
                        font_name="Segoe UI"):
    """Add multiline text - lines is a list of (text, size, color, bold) tuples."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    
    for i, (text, size, color, bold) in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.bold = bold
        run.font.name = font_name
        p.space_after = Pt(6)
    return txBox

def add_bullet_list(slide, left, top, width, height, items, 
                     font_size=14, color=LIGHT_GRAY, bullet_color=ROSE_GOLD):
    """Add a bulleted list."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        
        bullet_run = p.add_run()
        bullet_run.text = "◆  "
        bullet_run.font.size = Pt(font_size - 2)
        bullet_run.font.color.rgb = bullet_color
        bullet_run.font.name = "Segoe UI"
        
        text_run = p.add_run()
        text_run.text = item
        text_run.font.size = Pt(font_size)
        text_run.font.color.rgb = color
        text_run.font.name = "Segoe UI"
        
        p.space_after = Pt(8)
    return txBox

def add_tag(slide, left, top, text, bg_color=ACCENT_PURPLE, text_color=WHITE, font_size=10):
    """Add a small tag/badge."""
    width = Pt(len(text) * 7 + 20)
    height = Pt(24)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.fill.background()
    shape.adjustments[0] = 0.3
    
    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = text_color
    p.font.bold = True
    p.font.name = "Segoe UI"
    p.alignment = PP_ALIGN.CENTER
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    return shape

def add_phone_frame(slide, left, top, width, height, img_path):
    """Add an image with a subtle phone-like frame."""
    frame = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, 
        left - Inches(0.08), top - Inches(0.08), 
        width + Inches(0.16), height + Inches(0.16)
    )
    frame.fill.solid()
    frame.fill.fore_color.rgb = RGBColor(0x2A, 0x2A, 0x3E)
    frame.line.fill.background()
    frame.adjustments[0] = 0.05
    
    pic = slide.shapes.add_picture(img_path, left, top, width, height)
    return pic

# ─── Create Presentation ─────────────────────────────────────────────────────

prs = Presentation()
prs.slide_width = Emu(12192000)
prs.slide_height = Emu(6858000)

SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1: TITLE SLIDE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

bg_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
bg_rect.fill.gradient()
bg_rect.fill.gradient_stops[0].color.rgb = RGBColor(0x0A, 0x05, 0x1A)
bg_rect.fill.gradient_stops[0].position = 0.0
bg_rect.fill.gradient_stops[1].color.rgb = RGBColor(0x1A, 0x0A, 0x30)
bg_rect.fill.gradient_stops[1].position = 1.0
bg_rect.line.fill.background()

add_circle(slide, Inches(-1), Inches(-1), Inches(4), RGBColor(0x6C, 0x5C, 0xE7))
add_circle(slide, Inches(-0.5), Inches(-0.5), Inches(3), RGBColor(0x1A, 0x10, 0x35))
add_circle(slide, Inches(8), Inches(3.5), Inches(5), RGBColor(0xC4, 0x8B, 0x9F))
add_circle(slide, Inches(8.5), Inches(4), Inches(4), RGBColor(0x1A, 0x10, 0x30))

add_accent_line(slide, 0, Pt(0), SLIDE_W, ROSE_GOLD)

add_text_box(slide, Inches(0.8), Inches(1.2), Inches(8), Inches(1.2),
             "ChatZ", font_size=60, color=ROSE_GOLD, bold=True, font_name="Segoe UI Black")

add_text_box(slide, Inches(0.8), Inches(2.2), Inches(7), Inches(0.8),
             "A Real-Time Social Messaging Platform", 
             font_size=28, color=WHITE, bold=False, font_name="Segoe UI Light")

add_text_box(slide, Inches(0.8), Inches(3.0), Inches(7), Inches(0.5),
             "End-to-End Encrypted  \u2022  PostgreSQL Powered  \u2022  Cross-Platform",
             font_size=14, color=DIM_TEXT, bold=False)

add_accent_line(slide, Inches(0.8), Inches(3.8), Inches(2), ROSE_GOLD)

add_multiline_text(slide, Inches(0.8), Inches(4.1), Inches(5), Inches(1.8), [
    ("Software Development Project", 12, DIM_TEXT, False),
    ("", 6, DIM_TEXT, False),
    ("Developed by: S. Siyam", 16, WHITE, True),
    ("Supervisor: [Supervisor Name]", 14, LIGHT_GRAY, False),
    ("Department of Computer Science & Engineering", 12, DIM_TEXT, False),
])

add_phone_frame(slide, Inches(7.8), Inches(0.8), Inches(2.2), Inches(4.8), IMAGES["feed"])

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2: PROJECT OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_tag(slide, Inches(0.6), Inches(0.4), "01  PROJECT OVERVIEW", ACCENT_PURPLE)

add_text_box(slide, Inches(0.6), Inches(0.9), Inches(9), Inches(0.7),
             "What is ChatZ?", font_size=36, color=WHITE, bold=True, font_name="Segoe UI Black")

add_accent_line(slide, Inches(0.6), Inches(1.6), Inches(1.5), ROSE_GOLD)

add_text_box(slide, Inches(0.6), Inches(1.9), Inches(9), Inches(1.2),
             "ChatZ is a full-stack, real-time social messaging platform built with React Native (Expo) and Node.js/Express. It combines secure private messaging with social networking features, powered by PostgreSQL and WebSocket technology.",
             font_size=14, color=LIGHT_GRAY)

card_data = [
    ("\U0001f512", "E2E Encryption", "NaCl/TweetNaCl cryptography\nfor all private messages"),
    ("\U0001f4ac", "Real-Time Chat", "Socket.IO powered instant\nmessaging with typing indicators"),
    ("\U0001f465", "Social Feed", "Posts, reactions, comments,\nshares & stories"),
    ("\U0001f6e1\ufe0f", "Admin Panel", "Next.js admin dashboard\nwith role-based access"),
]

card_w = Inches(2.3)
card_h = Inches(2.2)
start_x = Inches(0.6)
gap = Inches(0.25)

for i, (icon, title, desc) in enumerate(card_data):
    x = start_x + i * (card_w + gap)
    y = Inches(3.3)
    
    card = add_dark_rect(slide, x, y, card_w, card_h, RGBColor(0x16, 0x16, 0x2B))
    
    add_text_box(slide, x + Inches(0.2), y + Inches(0.2), Inches(0.6), Inches(0.5),
                 icon, font_size=28, color=WHITE)
    
    add_text_box(slide, x + Inches(0.2), y + Inches(0.7), card_w - Inches(0.4), Inches(0.4),
                 title, font_size=15, color=ROSE_LIGHT, bold=True)
    
    add_text_box(slide, x + Inches(0.2), y + Inches(1.1), card_w - Inches(0.4), Inches(1.0),
                 desc, font_size=11, color=DIM_TEXT)

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3: SYSTEM ARCHITECTURE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_tag(slide, Inches(0.6), Inches(0.4), "02  ARCHITECTURE", ACCENT_PURPLE)
add_text_box(slide, Inches(0.6), Inches(0.9), Inches(9), Inches(0.7),
             "System Architecture", font_size=36, color=WHITE, bold=True, font_name="Segoe UI Black")
add_accent_line(slide, Inches(0.6), Inches(1.6), Inches(1.5), ROSE_GOLD)

tiers = [
    ("\U0001f4f1 Frontend", "React Native + Expo", [
        "React Navigation (Stack/Tabs)",
        "Zustand State Management",
        "Socket.IO Client",
        "TweetNaCl Encryption",
        "Expo Image Picker",
        "Dark/Light Theme Support",
    ], ROSE_GOLD),
    ("\u2699\ufe0f Backend", "Node.js + Express v5", [
        "RESTful API (v1)",
        "Socket.IO WebSocket Server",
        "JWT Authentication (30d)",
        "Zod Request Validation",
        "Helmet Security Headers",
        "Rate Limiting Protection",
    ], ACCENT_PURPLE),
    ("\U0001f5c4\ufe0f Database", "PostgreSQL + Cloud", [
        "15+ Tables with UUIDs",
        "Indexed Queries (B-tree)",
        "Migration System (10+)",
        "Cloudinary Media Storage",
        "Neon Serverless Postgres",
        "Render.com Deployment",
    ], ACCENT_BLUE),
]

tier_w = Inches(3.1)
tier_h = Inches(3.8)

for i, (title, subtitle, items, color) in enumerate(tiers):
    x = Inches(0.5) + i * (tier_w + Inches(0.2))
    y = Inches(2.0)
    
    card = add_dark_rect(slide, x, y, tier_w, tier_h, RGBColor(0x14, 0x14, 0x28))
    add_accent_line(slide, x, y, tier_w, color)
    
    add_text_box(slide, x + Inches(0.2), y + Inches(0.15), tier_w - Inches(0.4), Inches(0.4),
                 title, font_size=18, color=color, bold=True)
    add_text_box(slide, x + Inches(0.2), y + Inches(0.55), tier_w - Inches(0.4), Inches(0.3),
                 subtitle, font_size=11, color=DIM_TEXT)
    
    add_bullet_list(slide, x + Inches(0.15), y + Inches(0.9), tier_w - Inches(0.3), tier_h - Inches(1.2),
                    items, font_size=11, color=LIGHT_GRAY, bullet_color=color)

for i in range(2):
    x = Inches(0.5) + (i + 1) * (tier_w + Inches(0.2)) - Inches(0.3)
    add_text_box(slide, x, Inches(3.5), Inches(0.5), Inches(0.5),
                 "\u27f7", font_size=24, color=ROSE_GOLD, alignment=PP_ALIGN.CENTER)

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4: TECHNOLOGY STACK
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_tag(slide, Inches(0.6), Inches(0.4), "03  TECH STACK", ACCENT_PURPLE)
add_text_box(slide, Inches(0.6), Inches(0.9), Inches(9), Inches(0.7),
             "Technology Stack", font_size=36, color=WHITE, bold=True, font_name="Segoe UI Black")
add_accent_line(slide, Inches(0.6), Inches(1.6), Inches(1.5), ROSE_GOLD)

tech_groups = [
    ("Frontend Technologies", [
        ("React Native", "v0.81.5", ROSE_GOLD),
        ("Expo SDK", "v54.0", ROSE_GOLD),
        ("React Navigation", "v7.x", ROSE_GOLD),
        ("Zustand", "v5.0", ROSE_GOLD),
        ("Socket.IO Client", "v4.8", ROSE_GOLD),
        ("TweetNaCl", "v1.0", ROSE_GOLD),
    ]),
    ("Backend Technologies", [
        ("Node.js", "Runtime", ACCENT_PURPLE),
        ("Express.js", "v5.2.1", ACCENT_PURPLE),
        ("PostgreSQL", "via pg", ACCENT_PURPLE),
        ("Socket.IO", "v4.8.3", ACCENT_PURPLE),
        ("JWT Auth", "v9.0.3", ACCENT_PURPLE),
        ("Zod", "v4.4", ACCENT_PURPLE),
    ]),
    ("DevOps & Cloud", [
        ("Render.com", "Hosting", ACCENT_BLUE),
        ("Neon DB", "Serverless PG", ACCENT_BLUE),
        ("Cloudinary", "Media CDN", ACCENT_BLUE),
        ("GitHub", "Version Control", ACCENT_BLUE),
        ("Helmet", "Security", ACCENT_BLUE),
        ("Nodemon", "Dev Server", ACCENT_BLUE),
    ]),
]

col_w = Inches(3.1)
for ci, (group_title, techs) in enumerate(tech_groups):
    x = Inches(0.5) + ci * (col_w + Inches(0.2))
    
    add_text_box(slide, x, Inches(2.0), col_w, Inches(0.4),
                 group_title, font_size=16, color=techs[0][2], bold=True)
    
    add_accent_line(slide, x, Inches(2.4), Inches(1), techs[0][2])
    
    for ti, (name, version, color) in enumerate(techs):
        y = Inches(2.6) + ti * Inches(0.6)
        
        dot = add_circle(slide, x + Inches(0.05), y + Inches(0.12), Pt(8), color)
        
        add_text_box(slide, x + Inches(0.3), y, Inches(1.8), Inches(0.35),
                     name, font_size=13, color=WHITE, bold=True)
        add_text_box(slide, x + Inches(2.1), y, Inches(0.9), Inches(0.35),
                     version, font_size=11, color=DIM_TEXT)

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5: DATABASE DESIGN
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_tag(slide, Inches(0.6), Inches(0.4), "04  DATABASE DESIGN", ACCENT_PURPLE)
add_text_box(slide, Inches(0.6), Inches(0.9), Inches(9), Inches(0.7),
             "PostgreSQL Schema Design", font_size=36, color=WHITE, bold=True, font_name="Segoe UI Black")
add_accent_line(slide, Inches(0.6), Inches(1.6), Inches(1.5), ROSE_GOLD)

tables = [
    ("users", "id, name, email, password,\navatar, bio, public_key,\ncover_photo", ROSE_GOLD),
    ("messages", "id, sender_id, receiver_id,\nciphertext, nonce,\nis_encrypted", ACCENT_PURPLE),
    ("groups", "id, name, creator_id,\navatar", ACCENT_BLUE),
    ("posts", "id, user_id, content,\nimage, images (JSONB),\noriginal_post_id", SUCCESS_GREEN),
    ("stories", "id, user_id, image, text\n(24hr auto-cleanup)", WARN_AMBER),
    ("friend_requests", "id, sender_id, receiver_id,\nstatus (ENUM)", ROSE_GOLD),
]

tbl_w = Inches(3.0)
tbl_h = Inches(1.6)

for i, (tbl_name, fields, color) in enumerate(tables):
    row = i // 3
    col = i % 3
    x = Inches(0.5) + col * (tbl_w + Inches(0.2))
    y = Inches(2.0) + row * (tbl_h + Inches(0.2))
    
    card = add_dark_rect(slide, x, y, tbl_w, tbl_h, RGBColor(0x14, 0x14, 0x28))
    add_accent_line(slide, x, y, tbl_w, color)
    
    add_text_box(slide, x + Inches(0.15), y + Inches(0.1), tbl_w - Inches(0.3), Inches(0.35),
                 tbl_name, font_size=15, color=color, bold=True, font_name="Consolas")
    
    add_text_box(slide, x + Inches(0.15), y + Inches(0.5), tbl_w - Inches(0.3), tbl_h - Inches(0.6),
                 fields, font_size=10, color=DIM_TEXT, font_name="Consolas")

stats_bar = add_dark_rect(slide, Inches(0.5), Inches(5.6), Inches(9.2), Inches(0.7), RGBColor(0x12, 0x12, 0x25))

stats = [
    ("15+", "Tables"),
    ("10+", "Migrations"),
    ("UUID", "Primary Keys"),
    ("B-tree", "Indexes"),
    ("ENUM", "Types"),
]

for si, (num, label) in enumerate(stats):
    sx = Inches(0.8) + si * Inches(1.85)
    add_text_box(slide, sx, Inches(5.65), Inches(0.8), Inches(0.35),
                 num, font_size=16, color=ROSE_GOLD, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, sx, Inches(5.95), Inches(0.8), Inches(0.25),
                 label, font_size=9, color=DIM_TEXT, alignment=PP_ALIGN.CENTER)

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6: SECURITY
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_tag(slide, Inches(0.6), Inches(0.4), "05  SECURITY", ACCENT_PURPLE)
add_text_box(slide, Inches(0.6), Inches(0.9), Inches(6), Inches(0.7),
             "Security & Encryption", font_size=36, color=WHITE, bold=True, font_name="Segoe UI Black")
add_accent_line(slide, Inches(0.6), Inches(1.6), Inches(1.5), ROSE_GOLD)

security_features = [
    ("\U0001f510 End-to-End Encryption", "All private messages are encrypted using TweetNaCl\n(NaCl) public-key cryptography. Messages are stored\nas ciphertext \u2014 the server never sees plaintext."),
    ("\U0001f6e1\ufe0f JWT Authentication", "JSON Web Tokens with 30-day expiry, bcrypt password\nhashing (12 salt rounds), and secure token validation."),
    ("\u26a1 Rate Limiting", "Global rate limit: 1000 req/15min. Auth endpoints:\n20 req/15min. Prevents brute-force and DDoS attacks."),
    ("\U0001f512 Helmet Security", "HTTP security headers including CSP, X-Frame-Options,\nand HSTS to prevent common web vulnerabilities."),
]

for i, (title, desc) in enumerate(security_features):
    y = Inches(1.9) + i * Inches(1.15)
    
    card = add_dark_rect(slide, Inches(0.5), y, Inches(6.2), Inches(1.0), RGBColor(0x14, 0x14, 0x28))
    
    add_text_box(slide, Inches(0.7), y + Inches(0.05), Inches(5.8), Inches(0.35),
                 title, font_size=14, color=ROSE_LIGHT, bold=True)
    add_text_box(slide, Inches(0.7), y + Inches(0.35), Inches(5.8), Inches(0.6),
                 desc, font_size=10, color=DIM_TEXT)

add_phone_frame(slide, Inches(7.5), Inches(1.0), Inches(2.2), Inches(4.8), IMAGES["dm_encrypted"])

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7: UI SHOWCASE - SOCIAL
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_tag(slide, Inches(0.6), Inches(0.4), "06  UI SHOWCASE", ACCENT_PURPLE)
add_text_box(slide, Inches(0.6), Inches(0.9), Inches(9), Inches(0.7),
             "User Interface \u2014 Social Features", font_size=36, color=WHITE, bold=True, font_name="Segoe UI Black")
add_accent_line(slide, Inches(0.6), Inches(1.6), Inches(1.5), ROSE_GOLD)

screenshots = [
    (IMAGES["feed"], "Social Feed", "Posts, Stories, Reactions"),
    (IMAGES["friends"], "Friends & People", "Search, Add, Manage"),
    (IMAGES["profile"], "User Profile", "Bio, Posts, Statistics"),
]

for i, (img_path, title, sub) in enumerate(screenshots):
    x = Inches(0.4) + i * Inches(3.3)
    
    add_phone_frame(slide, x + Inches(0.3), Inches(2.0), Inches(1.9), Inches(4.0), img_path)
    
    add_text_box(slide, x, Inches(6.1), Inches(2.5), Inches(0.3),
                 title, font_size=13, color=ROSE_LIGHT, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x, Inches(6.4), Inches(2.5), Inches(0.25),
                 sub, font_size=10, color=DIM_TEXT, alignment=PP_ALIGN.CENTER)

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8: UI SHOWCASE - MESSAGING
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_tag(slide, Inches(0.6), Inches(0.4), "07  MESSAGING", ACCENT_PURPLE)
add_text_box(slide, Inches(0.6), Inches(0.9), Inches(9), Inches(0.7),
             "Messaging System", font_size=36, color=WHITE, bold=True, font_name="Segoe UI Black")
add_accent_line(slide, Inches(0.6), Inches(1.6), Inches(1.5), ROSE_GOLD)

msg_screenshots = [
    (IMAGES["chat_list"], "Chat List", "Conversations Overview"),
    (IMAGES["dm_encrypted"], "Private Chat", "E2E Encrypted DMs"),
    (IMAGES["group_chat"], "Group Chat", "Multi-user Messaging"),
]

for i, (img_path, title, sub) in enumerate(msg_screenshots):
    x = Inches(0.4) + i * Inches(3.3)
    
    add_phone_frame(slide, x + Inches(0.3), Inches(2.0), Inches(1.9), Inches(4.0), img_path)
    
    add_text_box(slide, x, Inches(6.1), Inches(2.5), Inches(0.3),
                 title, font_size=13, color=ROSE_LIGHT, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x, Inches(6.4), Inches(2.5), Inches(0.25),
                 sub, font_size=10, color=DIM_TEXT, alignment=PP_ALIGN.CENTER)

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9: API & REAL-TIME
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_tag(slide, Inches(0.6), Inches(0.4), "08  API DESIGN", ACCENT_PURPLE)
add_text_box(slide, Inches(0.6), Inches(0.9), Inches(9), Inches(0.7),
             "RESTful API & Real-Time Events", font_size=36, color=WHITE, bold=True, font_name="Segoe UI Black")
add_accent_line(slide, Inches(0.6), Inches(1.6), Inches(1.5), ROSE_GOLD)

add_text_box(slide, Inches(0.6), Inches(2.0), Inches(4.5), Inches(0.4),
             "REST API Endpoints (v1)", font_size=16, color=ROSE_LIGHT, bold=True)

api_endpoints = [
    ("POST", "/api/v1/auth/register", "User Registration"),
    ("POST", "/api/v1/auth/login", "JWT Login"),
    ("GET", "/api/v1/chat/messages/:id", "Fetch Messages"),
    ("POST", "/api/v1/posts", "Create Post"),
    ("GET", "/api/v1/friends", "List Friends"),
    ("POST", "/api/v1/groups", "Create Group"),
    ("POST", "/api/v1/upload", "Media Upload"),
    ("GET", "/api/v1/stories", "Get Stories"),
    ("POST", "/api/v1/reports", "Report Content"),
    ("GET", "/api/v1/admin/users", "Admin: Users"),
]

for i, (method, path, desc) in enumerate(api_endpoints):
    y = Inches(2.5) + i * Inches(0.37)
    
    method_colors = {"GET": SUCCESS_GREEN, "POST": ACCENT_BLUE, "PUT": WARN_AMBER}
    mc = method_colors.get(method, WHITE)
    
    add_text_box(slide, Inches(0.6), y, Inches(0.6), Inches(0.3),
                 method, font_size=9, color=mc, bold=True, font_name="Consolas")
    add_text_box(slide, Inches(1.2), y, Inches(2.5), Inches(0.3),
                 path, font_size=9, color=LIGHT_GRAY, font_name="Consolas")
    add_text_box(slide, Inches(3.8), y, Inches(1.5), Inches(0.3),
                 desc, font_size=9, color=DIM_TEXT)

add_text_box(slide, Inches(5.8), Inches(2.0), Inches(4), Inches(0.4),
             "Socket.IO Events (Real-Time)", font_size=16, color=ACCENT_BLUE, bold=True)

socket_events = [
    ("\u26a1", "send_message", "Send encrypted DM"),
    ("\U0001f4e8", "receive_message", "Receive incoming DM"),
    ("\u2705", "message_sent", "Delivery confirmation"),
    ("\u2328\ufe0f", "typing", "Typing indicator"),
    ("\U0001f465", "send_group_message", "Group messaging"),
    ("\U0001f4e1", "receive_group_message", "Group msg receive"),
    ("\U0001f504", "conversation_update", "Update chat list"),
    ("\U0001f517", "connected", "Socket authenticated"),
]

for i, (icon, event, desc) in enumerate(socket_events):
    y = Inches(2.5) + i * Inches(0.43)
    
    add_text_box(slide, Inches(5.8), y, Inches(0.4), Inches(0.35),
                 icon, font_size=12, color=WHITE)
    add_text_box(slide, Inches(6.2), y, Inches(2.2), Inches(0.35),
                 event, font_size=11, color=WHITE, bold=True, font_name="Consolas")
    add_text_box(slide, Inches(8.5), y, Inches(1.5), Inches(0.35),
                 desc, font_size=9, color=DIM_TEXT)

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10: REQUIREMENTS
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)
add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_tag(slide, Inches(0.6), Inches(0.4), "09  REQUIREMENTS", ACCENT_PURPLE)
add_text_box(slide, Inches(0.6), Inches(0.9), Inches(9), Inches(0.7),
             "Software & Hardware Requirements", font_size=36, color=WHITE, bold=True, font_name="Segoe UI Black")
add_accent_line(slide, Inches(0.6), Inches(1.6), Inches(1.5), ROSE_GOLD)

sw_card = add_dark_rect(slide, Inches(0.5), Inches(2.0), Inches(4.6), Inches(4.2), RGBColor(0x14, 0x14, 0x28))
add_accent_line(slide, Inches(0.5), Inches(2.0), Inches(4.6), ROSE_GOLD)

add_text_box(slide, Inches(0.7), Inches(2.15), Inches(4), Inches(0.4),
             "\U0001f4bb  Software Requirements", font_size=16, color=ROSE_LIGHT, bold=True)

sw_items = [
    "Frontend: Expo (v54.0) + React Native (v0.81.5)",
    "Backend: Express.js (v5.2.1) + Node.js",
    "Database: PostgreSQL (Neon Serverless)",
    "Auth: JWT (v9.0.3) + Bcrypt (12 rounds)",
    "Real-Time: Socket.IO (v4.8.3)",
    "Encryption: TweetNaCl (v1.0.3)",
    "Media: Cloudinary CDN + Multer",
    "Validation: Zod (v4.4.3)",
    "Admin: Next.js Dashboard",
    "Deployment: Render.com (Docker)",
]

add_bullet_list(slide, Inches(0.65), Inches(2.55), Inches(4.2), Inches(3.5),
                sw_items, font_size=11, color=LIGHT_GRAY, bullet_color=ROSE_GOLD)

hw_card = add_dark_rect(slide, Inches(5.3), Inches(2.0), Inches(4.6), Inches(2.8), RGBColor(0x14, 0x14, 0x28))
add_accent_line(slide, Inches(5.3), Inches(2.0), Inches(4.6), ACCENT_BLUE)

add_text_box(slide, Inches(5.5), Inches(2.15), Inches(4), Inches(0.4),
             "\U0001f5a5\ufe0f  Hardware Requirements", font_size=16, color=ACCENT_BLUE, bold=True)

hw_items = [
    "Processor: Single-Core 1.0 GHz+",
    "RAM: 512 MB - 1 GB minimum",
    "Storage: 1 GB free space",
    "Android 6.0+ / iOS 13+",
    "Network: Stable internet connection",
]

add_bullet_list(slide, Inches(5.45), Inches(2.55), Inches(4.2), Inches(2.2),
                hw_items, font_size=12, color=LIGHT_GRAY, bullet_color=ACCENT_BLUE)

time_card = add_dark_rect(slide, Inches(5.3), Inches(5.0), Inches(4.6), Inches(1.2), RGBColor(0x14, 0x14, 0x28))
add_accent_line(slide, Inches(5.3), Inches(5.0), Inches(4.6), WARN_AMBER)

add_text_box(slide, Inches(5.5), Inches(5.15), Inches(4), Inches(0.4),
             "\u23f1\ufe0f  Development Timeline", font_size=16, color=WARN_AMBER, bold=True)
add_text_box(slide, Inches(5.5), Inches(5.5), Inches(4), Inches(0.5),
             "Total Duration: 3 Months\nMethodology: Agile / Iterative Development", 
             font_size=12, color=LIGHT_GRAY)

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 (BONUS): THANK YOU
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

bg_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
bg_rect.fill.gradient()
bg_rect.fill.gradient_stops[0].color.rgb = RGBColor(0x0A, 0x05, 0x1A)
bg_rect.fill.gradient_stops[0].position = 0.0
bg_rect.fill.gradient_stops[1].color.rgb = RGBColor(0x1A, 0x0A, 0x30)
bg_rect.fill.gradient_stops[1].position = 1.0
bg_rect.line.fill.background()

add_circle(slide, Inches(7.5), Inches(-1), Inches(5), RGBColor(0xC4, 0x8B, 0x9F))
add_circle(slide, Inches(8), Inches(-0.5), Inches(4), RGBColor(0x15, 0x0A, 0x28))
add_circle(slide, Inches(-2), Inches(3), Inches(5), RGBColor(0x6C, 0x5C, 0xE7))
add_circle(slide, Inches(-1.5), Inches(3.5), Inches(4), RGBColor(0x12, 0x08, 0x22))

add_accent_line(slide, 0, 0, SLIDE_W, ROSE_GOLD)

add_text_box(slide, Inches(1.5), Inches(1.5), Inches(7), Inches(1.0),
             "Thank You", font_size=54, color=ROSE_GOLD, bold=True, 
             font_name="Segoe UI Black", alignment=PP_ALIGN.CENTER)

add_text_box(slide, Inches(1.5), Inches(2.6), Inches(7), Inches(0.6),
             "Questions & Discussion", font_size=24, color=WHITE, 
             font_name="Segoe UI Light", alignment=PP_ALIGN.CENTER)

add_accent_line(slide, Inches(3.5), Inches(3.4), Inches(3), ROSE_GOLD)

summary_card = add_dark_rect(slide, Inches(2), Inches(3.8), Inches(6), Inches(2.0), RGBColor(0x14, 0x14, 0x28))

add_multiline_text(slide, Inches(2.3), Inches(3.9), Inches(5.4), Inches(1.8), [
    ("ChatZ \u2014 Real-Time Social Messaging Platform", 14, ROSE_LIGHT, True),
    ("", 4, DIM_TEXT, False),
    ("React Native  \u2022  Express.js  \u2022  PostgreSQL  \u2022  Socket.IO", 12, DIM_TEXT, False),
    ("End-to-End Encrypted  \u2022  Admin Dashboard  \u2022  Cloud Deployed", 12, DIM_TEXT, False),
    ("", 6, DIM_TEXT, False),
    ("GitHub: github.com/Ssiyam0123/chatz", 12, ACCENT_BLUE, False),
])

add_phone_frame(slide, Inches(8.2), Inches(2.0), Inches(1.5), Inches(3.3), IMAGES["login"])

add_accent_line(slide, 0, Emu(SLIDE_H.emu - Pt(4).emu), SLIDE_W, ROSE_GOLD)


# ─── Save ─────────────────────────────────────────────────────────────────────
prs.save(OUTPUT_PATH)
print(f"Presentation saved to: {OUTPUT_PATH}")
print(f"Total slides: {len(prs.slides)}")
