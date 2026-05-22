"""
clean_ui.py
  1. Remove todos os emojis dos templates (exceto ✓ ✔ ✕ ✖ ✘ — símbolos funcionais)
  2. Substitui degradês decorativos por cores sólidas
"""
import glob, os, re

BASE   = "templates"
SKIP_EMOJI = {"_translate.html"}   # mantém flags da tradução

# ─── EMOJIS A REMOVER ────────────────────────────────────────────────────────
EMOJIS = [
    '⏳','⏻','⚖','⚙','⚠','⛔','✅','✉','✍','✏','➕','⬇','❌',
    '🆕','🌐','🐛','👤','👥','💰','💾','📁','📄','📊','📋',
    '📌','📍','📎','📜','📝','📥','📧','📨','📩','📬','📷',
    '🔄','🔍','🔎','🔑','🔒','🔴','🗑','🚨','🚪','🚫','🟢',
    '🇦','🇧','🇨','🇩','🇪','🇫','🇳','🇷','🇸','🇺',
]

# ─── DEGRADÊS → COR SÓLIDA ───────────────────────────────────────────────────
# Formato: (padrão_exato, substituição)
# Ordem importa: mais específico primeiro
GRADIENT_SUBS = [
    # ── Fundos brancos/quase-brancos (cards, panels) ──────────────────────
    ('linear-gradient(180deg,#fff,#f8fafc)',   '#fff'),
    ('linear-gradient(180deg,#fff,#faf5ff)',   '#fff'),
    ('linear-gradient(180deg,#fff,#fcfcfd)',   '#fff'),
    ('linear-gradient(180deg,#fff,#fffdf8)',   '#fff'),
    ('linear-gradient(180deg,#fff,#fffef7)',   '#fff'),
    ('linear-gradient(180deg,#ffffff,#f8fafc)','#ffffff'),
    ('linear-gradient(180deg,#f9fafb,#f1f5f9)','#f9fafb'),
    ('linear-gradient(180deg,#fff8db,#ffefab)', '#fff8db'),
    ('linear-gradient(180deg,#fffdf5,#fff8db)', '#fffdf5'),
    ('linear-gradient(180deg,#fff8db,#ffd84d)', '#ffcc00'),

    # ── Botões vermelhos ───────────────────────────────────────────────────
    ('linear-gradient(135deg,#d40511,#93000a)', '#d40511'),
    ('linear-gradient(135deg,#d40511,#b1030d)', '#d40511'),
    ('linear-gradient(180deg,#d40511,#93000a)', '#d40511'),
    ('linear-gradient(180deg,#d40511,#b1030d)', '#d40511'),
    ('linear-gradient(180deg,#dc2626,#b91c1c)', '#dc2626'),
    ('linear-gradient(180deg,#b91c1c,#7f1d1d)', '#b91c1c'),

    # ── Botões verdes ──────────────────────────────────────────────────────
    ('linear-gradient(135deg,#16a34a,#15803d)', '#16a34a'),
    ('linear-gradient(180deg,#16a34a,#15803d)', '#16a34a'),
    ('linear-gradient(180deg,#15803d,#166534)', '#15803d'),

    # ── Botões amarelos ────────────────────────────────────────────────────
    ('linear-gradient(180deg,#ffcc00,#f1be00)', '#ffcc00'),
    ('linear-gradient(180deg,#f59e0b,#d97706)', '#f59e0b'),

    # ── Botões escuros/cinzas ──────────────────────────────────────────────
    ('linear-gradient(135deg,#1a1a1a,#2d2d2d)', '#1a1a1a'),
    ('linear-gradient(135deg,#1f2937,#111827)', '#1f2937'),
    ('linear-gradient(180deg,#374151,#1f2937)', '#374151'),
    ('linear-gradient(180deg,#6b7280,#4b5563)', '#6b7280'),
    ('linear-gradient(180deg,#3e434a,#262a2f)', '#3e434a'),

    # ── Azuis ─────────────────────────────────────────────────────────────
    ('linear-gradient(135deg,#0f2040,#1e3a5f)', '#0f2040'),
    ('linear-gradient(135deg,#1e3a5f,#0f2742)', '#1e3a5f'),
    ('linear-gradient(180deg,#0078d4,#005fa3)', '#0078d4'),

    # ── Marrom (warning/amber) ─────────────────────────────────────────────
    ('linear-gradient(135deg,#92400e,#451a03)', '#92400e'),
]

# ─── PROCESSAMENTO ────────────────────────────────────────────────────────────
files = (
    glob.glob(os.path.join(BASE, "*.html")) +
    glob.glob(os.path.join(BASE, "chaves", "*.html"))
)

total_files = 0
for path in sorted(files):
    name = os.path.basename(path)
    with open(path, encoding="utf-8") as f:
        raw = f.read()
    new_raw = raw

    # 1. Remove emojis (exceto _translate.html que tem bandeiras)
    if name not in SKIP_EMOJI:
        for emoji in EMOJIS:
            new_raw = new_raw.replace(emoji, '')
        # Remove espaços duplos deixados pelos emojis no meio de texto
        new_raw = re.sub(r'(\s) +(\s)', r'\1\2', new_raw)

    # 2. Substitui degradês
    for old, new in GRADIENT_SUBS:
        new_raw = new_raw.replace(old, new)

    if new_raw != raw:
        with open(path, "w", encoding="utf-8") as f:
            f.write(new_raw)
        total_files += 1
        print(f"  OK: {name}")

print(f"\nTotal: {total_files} arquivos modificados")
