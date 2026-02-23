#!/usr/bin/env python3
"""
PA Rap Vocal Engine — extract lyrics, apply cadence/rhyme/TTS formatting,
then render one AIFF per phrase via macOS `say`.
Liquid DnB mode (172 BPM): half-time grid, snare alignment, breath space.
Output: PA-Rap-Vocal-Phrases/001.aiff, 002.aiff, ...
"""
import re
import random
import subprocess
import tempfile
from pathlib import Path

try:
    import pronouncing  # optional: pip install pronouncing (CMU dict)
    HAS_PRONOUNCING = True
except ImportError:
    HAS_PRONOUNCING = False

# ==============================
# CONFIG
# ==============================

MODULES = [
    "PA-Module1-Rap.md",
    "PA-Module2-Rap.md",
    "PA-Module3-Rap.md",
    "PA-Module4-Rap.md",
    "PA-Module5-Rap.md",
]

DIR = Path(__file__).resolve().parent
PHRASES_DIR = DIR / "PA-Rap-Vocal-Phrases"
MAX_WORDS_PER_PHRASE = 5

# Enhanced flow (set False to use raw lyrics only)
BPM = 92
BEATS_PER_BAR = 4
SYLLABLE_TARGET = 10
BARS_PER_SECTION = 8
PUNCH_WORD_PROB = 0.25
INTERNAL_RHYME_PROB = 0.2
TIGHTEN_LINES = True
USE_PHONETIC_FOR_SAY = False  # capitalizing long words can make `say` mispronounce

PUNCH_WORDS = [
    "facts", "clean", "tight", "locked", "sharp",
    "real talk", "straight",
]

INTERNAL_RHYME_PAIRS = [
    ("model", "nodal"),
    ("data", "later"),
    ("bias", "science"),
    ("trade", "made"),
    ("fit", "split"),
    ("flow", "grow"),
    ("mean", "lean"),
]

RANDOM_SEED = 42  # reproducible output

# ==============================
# LIQUID DnB @ 172 BPM
# Half-time feel (~86 BPM cadence), float + glide + space
# ==============================

ENABLE_LIQUID_DNB = True
DNB_BPM = 172
# 1 beat ≈ 0.349 s, 1 bar (4 beats) ≈ 1.395 s
LIQUID_WORDS_PER_BAR = 8  # half-time: 8 active slots per bar
LIQUID_SLOT_POSITIONS = [0, 3, 4, 7, 8, 11, 12, 15]  # 16th grid, snare on 4 & 12
SNARE_16TH_INDICES = [4, 12]  # beat 2 & 4

LIQUID_SETTINGS = {
    "target_syllables": 9,
    "internal_rhyme_prob": 0.15,
    "breath_probability": 0.25,
    "stress_mode": "trochaic",
    "elongate_vowels": True,
}
USE_ELONGATION_FOR_SAY = False  # "flow..." can sound odd in TTS

# ==============================
# SYLLABLE ESTIMATION
# ==============================


def estimate_syllables(word: str) -> int:
    word = word.lower()
    vowels = "aeiouy"
    count = 0
    prev_char_was_vowel = False
    for char in word:
        if char in vowels:
            if not prev_char_was_vowel:
                count += 1
            prev_char_was_vowel = True
        else:
            prev_char_was_vowel = False
    if word.endswith("e"):
        count = max(1, count - 1)
    return max(count, 1)


def line_syllables(line: str) -> int:
    return sum(estimate_syllables(w) for w in line.split())


# ==============================
# CADENCE TIGHTENING
# ==============================


def tighten_line(line: str) -> str:
    words = line.split()
    while line_syllables(" ".join(words)) > SYLLABLE_TARGET and len(words) > 4:
        words.pop(-2)
    if random.random() < PUNCH_WORD_PROB:
        words.append(random.choice(PUNCH_WORDS))
    return " ".join(words)


# ==============================
# INTERNAL RHYME
# ==============================


def add_internal_rhyme(line: str) -> str:
    if random.random() < INTERNAL_RHYME_PROB:
        pair = random.choice(INTERNAL_RHYME_PAIRS)
        return f"{line} — {pair[0]}, {pair[1]}"
    return line


# ==============================
# PHONETIC STRESS (optional for say)
# ==============================


def phonetic_mark(line: str) -> str:
    words = line.split()
    stressed = []
    for w in words:
        if len(w) > 6:
            stressed.append(w.upper())
        else:
            stressed.append(w)
    return " ".join(stressed)


# ==============================
# TTS-READY FORMATTING (say-safe: no XML)
# ==============================


def tts_format(line: str, engine: str = "say") -> str:
    if engine == "say":
        return line
    if engine == "elevenlabs":
        return line.replace(",", ", ... ").replace("—", "... ")
    if engine == "openai":
        return line  # SSML would go here; say can't use it
    if engine == "coqui":
        return line
    return line


# ==============================
# BEAT-ALIGNED BAR LABELS (for export only, not spoken)
# ==============================


def format_bars(lines: list[str]) -> list[str]:
    return [f"[Bar {(i % BARS_PER_SECTION) + 1:02d}] {line}" for i, line in enumerate(lines)]


# ==============================
# LIQUID DnB FLOW OPTIMIZER (172 BPM, half-time)
# ==============================


def quantize_liquid(words: list[str]) -> list[str]:
    """Half-time grid: 8 word slots at [0,3,4,7,8,11,12,15], rest dots."""
    grid = ["."] * 16
    for i, word in enumerate(words[:LIQUID_WORDS_PER_BAR]):
        if i < len(LIQUID_SLOT_POSITIONS):
            grid[LIQUID_SLOT_POSITIONS[i]] = word
    return grid


def liquid_snare_accent(grid: list[str]) -> list[str]:
    """Soft accent on snare (2 & 4): *word* not ALL CAPS."""
    for pos in SNARE_16TH_INDICES:
        if pos < len(grid) and grid[pos] not in (".", "_"):
            grid[pos] = f"*{grid[pos]}*"
    return grid


def liquid_breath(grid: list[str]) -> list[str]:
    """Random breath space for float; don't fill every slot."""
    p = LIQUID_SETTINGS["breath_probability"]
    for i in range(len(grid)):
        if grid[i] not in (".", "_") and random.random() < p:
            grid[i] = "."
    return grid


def elongate(line: str) -> str:
    """Subtle vowel stretch for liquid vibe (word-final vowels)."""
    return re.sub(r"([aeiou])\b", r"\1...", line, flags=re.IGNORECASE)


def soft_internal_rhyme(line: str, probability: float | None = None) -> str:
    """Gentle internal rhyme; uses pronouncing if available else INTERNAL_RHYME_PAIRS."""
    prob = probability if probability is not None else LIQUID_SETTINGS["internal_rhyme_prob"]
    words = line.split()
    if not HAS_PRONOUNCING:
        if random.random() < prob:
            pair = random.choice(INTERNAL_RHYME_PAIRS)
            return f"{line} — {pair[0]}, {pair[1]}"
        return line
    result = []
    for word in words:
        if random.random() >= prob:
            result.append(word)
            continue
        clean = re.sub(r"[^a-zA-Z]", "", word).lower()
        if clean:
            rhymes = pronouncing.rhymes(clean)
            if rhymes:
                result.append(f"{word}-{random.choice(rhymes)}")
            else:
                result.append(word)
        else:
            result.append(word)
    return " ".join(result)


def liquid_optimize_line(line: str, for_say: bool = False) -> str:
    """Single line: soft rhyme, optional elongate. Returns text for spoken/chunking."""
    line = soft_internal_rhyme(line)
    if LIQUID_SETTINGS["elongate_vowels"] and (not for_say or USE_ELONGATION_FOR_SAY):
        line = elongate(line)
    return line


def liquid_bar_grids_for_line(line: str) -> list[str]:
    """Turn a line into 16th-note bar grids (for display/performance guide)."""
    words = line.split()
    grids = []
    for start in range(0, len(words), LIQUID_WORDS_PER_BAR):
        chunk = words[start : start + LIQUID_WORDS_PER_BAR]
        grid = quantize_liquid(chunk)
        grid = liquid_snare_accent(grid)
        grid = liquid_breath(grid)
        grids.append(" ".join(grid))
    return grids


def liquid_optimize(lines: list[str], for_say: bool = True) -> tuple[list[str], list[str]]:
    """
    Process lines for liquid DnB. Returns (spoken_lines, bar_grid_lines).
    spoken_lines: for chunking and say (soft rhyme, optional elongate).
    bar_grid_lines: one string per bar with 16th grid + snare + breath for PA-Rap-Lyrics-WithBars.
    """
    spoken = []
    bar_lines = []
    for line in lines:
        text = liquid_optimize_line(line, for_say=for_say)
        spoken.append(text)
        for grid_str in liquid_bar_grids_for_line(text):
            bar_lines.append(grid_str)
    return spoken, bar_lines


# ==============================
# PROCESSING PIPELINE
# ==============================


def process_rap(lines: list[str], for_say: bool = True) -> list[str]:
    processed = []
    for line in lines:
        if TIGHTEN_LINES:
            line = tighten_line(line)
        line = add_internal_rhyme(line)
        if USE_PHONETIC_FOR_SAY or not for_say:
            line = phonetic_mark(line)
        line = tts_format(line, engine="say")
        processed.append(line)
    return processed


def process_rap_with_liquid(lines: list[str], for_say: bool = True) -> tuple[list[str], list[str]]:
    """When ENABLE_LIQUID_DNB: optional tighten first, then liquid optimize. Returns (spoken_lines, bar_grid_lines)."""
    if TIGHTEN_LINES:
        lines = [tighten_line(ln) for ln in lines]
    return liquid_optimize(lines, for_say=for_say)


# ==============================
# EXTRACT FROM MARKDOWN
# ==============================


def strip_bold(s: str) -> str:
    return re.sub(r"\*\*([^*]+)\*\*", r"\1", s).strip()


def is_skip_line(line: str) -> bool:
    line = line.strip()
    if not line or line == "---":
        return True
    if line.startswith("# ") or (line.startswith("*") and line.endswith("*")):
        return True
    if line.startswith("**[") and ("]**" in line or line.endswith("]**")):
        return True
    if line.startswith("[Intro]**") or line.startswith("[Outro]**"):
        return True
    if re.match(r"^\[\w+\s*\d*", line) or re.match(r"^\*\*\[", line):
        return True
    return False


def extract_lines() -> list[str]:
    lines = []
    for name in MODULES:
        path = DIR / name
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8")
        for raw in text.splitlines():
            line = raw.strip()
            if is_skip_line(line):
                continue
            clean = strip_bold(line)
            if clean:
                lines.append(clean)
    return lines


# ==============================
# MAIN
# ==============================


def main():
    random.seed(RANDOM_SEED)
    lines = extract_lines()

    if ENABLE_LIQUID_DNB:
        processed, bar_grid_lines = process_rap_with_liquid(lines, for_say=True)
        with_bars = [f"[Bar {(i % BARS_PER_SECTION) + 1:02d}] {g}" for i, g in enumerate(bar_grid_lines)]
    else:
        processed = process_rap(lines, for_say=True)
        with_bars = format_bars(processed)

    # Chunk into phrases by word count for say
    all_words = " ".join(processed).split()
    PHRASES_DIR.mkdir(exist_ok=True)
    phrases = []
    for i in range(0, len(all_words), MAX_WORDS_PER_PHRASE):
        chunk = all_words[i : i + MAX_WORDS_PER_PHRASE]
        phrases.append(" ".join(chunk))

    for i, phrase in enumerate(phrases):
        out_path = PHRASES_DIR / f"{i + 1:03d}.aiff"
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", delete=False, encoding="utf-8"
        ) as f:
            f.write(phrase)
            tmp = f.name
        try:
            subprocess.run(
                ["say", "-f", tmp, "-o", str(out_path)],
                check=True,
                capture_output=True,
            )
        finally:
            Path(tmp).unlink(missing_ok=True)
        if (i + 1) % 50 == 0:
            print(f"  {i + 1}/{len(phrases)} phrases...")

    (DIR / "PA-Rap-Lyrics-WithBars.txt").write_text("\n".join(with_bars), encoding="utf-8")

    mode = "Liquid DnB (172 BPM)" if ENABLE_LIQUID_DNB else "standard"
    print(f"Done: {len(phrases)} phrases (max {MAX_WORDS_PER_PHRASE} words) in {PHRASES_DIR} [{mode}]")
    print("Bar-aligned lyrics: PA-Rap-Lyrics-WithBars.txt (16th grid + snare on 2 & 4)")
    print("Run the Brandenburg patch; phrase spacing is set in the .scd.")


if __name__ == "__main__":
    main()
