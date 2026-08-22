from fontTools import ttLib
import sys

# Carlito carries Calibri's typographic metrics (sTypoAscender 1536, sTypoDescender -512) but does
# not set the USE_TYPO_METRICS flag, so consumers read the hhea/usWin values (1950/-550) instead.
# Calibri sets the flag, so a reader picks 1536/-512 there. Without this the metric compatible
# substitute would differ from Calibri in descent and text height, and match only in advance widths.
#
# The flag is only defined from OS/2 version 4 onwards, so the table version is raised as well.
# Versions 3 and 4 hold the same fields, so this is purely a declaration that bits 7-9 are meaningful.
USE_TYPO_METRICS = 1 << 7

font_path = sys.argv[1]
font_output = sys.argv[2]

font = ttLib.TTFont(font_path)
os2 = font["OS/2"]
if os2.version < 4:
    os2.version = 4
os2.fsSelection |= USE_TYPO_METRICS
font.save(font_output)
