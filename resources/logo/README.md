# XlsxSharp logo

The mark is a rounded tile with a grid that reads both as a spreadsheet and as a
sharp sign (`#`) - the two halves of the name. The wordmark pairs `Xlsx` in the
brand colour with `Sharp` in ink.

## Colours

| Role         | Hex       | Used for                          |
|--------------|-----------|-----------------------------------|
| Brand green  | `#21A366` | mark, `Xlsx` (primary variant)    |
| .NET purple  | `#512BD4` | mark, `Xlsx` (secondary variant)  |
| Ink          | `#243746` | `Sharp`                           |

## Files

| File                     | Content                                             |
|--------------------------|-----------------------------------------------------|
| `favicon-01..04`         | Mark only: green, purple, black outline, white outline |
| `logotype-a-05..08`      | Mark and wordmark side by side, same four flavours  |
| `logotype-b-09..12`      | Mark above the wordmark, same four flavours         |
| `nuget-logo.png`         | Package icon (see `Directory.Build.props`)          |
| `readme.png`             | Banner used at the top of `README.md`               |

The black and white flavours are meant for single-colour usage; the white one
needs a dark background.

## Regenerating

All files are generated - edit `resources/scripts/generate-logos.py` instead of
touching the SVGs, then run:

```sh
python3 resources/scripts/generate-logos.py
```

The wordmark is converted to outlines from `resources/fonts/OpenSans-Regular.ttf`
(Open Sans, Apache-2.0), so the SVGs render without that font installed.
