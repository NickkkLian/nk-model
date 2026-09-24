# Fonts shipped with this skill

preview.py inlines them into every preview page it writes, as `data:` URIs inside one `<style id="brand-fonts">` block, so the file draws its own type and still
fetches nothing when it opens. They are the same three files the author's own website serves.

| file | face | licence |
|---|---|---|
| `fraunces-latin-wght.woff2` | Fraunces, variable weight 100–900, Latin subset | SIL OFL 1.1, `OFL-Fraunces.txt` |
| `inter-latin-wght.woff2` | Inter, variable weight 100–900, Latin subset | SIL OFL 1.1, `OFL-Inter.txt` |
| `space-mono-latin-400.woff2` | Space Mono Regular, Latin subset | SIL OFL 1.1, `OFL-SpaceMono.txt` |

Where they come from: the Fontsource packages `@fontsource-variable/fraunces` 5.2.9 (`fraunces-latin-wght-normal.woff2`),
`@fontsource-variable/inter` 5.2.8 (`inter-latin-wght-normal.woff2`) and `@fontsource/space-mono` 5.2.9
(`space-mono-latin-400-normal.woff2`). They are unchanged; only the names are shorter.

| file | bytes | sha256 |
|---|---|---|
| `fraunces-latin-wght.woff2` | 36,620 | `7f9d191d999336d3b9790afa72e1358e50a13b06d4f289341e92a311967a80f9` |
| `inter-latin-wght.woff2` | 48,256 | `3100e775e8616cd2611beecfa23a4263d7037586789b43f035236a2e6fbd4c62` |
| `space-mono-latin-400.woff2` | 16,520 | `fb4a81a2d0a893e5c38c394a7e716a1cef0b24610a0af49c96f6d529bd66bf2b` |

Latin subsets only. Characters outside them (Chinese, arrows, most symbols) are drawn by the next font the
token file names, as before.
