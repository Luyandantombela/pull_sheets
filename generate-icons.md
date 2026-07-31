# Icon Instructions

The manifest references three PNG icons at these paths:

| File | Size | Used for |
|------|------|----------|
| `icon-16.png` | 16 × 16 px | Small ribbon icon |
| `icon-32.png` | 32 × 32 px | Medium ribbon icon |
| `icon-80.png` | 80 × 80 px | Large ribbon icon, add-in store |

## Quick method — generate with ImageMagick

```bash
# Install ImageMagick if needed: https://imagemagick.org/
# Run from the repo root

for size in 16 32 80; do
  magick -size ${size}x${size} \
    -define gradient:direction=East \
    gradient:'#107c10-#0078d4' \
    -gravity Center \
    -font Arial-Bold \
    -pointsize $((size / 2)) \
    -fill white \
    -annotate 0 "S" \
    assets/images/icon-${size}.png
done
```

## Online alternative

Use any icon generator (e.g. https://favicon.io) to create PNGs of the required sizes and drop them in this folder. The add-in will work without icons but the ribbon button will show a broken image.
