---
name: gif-maker
description: Create animated GIFs from images or videos using ffmpeg
homepage: 
metadata:
  {
    "openclaw": {
      "emoji": "🎞️",
      "requires": { "bins": ["ffmpeg"] },
      "install": []
    }
  }
---

# GIF Maker

Create animated GIFs using ffmpeg.

## Quick start

Create a simple animated GIF from images:
```bash
ffmpeg -framerate 10 -i frame_%03d.png -loop 0 output.gif
```

Create a GIF from a single image with zoom/motion effect:
```bash
ffmpeg -i input.png -filter_complex "split [a][b];[a] scale=iw*1.1:-1 [p];[b] blur=5 [b];[p][b] overlay=0:0" -loop 0 output.gif
```

## Create swimming animation from single image
This script creates a simple bobbing/zooming animation effect:
```bash
./scripts/swim-loop.sh /path/to/image.png output.gif
```
