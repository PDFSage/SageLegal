#!/bin/bash
set -euo pipefail

# Usage: ./watermark_video.sh <input_video> <output_video> [font_file]
# If no font_file is provided, it defaults to DejaVuSans-Bold.
if [ "$#" -lt 2 ]; then
    echo "Usage: $0 <input_video> <output_video> [font_file]"
    exit 1
fi

INPUT_VIDEO="$1"
OUTPUT_VIDEO="$2"
FONTFILE="${3:-/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf}"

# Verify that FFmpeg is installed.
if ! command -v ffmpeg &> /dev/null; then
    echo "Error: ffmpeg is not installed. Please install ffmpeg and try again."
    exit 1
fi

# Watermark text.
WATERMARK_TEXT="erosolar is cool"

# Apply the watermark at the bottom-right corner with a margin of 10 pixels.
# - `tw` and `th` are FFmpeg expressions for text width and text height.
# - The box option draws a semi-transparent background for better readability.
ffmpeg -y -i "$INPUT_VIDEO" \
  -vf "drawtext=fontfile='$FONTFILE':text='$WATERMARK_TEXT':fontcolor=white:fontsize=24:box=1:boxcolor=black@0.5:boxborderw=5:x=w-tw-10:y=h-th-10" \
  -codec:a copy "$OUTPUT_VIDEO"