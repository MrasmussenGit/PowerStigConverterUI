
REM Generate clean per-size PNGs from the normalized master
magick AppIcon_norm.png -resize 256x256 App_256.png
magick AppIcon_norm.png -resize 128x128 App_128.png
magick AppIcon_norm.png -resize 64x64  App_64.png
magick AppIcon_norm.png -resize 48x48   App_48.png
magick AppIcon_norm.png -resize 32x32   App_32.png
magick AppIcon_norm.png -resize 24x24   App_24.png
magick AppIcon_norm.png -resize 16x16   App_16.png

REM Bundle into a single ICO (note: output is LAST)
del /f App.ico 2>nul
magick App_256.png App_128.png App_64.png App_48.png App_32.png App_24.png App_16.png -colorspace sRGB -alpha on App.ico
