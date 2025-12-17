REM Open CMD in the folder that contains AppIcon_single.png
magick AppIcon_single.png -trim +repage -background none -gravity center -extent 1024x1024 AppIcon_norm.png
