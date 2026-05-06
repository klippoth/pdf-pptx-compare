# PowerPoint Native Slide Export Macro

This folder contains the VBA module for a PowerPoint-native slide PNG export path on macOS.

## Why this exists

The direct AppleScript command:

```applescript
save currentPresentation in outputPath as save as PNG
```

is unreliable on newer macOS / Office builds. The more promising native PowerPoint path is VBA:

```vb
Presentation.Export Path:="...", FilterName:="PNG", ScaleWidth:=1400, ScaleHeight:=788
```

## Intended usage

1. Create a PowerPoint add-in (`.ppam`) or macro-enabled deck (`.pptm`).
2. Import [`SlideExport.bas`](./SlideExport.bas) into that VBA project.
3. Load the add-in in PowerPoint.
4. Invoke it from AppleScript with PowerPoint's `run VB macro` command.

Example AppleScript shape:

```applescript
tell application "Microsoft PowerPoint"
    set exportResult to run VB macro macro name "SlideExport.ExportSlidesAsPng" ¬
        list of parameters {inputPath, outputFolder, "1400", "788"}
end tell
```

## Current blocker

The app can call a PowerPoint VBA macro once the macro exists, but it does not yet create the `.ppam` automatically. PowerPoint VBA projects are compiled binary artifacts, so a real add-in must be created once inside PowerPoint.
