# PowerPoint Native Reference Slide Macro

This folder contains the VBA module for the PowerPoint-native reference-slide insertion path on macOS.

## Why this exists

The slide PNG export path is now intentionally kept separate:

- `ImageExport.ppam`
  Owns the simple, reliable slide-image export macro:

  ```vb
  Sub ExportSlidesToFolder(folderPath As String)
      Dim sld As Slide
      Dim sName As String
      Dim fullPath As String

      If Right(folderPath, 1) <> "/" Then folderPath = folderPath & "/"

      For Each sld In ActivePresentation.Slides
          sName = "Slide_" & sld.SlideIndex & ".png"
          fullPath = folderPath & sName
          sld.Export fullPath, "PNG"
      Next sld
  End Sub
  ```

- `PDF-PPT_Helper.ppam`
  Owns only the reference-slide insertion macro:

  - `InsertReferenceSlidesFromManifest`

This split avoids having two add-ins define `ExportSlidesToFolder`, which turned out to be fragile on Mac PowerPoint.

## Intended usage

1. Keep `ImageExport.ppam` loaded in PowerPoint for slide PNG export.
2. Import [`SlideExport.bas`](./SlideExport.bas) into `PDF-PPT_Helper.ppam` or another helper add-in.
3. Load that helper add-in in PowerPoint too.
4. Invoke the macros from AppleScript with PowerPoint's `run VB macro` command.

For slide-image export, the app calls the export add-in's simple macro:

```applescript
tell application "Microsoft PowerPoint"
    run VB macro macro name "ExportSlidesToFolder" ¬
        list of parameters {outputFolder}
end tell
```

For reference-slide insertion, the app calls:

```applescript
tell application "Microsoft PowerPoint"
    run VB macro macro name "InsertReferenceSlidesFromManifest" ¬
        list of parameters {manifestPath}
end tell
```

For manual validation inside a `.pptm`, the helper module includes:

- `TestInsertReferenceSlidesFromManifest`

The export add-in should keep its own:

- `TestExportSlidesToFolder`

The manifest is a tab-delimited text file written by the app. Each line contains:

```text
INSERT    <zero-based-slide-index>    /absolute/path/to/reference-image.png
APPEND        /absolute/path/to/reference-image.png
```

## Important note

The app can call a PowerPoint VBA macro once the macro exists, but it does not create the `.ppam` automatically. PowerPoint VBA projects are compiled binary artifacts, so a real add-in must still be created once inside PowerPoint and updated when this module changes.
