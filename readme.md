# Fit to Letter and Scale

`FitLetterAndScale` formats the active Visio page for Letter paper. It selects portrait or landscape to match the drawing, adjusts the drawing scale to fit within 1/4-inch margins, and centers the drawing on the page.

The macro measures the drawing directly rather than running Visio's **Fit to Drawing** command. The printed page remains Letter size (8.5 × 11 inches); the drawing scale and page dimensions are adjusted to make the drawing fit.

## Install and run

1. In Visio, open **Developer → Visual Basic**.
2. In the VBA editor, choose **Insert → Module**.
3. Paste the contents of `letter-autosize.vsdm` into the module.
4. Run `FitLetterAndScale` from the VBA editor, or press **F5** while the cursor is inside the macro.

The macro runs as one undoable operation. Press **Ctrl+Z** to undo it.

To keep the macro in the drawing, save the file as a Visio macro-enabled drawing (`.vsdm`). To make it available while working in other drawings, store it in a macro-enabled stencil (`.vssm`).

### Optional: Add a Quick Access Toolbar button

Open **File → Options → Quick Access Toolbar**, choose **Macros** from the command list, and add `FitLetterAndScale`.

## Notes

- **The whole page is measured.** Title blocks, date boxes, and other shapes count toward the fit. Move them to a background page or use Visio headers and footers if they should not affect the drawing scale.
- **Text and line weights do not scale with shapes.** The geometry changes size, but font sizes and line weights retain their values.
- **The resulting scale may not be a standard architectural scale.** It is calculated to fit the drawing precisely, so it may produce a scale such as 1 inch = 7.3 feet.
- **The macro scales up as well as down.** Small drawings are enlarged to use the available page area; large drawings are reduced.