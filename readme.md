# Fit to Letter and Scale

`FitLetterAndScale` formats the active Visio page for Letter paper. It selects portrait or landscape to match the drawing, adjusts the drawing scale to fit within 1/4-inch margins, and centers the drawing on the page.

The macro measures the drawing directly rather than running Visio's **Fit to Drawing** command. The printed page remains Letter size (8.5 × 11 inches); the drawing scale and page dimensions are adjusted to make the drawing fit.

## Install and run in one drawing

1. Open the drawing in Visio and choose **Developer → Visual Basic**.
2. In the VBA editor, select the drawing's project in Project Explorer. If Project Explorer is hidden, press **Ctrl+R**.
3. Choose **File → Import File** and select `letter-autosize.bas`.
4. Run `FitLetterAndScale` from the VBA editor, or press **F5** while the cursor is inside the macro.

The imported module is named `Auto_Size_Letter`. VBA module names must be valid identifiers, so they cannot contain hyphens; the macro itself is still named `FitLetterAndScale`.

The macro runs as a single operation and can be undone by pressing **Ctrl+Z**it. 

To keep the macro in the drawing, you must save the Save the drawing as a Visio macro-enabled drawing (`.vsdm`).

## Make it available in other drawings

Install the macro in a macro-enabled Visio stencil so you do not need to import the module into every drawing:

1. In Visio, create a new stencil and save it as a Visio Macro-Enabled Stencil (`.vssm`).
2. Open that stencil's VBA project with **Developer → Visual Basic**.
3. Select the stencil's project in Project Explorer, then choose **File → Import File** and select `letter-autosize.bas`.
4. Save the `.vssm` stencil.
5. In Visio, open **File → Options → Advanced → File Locations** and find the **Startup** folder location.
6. Put the `.vssm` file in that folder, then restart Visio. The stencil will load at startup, and `FitLetterAndScale` will be available from **Developer → Macros** in other drawings.

If Visio blocks the macro, use your organization's trusted-publisher or trusted-location policy rather than lowering macro security globally.

### Optional: Add a Quick Access Toolbar button

Open **File → Options → Quick Access Toolbar**, choose **Macros** from the command list, and add `FitLetterAndScale`.

## Notes

- **The whole page is measured.** Title blocks, date boxes, and other shapes count toward the fit. Move them to a background page or use Visio headers and footers if they should not affect the drawing scale.
- **Text and line weights do not scale with shapes.** The geometry changes size, but font sizes and line weights retain their values.
- **The resulting scale may not be a standard architectural scale.** It is calculated to fit the drawing precisely, so it may produce a scale such as 1 inch = 7.3 feet.
- **The macro scales up as well as down.** Small drawings are enlarged to use the available page area; large drawings are reduced.