import * as vscode from 'vscode';
import { inlineImportsToText } from './inliner';
import { isOfficeScriptFile } from './marker';

/**
 * Opens an Excel-ready version of the given Office Script in a new untitled
 * editor, with every relative import replaced by inlined copies of the
 * referenced declarations. See src/inliner.ts for how inlining works and
 * what it doesn't support.
 *
 * Accepts `.osts` files and `.ts` files tagged with `/** @OfficeScript *\/`.
 */
export async function inlineImports(doc: vscode.TextDocument): Promise<void> {
    const source = doc.getText();
    if (!isOfficeScriptFile(doc.fileName, source)) {
        vscode.window.showErrorMessage(
            'Inline Imports requires an .osts file or a .ts file tagged with /** @OfficeScript */.',
        );
        return;
    }

    const { output, warnings } = await inlineImportsToText(doc.fileName, source);

    const newDoc = await vscode.workspace.openTextDocument({
        content: output,
        language: doc.languageId,
    });
    await vscode.window.showTextDocument(newDoc);

    if (warnings.length > 0) {
        vscode.window.showWarningMessage(`Inline Imports: ${warnings.join('; ')}`);
    }
}
