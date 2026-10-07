import * as assert from 'assert';
import * as path from 'path';
import * as vscode from 'vscode';

/**
 * Re-runs `probe` until it returns a value, or gives up after `timeoutMs`.
 * tsserver loads the plugin and indexes the ambient .d.ts asynchronously, at
 * a speed that varies a lot between machines, so polling beats a fixed sleep.
 */
async function waitFor<T>(probe: () => Thenable<T | undefined>, timeoutMs: number): Promise<T | undefined> {
    const deadline = Date.now() + timeoutMs;
    for (;;) {
        const result = await probe();
        if (result !== undefined || Date.now() > deadline) {
            return result;
        }
        await new Promise(resolve => setTimeout(resolve, 250));
    }
}

function targetUri(def: vscode.Location | vscode.LocationLink): vscode.Uri {
    return 'targetUri' in def ? def.targetUri : def.uri;
}

suite('Office Scripts extension', () => {
    suiteSetup(() => {
        vscode.window.showInformationMessage('Starting Office Scripts tests.');
    });

    test('ExcelScript types resolve in .osts files', async () => {
        // Compiled tests live in out/test/; source-of-truth fixture is src/test-usage.osts.
        const fixture = path.resolve(__dirname, '..', '..', 'src', 'test-usage.osts');
        const doc = await vscode.workspace.openTextDocument(fixture);
        await vscode.window.showTextDocument(doc);

        // Go to Definition on `getActiveWorksheet` can only land in
        // excel-script.d.ts if the plugin injected the ExcelScript types —
        // without them, `workbook.getActiveWorksheet` has no definition at all.
        const callSite = doc.positionAt(doc.getText().indexOf('getActiveWorksheet') + 1);
        const definition = await waitFor(async () => {
            const defs = await vscode.commands.executeCommand<(vscode.Location | vscode.LocationLink)[]>(
                'vscode.executeDefinitionProvider', doc.uri, callSite,
            );
            return defs?.find(d => targetUri(d).fsPath.endsWith('excel-script.d.ts'));
        }, 15000);
        assert.ok(definition, 'Expected getActiveWorksheet to resolve to types/excel-script.d.ts.');

        // Custom rules from diagnostics.ts must fire on `const x: any` and on
        // console.warn / console.error.
        const diagnostics = await waitFor(async () => {
            const diags = vscode.languages.getDiagnostics(doc.uri);
            const hasAny = diags.some(d => d.message.includes('"any" type is forbidden'));
            const hasConsole = diags.some(d => d.message.includes('is not supported in Office Scripts'));
            return hasAny && hasConsole ? diags : undefined;
        }, 5000);
        assert.ok(diagnostics, 'Expected the "any is forbidden" and console.warn/error custom diagnostics to fire.');

        // TS2339 = "Property X does not exist on type Y". The fixture has exactly one
        // such call (guarded by @ts-expect-error), which tsserver SUPPRESSES — so we
        // expect zero surfaced TS2339 diagnostics.
        const ts2339 = diagnostics.filter(d => d.code === 2339);
        assert.strictEqual(
            ts2339.length,
            0,
            `Expected no TS2339 diagnostics, got: ${ts2339.map(d => d.message).join(' | ')}`
        );
    });
});
