import * as assert from 'assert';
import { hasOfficeScriptMarker, isOfficeScriptFile } from '../../marker';

suite('marker', () => {
    const tagged = '/** @OfficeScript */\nfunction main(workbook: ExcelScript.Workbook) {}\n';
    const untagged = 'export function formatBatchNumber(n: number): string { return `B-${n}`; }\n';

    test('detects the @OfficeScript JSDoc tag', () => {
        assert.strictEqual(hasOfficeScriptMarker(tagged), true);
        assert.strictEqual(hasOfficeScriptMarker(untagged), false);
    });

    test('ignores the tag outside a JSDoc block', () => {
        assert.strictEqual(hasOfficeScriptMarker('// @OfficeScript\nfunction main() {}'), false);
    });

    test('.osts files always qualify, regardless of content', () => {
        assert.strictEqual(isOfficeScriptFile('inventory.osts', untagged), true);
        assert.strictEqual(isOfficeScriptFile('INVENTORY.OSTS', ''), true);
    });

    test('.ts files qualify only when tagged', () => {
        assert.strictEqual(isOfficeScriptFile('updateInventory.ts', tagged), true);
        assert.strictEqual(isOfficeScriptFile('helpers.ts', untagged), false);
    });

    test('.d.ts and other extensions never qualify', () => {
        assert.strictEqual(isOfficeScriptFile('excel-script.d.ts', tagged), false);
        assert.strictEqual(isOfficeScriptFile('notes.md', tagged), false);
    });
});
