import * as assert from 'assert';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { harvestItems } from '../../harvest';

suite('harvestItems', () => {
    let dir: string;

    setup(() => {
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'osts-harvest-'));
    });

    teardown(() => {
        fs.rmSync(dir, { recursive: true, force: true });
    });

    test('extracts top-level helpers and interfaces, skipping main', async () => {
        fs.writeFileSync(path.join(dir, 'restock.osts'), [
            'interface InventoryRow { sku: string; qty: number; }',
            '',
            '/** Returns the inventory sheet. */',
            'export function getInventorySheet(workbook: ExcelScript.Workbook) {',
            '    return workbook.getWorksheet("Inventory");',
            '}',
            '',
            'function main(workbook: ExcelScript.Workbook) {',
            '    getInventorySheet(workbook);',
            '}',
            '',
        ].join('\n'));

        const items = await harvestItems(dir);
        const names = items.map(i => `${i.kind}:${i.name}`).sort();
        assert.deepStrictEqual(names, ['function:getInventorySheet', 'interface:InventoryRow']);

        const helper = items.find(i => i.name === 'getInventorySheet')!;
        // JSDoc is preserved and `export` is stripped even when the JSDoc precedes it.
        assert.strictEqual(helper.source, [
            '/** Returns the inventory sheet. */',
            'function getInventorySheet(workbook: ExcelScript.Workbook) {',
            '    return workbook.getWorksheet("Inventory");',
            '}',
        ].join('\n'));
        assert.ok(!/\s{2,}/.test(helper.normalized), 'normalized form collapses whitespace');
    });

    test('walks subfolders and skips dot-folders and node_modules', async () => {
        const helper = 'function sumQty(): number { return 0; }\n';
        fs.mkdirSync(path.join(dir, 'warehouse'));
        fs.mkdirSync(path.join(dir, '.git'));
        fs.mkdirSync(path.join(dir, 'node_modules'));
        fs.writeFileSync(path.join(dir, 'warehouse', 'count.osts'), helper);
        fs.writeFileSync(path.join(dir, '.git', 'ignored.osts'), helper);
        fs.writeFileSync(path.join(dir, 'node_modules', 'ignored.osts'), helper);

        const items = await harvestItems(dir);
        assert.deepStrictEqual(items.map(i => path.relative(dir, i.sourceFile)), [path.join('warehouse', 'count.osts')]);
    });
});
