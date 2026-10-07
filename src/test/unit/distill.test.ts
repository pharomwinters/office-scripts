import * as assert from 'assert';
import { distillItems } from '../../distill';
import type { HarvestedItem } from '../../harvest';

function fn(name: string, sourceFile: string, source: string): HarvestedItem {
    return { kind: 'function', name, sourceFile, source, normalized: source.replace(/\s+/g, ' ').trim() };
}

suite('distillItems', () => {
    test('unique names go to core', () => {
        const { core, conflict } = distillItems([
            fn('getInventorySheet', 'a.osts', 'function getInventorySheet() {}'),
        ]);
        assert.deepStrictEqual([...core.keys()], ['getInventorySheet']);
        assert.strictEqual(conflict.size, 0);
    });

    test('identical duplicates collapse into one core entry with every origin', () => {
        const { core, conflict } = distillItems([
            fn('lastRow', 'receiving.osts', 'function lastRow() { return 1; }'),
            fn('lastRow', 'shipping.osts', 'function  lastRow()  {\n  return 1;\n}'),
        ]);
        assert.strictEqual(conflict.size, 0);
        assert.deepStrictEqual(core.get('lastRow')?.originFiles, ['receiving.osts', 'shipping.osts']);
        // The first-seen source text is kept verbatim.
        assert.strictEqual(core.get('lastRow')?.source, 'function lastRow() { return 1; }');
    });

    test('same name with different bodies goes to conflict, one entry per variant', () => {
        const { core, conflict } = distillItems([
            fn('formatBatch', 'a.osts', 'function formatBatch(n: number) { return `B-${n}`; }'),
            fn('formatBatch', 'b.osts', 'function formatBatch(n: number) { return `LOT-${n}`; }'),
            fn('formatBatch', 'c.osts', 'function formatBatch(n: number) { return `B-${n}`; }'),
        ]);
        assert.strictEqual(core.size, 0);
        const variants = conflict.get('formatBatch');
        assert.strictEqual(variants?.length, 2);
        assert.deepStrictEqual(variants?.[0].originFiles, ['a.osts', 'c.osts']);
        assert.deepStrictEqual(variants?.[1].originFiles, ['b.osts']);
    });
});
