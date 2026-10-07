import * as assert from 'assert';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { inlineImportsToText, resolveImportsToDeclarations } from '../../inliner';

suite('inliner', () => {
    let dir: string;

    setup(() => {
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'osts-inline-'));
    });

    teardown(() => {
        fs.rmSync(dir, { recursive: true, force: true });
    });

    /** Writes `files` (relative path → lines) into the temp project. */
    function project(files: Record<string, string[]>): void {
        for (const [rel, lines] of Object.entries(files)) {
            const full = path.join(dir, rel);
            fs.mkdirSync(path.dirname(full), { recursive: true });
            fs.writeFileSync(full, lines.join('\n') + '\n');
        }
    }

    const inline = (source: string[]) =>
        inlineImportsToText(path.join(dir, 'restock.osts'), source.join('\n') + '\n');

    suite('inlineImportsToText', () => {
        test('replaces the import with the declaration, types first and helpers last', async () => {
            project({
                'inventory.ts': [
                    'export interface InventoryRow { sku: string; qty: number; }',
                    '',
                    'export function lowStock(rows: InventoryRow[]): InventoryRow[] {',
                    '    return rows.filter(r => r.qty < 10);',
                    '}',
                ],
            });

            const { output, warnings } = await inline([
                'import { InventoryRow, lowStock } from "./inventory";',
                '',
                'function main(workbook: ExcelScript.Workbook) {',
                '    const rows: InventoryRow[] = [];',
                '    console.log(lowStock(rows).length);',
                '}',
            ]);

            assert.deepStrictEqual(warnings, []);
            assert.strictEqual(output, [
                'interface InventoryRow { sku: string; qty: number; }',
                '',
                'function main(workbook: ExcelScript.Workbook) {',
                '    const rows: InventoryRow[] = [];',
                '    console.log(lowStock(rows).length);',
                '}',
                '',
                'function lowStock(rows: InventoryRow[]): InventoryRow[] {',
                '    return rows.filter(r => r.qty < 10);',
                '}',
                '',
            ].join('\n'));
        });

        test('pulls in non-exported helpers the imported function depends on', async () => {
            project({
                'batch.ts': [
                    'const PREFIX = "LOT";',
                    'function pad(n: number): string { return String(n).padStart(4, "0"); }',
                    'function unused(): void {}',
                    'export function formatBatchNumber(n: number): string { return `${PREFIX}-${pad(n)}`; }',
                ],
            });

            const { output } = await inline([
                'import { formatBatchNumber } from "./batch";',
                'function main(workbook: ExcelScript.Workbook) { console.log(formatBatchNumber(7)); }',
            ]);

            assert.ok(output.includes('const PREFIX = "LOT";'));
            assert.ok(output.includes('function pad(n: number)'));
            assert.ok(!output.includes('unused'), 'unreferenced helpers are left out');
            assert.ok(!output.includes('import '), 'no import statements remain');
        });

        test('follows imports inside helper files, emitting a shared helper once', async () => {
            project({
                'shared/sheets.ts': [
                    'export function getSheet(wb: ExcelScript.Workbook, name: string) { return wb.getWorksheet(name); }',
                ],
                'receiving.ts': [
                    'import { getSheet } from "./shared/sheets";',
                    'export function receive(wb: ExcelScript.Workbook) { getSheet(wb, "Receiving"); }',
                ],
                'shipping.ts': [
                    'import { getSheet } from "./shared/sheets";',
                    'export function ship(wb: ExcelScript.Workbook) { getSheet(wb, "Shipping"); }',
                ],
            });

            const { output, warnings } = await inline([
                'import { receive } from "./receiving";',
                'import { ship } from "./shipping";',
                'function main(workbook: ExcelScript.Workbook) { receive(workbook); ship(workbook); }',
            ]);

            assert.deepStrictEqual(warnings, []);
            assert.strictEqual(output.split('function getSheet(').length - 1, 1);
            assert.ok(output.includes('function receive('));
            assert.ok(output.includes('function ship('));
        });

        test('resolves imports of .osts files and folder index files', async () => {
            project({
                'helpers.osts': ['export function logSku(sku: string) { console.log(sku); }'],
                'units/index.ts': ['export const CASE_SIZE = 12;'],
            });

            const { output, warnings } = await inline([
                'import { logSku } from "./helpers";',
                'import { CASE_SIZE } from "./units";',
                'function main(workbook: ExcelScript.Workbook) { logSku(String(CASE_SIZE)); }',
            ]);

            assert.deepStrictEqual(warnings, []);
            assert.ok(output.includes('function logSku('));
            assert.ok(output.includes('const CASE_SIZE = 12;'));
        });

        test('renames a non-exported helper that collides with a name in the script', async () => {
            project({
                'batch.ts': [
                    'function pad(n: number): string { return String(n).padStart(4, "0"); }',
                    'export function formatBatchNumber(n: number): string { return pad(n); }',
                ],
            });

            const { output, warnings } = await inline([
                'import { formatBatchNumber } from "./batch";',
                'function pad(s: string): string { return ` ${s} `; }',
                'function main(workbook: ExcelScript.Workbook) { console.log(pad(formatBatchNumber(1))); }',
            ]);

            assert.ok(output.includes('function pad_1(n: number)'));
            assert.ok(output.includes('return pad_1(n);'));
            assert.ok(output.includes('function pad(s: string)'), "the script's own pad is untouched");
            assert.deepStrictEqual(warnings, [
                'Renamed non-exported "pad" → "pad_1" (collision while inlining batch.ts)',
            ]);
        });

        test('warns about every unsupported or broken import', async () => {
            project({
                'batch.ts': [
                    'function pad(n: number): string { return String(n); }',
                    'export default function format(): void {}',
                ],
            });

            const { warnings } = await inline([
                'import { chunk } from "lodash";',
                'import format from "./batch";',
                'import * as batch from "./batch";',
                'import { pad } from "./batch";',
                'import { missing } from "./batch";',
                'import { nothing } from "./nowhere";',
                'function main(workbook: ExcelScript.Workbook) {}',
            ]);

            assert.deepStrictEqual(warnings, [
                'Skipped non-relative import "lodash"',
                'Skipped default import from "./batch"',
                'Skipped namespace import from "./batch"',
                '"pad" is not exported from batch.ts',
                '"missing" not found in batch.ts',
                `Cannot resolve "./nowhere" from ${dir}`,
            ]);
        });

        test('refuses aliased imports instead of emitting an undefined name', async () => {
            project({
                'batch.ts': ['export function formatBatchNumber(n: number): string { return `B-${n}`; }'],
            });

            const { output, warnings } = await inline([
                'import { formatBatchNumber as fmt } from "./batch";',
                'function main(workbook: ExcelScript.Workbook) { console.log(fmt(42)); }',
            ]);

            assert.deepStrictEqual(warnings, [
                'Skipped aliased import "formatBatchNumber as fmt" from "./batch" — import it by its original name',
            ]);
            // The import stays visible in the output rather than silently vanishing.
            assert.ok(output.includes('import { formatBatchNumber as fmt } from "./batch";'));
        });
    });

    suite('resolveImportsToDeclarations', () => {
        test('returns declarations keyed by name, types before values', async () => {
            project({
                'inventory.ts': [
                    'export function restock(row: InventoryRow): void { row.qty += 10; }',
                    'export interface InventoryRow { sku: string; qty: number; }',
                ],
            });

            const decls = await resolveImportsToDeclarations(
                dir,
                [{ specifier: './inventory', names: ['restock'] }],
                [],
            );

            assert.deepStrictEqual([...decls.keys()], ['InventoryRow', 'restock']);
            assert.strictEqual(decls.get('InventoryRow'), 'interface InventoryRow { sku: string; qty: number; }');
        });
    });
});
