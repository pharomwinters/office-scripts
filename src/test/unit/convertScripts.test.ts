import * as assert from 'assert';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { convertScriptsToOsts } from '../../convertScripts';

suite('convertScriptsToOsts', () => {
    let src: string;
    let out: string;

    setup(() => {
        const root = fs.mkdtempSync(path.join(os.tmpdir(), 'osts-convert-'));
        src = path.join(root, 'downloaded');
        out = path.join(root, 'converted');
        fs.mkdirSync(src);
    });

    teardown(() => {
        fs.rmSync(path.dirname(src), { recursive: true, force: true });
    });

    const read = (name: string) => fs.readFileSync(path.join(out, name), 'utf8');

    test('unwraps the JSON envelope and tags main', () => {
        const body = 'function main(workbook: ExcelScript.Workbook) {\n    console.log("restock");\n}';
        fs.writeFileSync(path.join(src, 'restock.osts'), JSON.stringify({ version: '0.3.0', body }));

        const written = convertScriptsToOsts(src, out);

        assert.deepStrictEqual(written, [path.join(out, 'restock.ts')]);
        assert.strictEqual(read('restock.ts'), `/** @OfficeScript */\n${body}\n`);
    });

    test('keeps plain-text scripts and an existing JSDoc above main', () => {
        const text = '/** Logs batch numbers. */\nasync function main(workbook: ExcelScript.Workbook) {}';
        fs.writeFileSync(path.join(src, 'batches.osts'), text);

        convertScriptsToOsts(src, out);

        assert.strictEqual(
            read('batches.ts'),
            '/** Logs batch numbers. */\n/** @OfficeScript */\nasync function main(workbook: ExcelScript.Workbook) {}\n',
        );
    });

    test('does not double-tag an already-marked script', () => {
        const text = '/** @OfficeScript */\nfunction main(workbook: ExcelScript.Workbook) {}\n';
        fs.writeFileSync(path.join(src, 'tagged.osts'), text);

        convertScriptsToOsts(src, out);

        assert.strictEqual(read('tagged.ts'), text);
    });
});
