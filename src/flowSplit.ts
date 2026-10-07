import * as vscode from 'vscode';
import * as ts from 'typescript';
import * as fs from 'fs';
import * as path from 'path';
import { ImportBinding, resolveImportsToDeclarations } from './inliner';
import { isOfficeScriptFile } from './marker';

/**
 * Splits every function annotated with a `/** @FlowSplit *\/` JSDoc tag in
 * the current .osts file into its own .osts file. The split files land in a
 * sibling folder named after the source file (minus .osts):
 *
 *     reports.osts          (source — stays intact)
 *     reports/
 *       syncPlazas.osts     (was function syncPlazas, now `main`)
 *       updateInventory.osts
 *
 * Each split file carries forward only the helpers and imports it actually
 * uses, computed by walking the split function's body and every helper it
 * transitively references.
 *
 * Limitations:
 *   - Helper dependency detection is name-based. A local variable that
 *     shadows a top-level name will still cause the top-level declaration
 *     to be copied (harmless — produces extra code, not wrong code).
 *   - Only top-level function/class/interface/type/enum/variable
 *     declarations are considered as helpers; nested declarations are not.
 */
export async function splitFlows(doc: vscode.TextDocument): Promise<void> {
    const source = doc.getText();
    if (!isOfficeScriptFile(doc.fileName, source)) {
        vscode.window.showErrorMessage(
            'Split Flows requires an .osts file or a .ts file tagged with /** @OfficeScript */.',
        );
        return;
    }

    const sf = ts.createSourceFile(doc.fileName, source, ts.ScriptTarget.ES2020, true);

    const flows = collectFlowFunctions(sf);
    if (flows.length === 0) {
        vscode.window.showInformationMessage('No /** @FlowSplit */ functions found.');
        return;
    }

    const topDecls = collectTopDeclarations(sf, flows);
    const importsByName = collectImportsByName(sf);

    const sourceExt = path.extname(doc.fileName);
    const sourceDir = path.dirname(doc.fileName);
    const sourceBase = path.basename(doc.fileName, sourceExt);
    const outDir = path.join(sourceDir, sourceBase);
    fs.mkdirSync(outDir, { recursive: true });

    const warnings: string[] = [];
    const written: string[] = [];
    for (const flow of flows) {
        const { helperNames, importedBySpecifier } = resolveDependencies(
            flow,
            topDecls,
            importsByName,
            flows,
        );
        const importedDecls = await resolveImportsToDeclarations(
            sourceDir,
            [...importedBySpecifier].map(([specifier, bindings]) => ({
                specifier,
                bindings: [...bindings.values()],
            })),
            warnings,
        );
        const content = buildSplitFile(flow, helperNames, topDecls, importedDecls, source);
        // Split output matches the source extension so a `.ts` authoring
        // file produces `.ts` splits (still carrying the `@OfficeScript`
        // marker so the extension keeps treating them specially).
        const markedContent = sourceExt === '.ts' ? `/** @OfficeScript */\n${content}` : content;
        const outPath = path.join(outDir, `${flow.name}${sourceExt}`);
        fs.writeFileSync(outPath, markedContent);
        written.push(path.basename(outPath));
    }

    vscode.window.showInformationMessage(
        `Split ${written.length} flow(s) into ${path.relative(sourceDir, outDir)}/: ${written.join(', ')}`,
    );
    if (warnings.length > 0) {
        vscode.window.showWarningMessage(`Split Flows: ${warnings.join('; ')}`);
    }
}

interface FlowFn {
    name: string;
    stmt: ts.FunctionDeclaration;
}

interface TopDecl {
    name: string;
    stmt: ts.Statement;
}

interface ImportRef {
    specifier: string;
    binding: ImportBinding;
}

function collectFlowFunctions(sf: ts.SourceFile): FlowFn[] {
    const flows: FlowFn[] = [];
    for (const stmt of sf.statements) {
        if (!ts.isFunctionDeclaration(stmt) || !stmt.name) continue;
        if (!hasFlowSplitTag(stmt)) continue;
        flows.push({ name: stmt.name.text, stmt });
    }
    return flows;
}

function hasFlowSplitTag(stmt: ts.Statement): boolean {
    const jsDocs = (stmt as unknown as { jsDoc?: ts.JSDoc[] }).jsDoc;
    if (!jsDocs) return false;
    return jsDocs.some(doc =>
        doc.tags?.some(tag => tag.tagName.text === 'FlowSplit'),
    );
}

function collectTopDeclarations(sf: ts.SourceFile, flows: FlowFn[]): Map<string, TopDecl> {
    const flowNames = new Set(flows.map(f => f.name));
    const map = new Map<string, TopDecl>();

    for (const stmt of sf.statements) {
        if (ts.isFunctionDeclaration(stmt) && stmt.name && !flowNames.has(stmt.name.text)) {
            map.set(stmt.name.text, { name: stmt.name.text, stmt });
        } else if (ts.isClassDeclaration(stmt) && stmt.name) {
            map.set(stmt.name.text, { name: stmt.name.text, stmt });
        } else if (ts.isInterfaceDeclaration(stmt)) {
            map.set(stmt.name.text, { name: stmt.name.text, stmt });
        } else if (ts.isTypeAliasDeclaration(stmt)) {
            map.set(stmt.name.text, { name: stmt.name.text, stmt });
        } else if (ts.isEnumDeclaration(stmt)) {
            map.set(stmt.name.text, { name: stmt.name.text, stmt });
        } else if (ts.isVariableStatement(stmt)) {
            for (const decl of stmt.declarationList.declarations) {
                if (ts.isIdentifier(decl.name)) {
                    map.set(decl.name.text, { name: decl.name.text, stmt });
                }
            }
        }
    }
    return map;
}

function collectImportsByName(sf: ts.SourceFile): Map<string, ImportRef> {
    const map = new Map<string, ImportRef>();

    for (const stmt of sf.statements) {
        if (!ts.isImportDeclaration(stmt)) continue;
        if (!ts.isStringLiteral(stmt.moduleSpecifier)) continue;
        const clause = stmt.importClause;
        if (!clause) continue;
        const specifier = stmt.moduleSpecifier.text;

        if (clause.name) {
            const name = clause.name.text;
            map.set(name, { specifier, binding: { imported: name, local: name } });
        }
        if (clause.namedBindings) {
            if (ts.isNamespaceImport(clause.namedBindings)) {
                const name = clause.namedBindings.name.text;
                map.set(name, { specifier, binding: { imported: name, local: name } });
            } else {
                for (const el of clause.namedBindings.elements) {
                    // propertyName is the exported name; `name` is the local
                    // binding. The resolver needs both: it looks up the
                    // exported name and aliases it when the two differ.
                    map.set(el.name.text, {
                        specifier,
                        binding: { imported: (el.propertyName ?? el.name).text, local: el.name.text },
                    });
                }
            }
        }
    }
    return map;
}

function resolveDependencies(
    flow: FlowFn,
    topDecls: Map<string, TopDecl>,
    importsByName: Map<string, ImportRef>,
    flows: FlowFn[],
): { helperNames: string[]; importedBySpecifier: Map<string, Map<string, ImportBinding>> } {
    const flowNames = new Set(flows.map(f => f.name));
    const helperNames = new Set<string>();
    const importedBySpecifier = new Map<string, Map<string, ImportBinding>>();
    const queue: ts.Node[] = [flow.stmt];
    const visited = new Set<string>();

    while (queue.length > 0) {
        const node = queue.shift()!;
        collectReferencedIdentifiers(node, (name) => {
            if (visited.has(name)) return;
            visited.add(name);

            if (flowNames.has(name)) return; // don't pull in other flows
            if (name === 'main') return;

            const ref = importsByName.get(name);
            if (ref) {
                let bindings = importedBySpecifier.get(ref.specifier);
                if (!bindings) {
                    bindings = new Map<string, ImportBinding>();
                    importedBySpecifier.set(ref.specifier, bindings);
                }
                bindings.set(name, ref.binding);
            }
            if (topDecls.has(name)) {
                helperNames.add(name);
                queue.push(topDecls.get(name)!.stmt);
            }
        });
    }

    return { helperNames: [...helperNames], importedBySpecifier };
}

function collectReferencedIdentifiers(node: ts.Node, visit: (name: string) => void): void {
    const walk = (n: ts.Node) => {
        if (ts.isIdentifier(n) && isReferencePosition(n)) {
            visit(n.text);
        }
        ts.forEachChild(n, walk);
    };
    walk(node);
}

function isReferencePosition(id: ts.Identifier): boolean {
    const parent = id.parent;
    if (!parent) return true;

    // Property access: `obj.foo` — `foo` is a property name, not a reference.
    if (ts.isPropertyAccessExpression(parent) && parent.name === id) return false;
    // Qualified name: `Foo.Bar` — `Bar` is a namespace member name.
    if (ts.isQualifiedName(parent) && parent.right === id) return false;
    // Object literal key: `{ foo: 1 }` — `foo` is a property name.
    if (ts.isPropertyAssignment(parent) && parent.name === id) return false;
    if (ts.isShorthandPropertyAssignment(parent) && parent.name === id) {
        // Shorthand `{ foo }` IS a reference to `foo` in scope.
        return true;
    }
    // Declaration names: function foo / class foo / const foo / parameter foo.
    if (
        (ts.isFunctionDeclaration(parent) ||
            ts.isFunctionExpression(parent) ||
            ts.isClassDeclaration(parent) ||
            ts.isInterfaceDeclaration(parent) ||
            ts.isTypeAliasDeclaration(parent) ||
            ts.isEnumDeclaration(parent) ||
            ts.isVariableDeclaration(parent) ||
            ts.isParameter(parent) ||
            ts.isBindingElement(parent) ||
            ts.isMethodDeclaration(parent) ||
            ts.isPropertyDeclaration(parent)) &&
        (parent as unknown as { name?: ts.Node }).name === id
    ) {
        return false;
    }
    // Import/export clause names are bindings, not references.
    if (ts.isImportSpecifier(parent) || ts.isImportClause(parent) || ts.isNamespaceImport(parent)) {
        return false;
    }
    return true;
}

function buildSplitFile(
    flow: FlowFn,
    helperNames: string[],
    topDecls: Map<string, TopDecl>,
    importedDecls: Map<string, string>,
    source: string,
): string {
    const parts: string[] = [renameFlowToMain(flow, source)];

    for (const name of helperNames) {
        parts.push(extractStatementText(topDecls.get(name)!.stmt, source));
    }
    for (const snippet of importedDecls.values()) {
        parts.push(snippet);
    }

    return parts.join('\n\n').trimEnd() + '\n';
}

function renameFlowToMain(flow: FlowFn, source: string): string {
    const stmt = flow.stmt;
    const bodyStart = stmt.getStart(); // skips leading trivia, including the @FlowSplit JSDoc
    const bodyEnd = stmt.getEnd();
    const text = source.slice(bodyStart, bodyEnd);

    const nameNode = stmt.name!;
    const relNameStart = nameNode.getStart() - bodyStart;
    const relNameEnd = nameNode.getEnd() - bodyStart;

    return text.slice(0, relNameStart) + 'main' + text.slice(relNameEnd);
}

function extractStatementText(stmt: ts.Statement, source: string): string {
    // Use getStart (not getFullStart) to skip the cascading leading trivia
    // from the previous statement, but keep any JSDoc attached to this stmt.
    const jsDocs = (stmt as unknown as { jsDoc?: ts.JSDoc[] }).jsDoc;
    const start = jsDocs && jsDocs.length > 0 ? jsDocs[0].getStart() : stmt.getStart();
    return source.slice(start, stmt.getEnd());
}
