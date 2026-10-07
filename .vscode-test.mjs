import { defineConfig } from '@vscode/test-cli';

export default defineConfig({
	// Integration tests only; unit tests in out/test/unit run under plain mocha (npm run test:unit).
	files: 'out/test/*.test.js',
	// extension.test.ts waits 2.5s for tsserver to load the plugin, which
	// mocha's default 2s timeout would cut short.
	mocha: { timeout: 20000 },
});
