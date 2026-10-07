import { defineConfig } from '@vscode/test-cli';

export default defineConfig({
	// Integration tests only; unit tests in out/test/unit run under plain mocha (npm run test:unit).
	files: 'out/test/*.test.js',
	// extension.test.ts polls up to ~20s for tsserver to load the plugin; leave
	// headroom so a slow runner fails on its assertion, not on mocha's timeout.
	mocha: { timeout: 30000 },
});
