import { defineConfig } from '@vscode/test-cli';

export default defineConfig({
	// Integration tests only; unit tests in out/test/unit run under plain mocha (npm run test:unit).
	files: 'out/test/*.test.js',
});
